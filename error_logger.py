# -*- coding: utf-8 -*-
"""error_logger — структурированный журнал ошибок и API-вызовов пайплайна.

JSONL-файл рядом с ботом. Никогда не бросает исключения наружу:
логирование не должно ломать генерацию.
"""

from __future__ import annotations

import json
import os
import time
import traceback
from collections import Counter
from typing import Any

_LOG_DIR = os.path.dirname(os.path.abspath(__file__))
ERROR_LOG_FILE = os.path.join(_LOG_DIR, "pipeline_errors.jsonl")
API_LOG_FILE = os.path.join(_LOG_DIR, "api_calls.jsonl")

# Ограничение размера лог-файлов (10 МБ) — при превышении файл усечётся.
_MAX_LOG_BYTES = 10 * 1024 * 1024


class PipelineError(Exception):
    """Ошибка этапа генерации (структура, блок, сборка DOCX)."""


class APIFailure(Exception):
    """Ошибка вызова LLM-провайдера."""


# ═══════════════════════════════════════════════════════════════
#  FIX #5: ОПОВЕЩЕНИЕ АДМИНИСТРАТОРА В TELEGRAM О КРИТИЧЕСКИХ ОШИБКАХ
#  Раньше ошибки только писались в JSONL — о падениях бота админ узнавал
#  от пользователей. Теперь критические события летят в Telegram сразу.
# ═══════════════════════════════════════════════════════════════

# Настраивается из бота через configure_error_alerts() либо через env.
_ALERT_BOT_TOKEN: str = os.environ.get("BOT_TOKEN", "")
_ALERT_CHAT_ID: str = os.environ.get("ADMIN_CHAT_ID", "")

# Этапы, при которых оповещение отправляется всегда.
_CRITICAL_STAGES = {
    "fatal", "startup_failed", "generation_failed", "docx_build_failed",
    "quality_gate_failed", "payment_error", "unhandled_exception",
}

# Антиспам: не чаще 1 оповещения на (stage) за 5 минут.
_ALERT_COOLDOWN_SEC = 300
_last_alert_ts: dict[str, float] = {}


def configure_error_alerts(bot_token: str, admin_chat_id: str | int) -> None:
    """Вызывается из бота при старте: включает Telegram-оповещения."""
    global _ALERT_BOT_TOKEN, _ALERT_CHAT_ID
    if bot_token:
        _ALERT_BOT_TOKEN = str(bot_token)
    if admin_chat_id:
        _ALERT_CHAT_ID = str(admin_chat_id)


def _send_telegram_alert(text: str) -> None:
    """Отправка через Bot API в отдельном потоке — никогда не блокирует
    event loop и не бросает исключений (алерт не должен ломать генерацию)."""
    if not (_ALERT_BOT_TOKEN and _ALERT_CHAT_ID):
        return

    def _worker() -> None:
        try:
            import urllib.request
            import urllib.parse
            url = f"https://api.telegram.org/bot{_ALERT_BOT_TOKEN}/sendMessage"
            payload = urllib.parse.urlencode({
                "chat_id": _ALERT_CHAT_ID,
                "text": text[:4000],
            }).encode()
            urllib.request.urlopen(url, data=payload, timeout=10)
        except Exception:
            pass  # алерт — best effort

    try:
        import threading
        threading.Thread(target=_worker, daemon=True).start()
    except Exception:
        pass


def _maybe_alert(stage: str, message: str, record: dict, force: bool) -> None:
    if not force and stage not in _CRITICAL_STAGES:
        return
    now = time.time()
    if now - _last_alert_ts.get(stage, 0.0) < _ALERT_COOLDOWN_SEC:
        return
    _last_alert_ts[stage] = now
    lines = [f"🚨 ОШИБКА БОТА: {stage}", message[:1500]]
    for key in ("model", "user_id", "topic", "doc_type"):
        if record.get(key):
            lines.append(f"{key}: {record[key]}")
    if record.get("exception"):
        lines.append(f"exception: {record['exception'][:300]}")
    _send_telegram_alert("\n".join(lines))


def _rotate_if_needed(path: str) -> None:
    """Усекает лог до последней половины БЕЗ загрузки всего файла в память.

    Раньше здесь был f.readlines() — на большом логе (десятки/сотни МБ) это
    поднимало весь файл в RAM. Теперь читаем только «хвост» через seek:
    прыгаем на середину файла по байтам, отбрасываем возможную обрезанную
    первую строку и переписываем остаток потоково.
    """
    try:
        if not (os.path.exists(path) and os.path.getsize(path) > _MAX_LOG_BYTES):
            return
        size = os.path.getsize(path)
        keep_from = size // 2
        tmp = path + ".rot.tmp"
        with open(path, "rb") as src, open(tmp, "wb") as dst:
            src.seek(keep_from)
            src.readline()  # отбрасываем обрезанную первую строку
            while True:
                chunk = src.read(1024 * 1024)  # 1 МБ за раз
                if not chunk:
                    break
                dst.write(chunk)
        os.replace(tmp, path)
    except Exception:
        try:
            if os.path.exists(path + ".rot.tmp"):
                os.remove(path + ".rot.tmp")
        except Exception:
            pass


def _append_jsonl(path: str, record: dict) -> None:
    try:
        _rotate_if_needed(path)
        with open(path, "a", encoding="utf-8") as f:
            f.write(json.dumps(record, ensure_ascii=False) + "\n")
    except Exception:
        pass


def log_error(
    stage: str,
    message: str,
    *,
    model: str = "",
    user_id: int = 0,
    topic: str = "",
    doc_type: str = "",
    exc_info: BaseException | None = None,
    extra: dict[str, Any] | None = None,
    send_alert: bool = False,
) -> None:
    """Пишет событие пайплайна (не только ошибки — и вехи success/info).

    FIX #5: при send_alert=True или критическом stage (см. _CRITICAL_STAGES)
    дополнительно отправляет оповещение администратору в Telegram —
    в фоновом потоке, с антиспамом 1 сообщение/stage/5 мин.
    """
    record: dict[str, Any] = {
        "ts": time.time(),
        "time": time.strftime("%Y-%m-%d %H:%M:%S"),
        "stage": str(stage),
        "message": str(message)[:2000],
    }
    if model:
        record["model"] = model
    if user_id:
        record["user_id"] = user_id
    if topic:
        record["topic"] = str(topic)[:200]
    if doc_type:
        record["doc_type"] = doc_type
    if exc_info is not None:
        record["exception"] = repr(exc_info)[:500]
        try:
            record["traceback"] = "".join(
                traceback.format_exception(type(exc_info), exc_info, exc_info.__traceback__)
            )[-3000:]
        except Exception:
            pass
    if extra:
        try:
            record["extra"] = json.loads(json.dumps(extra, ensure_ascii=False, default=str))
        except Exception:
            record["extra"] = str(extra)[:1000]
    _append_jsonl(ERROR_LOG_FILE, record)
    try:
        _maybe_alert(str(stage), str(message), record, send_alert)
    except Exception:
        pass


def log_api_call(
    model: str,
    success: bool,
    duration_ms: int,
    stage: str = "",
    error_msg: str = "",
) -> None:
    _append_jsonl(API_LOG_FILE, {
        "ts": time.time(),
        "time": time.strftime("%Y-%m-%d %H:%M:%S"),
        "model": model,
        "success": bool(success),
        "duration_ms": int(duration_ms),
        "stage": stage,
        "error": str(error_msg)[:500],
    })


def _read_recent(path: str, days: int) -> list[dict]:
    cutoff = time.time() - days * 86400
    out: list[dict] = []
    try:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                try:
                    rec = json.loads(line)
                    if float(rec.get("ts", 0)) >= cutoff:
                        out.append(rec)
                except Exception:
                    continue
    except FileNotFoundError:
        pass
    except Exception:
        pass
    return out


# Этапы-вехи (не ошибки) — исключаются из сводки ошибок.
_INFO_STAGES = {
    "api_success", "structure_generated", "human_in_loop",
    "quality_gate", "quality_gate_autofix",
}


def get_error_summary(days: int = 7) -> dict:
    """Сводка ошибок: {"total": int, "stages": {..}, "models": {..}}."""
    records = [r for r in _read_recent(ERROR_LOG_FILE, days)
               if r.get("stage") not in _INFO_STAGES]
    stages: Counter = Counter(r.get("stage", "?") for r in records)
    models: Counter = Counter(r.get("model", "") for r in records if r.get("model"))
    return {
        "total": len(records),
        "stages": dict(stages.most_common()),
        "models": dict(models.most_common()),
    }


def print_error_summary(days: int = 7) -> None:
    s = get_error_summary(days)
    print(f"[ERRORS] за {days} дн.: всего {s['total']}")
    for stage, cnt in s["stages"].items():
        print(f"  {stage}: {cnt}")
