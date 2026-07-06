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
from typing import Any, Optional

_LOG_DIR = os.path.dirname(os.path.abspath(__file__))
ERROR_LOG_FILE = os.path.join(_LOG_DIR, "pipeline_errors.jsonl")
API_LOG_FILE = os.path.join(_LOG_DIR, "api_calls.jsonl")

# Ограничение размера лог-файлов (10 МБ) — при превышении файл усечётся.
_MAX_LOG_BYTES = 10 * 1024 * 1024


class PipelineError(Exception):
    """Ошибка этапа генерации (структура, блок, сборка DOCX)."""


class APIFailure(Exception):
    """Ошибка вызова LLM-провайдера."""


def _rotate_if_needed(path: str) -> None:
    try:
        if os.path.exists(path) and os.path.getsize(path) > _MAX_LOG_BYTES:
            # Оставляем последнюю половину строк.
            with open(path, "r", encoding="utf-8") as f:
                lines = f.readlines()
            with open(path, "w", encoding="utf-8") as f:
                f.writelines(lines[len(lines) // 2:])
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
    exc_info: Optional[BaseException] = None,
    extra: Optional[dict[str, Any]] = None,
) -> None:
    """Пишет событие пайплайна (не только ошибки — и вехи success/info)."""
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
