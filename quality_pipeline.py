# -*- coding: utf-8 -*-
"""quality_pipeline — оркестратор качества: JSON-структура работы,
Human-in-the-loop утверждение и валидация блоков.

Ключевая идея «с первого раза»: структура генерируется как строгий JSON,
валидируется машинно (число глав, непустые названия, уникальность),
при браке — до 2 повторных попыток с указанием конкретной ошибки модели.
"""

from __future__ import annotations

import html
import json
import re
from dataclasses import dataclass, field
from typing import Any, Awaitable, Callable

# ── Статусы этапов ──────────────────────────────────────────────────
_STATUS_EMOJI = {
    "pending": "⏳", "running": "🧠", "done": "✅",
    "failed": "❌", "approved": "👍", "waiting_approval": "🙋",
}


def get_pipeline_status_emoji(status: str) -> str:
    return _STATUS_EMOJI.get(status, "▫️")


@dataclass
class StructureApproval:
    """Данные структуры, отправленной пользователю на утверждение."""
    structure: list = field(default_factory=list)
    doc_type: str = ""
    topic: str = ""
    subject: str = ""
    pages: int = 0
    style: str = "classic"
    approved: bool = False


# ── Валидация JSON-структуры ────────────────────────────────────────

def validate_json_structure(structure: Any, num_chapters: int = 0) -> tuple[bool, str]:
    """Проверяет список глав: [{id,title,subs:[{title}]}]. → (ok, причина)."""
    if not isinstance(structure, list) or not structure:
        return False, "структура не является непустым списком"
    titles_seen: set[str] = set()
    chapters = 0
    for item in structure:
        if not isinstance(item, dict):
            return False, "элемент структуры не объект"
        title = str(item.get("title", "")).strip()
        if not title or len(title) < 5:
            return False, f"пустое/слишком короткое название: «{title}»"
        if re.search(r"(?i)(глава\s*\d+\s*$|раздел\s*\d+\s*$|\bTODO\b|\.\.\.$)", title):
            return False, f"шаблонное название без содержания: «{title}»"
        key = title.lower()
        if key in titles_seen:
            return False, f"дубликат названия: «{title}»"
        titles_seen.add(key)
        if item.get("id") not in ("intro", "conclusion", "literature"):
            chapters += 1
            subs = item.get("subs", [])
            if not isinstance(subs, list):
                return False, "subs не является списком"
            for s in subs:
                st = str((s or {}).get("title", "")).strip() if isinstance(s, dict) else str(s).strip()
                if not st or len(st) < 5:
                    return False, "пустое название подраздела"
    if num_chapters and chapters not in (num_chapters, num_chapters + 1, num_chapters - 1):
        return False, f"глав {chapters}, требовалось {num_chapters}"
    return True, ""


def validate_block_content(text: str, *, min_chars: int = 400,
                           require_citations: bool = True) -> tuple[bool, str]:
    """Проверка сгенерированного блока текста перед принятием."""
    t = (text or "").strip()
    if len(t) < min_chars:
        return False, f"слишком короткий блок: {len(t)} < {min_chars} знаков"
    if re.search(r"(^|\s)(#{1,6}\s|\*\*|```)", t):
        return False, "markdown-разметка в тексте"
    low = t.lower()
    for marker in ("как ии", "как языковая модель", "конечно, вот", "надеюсь, это поможет"):
        if marker in low:
            return False, f"фраза-маркер ИИ: «{marker}»"
    if require_citations and not re.search(r"\[\d+,\s*с\.\s*\d+", t):
        return False, "нет ни одной ссылки формата [N, с. X]"
    if re.search(r"\[\d+,\s*с\.\s*(?=[^\d\s])", t):
        return False, "оборванная ссылка [N, с. ..."
    return True, ""


def _extract_json_array(raw: str) -> list | None:
    """Достаёт JSON-массив из ответа LLM (терпимо к болтовне вокруг)."""
    raw = (raw or "").strip()
    raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.M).strip()
    try:
        data = json.loads(raw)
        if isinstance(data, list):
            return data
        if isinstance(data, dict):
            for key in ("structure", "chapters", "items"):
                if isinstance(data.get(key), list):
                    return data[key]
    except Exception:
        pass
    m = re.search(r"\[[\s\S]*\]", raw)
    if m:
        try:
            data = json.loads(m.group(0))
            if isinstance(data, list):
                return data
        except Exception:
            return None
    return None


def format_structure_for_approval(data: dict) -> str:
    """HTML-сообщение для Telegram с планом работы на утверждение."""
    structure = data.get("structure", [])
    lines = [
        "📋 <b>План работы готов — утвердите структуру</b>",
        f"📄 <b>Тема:</b> {html.escape(str(data.get('topic', '')))}",
        f"📚 <b>Дисциплина:</b> {html.escape(str(data.get('subject', '')))}",
        f"📏 <b>Объём:</b> {data.get('pages', '?')} стр.",
        "━━━━━━━━━━━━━━━━━━━━━━",
    ]
    n = 0
    for item in structure:
        title = html.escape(str(item.get("title", "")).strip())
        if item.get("id") in ("intro", "conclusion", "literature"):
            lines.append(f"▫️ <b>{title.upper()}</b>")
            continue
        n += 1
        lines.append(f"<b>{n} {title}</b>")
        for i, sub in enumerate(item.get("subs", []), 1):
            st = sub.get("title", "") if isinstance(sub, dict) else str(sub)
            lines.append(f"   {n}.{i} {html.escape(str(st).strip())}")
    lines.append("━━━━━━━━━━━━━━━━━━━━━━")
    lines.append("Если структура устраивает — нажмите «✅ Утвердить структуру».")
    return "\n".join(lines)


class PipelineOrchestrator:
    """Оркестратор этапов генерации одной работы."""

    def __init__(self) -> None:
        self.context: dict = {}
        self.stage = "pending"

    def start(self, *, user_id: int = 0, topic: str = "", subject: str = "",
              doc_type: str = "", pages: int = 0, model_key: str = "") -> None:
        self.context = {
            "user_id": user_id, "topic": topic, "subject": subject,
            "doc_type": doc_type, "pages": pages, "model_key": model_key,
        }
        self.stage = "running"

    async def generate_structure(
        self,
        chat_fn: Callable[..., Awaitable[tuple[str, str]]],
        doc_type: str,
        topic: str,
        subject: str,
        num_chapters: int,
        max_retries: int = 2,
    ) -> list[dict]:
        """Генерирует валидированную JSON-структуру. [] при полном провале
        (вызывающий код тогда падает на классический generate_chapter_titles)."""
        model_key = self.context.get("model_key", "deepseek")
        subs_per_chapter = 2 if num_chapters >= 3 else 3

        base_prompt = (
            f"Составь структуру работы типа «{doc_type}» на тему «{topic}», "
            f"дисциплина «{subject}». Ровно {num_chapters} глав(ы), "
            f"в каждой {subs_per_chapter}–3 подраздела.\n"
            "ТРЕБОВАНИЯ К НАЗВАНИЯМ: конкретные, содержательные, отражают тему; "
            "первая глава — теоретическая, последняя — практическая/аналитическая; "
            "запрещены шаблоны «Глава 1», «Основная часть», многоточия.\n"
            "ОТВЕТ — ТОЛЬКО валидный JSON-массив без пояснений и markdown:\n"
            '[{"id": "ch1", "title": "Название главы", '
            '"subs": [{"title": "Название подраздела"}]}]'
        )

        error_hint = ""
        for attempt in range(max_retries + 1):
            prompt = base_prompt + (
                f"\n\nПРЕДЫДУЩАЯ ПОПЫТКА ОТКЛОНЕНА: {error_hint}. Исправь это."
                if error_hint else ""
            )
            try:
                raw, _model = await chat_fn(
                    model_key,
                    [
                        {"role": "system",
                         "content": "Ты генерируешь строго валидный JSON. Никакого текста вне JSON."},
                        {"role": "user", "content": prompt},
                    ],
                    2048,
                )
            except Exception as e:
                error_hint = f"ошибка API: {e}"
                continue

            structure = _extract_json_array(raw)
            if structure is None:
                error_hint = "ответ не является валидным JSON-массивом"
                continue
            # Нормализация: id и subs всегда присутствуют
            for i, item in enumerate(structure):
                if isinstance(item, dict):
                    item.setdefault("id", f"ch{i + 1}")
                    item.setdefault("subs", [])
            ok, reason = validate_json_structure(structure, num_chapters)
            if ok:
                self.stage = "done"
                return structure
            error_hint = reason

        self.stage = "failed"
        print(f"[PIPELINE] Структура не сгенерирована: {error_hint}")
        return []

    def build_approval_data(self, structure: list, doc_type: str, topic: str,
                            subject: str, pages: int, style: str = "classic") -> dict:
        self.stage = "waiting_approval"
        return {
            "structure": structure,
            "doc_type": doc_type,
            "topic": topic,
            "subject": subject,
            "pages": pages,
            "style": style,
        }
