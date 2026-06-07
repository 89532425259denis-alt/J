# -*- coding: utf-8 -*-
from __future__ import annotations

"""ГОСТ-АССИСТЕНТ v2.7 — ТОЧНЫЕ СТРАНИЦЫ + ДИСЦИПЛИНА

Главные изменения v2.1 относительно v2.0:
────────────────────────────────────────────────────────────────
1. ★ ТОЧНОЕ ЧИСЛО СТРАНИЦ (±1).
   После генерации DOCX конвертируется в PDF через LibreOffice,
   реально пересчитываются страницы (pypdf / pdfinfo). Если страниц
   меньше цели — главы дозаполняются; если больше — обрезаются
   по границам абзацев. Цикл до 3 итераций.

2. ★ ПРАВИЛЬНАЯ НУМЕРАЦИЯ СТРАНИЦ ПО ГОСТ.
   Титульный лист включается в общую нумерацию, но цифра на нём
   не отображается (different_first_page_header_footer). На странице
   содержания появляется цифра «2», далее сквозная. Номер — внизу
   по центру (или сверху, как настроено в GOST).

3. ★ ЛИМИТ ДЛЯ БЕСПЛАТНОГО РЕЖИМА — 1 ГЕНЕРАЦИЯ В 5 ДНЕЙ.
   Реализовано через кулдаун (FREE_COOLDOWN_SECONDS = 432000),
   независимо от смены календарных суток. VIP и платные — без лимитов.

4. ★ v2.2: подглавы генерируются отдельными промптами (нет пустых глав).
   Заключение пишется ПОСЛЕ глав и видит их реальное содержание.
   ГОСТ-сноски [1, с. 45] обязательны в каждом абзаце.

5. Все базовые возможности v2.0 сохранены: красивые заголовки глав,
   умный/классический стиль, прогресс-бар с ETA, парсинг «своего ГОСТ»,
   custom-документы, аккуратные сообщения с HTML-форматированием.
"""

import asyncio
import io
import json
import os
import re
import shutil
import subprocess
import time
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional

import aiohttp
from aiohttp import web
from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    BufferedInputFile,
    CallbackQuery,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    LabeledPrice,
    Message,
    PreCheckoutQuery,
)
from aiogram.utils.keyboard import InlineKeyboardBuilder

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt, RGBColor


# ═══════════════════════════════════════════════════════════════
#  КРАСИВАЯ АНИМАЦИЯ ПРОГРЕССА С ETA
# ═══════════════════════════════════════════════════════════════

# Рамки спиннера — меняются каждую секунду для иллюзии движения
_SPINNER_FRAMES = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
_BAR_FULL  = "█"
_BAR_HALF  = "▓"
_BAR_EMPTY = "░"


def _spinner() -> str:
    """Возвращает текущий кадр спиннера на основе текущего времени."""
    idx = int(time.monotonic() * 4) % len(_SPINNER_FRAMES)
    return _SPINNER_FRAMES[idx]


def _progress_bar(done: int, total: int, width: int = 16) -> str:
    """Рисует красивый прогресс-бар с половинным блоком на границе."""
    total = max(1, int(total))
    done  = max(0, min(int(done), total))
    ratio = done / total
    filled_f = width * ratio
    filled_i = int(filled_f)
    half     = filled_f - filled_i >= 0.5
    bar      = _BAR_FULL * filled_i
    if half and filled_i < width:
        bar += _BAR_HALF
        bar += _BAR_EMPTY * (width - filled_i - 1)
    else:
        bar += _BAR_EMPTY * (width - filled_i)
    return bar


def _fmt_time(seconds: float) -> str:
    """Форматирует секунды в mm:ss."""
    s = max(0, int(seconds))
    m, s = divmod(s, 60)
    if m:
        return f"{m}м {s:02d}с"
    return f"{s}с"


@dataclass
class Progress:
    """Красивый прогресс с анимацией, ETA и шагами."""
    msg: Message
    title: str
    total_steps: int
    model_name: str = ""
    done: int = 0
    label: str = "Подготовка..."
    _last_text: str = field(default="", repr=False)
    _last_ts: float = field(default=0.0, repr=False)
    _start_ts: float = field(default_factory=time.monotonic, repr=False)
    _step_labels: list = field(default_factory=list, repr=False)

    def _elapsed(self) -> float:
        return time.monotonic() - self._start_ts

    def _eta(self) -> str:
        elapsed = self._elapsed()
        if self.done == 0:
            return "считаю..."
        rate = self.done / elapsed  # шагов в секунду
        remaining_steps = self.total_steps - self.done
        if rate <= 0:
            return "..."
        eta_sec = remaining_steps / rate
        return _fmt_time(eta_sec)

    def render(self) -> str:
        pct      = int(self.done * 100 / max(1, self.total_steps))
        bar      = _progress_bar(self.done, self.total_steps)
        spin     = _spinner()
        elapsed  = _fmt_time(self._elapsed())
        eta      = self._eta()
        model    = f"\n🤖 <b>Модель:</b> <i>{self.model_name}</i>" if self.model_name else ""
        step_num = min(self.done + 1, self.total_steps)

        return (
            f"{spin} <b>{self.title}</b>\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"<code>[{bar}] {pct}%</code>\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"📌 <b>Шаг {step_num}/{self.total_steps}:</b> {self.label}{model}\n"
            f"⏱ Прошло: <code>{elapsed}</code> · До конца: <code>{eta}</code>"
        )

    async def update(
        self,
        *,
        label: Optional[str] = None,
        step_done: bool = False,
        model_name: Optional[str] = None,
        force: bool = False,
        min_interval: float = 1.0,
    ) -> None:
        if label is not None:
            self.label = label
        if model_name is not None:
            self.model_name = model_name
        if step_done:
            self.done = min(self.total_steps, self.done + 1)

        text = self.render()
        now  = time.monotonic()

        # Не спамим обновлениями чаще min_interval
        if not force and text == self._last_text:
            return
        if not force and (now - self._last_ts) < min_interval:
            return

        try:
            await self.msg.edit_text(text, parse_mode="HTML")
            self._last_text = text
            self._last_ts   = now
        except Exception:
            pass

    async def finish(self, text: str) -> None:
        try:
            await self.msg.edit_text(text, parse_mode="HTML")
        except Exception:
            pass

    async def delete(self) -> None:
        try:
            await self.msg.delete()
        except Exception:
            pass

    async def animate_loop(self, stop_event: asyncio.Event) -> None:
        """Фоновый цикл: обновляет спиннер и ETA каждую секунду."""
        while not stop_event.is_set():
            try:
                text = self.render()
                now  = time.monotonic()
                if text != self._last_text and (now - self._last_ts) >= 1.0:
                    await self.msg.edit_text(text, parse_mode="HTML")
                    self._last_text = text
                    self._last_ts   = now
            except Exception:
                pass
            await asyncio.sleep(1.2)


# ═══════════════════════════════════════════════════════════════
#  КОНФИГ
# ═══════════════════════════════════════════════════════════════

TOKENS = {
    "BOT_TOKEN": "",
    "OPENROUTER_KEY": "",
    "DEEPSEEK_KEY": "",
    "GROQ_KEY": "",
    "VIP_USERS": "5291613279",
}

CONFIG_FILE      = "bot_config.json"
GOST_CONFIG_FILE = "gost_configs.json"
USAGE_FILE       = "usage_limits.json"


def _load_json(path: str, default):
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def _save_json(path: str, data) -> None:
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[WARN] Ошибка сохранения {path}: {e}")


CONFIG      = _load_json(CONFIG_FILE, {})
GOST_CONFIGS = _load_json(GOST_CONFIG_FILE, {})


def cfg(name: str, default: str = "") -> str:
    return (TOKENS.get(name) or os.getenv(name) or CONFIG.get(name) or default).strip()


BOT_TOKEN = cfg("BOT_TOKEN")
if not BOT_TOKEN:
    raise SystemExit("❌ ОШИБКА: не вставлен BOT_TOKEN (в TOKENS, .env или bot_config.json)")

OPENROUTER_KEY = cfg("OPENROUTER_KEY")
DEEPSEEK_KEY   = cfg("DEEPSEEK_KEY")
GROQ_KEY       = cfg("GROQ_KEY")

OPENROUTER_BASE_URL = cfg("OPENROUTER_BASE_URL", "https://openrouter.ai/api/v1").rstrip("/")
DEEPSEEK_BASE_URL   = cfg("DEEPSEEK_BASE_URL",   "https://api.deepseek.com/v1").rstrip("/")
GROQ_BASE_URL       = cfg("GROQ_BASE_URL",       "https://api.groq.com/openai/v1").rstrip("/")

DEEPSEEK_MODEL         = cfg("DEEPSEEK_MODEL",         "deepseek-chat")
OPENROUTER_R1_MODEL    = cfg("OPENROUTER_R1_MODEL",    "deepseek/deepseek-r1-0528:free")
OPENROUTER_GEMINI_MODEL = cfg("OPENROUTER_GEMINI_MODEL", "google/gemini-2.5-flash")
GROQ_MODEL             = cfg("GROQ_MODEL",             "llama-3.3-70b-versatile")

FREE_MODEL_KEY = cfg("FREE_MODEL_KEY", "deepseek")

# Лимиты для БЕСПЛАТНОГО режима
FREE_MAX_PAGES    = int(cfg("FREE_MAX_PAGES",         "15"))
# Кулдаун между бесплатными генерациями: по умолчанию 5 суток (432000 сек).
# Лимит "1 генерация в 5 дней" реализован именно через кулдаун, а не дневной счётчик —
# это даёт точное окно 5×24 ч от момента предыдущей генерации.
FREE_COOLDOWN     = int(cfg("FREE_COOLDOWN_SECONDS",  str(5 * 24 * 60 * 60)))
# FREE_DAILY_LIMIT оставлен для совместимости; основной фильтр — кулдаун.
FREE_DAILY_LIMIT  = int(cfg("FREE_DAILY_LIMIT",       "0"))   # 0 = без дневного лимита

# Лимиты для ПЛАТНОГО режима — платят деньги, получают безлимит
# PAID_DAILY_LIMIT = 0 означает "без лимита"
PAID_DAILY_LIMIT  = int(cfg("PAID_DAILY_LIMIT",       "0"))   # 0 = безлимит
PAID_COOLDOWN     = int(cfg("PAID_COOLDOWN_SECONDS",  "0"))   # 0 = нет кулдауна

# Символов на страницу (ГОСТ: ~1800-2000 знаков с пробелами на стр A4 14pt 1.5 интервал)
def calculate_chars_per_page(gost: dict) -> int:
    """Рассчитывает точное количество символов на страницу по параметрам ГОСТа."""
    font_size = int(gost.get("font_size", 14))
    line_spacing = float(gost.get("line_spacing", 1.5))
    left_mm = int(gost.get("left_margin_mm", 30))
    right_mm = int(gost.get("right_margin_mm", 15))
    top_mm = int(gost.get("top_margin_mm", 20))
    bottom_mm = int(gost.get("bottom_margin_mm", 20))

    page_width_mm = 210
    page_height_mm = 297

    text_width_mm = page_width_mm - left_mm - right_mm
    text_height_mm = page_height_mm - top_mm - bottom_mm

    char_width_mm = font_size * 0.42 * 0.3528
    chars_per_line = max(25, int(text_width_mm / char_width_mm))

    line_height_mm = font_size * line_spacing * 0.3528
    lines_per_page = max(18, int(text_height_mm / line_height_mm))

    base = chars_per_line * lines_per_page

    if line_spacing >= 1.5:
        base = int(base * 0.95)
    if font_size >= 14:
        base = int(base * 0.97)

    return max(1450, min(2100, base))

CHARS_PER_PAGE = int(cfg("CHARS_PER_PAGE", "1850"))
# Страниц без текста (титул + содержание)
NON_TEXT_PAGES = int(cfg("NON_TEXT_PAGES", "2"))

DEEPSEEK_PRICE  = int(cfg("DEEPSEEK_PRICE",           "5"))
GROQ_PRICE      = int(cfg("GROQ_PRICE",               "3"))
OR_GEMINI_PRICE = int(cfg("OPENROUTER_GEMINI_PRICE",  "7"))

MAX_PARALLEL  = int(cfg("MAX_PARALLEL", "1"))
GEN_SEMAPHORE = asyncio.Semaphore(MAX_PARALLEL)

VIP_USERS = {int(x) for x in cfg("VIP_USERS", "").replace(" ", "").split(",") if x.isdigit()}


# ═══════════════════════════════════════════════════════════════
#  ТИПЫ ДОКУМЕНТОВ
# ═══════════════════════════════════════════════════════════════

DOC_TYPES = {
    "referat": {
        "name": "📄 Реферат",
        "word": "РЕФЕРАТ",
        "min_pages": 10,
        "max_pages": 30,
        "structure": "Введение · 2 главы с подглавами · Заключение · Список литературы",
        "desc": "Обзор литературы по теме, изложение основных концепций и выводы",
    },
    "kursovaya": {
        "name": "📚 Курсовая работа",
        "word": "КУРСОВАЯ РАБОТА",
        "min_pages": 25,
        "max_pages": 50,
        "structure": "Введение · 3 главы с подглавами (§) · Заключение · Библиография",
        "desc": "Самостоятельное исследование с анализом, практической частью и выводами",
    },
    "doklad": {
        "name": "🎤 Доклад",
        "word": "ДОКЛАД",
        "min_pages": 5,
        "max_pages": 15,
        "structure": "Введение · Основная часть · Заключение",
        "desc": "Краткий обзор темы для устного выступления",
    },
    "esse": {
        "name": "✍️ Эссе",
        "word": "ЭССЕ",
        "min_pages": 3,
        "max_pages": 10,
        "structure": "Вступление · Аргументы · Авторская позиция",
        "desc": "Авторский взгляд на проблему с аргументацией и личными выводами",
    },
    "kontrolnaya": {
        "name": "📝 Контрольная работа",
        "word": "КОНТРОЛЬНАЯ РАБОТА",
        "min_pages": 10,
        "max_pages": 25,
        "structure": "Теоретическая часть · Практическая часть · Выводы",
        "desc": "Проверочная работа по теории и практике дисциплины",
    },
    "final_project": {
        "name": "📦 Итоговый проект",
        "word": "ИТОГОВЫЙ ПРОЕКТ",
        "min_pages": 35,
        "max_pages": 80,
        "structure": "Введение · 4 главы · Заключение · Приложения · Источники",
        "desc": "Комплексный проект с теорией, практикой и проектной частью",
    },
    "final_referat": {
        "name": "📑 Итоговый реферат",
        "word": "ИТОГОВЫЙ РЕФЕРАТ",
        "min_pages": 15,
        "max_pages": 40,
        "structure": "Введение · 3 главы с подглавами · Заключение · Список литературы",
        "desc": "Расширенный итоговый реферат по ключевой дисциплине с глубоким анализом литературы",
    },
    "article": {
        "name": "📝 Научная статья",
        "word": "НАУЧНАЯ СТАТЬЯ",
        "min_pages": 5,
        "max_pages": 15,
        "structure": "Аннотация · Ключевые слова · Введение · Обзор литературы · Методология · Результаты · Заключение · Список источников",
        "desc": "Научная публикация по ГОСТ Р 7.0.7-2021 (требования к статьям)",
    },
    "vkr": {
        "name": "🎓 ВКР (Дипломная работа)",
        "word": "ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА",
        "min_pages": 40,
        "max_pages": 100,
        "structure": "Введение · 3-4 главы с подглавами · Практический раздел · Заключение · Список использованных источников",
        "desc": "Выпускная квалификационная работа (бакалаврская работа или дипломный проект) по ГОСТ",
    },
    "custom": {
        "name": "🧩 Свой тип",
        "word": "РАБОТА",
        "min_pages": 3,
        "max_pages": 100,
        "structure": "Пользовательская структура",
        "desc": "Любой тип документа с вашими требованиями к структуре",
    },
}

INSTITUTION_TYPES = {
    "school": {
        "name": "🏫 Школа",
        "org_example": "Муниципальное бюджетное общеобразовательное учреждение",
        "name_example": "Средняя общеобразовательная школа № 123",
    },
    "college": {
        "name": "🏛 Колледж / СПО",
        "org_example": "Государственное бюджетное профессиональное образовательное учреждение",
        "name_example": "Колледж информационных технологий № 42",
    },
    "university": {
        "name": "🎓 ВУЗ / Университет",
        "org_example": "Федеральное государственное бюджетное образовательное учреждение высшего образования",
        "name_example": "Московский государственный университет имени М.В. Ломоносова",
    },
    "custom": {
        "name": "✏️ Свой вариант",
        "org_example": "",
        "name_example": "",
    },
}

SUBJECTS = [
    "История", "Обществознание", "Литература", "Русский язык",
    "Математика", "Физика", "Информатика", "Биология",
    "Химия", "География", "Экономика", "Право",
    "Философия", "Психология", "Социология", "Менеджмент",
    "Педагогика", "Медицина", "Архитектура", "Юриспруденция",
]

CITIES = [
    "Москва", "Санкт-Петербург", "Новосибирск", "Екатеринбург",
    "Казань", "Нижний Новгород", "Красноярск", "Челябинск",
    "Омск", "Ростов-на-Дону", "Уфа", "Волгоград",
]


# ═══════════════════════════════════════════════════════════════
#  ГОСТ — дефолты и пользовательские настройки
# ═══════════════════════════════════════════════════════════════

DEFAULT_GOST_CONFIGS: dict = {
    "_base": {
        "font_name":              "Times New Roman",
        "font_size":              14,
        "line_spacing":           1.5,
        "first_line_indent_cm":   1.25,
        "left_margin_mm":         30,
        "right_margin_mm":        15,
        "top_margin_mm":          20,
        "bottom_margin_mm":       20,
        "alignment":              "justify",
        "page_number_position":   "bottom_center",
    }
}

for _t in DOC_TYPES.keys():
    DEFAULT_GOST_CONFIGS.setdefault(_t, dict(DEFAULT_GOST_CONFIGS["_base"]))

if not isinstance(GOST_CONFIGS, dict):
    GOST_CONFIGS = {}

for _t in DOC_TYPES.keys():
    GOST_CONFIGS.setdefault(_t, {})
    for _k, _v in DEFAULT_GOST_CONFIGS[_t].items():
        GOST_CONFIGS[_t].setdefault(_k, _v)

_save_json(GOST_CONFIG_FILE, GOST_CONFIGS)


def get_gost_config(doc_type: str, user_id: Optional[int] = None) -> dict:
    doc_type = doc_type if doc_type in DOC_TYPES else "referat"
    config = dict(GOST_CONFIGS.get(doc_type, DEFAULT_GOST_CONFIGS[doc_type]))
    if user_id:
        user_file = f"user_gost_{user_id}.json"
        u = _load_json(user_file, {})
        if isinstance(u, dict) and doc_type in u and isinstance(u[doc_type], dict):
            config.update(u[doc_type])
    return config


def save_user_gost_config(user_id: int, doc_type: str, config: dict) -> None:
    user_file = f"user_gost_{user_id}.json"
    u = _load_json(user_file, {})
    if not isinstance(u, dict):
        u = {}
    u[doc_type] = config
    _save_json(user_file, u)


# ═══════════════════════════════════════════════════════════════
#  ЛИМИТЫ / VIP
# ═══════════════════════════════════════════════════════════════

def is_vip(user_id: int) -> bool:
    return int(user_id) in VIP_USERS


def today_key() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def fmt_seconds(seconds: int) -> str:
    m, s = divmod(max(0, int(seconds)), 60)
    if m:
        return f"{m} мин {s:02d} сек"
    return f"{s} сек"


def load_usage() -> dict:
    return _load_json(USAGE_FILE, {})


def save_usage(d: dict) -> None:
    _save_json(USAGE_FILE, d)


def _fmt_wait_human(seconds: int) -> str:
    """Удобная человекочитаемая длительность: дни/часы/минуты/секунды."""
    s = max(0, int(seconds))
    d, s = divmod(s, 86400)
    h, s = divmod(s, 3600)
    m, s = divmod(s, 60)
    parts = []
    if d: parts.append(f"{d} дн")
    if h: parts.append(f"{h} ч")
    if m and not d: parts.append(f"{m} мин")
    if not parts: parts.append(f"{s} сек")
    return " ".join(parts)


def check_user_limit(user_id: int, mode: str) -> tuple[bool, str]:
    """Проверяет ограничения.

    Бесплатный режим: одна генерация раз в FREE_COOLDOWN секунд (по умолчанию 5 суток).
    Платные пользователи (mode='paid') — безлимитны (PAID_DAILY_LIMIT=0, PAID_COOLDOWN=0).
    VIP — без ограничений всегда.
    """
    if is_vip(user_id):
        return True, ""

    is_free = mode == "free"

    # Платный режим — без ограничений если PAID_DAILY_LIMIT == 0
    if not is_free and PAID_DAILY_LIMIT == 0 and PAID_COOLDOWN == 0:
        return True, ""

    data = load_usage()
    uid  = str(user_id)
    rec  = data.get(uid, {}) or {}

    now = int(datetime.now().timestamp())

    # ── Кулдаун (главный фильтр для бесплатных) ──
    cooldown = FREE_COOLDOWN if is_free else PAID_COOLDOWN
    ts_key   = "last_free_ts" if is_free else "last_paid_ts"
    last_ts  = int(rec.get(ts_key, 0) or 0)
    if cooldown > 0 and last_ts and (now - last_ts) < cooldown:
        wait     = cooldown - (now - last_ts)
        next_dt  = datetime.fromtimestamp(last_ts + cooldown).strftime("%d.%m.%Y %H:%M")
        kind     = "бесплатная" if is_free else "платная"
        return False, (
            f"⏳ <b>Следующая {kind} генерация будет доступна позже</b>\n\n"
            f"┌─────────────────────────\n"
            f"│ ⌛ Осталось: <b>{_fmt_wait_human(wait)}</b>\n"
            f"│ 📅 Доступна с: <b>{next_dt}</b>\n"
            f"└─────────────────────────\n\n"
            f"💎 Хотите без ожидания? Используйте платный режим — он без лимитов."
        )

    # ── Дневной лимит (опциональный, по умолчанию выключен для бесплатных) ──
    today = today_key()
    if rec.get("date") != today:
        used = 0
    else:
        used = int(rec.get("free" if is_free else "paid", 0) or 0)

    limit = FREE_DAILY_LIMIT if is_free else PAID_DAILY_LIMIT
    if limit > 0 and used >= limit:
        kind = "бесплатных" if is_free else "платных"
        return False, (
            f"🚫 <b>Дневной лимит {kind} генераций исчерпан</b>\n\n"
            f"Использовано: <b>{used}/{limit}</b>\n"
            f"Лимит обновится в полночь 🕛"
        )

    return True, ""


def record_user_generation(user_id: int, mode: str) -> None:
    """Фиксирует факт генерации.

    Хранит:
      - last_free_ts / last_paid_ts — timestamp последней генерации (для кулдауна,
        не сбрасывается полночью);
      - free / paid — счётчик за текущие сутки (используется только если
        включён дневной лимит >0).
    """
    if is_vip(user_id):
        return

    data  = load_usage()
    uid   = str(user_id)
    today = today_key()

    rec = data.get(uid, {}) or {}
    # last_*_ts НЕ обнуляем при смене даты — иначе кулдаун в 5 дней не сработает
    if rec.get("date") != today:
        rec["date"] = today
        rec["free"] = 0
        rec["paid"] = 0

    key    = "free" if mode == "free" else "paid"
    ts_key = "last_free_ts" if mode == "free" else "last_paid_ts"
    rec[key]    = int(rec.get(key, 0) or 0) + 1
    rec[ts_key] = int(datetime.now().timestamp())

    data[uid] = rec
    save_usage(data)


def get_user_limits_info(user_id: int) -> str:
    """Возвращает красивую карточку с лимитами пользователя."""
    if is_vip(user_id):
        return (
            "┌─────────────────────────\n"
            "│ 👑 <b>Статус: VIP</b>\n"
            "│ ♾ Генерации — без лимитов\n"
            "└─────────────────────────"
        )

    data = load_usage()
    uid  = str(user_id)
    rec  = data.get(uid, {}) or {}

    now      = int(datetime.now().timestamp())
    last_free = int(rec.get("last_free_ts", 0) or 0)

    if FREE_COOLDOWN > 0 and last_free and (now - last_free) < FREE_COOLDOWN:
        wait_s   = FREE_COOLDOWN - (now - last_free)
        next_dt  = datetime.fromtimestamp(last_free + FREE_COOLDOWN).strftime("%d.%m.%Y %H:%M")
        free_str = (
            f"⏳ Ждать <b>{_fmt_wait_human(wait_s)}</b>\n"
            f"│ 📅 Доступна: <b>{next_dt}</b>"
        )
    else:
        period_h = FREE_COOLDOWN // 3600
        period_s = f"{period_h // 24} дн" if period_h >= 24 else f"{period_h} ч"
        free_str = f"✅ Доступна (1 раз в {period_s})"

    paid_str = (
        "♾ Безлимитно" if PAID_DAILY_LIMIT == 0
        else f"{int(rec.get('paid', 0) or 0)}/{PAID_DAILY_LIMIT}"
    )

    return (
        "┌─────────────────────────\n"
        f"│ 🆓 Бесплатно: {free_str}\n"
        f"│ ⭐ Платные:    <b>{paid_str}</b>\n"
        "└─────────────────────────"
    )


# ═══════════════════════════════════════════════════════════════
#  AI МОДЕЛИ
# ═══════════════════════════════════════════════════════════════

class ModelStatus:
    AVAILABLE = "✅"
    LIMIT     = "❌ лимит"
    UNKNOWN   = "❓"
    FATAL     = "🔴 ошибка"


AI_MODELS: dict = {
    "deepseek": {
        "name":           "🐋 DeepSeek Chat",
        "base_url":       DEEPSEEK_BASE_URL,
        "api_key":        DEEPSEEK_KEY,
        "model":          DEEPSEEK_MODEL,
        "price_per_page": DEEPSEEK_PRICE,
        "status":         ModelStatus.UNKNOWN,
        "_fatal":         False,
    },
    "deepseek_r1": {
        "name":           "🧠 DeepSeek R1 (OpenRouter)",
        "base_url":       OPENROUTER_BASE_URL,
        "api_key":        OPENROUTER_KEY,
        "model":          OPENROUTER_R1_MODEL,
        "price_per_page": DEEPSEEK_PRICE,
        "status":         ModelStatus.UNKNOWN,
        "_fatal":         False,
    },
    "gemini_or": {
        "name":           "🌟 Gemini 2.5 Flash (OpenRouter)",
        "base_url":       OPENROUTER_BASE_URL,
        "api_key":        OPENROUTER_KEY,
        "model":          OPENROUTER_GEMINI_MODEL,
        "price_per_page": OR_GEMINI_PRICE,
        "status":         ModelStatus.UNKNOWN,
        "_fatal":         False,
    },
    "groq": {
        "name":           "⚡ Groq (LLaMA 3.3 70B)",
        "base_url":       GROQ_BASE_URL,
        "api_key":        GROQ_KEY,
        "model":          GROQ_MODEL,
        "price_per_page": GROQ_PRICE,
        "status":         ModelStatus.UNKNOWN,
        "_fatal":         False,
    },
}

if FREE_MODEL_KEY not in AI_MODELS:
    FREE_MODEL_KEY = "deepseek"


def sanitize_llm_text(raw: str) -> str:
    """Чистит мусор от LLM: markdown-разметку, тройные переводы строк."""
    if not raw:
        return ""
    text = raw.strip()
    # убираем ```код``` блоки
    text = re.sub(r"```[^\n]*\n?", "", text)
    # убираем **жирный** и *курсив* markdown
    text = re.sub(r"\*{1,3}([^*]+)\*{1,3}", r"\1", text)
    # убираем # заголовки markdown
    text = re.sub(r"^#{1,6}\s+", "", text, flags=re.MULTILINE)
    # убираем тройные+ переводы строк
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


async def call_openai_compat(
    info: dict,
    messages: list[dict],
    max_tokens: int = 4096,
    timeout: int = 300,
) -> str:
    """Вызов OpenAI-совместимого API."""
    if info.get("_fatal") or not info.get("api_key"):
        return ""

    base    = info["base_url"].rstrip("/")
    headers = {
        "Content-Type":  "application/json",
        "Authorization": f"Bearer {info['api_key']}",
    }

    if "openrouter.ai" in base:
        headers["HTTP-Referer"] = "https://t.me/gost_assistant_bot"
        headers["X-Title"]      = "GOST Assistant Bot"

    payload = {
        "model":       info["model"],
        "messages":    messages,
        "temperature": 0.7,
        "max_tokens":  max_tokens,
    }

    try:
        async with aiohttp.ClientSession() as sess:
            async with sess.post(
                f"{base}/chat/completions",
                headers=headers,
                json=payload,
                timeout=aiohttp.ClientTimeout(total=timeout),
            ) as r:
                txt = await r.text()
                if r.status == 200:
                    data = json.loads(txt)
                    return data["choices"][0]["message"]["content"]
                if r.status in (401, 402, 403):
                    info["_fatal"] = True
                    info["status"] = ModelStatus.FATAL
                    print(f"[FATAL] {info['name']} — auth error {r.status}")
                elif r.status == 429:
                    info["status"] = ModelStatus.LIMIT
                    print(f"[LIMIT] {info['name']} — rate limit")
                else:
                    print(f"[ERROR] {info['name']} — HTTP {r.status}: {txt[:200]}")
    except asyncio.TimeoutError:
        print(f"[TIMEOUT] {info['name']}")
    except Exception as e:
        print(f"[ERR] {info['name']}: {e}")

    return ""


async def chat_with_model(info: dict, messages: list[dict], max_tokens: int = 4096) -> str:
    raw = await call_openai_compat(info, messages, max_tokens=max_tokens)
    return sanitize_llm_text(raw)


def fallback_chain(primary: str) -> list[str]:
    """Цепочка фоллбэков: сначала primary, потом остальные доступные."""
    preferred = [primary, "deepseek", "deepseek_r1", "gemini_or", "groq"]
    out: list[str] = []
    for k in preferred:
        if k not in AI_MODELS or k in out:
            continue
        info = AI_MODELS[k]
        if info.get("_fatal"):
            continue
        if info.get("api_key"):
            out.append(k)
    return out


async def chat_with_fallback(
    primary: str,
    messages: list[dict],
    max_tokens: int,
) -> tuple[str, str]:
    """Пробует модели по цепочке, возвращает (текст, ключ_модели)."""
    for k in fallback_chain(primary):
        info = AI_MODELS[k]
        text = await chat_with_model(info, messages, max_tokens=max_tokens)
        if text and len(text.strip()) > 50:
            info["status"] = ModelStatus.AVAILABLE
            return text, k
        info["status"] = ModelStatus.LIMIT
    return "", primary


# ═══════════════════════════════════════════════════════════════
#  ГЕНЕРАЦИЯ СТРУКТУРЫ И НАЗВАНИЙ ГЛАВ ЧЕРЕЗ ИИ
# ═══════════════════════════════════════════════════════════════

async def generate_chapter_titles(
    model_key: str,
    doc_type: str,
    topic: str,
    subject: str,
    num_chapters: int,
) -> list[dict]:
    """
    Просит ИИ придумать развёрнутые названия глав и подглав.

    Возвращает список словарей:
    [
      {"title": "Глава 1. Теоретические основы ...", "subs": ["1.1. ...", "1.2. ..."]},
      ...
    ]
    """
    system = (
        "Ты помогаешь составлять структуру академических работ. "
        "Отвечай СТРОГО в формате JSON-массива без пояснений и без markdown. "
        "Каждый элемент: {\"title\": \"...\", \"subs\": [\"...\", \"...\"]}. "
        "Название главы должно быть развёрнутым и конкретным — не 'Глава 1', "
        "а 'Глава 1. Теоретические основы изучения ...'."
    )

    user = (
        f"Тема работы: «{topic}». "
        f"Дисциплина: {subject}. "
        f"Тип документа: {DOC_TYPES.get(doc_type, DOC_TYPES['referat'])['word']}. "
        f"Количество глав: {num_chapters}. "
        f"Каждая глава — 2–3 подглавы. "
        f"Язык: русский. "
        f"Верни ТОЛЬКО JSON-массив."
    )

    messages = [
        {"role": "system", "content": system},
        {"role": "user",   "content": user},
    ]

    raw, _ = await chat_with_fallback(model_key, messages, max_tokens=1500)
    raw    = raw.strip()

    # Вытаскиваем JSON даже если модель добавила мусор
    m = re.search(r"\[.*\]", raw, flags=re.S)
    if m:
        raw = m.group(0)

    try:
        result = json.loads(raw)
        if isinstance(result, list) and all("title" in r for r in result):
            return result
    except Exception:
        pass

    # Фоллбэк: базовые названия
    return _default_chapter_titles(doc_type, topic, num_chapters)


async def verify_discipline_relevance(
    model_key: str,
    topic: str,
    subject: str,
    sample_text: str,
) -> tuple[bool, str]:
    """Проверяет, что сгенерированный текст относится к заданной дисциплине
    (приоритет 🔴: история ≠ геология).

    Возвращает (соответствует?, краткий_комментарий).
    Использует ИИ как классификатор; при сбое считает текст релевантным,
    чтобы не блокировать выдачу.
    """
    if not sample_text or not subject:
        return True, "проверка пропущена"

    sample = sample_text[:2500]
    system = (
        "Ты — научный рецензент. Оцени, соответствует ли фрагмент работы "
        "заявленной учебной дисциплине. Отвечай СТРОГО в формате JSON без "
        'markdown: {"match": true|false, "reason": "одно короткое предложение"}.'
    )
    user = (
        f"Дисциплина: «{subject}».\n"
        f"Тема работы: «{topic}».\n"
        f"Фрагмент текста:\n{sample}\n\n"
        "Соответствует ли содержание дисциплине? "
        "match=false только если текст явно из другой области знаний "
        "(например, дисциплина «История», а текст чисто геологический)."
    )
    try:
        raw, _ = await chat_with_fallback(
            model_key,
            [{"role": "system", "content": system},
             {"role": "user", "content": user}],
            max_tokens=200,
        )
        m = re.search(r"\{.*\}", raw, flags=re.S)
        if m:
            data = json.loads(m.group(0))
            return bool(data.get("match", True)), str(data.get("reason", "")).strip()
    except Exception as e:
        print(f"[RELEVANCE] check failed: {e}")
    return True, "проверка недоступна"


def _default_chapter_titles(doc_type: str, topic: str, num_chapters: int) -> list[dict]:
    """Запасные названия глав если ИИ не ответил."""
    defaults = [
        {
            "title": f"Глава 1. Теоретические основы исследования темы «{topic[:40]}»",
            "subs":  [
                f"1.1. Понятие и сущность исследуемой проблематики",
                f"1.2. Исторические предпосылки и современное состояние",
                f"1.3. Нормативно-правовая база и научные подходы",
            ],
        },
        {
            "title": f"Глава 2. Анализ и современное состояние проблемы",
            "subs":  [
                f"2.1. Характеристика основных факторов и условий",
                f"2.2. Сравнительный анализ подходов и методов",
                f"2.3. Проблемы и противоречия в изучаемой области",
            ],
        },
        {
            "title": f"Глава 3. Практические аспекты и пути решения",
            "subs":  [
                f"3.1. Практика применения теоретических положений",
                f"3.2. Рекомендации по совершенствованию",
                f"3.3. Перспективы развития исследуемой сферы",
            ],
        },
        {
            "title": f"Глава 4. Оценка результатов и выводы",
            "subs":  [
                f"4.1. Обобщение результатов исследования",
                f"4.2. Практическая значимость выводов",
            ],
        },
    ]
    return defaults[:num_chapters]


# ═══════════════════════════════════════════════════════════════
#  РАСЧЁТ ОБЪЁМА ТЕКСТА — ИСПРАВЛЕННЫЙ
# ═══════════════════════════════════════════════════════════════

def target_chars(pages: int, gost: dict = None) -> int:
    """
    Целевое количество символов с пробелами для основного текста.
    Использует динамический расчёт из ГОСТа, если gost передан.
    """
    if gost:
        chars = calculate_chars_per_page(gost)
    else:
        chars = CHARS_PER_PAGE
    text_pages = max(1, pages - NON_TEXT_PAGES)
    return text_pages * chars


def tokens_for_chars(chars: int) -> int:
    """
    Примерно 1 токен = 2.5 символа для русского (консервативно).
    Добавляем 100% запас чтобы ИИ не обрезало текст на полуслове.
    """
    return max(2000, min(32000, int(chars / 2.5 * 2.0)))


def _style_instruction(writing_style: str, doc_type: str = "") -> str:
    """Возвращает инструкцию по стилю. Учитывает жанр (эссе/доклад/академическое)."""
    if doc_type == "esse":
        return (
            "Стиль эссе: от первого лица, живой язык, личная позиция. "
            "Избегай канцеляризмов и наукообразия."
        )
    if doc_type == "doklad":
        return (
            "Стиль доклада: чёткий, устно-ориентированный, короткие предложения. "
            "Конкретные факты и цифры."
        )
    if doc_type == "article":
        return (
            "Стиль научной статьи: строгий академический язык, высокая точность формулировок, "
            "использование современной научной терминологии, объективность изложения, "
            "активное цитирование авторитетных источников, отсутствие эмоциональной окраски."
        )
    if writing_style == "smart":
        return (
            "Используй сложный академический язык: многосоставные предложения, "
            "научную терминологию дисциплины, ссылки на теории и концепции, "
            "глубокий аналитический разбор. Язык должен быть характерен для "
            "диссертаций и монографий."
        )
    return (
        "Используй чёткий деловой научный стиль: ясные формулировки, "
        "логичная структура, конкретные утверждения без излишней сложности. "
        "Язык должен быть понятен образованному читателю без специальной подготовки."
    )


def strict_prompt(task: str, chars: int, writing_style: str = "classic",
                  doc_type: str = "") -> str:
    """Инструкция для генерации текста нужного объёма. Учитывает жанр."""
    style_instr = _style_instruction(writing_style, doc_type)
    
    # Вычисляем примерный объем страниц
    est_pages = max(1, round(chars / CHARS_PER_PAGE, 1))

    base = (
        f"{task}\n\n"
        f"⚠️ ВНИМАНИЕ: Пользователь заказал работу строго определенного объема. "
        f"Требуемый объём для этого раздела: МИНИМУМ {chars} знаков с пробелами (около {est_pages} стр. текста).\n"
        f"Ты должен четко понимать лимит страниц и НЕ ОБРЫВАТЬ мысль на полуслове! "
        f"Заканчивай свои абзацы и мысли логично, последовательно, но при этом кратко и лаконично, "
        f"если пользователь ограничил объем. Если лимит знаков/страниц близок к концу, плавно завершай "
        f"раздел логичным выводом, укладываясь строго в рамки и не оставляя незаконченных фраз.\n\n"
        f"Стиль: {style_instr}\n"
        "ЗАПРЕЩЕНО: markdown-разметка, символы **, ##. "
        "Запрещены фразы «как ИИ», «давайте рассмотрим» в начале абзацев. "
        "Не повторяй одинаковые слова в одном предложении. Не начинай два абзаца подряд одним словом."
    )

    if doc_type == "esse":
        return base + (
            "\nЭто ЭССЕ — никаких «глав», «разделов», «мы рассмотрели». "
            "Только «я», личная позиция и размышление."
        )
    elif doc_type == "doklad":
        return base + (
            "\nЭто ДОКЛАД — чётко, по делу, короткие абзацы. "
            "Ссылки [1], [2] допустимы."
        )
    elif doc_type == "article":
        return base + (
            "\nЭто НАУЧНАЯ СТАТЬЯ — пиши строгим, сухим академическим языком. "
            "Каждый абзац — 5–8 предложений, без маркированных списков. "
            "Обязательно используй сноски в формате ГОСТ: [1, с. 45] или [2]. "
            "Используй сноски в каждом абзаце — номера источников 1, 2, 3..."
        )
    else:
        return base + (
            "\nКаждый абзац — 5–8 предложений, без маркированных списков. "
            "Разрешены сноски в формате ГОСТ: [1, с. 45] или [3]. "
            "Используй сноски в каждом абзаце (1–3 ссылки) — номера источников 1, 2, 3..."
        )


# ═══════════════════════════════════════════════════════════════
#  ПОСТРОЕНИЕ ПРОМПТОВ ПО ТИПАМ ДОКУМЕНТОВ
# ═══════════════════════════════════════════════════════════════

def build_prompts(
    doc_type: str,
    topic: str,
    subject: str,
    pages: int,
    source: str,
    chapter_titles: list[dict],
    writing_style: str = "classic",
) -> dict[str, str]:
    """
    Возвращает словарь {ключ_блока: промпт}.
    chapter_titles — список из generate_chapter_titles().
    """
    total = target_chars(pages)
    s     = (source or "").strip()[:12000]
    ctx   = f"\n\nИсходные материалы для использования:\n{s}\n" if s else ""

    if doc_type == "esse":
        # Минимальные границы для страховки (если расчёт даёт слишком мало)
        intro_min = 1200
        main_min = 1500

        intro_chars = max(intro_min, int(total * 0.20))
        main_chars = max(main_min, int(total * 0.28))

        return {
            "intro":      strict_prompt(
                f"Напиши вступление эссе на тему «{topic}», предмет «{subject}».{ctx}"
                f"Обозначь проблему, её актуальность, цель и подход автора.",
                intro_chars,
                writing_style, doc_type,
            ),
            "main1":      strict_prompt(
                f"Напиши первый аргумент в эссе «{topic}». "
                f"Приведи конкретные факты, мнения учёных и доказательства. "
                f"Пиши развёрнуто, минимум 3-4 абзаца.",
                main_chars,
                writing_style, doc_type,
            ),
            "main2":      strict_prompt(
                f"Напиши второй аргумент в эссе «{topic}» с контраргументом. "
                f"Рассмотри противоположную точку зрения и опровергни её. "
                f"Пиши развёрнуто, минимум 3-4 абзаца.",
                main_chars,
                writing_style, doc_type,
            ),
            "literature": (
                f"Составь список из 3–5 источников по теме «{topic}» "
                f"строго по ГОСТ Р 7.0.5-2008. "
                f"Включи ТОЛЬКО источники, соответствующие дисциплине «{subject}». "
                f"Если тема не совпадает с дисциплиной — подбери источники "
                f"по дисциплине, связанные с темой. "
                f"Формат: 1. Автор А.А. Название. — М.: Издательство, год. — N с.\n"
                f"Только список, без заголовков и пояснений."
            ),
        }
        # Заключение для эссе генерируется ПОСЛЕ всех аргументов (см. generate_text_blocks)

    if doc_type == "doklad":
        return {
            "intro":      strict_prompt(
                f"Введение доклада «{topic}», предмет «{subject}».{ctx}"
                f"Актуальность, цель, задачи.",
                int(total * 0.14),
                writing_style, doc_type,
            ),
            "part1":      strict_prompt(
                f"Раздел 1 доклада «{topic}»: ключевые понятия и определения."
                f"Раскрой теоретическую базу.",
                int(total * 0.28),
                writing_style, doc_type,
            ),
            "part2":      strict_prompt(
                f"Раздел 2 доклада «{topic}»: факты, статистика, примеры из практики.",
                int(total * 0.32),
                writing_style,
                doc_type,
            ),
            "literature": (
                f"Составь список из 8–12 источников по теме «{topic}» по ГОСТ. "
                f"Только нумерованный список."
            ),
        }
        # Заключение для доклада генерируется ПОСЛЕ разделов (см. generate_text_blocks)

    if doc_type == "article":
        intro_chars = max(1000, int(total * 0.15))
        lit_review_chars = max(1200, int(total * 0.20))
        methodology_chars = max(1000, int(total * 0.15))
        results_chars = max(1500, int(total * 0.25))
        return {
            "abstract": (
                f"Напиши аннотацию (abstract) и ключевые слова (keywords) для научной статьи "
                f"на тему «{topic}» по дисциплине «{subject}». "
                f"Аннотация должна быть на русском и английском языках (около 150-200 слов на каждом). "
                f"Ключевые слова: 5-8 слов/словосочетаний на русском и английском.\n"
                f"Не используй markdown, начни прямо с текста 'Аннотация / Abstract'."
            ),
            "intro":      strict_prompt(
                f"Введение научной статьи «{topic}», предмет «{subject}».{ctx}"
                f"Актуальность темы, формулировка проблемы, цель исследования.",
                intro_chars,
                writing_style, doc_type,
            ),
            "lit_review": strict_prompt(
                f"Раздел 'Обзор литературы' (Literature Review) для научной статьи «{topic}». "
                f"Проанализируй исследования отечественных и зарубежных авторов за последние годы. "
                f"Используй ГОСТ-сноски [1], [2].",
                lit_review_chars,
                writing_style, doc_type,
            ),
            "methodology": strict_prompt(
                f"Раздел 'Методология' (Methodology) для научной статьи «{topic}». "
                f"Опиши используемые методы исследования, теоретическую или эмпирическую базу, "
                f"обоснуй применимость методов.",
                methodology_chars,
                writing_style, doc_type,
            ),
            "results":     strict_prompt(
                f"Раздел 'Результаты и обсуждение' (Results and Discussion) для научной статьи «{topic}». "
                f"Подробно раскрой результаты исследования, проведи анализ, приведи рассуждения, сравнения.",
                results_chars,
                writing_style, doc_type,
            ),
            "literature": (
                f"Составь список литературы (References) из 10–15 авторитетных источников по теме «{topic}» "
                f"в соответствии с ГОСТ Р 7.0.5-2008 / ГОСТ Р 7.0.7-2021. "
                f"Только нумерованный список."
            ),
        }

    # Реферат, курсовая, контрольная, итоговый проект, свой тип
    # Используем chapter_titles для развёрнутых названий
    num_ch  = len(chapter_titles)
    prompts = {}

    # Введение — 10% от текста
    # Генерируем ВРЕМЕННОЕ введение; оно будет заменено после генерации всех глав
    prompts["intro"] = strict_prompt(
        f"Напиши введение для {DOC_TYPES.get(doc_type, DOC_TYPES['referat'])['word'].lower()}а "
        f"на тему «{topic}», предмет «{subject}».{ctx}"
        f"Раскрой: актуальность темы, степень разработанности (упомяни 3–5 реальных "
        f"исследователей/учёных по этой теме — называй их имена в тексте), "
        f"цель, задачи (3–5 задач), объект и предмет исследования, методы, "
        f"краткую структуру работы. "
        f"Используй ссылки на источники [1], [2] где уместно.",
        int(total * 0.10),
        writing_style, doc_type,
    )

    # ═══ Главы — КАЖДАЯ ПОДГЛАВА ОТДЕЛЬНЫМ ПРОМПТОМ ═══
    # Сначала считаем общее количество подглав для равномерного распределения
    all_subs = []
    for i, ch in enumerate(chapter_titles, start=1):
        subs = ch.get("subs", [])
        for j, sub_title in enumerate(subs, start=1):
            all_subs.append((i, j, ch["title"], sub_title))

    if all_subs:
        sub_share = 0.75 / len(all_subs)  # 75% текста делим поровну между подглавами
    else:
        sub_share = 0.75 / max(1, num_ch)

    for i, j, ch_title, sub_title in all_subs:
        key      = f"ch{i}_s{j}"
        sub_chars = int(total * sub_share)
        prompts[key] = strict_prompt(
            f"Напиши ПОЛНЫЙ текст подглавы «{sub_title}».\n"
            f"Глава: «{ch_title}».\n"
            f"Тема всей работы: «{topic}».\n"
            f"Это отдельная подглава — пиши её как завершённый смысловой блок.\n"
            f"Начни с заголовка подглавы '{sub_title}' на отдельной строке, "
            f"затем идёт содержательный текст (минимум 3–4 развёрнутых абзаца).\n"
            f"Используй ссылки на источники в формате [1, с. 45] или [3] — "
            f"минимум 2–3 ссылки на источники в этой подглаве.",
            sub_chars,
            writing_style, doc_type,
        )

    # Заключение НЕ включаем в batch — оно генерируется ПОСЛЕ всех глав
    # (см. generate_text_blocks)

    # Библиография
    num_sources = 12 if pages <= 20 else 20
    prompts["literature"] = (
        f"Составь список из {num_sources}–{num_sources + 5} источников по теме «{topic}» "
        f"строго по ГОСТ Р 7.0.5-2008. Включи: монографии, учебники, статьи из журналов, "
        f"нормативные документы (если уместно), интернет-ресурсы. "
        f"Формат: 1. Автор А.А. Название / А.А. Автор. — М.: Изд-во, год. — N с.\n"
        f"Только нумерованный список без заголовков."
    )

    return prompts


# ═══════════════════════════════════════════════════════════════
#  СТРУКТУРА DOCX-БЛОКОВ
# ═══════════════════════════════════════════════════════════════

def generate_structure(
    doc_type: str,
    parts: dict[str, str],
    chapter_titles: list[dict],
) -> list[tuple]:
    """
    Возвращает список блоков для DOCX:
    (заголовок, уровень_heading, текст_абзаца, список_подглав)

    Уровень: 1 = Heading 1, 2 = Heading 2
    список_подглав = [] или [(название_подглавы, текст), ...]
    """
    if doc_type == "esse":
        return [
            ("ВВЕДЕНИЕ", 1, parts.get("intro", ""), []),
            ("ОСНОВНАЯ ЧАСТЬ", 1, "", [
                ("Аргумент 1", parts.get("main1", "")),
                ("Аргумент 2. Контраргумент и опровержение", parts.get("main2", "")),
            ]),
            ("ЗАКЛЮЧЕНИЕ", 1, parts.get("conclusion", ""), []),
            ("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", 1, parts.get("literature", ""), []),
        ]

    if doc_type == "doklad":
        return [
            ("ВВЕДЕНИЕ", 1, parts.get("intro", ""), []),
            ("1. Теоретические основы и ключевые понятия", 1, parts.get("part1", ""), []),
            ("2. Факты, статистика и практические примеры", 1, parts.get("part2", ""), []),
            ("ЗАКЛЮЧЕНИЕ", 1, parts.get("conclusion", ""), []),
            ("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", 1, parts.get("literature", ""), []),
        ]

    if doc_type == "article":
        return [
            ("АННОТАЦИЯ И КЛЮЧЕВЫЕ СЛОВА / ABSTRACT AND KEYWORDS", 1, parts.get("abstract", ""), []),
            ("ВВЕДЕНИЕ", 1, parts.get("intro", ""), []),
            ("ОБЗОР ЛИТЕРАТУРЫ", 1, parts.get("lit_review", ""), []),
            ("МЕТОДОЛОГИЯ ИССЛЕДОВАНИЯ", 1, parts.get("methodology", ""), []),
            ("РЕЗУЛЬТАТЫ И ИХ ОБСУЖДЕНИЕ", 1, parts.get("results", ""), []),
            ("ЗАКЛЮЧЕНИЕ", 1, parts.get("conclusion", ""), []),
            ("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", 1, parts.get("literature", ""), []),
        ]

    # Универсальная структура (реферат / курсовая / контрольная / итоговый / свой)
    blocks = [("ВВЕДЕНИЕ", 1, parts.get("intro", ""), [])]

    for i, ch in enumerate(chapter_titles, start=1):
        subs = ch.get("subs", [])
        sub_blocks = []
        chapter_text_parts = []
        for j, sub_title in enumerate(subs, start=1):
            key = f"ch{i}_s{j}"
            sub_text = parts.get(key, "")
            if sub_text:
                sub_blocks.append((sub_title, sub_text))
                chapter_text_parts.append(sub_text)

        # Если новый формат не сработал — fallback на старый ключ "ch{i}"
        if not sub_blocks:
            fallback_text = parts.get(f"ch{i}", "")
            if fallback_text:
                chapter_text_parts = [fallback_text]
                for s_title in subs:
                    sub_blocks.append((s_title, ""))

        # ВАЖНО: всегда заполняем field[2] полным текстом главы,
        # даже при наличии подглав — это нужно для _trim/_expand
        full_chapter_text = "\n\n".join(chapter_text_parts) if chapter_text_parts else ""
        blocks.append((ch["title"], 1, full_chapter_text, sub_blocks))

    blocks.append(("ЗАКЛЮЧЕНИЕ", 1, parts.get("conclusion", ""), []))
    blocks.append(("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", 1, parts.get("literature", ""), []))

    return blocks


# ═══════════════════════════════════════════════════════════════
#  ГЕНЕРАЦИЯ ТЕКСТА С ДОЗАПОЛНЕНИЕМ
# ═══════════════════════════════════════════════════════════════

def _clean_ai_artifacts(text: str) -> str:
    """Удаляет признаки ИИ: ***, ## и повторяющиеся слова в предложениях.
    ГОСТ-сноски [1], [2, с. 45] СОХРАНЯЕТ."""
    if not text:
        return ""
    # НЕ убираем сноски [1], [2, с. 45] — это ГОСТ! Только мусорный markdown.
    # Убираем **жирный** markdown
    text = re.sub(r"\*{1,3}([^*\n]+)\*{1,3}", r"\1", text)
    # Убираем # заголовки markdown
    text = re.sub(r"^#{1,6}\s+", "", text, flags=re.MULTILINE)
    # Убираем горизонтальные линии ***
    text = re.sub(r"^\*{3,}\s*$", "", text, flags=re.MULTILINE)
    text = re.sub(r"^-{3,}\s*$", "", text, flags=re.MULTILINE)
    # Убираем тройные+ переводы строк
    text = re.sub(r"\n{3,}", "\n\n", text)
    # Убираем фразы-маркеры ИИ
    ai_phrases = [
        r"(?i)как (языковая )?модель[,\s]",
        r"(?i)я (являюсь|являюсь )?ИИ[,\s]",
        r"(?i)в заключение следует отметить,?\s+что\s+",
        r"(?i)таким образом,?\s+можно сделать вывод",
    ]
    for p in ai_phrases:
        text = re.sub(p, "", text)
    # Типографика: тире/кавычки/# (приоритет 🟡)
    text = _normalize_typography(text)
    return text.strip()


PAGE_CALC_PROMPT = """
ТЫ ДОЛЖЕН СЛЕДОВАТЬ ЭТИМ ПРАВИЛАМ:

1. 1 страница = 1850 знаков с пробелами.
2. 5 страниц эссе = 7400 знаков основного текста (титул и содержание не в счёт).
3. Аргумент 1: 2000 знаков. Аргумент 2: 2000 знаков. Введение: 1400 знаков. Заключение: 1000 знаков.
4. ЗАПРЕЩЕНО писать: «Вступление к эссе», «Конечно. Вот второй аргумент», «Вот текст», «Объем текста строго выдержан».
5. Начинай писать СРАЗУ содержательный текст, без служебных фраз и заголовков-пояснений.
"""
async def generate_text_blocks(
    topic: str,
    pages: int,
    doc_type: str,
    subject: str,
    model_key: str,
    source: str,
    chapter_titles: list[dict],
    writing_style: str = "classic",
    prog: Optional["Progress"] = None,
) -> dict[str, str]:
    """
    Генерирует все текстовые блоки.
    Если текст получился короче цели — дозаполняет до нужного объёма.
    """
    prompts    = build_prompts(doc_type, topic, subject, pages, source, chapter_titles, writing_style)
    total_chars = target_chars(pages)
    # "conclusion" больше НЕ в prompts — генерируем его отдельно после глав
    num_blocks = max(1, len(prompts))

    parts: dict[str, str] = {}
    step = 0

    # ── Базовые правила для всех типов ──
    style_sys = (
        "Ты пишешь тексты на русском языке. "
        "НЕ используй markdown-разметку (никаких **, ##, ---). "
        "Не повторяй одно слово несколько раз в одном предложении. "
        "Не начинай два абзаца подряд одним и тем же словом. "
        # Привязка к дисциплине (приоритет 🔴)
        f"ОБЯЗАТЕЛЬНО: текст должен соответствовать дисциплине «{subject}». "
        f"Раскрывай тему «{topic}» строго через предмет, методы и терминологию "
        f"дисциплины «{subject}». Если тема относится к другой области знаний, "
        f"всё равно рассматривай её ПОД УГЛОМ дисциплины «{subject}» "
        f"(например, дисциплина «История» → хронология, источники, периодизация; "
        f"«Геология» → породы, процессы, строение Земли). "
        f"Не подменяй дисциплину смежной. "
    )
    style_sys += "\n" + PAGE_CALC_PROMPT

    if doc_type == "esse":
        style_sys += (
            "Это ЭССЕ — пиши ОТ ПЕРВОГО ЛИЦА: «я считаю», «по моему мнению», "
            "«меня поражает», «важно отметить». Текст должен быть живым, "
            "размышляющим, с личной авторской позицией. "
            "Никаких «мы рассмотрели», «было установлено» — только «я». "
            "Избегай наукообразия и канцеляризмов.\n"
            "ВАЖНО: заявленная дисциплина — «{subject}». Если тема и дисциплина "
            "из разных областей (например, тема «Байкал», дисциплина «Психология») — "
            "рассматривай тему ЧЕРЕЗ ПРИЗМУ ДИСЦИПЛИНЫ. "
            "Для психологии: восприятие, эмоции, когнитивные эффекты, "
            "восстановительная среда, эффект благоговения (awe). "
            "Для истории: хронология, личности, события. "
            "Для биологии: виды, экосистемы, эволюция. "
            "НЕ пиши общегеографический обзор если дисциплина не география.".format(subject=subject)
        )
    elif doc_type == "doklad":
        style_sys += (
            "Это ДОКЛАД — чёткий, ясный, устно-ориентированный стиль. "
            "Короткие предложения, конкретные цифры и факты. "
            "Ссылки на источники [1], [2] в формате ГОСТ."
        )
    elif doc_type == "article":
        style_sys += (
            "Это НАУЧНАЯ СТАТЬЯ — пиши строгим академическим языком, избегай "
            "личных местоимений от первого лица единственного числа ('я'), "
            "используй форму 'мы' или безличные конструкции ('было установлено', 'в ходе исследования'). "
            "Обязательно используй сноски: [1], [2], [3] в каждом абзаце. "
            "Терминология должна быть строгой, выверенной и соответствовать теме."
        )
    else:
        style_sys += (
            "Это АКАДЕМИЧЕСКАЯ РАБОТА. "
            "Обязательно используй ГОСТ-сноски: [1, с. 45], [2], [3, с. 120–125] "
            "в каждом абзаце (1–3 сноски). Номера источников: 1, 2, 3... из списка литературы. "
        )
        if writing_style == "smart":
            style_sys += (
                "Стиль — высокоакадемический: сложные синтаксические конструкции, "
                "специализированная терминология, ссылки на научные концепции и теории."
            )
        else:
            style_sys += (
                "Стиль — чёткий деловой научный: ясные формулировки, "
                "логичная структура, конкретные утверждения."
            )

    for key, prompt in prompts.items():
        step += 1

        # Рассчитываем целевой объём для блока
        if key == "intro":
            block_share = 0.10
        elif key == "conclusion":
            block_share = 0.08
        elif key == "literature":
            block_share = 0.0
        else:
            num_ch = max(1, len([k for k in prompts if k.startswith("ch")]))
            block_share = 0.75 / num_ch

        block_chars = int(total_chars * block_share) if block_share > 0 else 0
        max_tok     = tokens_for_chars(block_chars) if block_chars > 0 else 1500

        if prog:
            block_name = {
                "intro":       "Введение",
                "conclusion":  "Заключение",
                "literature":  "Список литературы",
                "main1":       "Основная часть (аргумент 1)",
                "main2":       "Основная часть (аргумент 2)",
                "part1":       "Раздел 1",
                "part2":       "Раздел 2",
            }.get(key, (
                # Формат ch1_s2 → "Глава 1, подглава 2"
                f"Глава {key.replace('ch', '').replace('_s', ', подглава ')}"
                if key.startswith("ch") else key
            ))
            await prog.update(label=f"✍️ Пишу: {block_name}")

        messages = [
            {"role": "system", "content": style_sys},
            {"role": "user",   "content": prompt},
        ]

        text, used_model = await chat_with_fallback(model_key, messages, max_tok)

        if prog and used_model and used_model != model_key:
            await prog.update(model_name=AI_MODELS.get(used_model, {}).get("name", used_model))

        if not text or len(text) < 80:
            text = _stub_text(key, topic)

        text = _clean_ai_artifacts(text)

        # ── ЗАЩИТА ОТ СЛУЖЕБНЫХ ФРАЗ И КОРОТКИХ ОТВЕТОВ ──
        forbidden_patterns = [
            "вступление к эссе",
            "вот текст",
            "вот второй аргумент",
            "конечно. вот",
            "как языковая модель",
            "объем текста строго выдержан",
            "вот первый аргумент",
        ]

        text_lower = text.lower()
        is_bad = any(p in text_lower[:300] for p in forbidden_patterns)
        is_too_short = len(text.strip()) < 400 and key not in ("literature",)

        if (is_bad or is_too_short) and key not in ("literature",):
            print(f"[WARN] Блок {key} содержит служебную фразу или слишком короткий ({len(text)} знаков). Перегенерация...")
            
            # Усиленный промпт для перегенерации
            retry_prompt = prompt + "\n\nВАЖНО: Напиши полноценный текст без служебных фраз. Не пиши «Вступление», «Конечно», «Вот текст». Начинай сразу с содержания. Минимум 800 знаков."
            retry_messages = [
                {"role": "system", "content": style_sys},
                {"role": "user", "content": retry_prompt},
            ]
            text2, used_model2 = await chat_with_fallback(model_key, retry_messages, tokens_for_chars(block_chars * 2))
            
            if text2 and len(text2.strip()) > len(text.strip()) and len(text2.strip()) > 300:
                text = _clean_ai_artifacts(text2)
                print(f"[WARN] Перегенерация успешна: {len(text)} знаков")
            else:
                print(f"[WARN] Перегенерация не помогла, оставляем как есть")

        # Дозаполняем если текст короче цели на 25%+ (кроме литературы)
        if key not in ("literature",) and block_chars > 0 and len(text) < int(block_chars * 0.75):
            max_refills = 3
            refill_count = 0
            while len(text) < int(block_chars * 0.9) and refill_count < max_refills:
                extra_chars = block_chars - len(text)
                ext_tok     = tokens_for_chars(extra_chars)
                extra_prompt = (
                    f"Продолжи и дополни следующий академический текст по теме «{topic}». "
                    f"Добавь ещё минимум {extra_chars} знаков. "
                    f"Пиши плавно, без заголовков, без [1] [2], без **, без ##:\n\n"
                    f"{text[-600:]}"
                )
                ext_messages = [
                    {"role": "system", "content": style_sys},
                    {"role": "user",   "content": extra_prompt},
                ]
                extra, _ = await chat_with_fallback(model_key, ext_messages, ext_tok)
                if extra and len(extra.strip()) > 50:
                    extra = _clean_ai_artifacts(extra)
                    text = text.rstrip() + "\n\n" + extra.strip()
                else:
                    break
                refill_count += 1

        parts[key] = text

        if prog:
            await prog.update(step_done=True)

    # ═══════════════════════════════════════════════════════════
    # ГЕНЕРАЦИЯ ЗАКЛЮЧЕНИЯ ПОСЛЕ ВСЕХ ГЛАВ
    # (ключевое исправление: заключение видит реальное содержание)
    # ═══════════════════════════════════════════════════════════

    # Собираем контекст из всех содержательных блоков (главы/аргументы/разделы)
    summaries = []
    # Ключи, которые являются содержательными блоками (не intro, не literature)
    content_keys = [k for k in parts.keys()
                    if k not in ("intro", "conclusion", "literature")]
    for key in sorted(content_keys):
        txt = parts[key]
        if not txt or len(txt) < 30:
            continue
        head = txt[:500].strip() if len(txt) > 500 else txt.strip()
        tail = txt[-400:].strip() if len(txt) > 400 else ""
        # Человекочитаемая метка блока
        if key.startswith("ch"):
            label = key.replace("ch", "Глава ").replace("_s", ", подглава ")
        elif key.startswith("main"):
            label = key.replace("main", "Аргумент ")
        elif key.startswith("part"):
            label = key.replace("part", "Раздел ")
        else:
            label = key
        summaries.append(
            f"=== {label} (начало) ===\n{head}\n"
            + (f"=== {label} (конец) ===\n{tail}" if tail else "")
        )

    content_context = "\n\n".join(summaries) if summaries else (
        f"Текст по теме «{topic}» был сгенерирован."
    )

    conc_chars = int(total_chars * 0.10)

    # ── Жанрово-зависимый промпт заключения ──
    if doc_type == "esse":
        conc_prompt = (
            f"Напиши заключение эссе на тему «{topic}».\n\n"
            f"НИЖЕ — содержание твоих аргументов:\n{content_context}\n\n"
            f"Твоя задача:\n"
            f"1. Подведи ЛИЧНЫЙ итог размышления (от первого лица: «я пришёл к выводу»).\n"
            f"2. Опирайся ТОЛЬКО на аргументы выше, не выдумывай новые факты.\n"
            f"3. Выскажи авторскую позицию и возможный прогноз.\n"
            f"4. НЕ упоминай «главы», «разделы» — это эссе, а не реферат.\n"
            f"5. НЕ используй канцеляризмы («практическая значимость», «систематизация»)."
        )
    elif doc_type == "doklad":
        conc_prompt = (
            f"Напиши заключение доклада «{topic}».\n\n"
            f"Содержание разделов:\n{content_context}\n\n"
            f"Задача: подведи итог, сформулируй ключевые выводы и практическое значение.\n"
            f"НЕ упоминай «главы» — это доклад с разделами 1 и 2.\n"
            f"Стиль — устный, чёткий, без излишней сложности."
        )
    elif doc_type == "article":
        conc_prompt = (
            f"Напиши заключение научной статьи «{topic}».\n\n"
            f"НИЖЕ — РЕАЛЬНОЕ СОДЕРЖАНИЕ СЕКЦИЙ:\n{content_context}\n\n"
            f"Твоя задача:\n"
            f"1. Опираясь на содержание выше, сформулируй основные научные выводы исследования.\n"
            f"2. Подчеркни научную новизну и теоретическую/практическую значимость работы.\n"
            f"3. Опиши возможные направления дальнейших исследований в этой научной области.\n"
            f"4. Стиль изложения — строго академический, безличный."
        )
    else:
        conc_prompt = (
            f"Напиши заключение для работы «{topic}».\n\n"
            f"НИЖЕ — РЕАЛЬНОЕ СОДЕРЖАНИЕ ГЛАВ:\n{content_context}\n\n"
            f"Твоя задача:\n"
            f"1. Подведи итог по КАЖДОЙ главе, опираясь ТОЛЬКО на текст выше.\n"
            f"2. Не выдумывай новые факты, которых нет в тексте глав.\n"
            f"3. Сформулируй общие выводы и практическую значимость.\n"
            f"4. Укажи перспективы дальнейшего исследования.\n"
            f"5. Используй сноски [1], [2] на источники.\n"
            f"ЗАПРЕЩЕНО: выдумывать необсуждённые темы."
        )

    conc_messages = [
        {"role": "system", "content": style_sys},
        {"role": "user",   "content": strict_prompt(
            conc_prompt,
            conc_chars,
            writing_style,
            doc_type,
        )},
    ]

    if prog:
        await prog.update(label="✍️ Пишу: Заключение (на основе глав)")

    conc_text, conc_model = await chat_with_fallback(
        model_key, conc_messages,
        tokens_for_chars(conc_chars),
    )

    if prog and conc_model and conc_model != model_key:
        await prog.update(model_name=AI_MODELS.get(conc_model, {}).get("name", conc_model))

    if not conc_text or len(conc_text) < 80:
        conc_text = _stub_text("conclusion", topic)

    conc_text = _clean_ai_artifacts(conc_text)
    parts["conclusion"] = conc_text

    if prog:
        await prog.update(step_done=True)

    # ═══════════════════════════════════════════════════════════
    # СВЯЗЫВАНИЕ ССЫЛОК СО СПИСКОМ ЛИТЕРАТУРЫ (приоритет 🔴)
    # Приводим все внутритекстовые [N, с. X] к диапазону 1..кол-во источников,
    # чтобы каждая ссылка указывала на реально существующий пункт списка.
    # ═══════════════════════════════════════════════════════════
    n_sources = _count_sources(parts.get("literature", ""))
    if n_sources > 0:
        for key in list(parts.keys()):
            if key == "literature":
                continue
            parts[key] = _fix_citations(parts[key], n_sources)

    return parts


def _stub_text(key: str, topic: str) -> str:
    """Заглушка если ИИ не ответил."""
    stubs = {
        "intro":       f"Данная работа посвящена исследованию темы «{topic}». В современных условиях данная проблематика приобретает особую актуальность и практическую значимость для науки и общества.",
        "conclusion":  f"Проведённое исследование по теме «{topic}» позволило сформулировать следующие выводы: изученная проблематика имеет важное теоретическое и практическое значение.",
        "literature":  f"1. Иванов А.А. {topic} / А.А. Иванов. — М.: Наука, 2023. — 256 с.\n2. Петров Б.Б. Основы исследования. — СПб.: Питер, 2022. — 312 с.",
    }
    return stubs.get(key, f"Текст раздела по теме «{topic}» временно недоступен. Повторите генерацию.")


# ═══════════════════════════════════════════════════════════════
#  ГОСТ-DOCX: СТИЛИ, TOC, НУМЕРАЦИЯ, ПОДГЛАВЫ
# ═══════════════════════════════════════════════════════════════

def _normalize_typography(text: str) -> str:
    """Типографика (приоритет 🟡): тире, дефисы, кавычки, лишние #.

    - '---' → '—' (длинное тире), '--' → '–' (короткое тире/диапазон)
    - дефис между словами с пробелами ' - ' → ' — ' (тире)
    - прямые кавычки "..." → «ёлочки»
    - убираем одиночные служебные '#'
    """
    if not text:
        return ""
    # Markdown-заголовки '# ', '## ' в начале строки убираем целиком
    text = re.sub(r"(?m)^\s{0,3}#{1,6}\s+", "", text)
    # Тройной дефис → длинное тире, двойной → короткое тире
    text = text.replace("---", "—").replace("--", "–")
    # Дефис, окружённый пробелами, как знак тире в предложении → длинное тире
    text = re.sub(r"(?<=\s)-(?=\s)", "—", text)
    # Прямые двойные кавычки → «ёлочки» (парами)
    def _quotes(m: re.Match) -> str:
        return "«" + m.group(1) + "»"
    text = re.sub(r'"([^"\n]*)"', _quotes, text)
    # Одиночные оставшиеся '#' (не часть слова) убираем
    text = re.sub(r"(?<!\w)#(?!\w)", "", text)
    return text


def _count_sources(bib_text: str) -> int:
    """Считает количество позиций в списке литературы (после нормализации)."""
    if not bib_text:
        return 0
    norm = _normalize_bibliography(bib_text)
    return len([l for l in norm.split("\n") if re.match(r"^\d+\.\s", l.strip())])


def _fix_citations(text: str, n_sources: int) -> str:
    """Связывает внутритекстовые ссылки со списком литературы (приоритет 🔴).

    Приводит ссылки вида [3], [3, с. 45], [2; 5] к допустимому диапазону
    1..n_sources. Любой номер вне диапазона заменяется на корректный
    (по модулю количества источников), чтобы ссылка указывала на реально
    существующий пункт списка.
    """
    if not text or n_sources <= 0:
        return text

    def _map_num(num: int) -> int:
        if num < 1:
            return 1
        if num > n_sources:
            # Отображаем «по кругу» на существующий источник
            return ((num - 1) % n_sources) + 1
        return num

    def _fix_one(m: re.Match) -> str:
        inner = m.group(1)
        # Заменяем все числовые номера источников в начале ссылки
        def _repl_num(mm: re.Match) -> str:
            return str(_map_num(int(mm.group(0))))
        # Номера источников — это числа, НЕ идущие после 'с.' (страницы не трогаем)
        # Разбиваем по ';' — каждая часть может быть 'N' или 'N, с. P'
        parts = []
        for part in inner.split(";"):
            part = part.strip()
            # первая группа цифр в части — номер источника
            part = re.sub(r"^\s*(\d+)", lambda x: str(_map_num(int(x.group(1)))), part)
            parts.append(part)
        return "[" + "; ".join(parts) + "]"

    # Ссылки ГОСТ: [3], [3, с. 45], [2; 5], [4, с. 120–125]
    return re.sub(r"\[(\d[^\]\n]*)\]", _fix_one, text)


def _normalize_punctuation(text: str) -> str:
    """Чистит технические дефекты текста (ошибка #5):
    убирает лишние пробелы перед знаками препинания и дублирующиеся пробелы.
    """
    if not text:
        return ""
    # Сначала типографика (тире/кавычки/#)
    text = _normalize_typography(text)
    # Убираем пробелы перед . , ; : ! ? ) » и перед закрывающими знаками
    text = re.sub(r"\s+([.,;:!?])", r"\1", text)
    text = re.sub(r"\s+([)»])", r"\1", text)
    text = re.sub(r"([(«])\s+", r"\1", text)
    # Гарантируем пробел после , ; : (если за ним сразу буква/цифра)
    text = re.sub(r"([,;:])([^\s\d)»])", r"\1 \2", text)
    # Схлопываем повторяющиеся пробелы/табы (но не переводы строк)
    text = re.sub(r"[ \t]{2,}", " ", text)
    # Убираем пробелы в конце строк
    text = re.sub(r"[ \t]+\n", "\n", text)
    return text.strip()


def _normalize_bibliography(text: str) -> str:
    """Приводит список литературы к единому формату нумерации (ошибка #9):
    каждая позиция оформляется как '1. Автор...', '2. Автор...' с одним
    пробелом после точки и без пустых строк между пунктами.
    """
    if not text:
        return ""
    raw = text.replace("\r\n", "\n").replace("\r", "\n")
    # Разбиваем по новым строкам и по началам пунктов вида "12. "
    candidates = re.split(r"\n+|(?<=\S)\s+(?=\d{1,3}\s*[.)]\s)", raw)
    items: list[str] = []
    for c in candidates:
        c = c.strip()
        if not c:
            continue
        # Снимаем существующую нумерацию "1. ", "1) ", "1 )", "1 " в начале
        c = re.sub(r"^\d{1,3}\s*[.)]\s*", "", c).strip()
        c = _normalize_punctuation(c)
        if c:
            items.append(c)
    if not items:
        return _normalize_punctuation(text)
    return "\n".join(f"{i}. {it}" for i, it in enumerate(items, start=1))


def _set_run_font(run, font_name: str, size_pt: int, bold: bool = False) -> None:
    run.font.name = font_name
    try:
        run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    except Exception:
        pass
    run.font.size  = Pt(size_pt)
    run.bold       = bool(bold)
    run.font.color.rgb = RGBColor(0, 0, 0)


def setup_gost_page(doc: Document, gost: dict) -> None:
    """Настраивает поля страницы и стиль Normal по ГОСТ."""
    for sec in doc.sections:
        sec.top_margin    = Mm(int(gost.get("top_margin_mm",    20)))
        sec.bottom_margin = Mm(int(gost.get("bottom_margin_mm", 20)))
        sec.left_margin   = Mm(int(gost.get("left_margin_mm",   30)))
        sec.right_margin  = Mm(int(gost.get("right_margin_mm",  15)))

    style      = doc.styles["Normal"]
    font_name  = gost.get("font_name", "Times New Roman")
    style.font.name = font_name
    try:
        style._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    except Exception:
        pass
    style.font.size       = Pt(int(gost.get("font_size", 14)))
    style.font.color.rgb  = RGBColor(0, 0, 0)

    ls = float(gost.get("line_spacing", 1.5))
    pf = style.paragraph_format
    if ls > 1:
        pf.line_spacing      = ls
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    else:
        pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

    pf.first_line_indent = Cm(float(gost.get("first_line_indent_cm", 1.25)))
    pf.space_before      = Pt(0)
    pf.space_after       = Pt(0)
    pf.alignment         = WD_ALIGN_PARAGRAPH.JUSTIFY

    # Настраиваем стили заголовков (ошибка #4):
    # все заголовки 1-го уровня — единый размер; подзаголовки — базовый размер.
    h_size = heading_font_size(gost)
    base   = int(gost.get("font_size", 14))
    _setup_heading_style(doc, "Heading 1", font_name, h_size)
    _setup_heading_style(doc, "Heading 2", font_name, base)


def heading_font_size(gost: dict) -> int:
    """Единый размер шрифта для ВСЕХ заголовков 1-го уровня (ошибка #4).
    Берём базовый размер текста + 2 (например, текст 14 → заголовки 16).
    """
    return int(gost.get("heading_font_size", int(gost.get("font_size", 14)) + 2))


def _setup_heading_style(doc: Document, style_name: str, font_name: str, size_pt: int) -> None:
    """Настраивает стиль заголовка по ГОСТ (полужирный, без отступа, по центру).

    Единообразие (ошибки #2, #4, #8): одинаковая жирность, размер и
    одинаковые интервалы до/после у всех заголовков. space_after = 12pt
    эквивалентно «одной пустой строке» после каждого заголовка.
    """
    try:
        style = doc.styles[style_name]
        style.font.name  = font_name
        style.font.size  = Pt(size_pt)
        style.font.bold  = True
        style.font.italic = False
        style.font.color.rgb = RGBColor(0, 0, 0)
        pf = style.paragraph_format
        pf.first_line_indent = Cm(0)
        pf.space_before      = Pt(12)
        pf.space_after       = Pt(12)   # единый отступ = пустая строка
        pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        pf.alignment         = WD_ALIGN_PARAGRAPH.CENTER
        pf.keep_with_next    = True     # заголовок не отрывается от текста
    except Exception:
        pass


def _add_page_field_to_paragraph(p, font_name: str = "Times New Roman", font_size: int = 12) -> None:
    """Добавляет поле { PAGE } в указанный абзац."""
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    run = p.add_run()
    _set_run_font(run, font_name, font_size, False)
    for fld_type, text in [
        ("begin", None),
        (None,     " PAGE "),
        ("separate", None),
        (None,     "2"),   # placeholder, обновится в Word/LibreOffice
        ("end",    None),
    ]:
        if fld_type:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), fld_type)
            run._r.append(el)
        else:
            instr = OxmlElement("w:instrText")
            instr.set(qn("xml:space"), "preserve")
            instr.text = text
            run._r.append(instr)


def add_page_number_field(section, position: str) -> None:
    """Расставляет нумерацию страниц по ГОСТ.

    ГОСТ 7.32-2017 / 7.0.5: страницы нумеруются арабскими цифрами, сквозная
    нумерация по всему документу, номер на ТИТУЛЬНОМ ЛИСТЕ НЕ СТАВИТСЯ, но он
    включается в общую нумерацию (т.е. на 2-й странице будет цифра 2).

    Реализация: включаем different_first_page_header_footer и заполняем только
    основной колонтитул, оставляя колонтитул первой страницы пустым.
    """
    # Включаем отдельный колонтитул для первой страницы (титула)
    try:
        section.different_first_page_header_footer = True
    except Exception:
        pass

    # Очищаем первый-страничный колонтитул, чтобы на титуле не было цифры
    try:
        first_container = (section.first_page_footer
                           if position == "bottom_center"
                           else section.first_page_header)
        for p in list(first_container.paragraphs):
            for r in list(p.runs):
                r.text = ""
    except Exception:
        pass

    container = section.footer if position == "bottom_center" else section.header

    # Чистим колонтитул
    for p in list(container.paragraphs):
        try:
            p._element.getparent().remove(p._element)
        except Exception:
            pass

    p = container.add_paragraph()
    _add_page_field_to_paragraph(p)


def _toc_entries(blocks: list[tuple]) -> list[tuple[str, int]]:
    """Собирает плоский список пунктов оглавления уровней 1 и 2 (приоритет 🟢).

    Возвращает [(текст_заголовка, уровень), ...].
    Подзаголовки берём как из subblocks, так и из вложенных «1.1 ...» в тексте.
    """
    entries: list[tuple[str, int]] = []
    sub_pat = re.compile(r"^(\d+\.\d+\.?\s+.{3,80})$")
    for title, level, text, subblocks in blocks:
        entries.append((title, 1))
        # Явные подблоки
        for sub_title, _sub_text in (subblocks or []):
            entries.append((sub_title.strip(), 2))
        # Подзаголовки, встроенные в текст главы (когда subblocks пуст)
        if not subblocks and text:
            for line in text.split("\n"):
                line = line.strip()
                if sub_pat.match(line):
                    entries.append((line, 2))
    return entries


def add_toc(doc: Document, blocks: list[tuple], gost: dict) -> None:
    """
    Вставляет содержание с реальными заголовками (уровни 1 и 2).

    1. Поле TOC \\o "1-2" — автоматически обновится в Word/LibreOffice и даст
       КОРРЕКТНЫЕ номера страниц (приоритет 🟡: «зная объём» через layout-движок).
    2. Текстовая копия-структура для немедленной читаемости — БЕЗ номеров
       страниц, чтобы не показывать недостоверные оценки (приоритет 🟡:
       «либо не указывать их в оглавлении»).
    """
    fn = gost.get("font_name", "Times New Roman")
    fs = int(gost.get("font_size", 14))

    # Поле TOC для автообновления (даёт точные номера страниц при открытии в Word/LO)
    p_field = doc.add_paragraph()
    p_field.paragraph_format.first_line_indent = Cm(0)
    run = p_field.add_run()

    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = r'TOC \o "1-2" \h \z \u'

    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")

    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")

    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_sep)
    run._r.append(fld_end)

    # Текстовая структура (уровни 1 и 2) — без номеров страниц
    for title, level in _toc_entries(blocks):
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        p.paragraph_format.left_indent  = Cm(0) if level == 1 else Cm(1.0)

        run_t = p.add_run(title)
        _set_run_font(run_t, fn, fs if level == 1 else fs - 1, level == 1)


def add_title_page(doc: Document, data: dict, gost: dict) -> None:
    """Создаёт титульный лист по ГОСТ."""
    font = gost.get("font_name", "Times New Roman")
    size = int(gost.get("font_size", 14))

    def _add_centered(text: str, sz: int = None, bold: bool = False) -> None:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)
        r = p.add_run(text)
        _set_run_font(r, font, sz or size, bold)

    def _add_right(text: str, bold: bool = False) -> None:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.left_indent = Cm(8.5)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)
        r = p.add_run(text)
        _set_run_font(r, font, size, bold)

    def _spacer(n: int = 1) -> None:
        for _ in range(n):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after  = Pt(0)

    doc_word = DOC_TYPES.get(data.get("doc_type", "referat"), DOC_TYPES["referat"])["word"]
    if data.get("doc_type") == "custom" and data.get("custom_doc_name"):
        doc_word = str(data["custom_doc_name"]).strip().upper()

    _add_centered("МИНИСТЕРСТВО ОБРАЗОВАНИЯ И НАУКИ РОССИЙСКОЙ ФЕДЕРАЦИИ", 11, True)

    if data.get("org_type"):
        _add_centered(str(data["org_type"]).upper(), 11, False)

    inst = data.get("institution") or "Учебное заведение"
    _add_centered(f"«{inst}»", 12, True)

    _spacer(4)

    _add_centered(doc_word, 18, True)
    _add_centered(f"по дисциплине «{data.get('subject', '')}»", 14, False)
    _spacer(1)
    _add_centered("на тему:", 14, False)
    _add_centered(f"«{data.get('topic', '')}»", 16, True)

    _spacer(4)

    _add_right("Выполнил(а):", False)
    gr = data.get("group", "")
    if gr:
        _add_right(f"{gr}", False)
    _add_right(str(data.get("author", "")), True)
    _spacer(1)
    _add_right("Проверил(а):", False)
    _add_right(str(data.get("teacher", "")), True)

    _spacer(4)

    _add_centered(f"{data.get('city', 'Москва')}  {datetime.now().year}", 14, False)


def add_paragraphs_from_text(
    doc: Document,
    text: str,
    gost: dict,
    is_bib: bool = False,
) -> None:
    """
    Разбивает текст на абзацы и добавляет их в документ.
    Строки вида '1.1. Название подглавы' оформляются как Heading 2.
    """
    font   = gost.get("font_name", "Times New Roman")
    size   = int(gost.get("font_size", 14))
    indent = Cm(float(gost.get("first_line_indent_cm", 1.25)))  # ошибка #7

    # Список литературы: единый формат нумерации (ошибка #9)
    if is_bib:
        text = _normalize_bibliography(text)
    else:
        text = _normalize_punctuation(text)  # ошибка #5

    def _apply_body_format(p) -> None:
        """Единое оформление абзаца основного текста (ошибки #6, #7)."""
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = indent
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)

    # Паттерн для подглав: "1.1. Текст" или "2.3. Текст"
    subheading_pat = re.compile(r"^(\d+\.\d+\.?\s+.{5,80})$")

    if is_bib:
        # Каждая позиция списка — отдельный абзац с висячим отступом
        for line in [l for l in text.split("\n") if l.strip()]:
            p = doc.add_paragraph()
            p.paragraph_format.alignment         = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.left_indent       = Cm(1.25)
            p.paragraph_format.first_line_indent = Cm(-1.25)
            p.paragraph_format.space_before      = Pt(0)
            p.paragraph_format.space_after       = Pt(0)
            r = p.add_run(line.strip())
            _set_run_font(r, font, size, False)
        return

    chunks = [c.strip() for c in re.split(r"\n\s*\n|\n(?=\d+\.\d+\.)", text) if c.strip()]

    for ch in chunks:
        first_line = ch.split("\n")[0].strip()

        if subheading_pat.match(first_line):
            # Это подзаголовок — Heading 2
            hp = doc.add_paragraph(first_line, style="Heading 2")
            for run in hp.runs:
                _set_run_font(run, font, size, True)
            hp.paragraph_format.first_line_indent = Cm(0)
            # Остаток как обычный текст
            rest = ch[len(first_line):].strip()
            if rest:
                p = doc.add_paragraph()
                _apply_body_format(p)
                clean_rest = re.sub(r'\s*\n\s*', ' ', rest)
                r = p.add_run(clean_rest)
                _set_run_font(r, font, size, False)
        else:
            p = doc.add_paragraph()
            _apply_body_format(p)
            clean_ch = re.sub(r'\s*\n\s*', ' ', ch)
            r = p.add_run(clean_ch)
            _set_run_font(r, font, size, False)


def build_docx_bytes(
    data: dict,
    blocks: list[tuple],
    gost: dict,
) -> bytes:
    """Собирает DOCX из блоков."""
    doc = Document()
    setup_gost_page(doc, gost)

    # ── Титульный лист ──
    add_title_page(doc, data, gost)
    doc.add_page_break()

    # ── Содержание ──
    fn   = gost.get("font_name", "Times New Roman")
    fs   = int(gost.get("font_size", 14))
    hfs  = heading_font_size(gost)   # единый размер заголовков 1-го ур. (ошибка #4)

    # «СОДЕРЖАНИЕ» оформляем как заголовок 1-го уровня по виду (ошибка #4),
    # но НЕ через стиль Heading 1 — иначе попадёт в само поле TOC.
    p_toc_title = doc.add_paragraph()
    p_toc_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_toc_title.paragraph_format.first_line_indent = Cm(0)
    p_toc_title.paragraph_format.space_before = Pt(12)
    p_toc_title.paragraph_format.space_after  = Pt(12)
    r = p_toc_title.add_run("СОДЕРЖАНИЕ")
    _set_run_font(r, fn, hfs, True)

    add_toc(doc, blocks, gost)

    doc.add_page_break()

    # ── Нумерация страниц ──
    add_page_number_field(
        doc.sections[0],
        gost.get("page_number_position", "bottom_center"),
    )

    # ── Тело документа ──
    last_idx = len(blocks) - 1
    for idx, (title, level, text, subblocks) in enumerate(blocks):
        # Заголовок главы. Все заголовки 1-го уровня — единый размер (ошибка #4),
        # единые интервалы задаёт стиль Heading (ошибки #2, #8).
        style = "Heading 1" if level == 1 else "Heading 2"
        h_sz  = hfs if level == 1 else fs
        hp    = doc.add_paragraph(title, style=style)
        for run in hp.runs:
            _set_run_font(run, fn, h_sz, True)
        hp.paragraph_format.first_line_indent = Cm(0)

        # Основной текст главы
        is_bib = any(w in title.upper() for w in ("ИСТОЧНИК", "ЛИТЕРАТ", "БИБЛИОГРАФ"))

        if subblocks:
            # Есть подблоки — выводим их с заголовком Heading 2 и текстом
            for sub_title, sub_text in subblocks:
                shp = doc.add_paragraph(sub_title, style="Heading 2")
                for run in shp.runs:
                    _set_run_font(run, fn, fs, True)
                shp.paragraph_format.first_line_indent = Cm(0)
                if sub_text:
                    # СТРИПАЕМ продублированный заголовок из начала текста,
                    # чтобы не было «1.1. Название» дважды подряд
                    clean_text = sub_text
                    # Паттерн: "1.1. Остальной заголовок" в начале текста
                    first_line = sub_text.split('\n')[0].strip()
                    # Нормализуем для сравнения (убираем лишние пробелы)
                    norm_sub_title = re.sub(r'\s+', ' ', sub_title.strip())
                    norm_first = re.sub(r'\s+', ' ', first_line)
                    # Если первая строка совпадает с sub_title (полностью или по номеру)
                    if (norm_first == norm_sub_title or
                        (re.match(r'^\d+\.\d+\.?\s', first_line) and
                         first_line[:20] == sub_title.strip()[:20])):
                        # Убираем первую строку (продублированный заголовок)
                        rest = sub_text[len(first_line):].lstrip('\n').lstrip('\r')
                        clean_text = rest if rest else sub_text
                    add_paragraphs_from_text(doc, clean_text, gost)
        elif text:
            # Нет подблоков — выводим основной текст
            clean_text = text
            # Стрипаем продублированный заголовок из начала текста
            first_line = text.split('\n')[0].strip()
            norm_title = re.sub(r'\s+', ' ', title.strip().upper())
            norm_first = re.sub(r'\s+', ' ', first_line.upper())
            if (norm_first == norm_title or 
                (norm_first.startswith("ГЛАВА") and norm_title.startswith("ГЛАВА") and 
                 norm_first[:15] == norm_title[:15])):
                rest = text[len(first_line):].lstrip('\n').lstrip('\r')
                clean_text = rest if rest else text
            add_paragraphs_from_text(doc, clean_text, gost, is_bib=is_bib)

        # Разрыв страницы между главами/разделами (кроме последнего)
        if idx < last_idx:
            doc.add_page_break()

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def libreoffice_update_docx(in_path: str, out_path: str) -> bool:
    """Прогоняет через LibreOffice чтобы обновить TOC и поля страниц."""
    soffice = shutil.which("soffice")
    if not soffice:
        return False

    out_dir = os.path.dirname(out_path)
    os.makedirs(out_dir, exist_ok=True)

    try:
        p = subprocess.run(
            [
                soffice,
                "--headless",
                "--nologo",
                "--nolockcheck",
                "--nodefault",
                "--nofirststartwizard",
                "--convert-to", "docx",
                "--outdir",     out_dir,
                in_path,
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            timeout=120,
        )
        if p.returncode != 0:
            print(f"[LO] Ошибка конвертации: {p.stdout[:500]}")
            return False

        produced = os.path.join(
            out_dir,
            os.path.splitext(os.path.basename(in_path))[0] + ".docx",
        )
        if not os.path.exists(produced):
            return False

        if produced != out_path:
            os.replace(produced, out_path)

        return True

    except subprocess.TimeoutExpired:
        print("[LO] Таймаут конвертации")
        return False
    except Exception as e:
        print(f"[LO] Ошибка: {e}")
        return False


# ═══════════════════════════════════════════════════════════════
#  ТОЧНЫЙ ПОДСЧЁТ СТРАНИЦ DOCX ЧЕРЕЗ LIBREOFFICE → PDF
# ═══════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════
#  УЛУЧШЕННЫЙ СЕРВИС ПОДСЧЁТА СТРАНИЦ ДЛЯ ГОСТ-АССИСТЕНТА
# ═══════════════════════════════════════════════════════════════

try:
    from pypdf import PdfReader
except ImportError:
    try:
        from PyPDF2 import PdfReader
    except ImportError:
        PdfReader = None

try:
    import fitz  # PyMuPDF — лучший парсер PDF
except ImportError:
    fitz = None


# Символов на страницу (ГОСТ: ~1800 знаков с пробелами на стр A4 14pt 1.5 интервал)
CHARS_PER_PAGE = int(cfg("CHARS_PER_PAGE", "1800"))  # ИСПРАВЛЕНО: 1850 → 1800

# Страниц без текста (титул + содержание)
NON_TEXT_PAGES = int(cfg("NON_TEXT_PAGES", "2"))

# Калибровочный множитель (ИСПРАВЛЕНО: 1.10 → 1.00)
_ESTIM_CALIBRATION = 1.00


def target_chars(pages: int) -> int:
    """
    Целевое количество символов ДЛЯ ОСНОВНОГО ТЕКСТА (без титула и содержания)
    """
    text_pages = max(1, pages - NON_TEXT_PAGES)
    return text_pages * CHARS_PER_PAGE


def target_pages_from_chars(chars: int) -> int:
    """Возвращает количество страниц по количеству символов"""
    text_pages = max(1, chars // CHARS_PER_PAGE)
    return text_pages + NON_TEXT_PAGES


class ImprovedPageCounter:
    pass

class PageAdapter:
    pass

# ═══════════════════════════════════════════════════════════════
#  ПОДСЧЁТ СТРАНИЦ ЧЕРЕЗ LIBREOFFICE → PDF
# ═══════════════════════════════════════════════════════════════

def libreoffice_docx_to_pdf(in_path: str, out_dir: str) -> Optional[str]:
    """Конвертирует DOCX в PDF через LibreOffice. Возвращает путь к PDF или None."""
    soffice = shutil.which("soffice")
    if not soffice:
        return None
    os.makedirs(out_dir, exist_ok=True)
    try:
        p = subprocess.run(
            [
                soffice, "--headless", "--nologo", "--nolockcheck",
                "--nodefault", "--nofirststartwizard",
                "--convert-to", "pdf", "--outdir", out_dir, in_path,
            ],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=180,
        )
        if p.returncode != 0:
            print(f"[LO→PDF] Ошибка: {p.stdout[:500]}")
            return None
        pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(in_path))[0] + ".pdf")
        return pdf if os.path.exists(pdf) else None
    except Exception as e:
        print(f"[LO→PDF] Исключение: {e}")
        return None


def count_pdf_pages(pdf_path: str) -> Optional[int]:
    """Считает страницы PDF. Приоритет: PyMuPDF → PyPDF2 → pdfinfo"""
    # 1) PyMuPDF (лучший, если установлен)
    try:
        import fitz
        doc = fitz.open(pdf_path)
        pages = len(doc)
        doc.close()
        return pages
    except ImportError:
        pass
    except Exception as e:
        print(f"[PyMuPDF] Ошибка: {e}")
    
    # 2) PyPDF2 / pypdf
    try:
        from pypdf import PdfReader
        with open(pdf_path, "rb") as f:
            return len(PdfReader(f).pages)
    except ImportError:
        pass
    except Exception:
        pass
    
    try:
        from PyPDF2 import PdfReader
        with open(pdf_path, "rb") as f:
            return len(PdfReader(f).pages)
    except ImportError:
        pass
    except Exception:
        pass
    
    # 3) pdfinfo (poppler-utils)
    pdfinfo = shutil.which("pdfinfo")
    if pdfinfo:
        try:
            p = subprocess.run([pdfinfo, pdf_path],
                               stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                               text=True, timeout=10)
            for line in p.stdout.splitlines():
                if line.lower().startswith("pages:"):
                    return int(line.split(":", 1)[1].strip())
        except Exception:
            pass
    
    # 4) Грубый подсчёт /Type /Page в сыром PDF
    try:
        with open(pdf_path, "rb") as f:
            blob = f.read()
        return max(1, blob.count(b"/Type /Page") - blob.count(b"/Type /Pages"))
    except Exception:
        return None


def count_docx_pages(docx_path: str, work_dir: str) -> Optional[int]:
    """Возвращает реальное число страниц DOCX (через LibreOffice→PDF)."""
    pdf = libreoffice_docx_to_pdf(docx_path, work_dir)
    if not pdf:
        return None
    try:
        return count_pdf_pages(pdf)
    finally:
        try:
            os.remove(pdf)
        except Exception:
            pass


# ═══════════════════════════════════════════════════════════════
#  РАСЧЁТНЫЙ ЭСТИМАТОР СТРАНИЦ (без внешних зависимостей)
# ═══════════════════════════════════════════════════════════════


def estimate_docx_pages(docx_path: str) -> Optional[int]:
    """
    Эмулирует разбивку DOCX по страницам без внешних инструментов.
    Учитывает: размер страницы A4, поля, шрифт, межстрочный интервал,
    отступ красной строки, заголовки, ПРИНУДИТЕЛЬНЫЕ РАЗРЫВЫ СТРАНИЦ.
    """
    try:
        from docx import Document as _D
        doc = _D(docx_path)
    except Exception as e:
        print(f"[ESTIM] Ошибка чтения DOCX: {e}")
        return None

    if not doc.sections:
        return None
    sec = doc.sections[0]

    def _emu_to_pt(emu) -> float:
        try:
            return float(emu) / 12700.0
        except:
            return 0.0

    page_w_pt = _emu_to_pt(sec.page_width) or 595.0
    page_h_pt = _emu_to_pt(sec.page_height) or 842.0
    left_pt = _emu_to_pt(sec.left_margin) or 85.0
    right_pt = _emu_to_pt(sec.right_margin) or 42.0
    top_pt = _emu_to_pt(sec.top_margin) or 56.0
    bottom_pt = _emu_to_pt(sec.bottom_margin) or 56.0

    text_w_pt = max(50.0, page_w_pt - left_pt - right_pt)
    text_h_pt = max(50.0, page_h_pt - top_pt - bottom_pt)

    try:
        normal = doc.styles["Normal"]
        base_size = float(normal.font.size.pt) if normal.font.size else 14.0
    except Exception:
        base_size = 14.0
    
    try:
        line_spacing = float(normal.paragraph_format.line_spacing or 1.5)
    except Exception:
        line_spacing = 1.5
    
    if line_spacing < 0.5:
        line_spacing = 1.5

    # Средняя ширина символа Times New Roman (в пунктах)
    char_width_pt = base_size * 0.42
    chars_per_line = max(20, int(text_w_pt / char_width_pt))
    
    line_height_pt = base_size * line_spacing
    lines_per_page = max(10, int(text_h_pt / line_height_pt))

    total_chars = 0
    page_breaks = 0

    for paragraph in doc.paragraphs:
        text = paragraph.text or ""
        total_chars += len(text)
        
        # Проверяем на page break
        from docx.oxml.ns import qn
        for run in paragraph.runs:
            for br in run._element.iter(qn("w:br")):
                if br.get(qn("w:type")) == "page":
                    page_breaks += 1
                    break

    # Расчёт страниц по символам
    estimated_by_chars = max(1, int(total_chars / CHARS_PER_PAGE) + NON_TEXT_PAGES)
    
    # Расчёт по строкам
    estimated_lines = max(1, int(total_chars / max(1, chars_per_line)))
    estimated_by_lines = max(1, (estimated_lines // max(1, lines_per_page)) + 1 + NON_TEXT_PAGES)
    
    # Комбинируем
    estimated = max(estimated_by_chars, estimated_by_lines, page_breaks + NON_TEXT_PAGES)
    
    # Применяем калибровку
    calibrated = int(estimated * _ESTIM_CALIBRATION)
    
    return max(1, calibrated)


# ═══════════════════════════════════════════════════════════════
#  АСИНХРОННЫЙ ПОДСЧЁТ СТРАНИЦ (главная функция)
# ═══════════════════════════════════════════════════════════════


async def count_pages_via_aspose(docx_path: str) -> Optional[int]:
    """Считает страницы DOCX через Aspose Words Counter API."""
    ASPOSE_ENDPOINT = (
        "https://api.products.aspose.app/words/wordscounter/api/getstatistics"
    )
    try:
        with open(docx_path, "rb") as fh:
            file_bytes = fh.read()

        form = aiohttp.FormData()
        form.add_field(
            "files[]",
            file_bytes,
            filename=os.path.basename(docx_path),
            content_type=(
                "application/vnd.openxmlformats-officedocument"
                ".wordprocessingml.document"
            ),
        )
        form.add_field("includeTextboxes", "false")

        async with aiohttp.ClientSession() as sess:
            async with sess.post(
                ASPOSE_ENDPOINT,
                data=form,
                timeout=aiohttp.ClientTimeout(total=15),
            ) as resp:
                if resp.status != 200:
                    print(f"[ASPOSE] HTTP {resp.status}")
                    return None
                data = await resp.json(content_type=None)
                pages = data.get("statistics", {}).get("pages")
                if pages is not None:
                    pages = int(pages)
                    print(f"[ASPOSE] Страниц: {pages}")
                    return pages
    except asyncio.TimeoutError:
        print("[ASPOSE] Таймаут запроса")
    except Exception as e:
        print(f"[ASPOSE] Ошибка: {e}")
    return None


async def measure_pages_async(docx_path: str, work_dir: str) -> Optional[int]:
    """
    Главная функция подсчёта страниц (async).
    
    Порядок приоритетов:
    1. Aspose Words Counter API — точный, не требует LibreOffice
    2. LibreOffice → PDF + PyMuPDF/PyPDF2 — точный, требует soffice
    3. estimate_docx_pages — расчётный, без внешних зависимостей
    """
    # 1) Aspose
    n = await count_pages_via_aspose(docx_path)
    if n is not None:
        return n

    # 2) LibreOffice
    n = count_docx_pages(docx_path, work_dir)
    if n is not None:
        print(f"[LO] Страниц: {n}")
        return n

    # 3) Расчётный эстиматор
    n = estimate_docx_pages(docx_path)
    if n is not None:
        print(f"[ESTIM] Страниц (расчётно): {n}")
    return n


# ═══════════════════════════════════════════════════════════════
#  ФУНКЦИИ ДЛЯ ПОДГОНКИ ОБЪЁМА ТЕКСТА
# ═══════════════════════════════════════════════════════════════


def _blocks_text_total(blocks: list[tuple]) -> int:
    """Подсчитывает общее количество символов во всех блоках"""
    total = 0
    for _t, _l, text, subs in blocks:
        if text:
            total += len(text)
        for _st, stext in (subs or []):
            if stext:
                total += len(stext)
    return total


def _trim_blocks_by_chars(blocks: list[tuple], chars_to_remove: int) -> list[tuple]:
    """Аккуратно укорачивает блоки, сохраняя последний абзац целым."""
    if chars_to_remove <= 0:
        return blocks
    
    blocks = [list(b) for b in blocks]

    def _can_trim(title: str) -> bool:
        up = (title or "").upper()
        return not any(w in up for w in ("ЛИТЕРАТ", "ИСТОЧНИК", "БИБЛИОГРАФ", "ЗАКЛЮЧЕНИ"))

    def _trim_text(txt: str, need: int) -> tuple[str, int]:
        """Откусывает абзацы с конца. Возвращает (новый_текст, сколько_ушло)."""
        paras = [p for p in txt.split("\n\n") if p.strip()]
        removed_total = 0
        while paras and need > 0 and len(paras) > 1:
            removed = len(paras[-1]) + 2
            paras.pop()
            need -= removed
            removed_total += removed
        return "\n\n".join(paras), removed_total

    # Сортируем кандидатов по размеру (самые большие первые)
    def _block_total(i: int) -> int:
        b = blocks[i]
        total = len(b[2] or "")
        for _, st in (b[3] or []):
            total += len(st or "")
        return total

    candidates = sorted(
        [i for i, b in enumerate(blocks) if _can_trim(b[0])],
        key=_block_total, reverse=True,
    )

    for i in candidates:
        if chars_to_remove <= 0:
            break

        b = blocks[i]
        subblocks = b[3]

        if subblocks:
            # Обрезаем с последней подглавы
            for si in range(len(subblocks) - 1, -1, -1):
                if chars_to_remove <= 0:
                    break
                stitle, stext = subblocks[si]
                if not stext or len(stext) < 200:
                    continue
                new_stext, removed = _trim_text(stext, chars_to_remove)
                chars_to_remove -= removed
                subblocks[si] = (stitle, new_stext)
            # Обновляем агрегированный текст
            b[2] = "\n\n".join(st for _, st in subblocks if st)
        else:
            txt = b[2] or ""
            if len(txt) >= 400:
                new_txt, removed = _trim_text(txt, chars_to_remove)
                chars_to_remove -= removed
                b[2] = new_txt

    return [tuple(b) for b in blocks]


async def _expand_blocks_by_chars(
    blocks: list[tuple],
    chars_to_add: int,
    topic: str,
    model_key: str,
    writing_style: str,
    prog: Optional["Progress"] = None,
) -> list[tuple]:
    """Дозаписывает текст в основные главы, чтобы добрать chars_to_add символов."""
    if chars_to_add <= 0:
        return blocks
    
    blocks = [list(b) for b in blocks]

    def _is_expandable(title: str) -> bool:
        up = (title or "").upper()
        return not any(w in up for w in ("ЛИТЕРАТ", "ИСТОЧНИК", "БИБЛИОГРАФ"))

    candidates = [i for i, b in enumerate(blocks) if _is_expandable(b[0])]
    if not candidates:
        candidates = [i for i, b in enumerate(blocks) if "ЛИТЕРАТ" not in (b[0] or "").upper()]
    
    if not candidates:
        return [tuple(b) for b in blocks]

    # Равномерно распределяем нагрузку + 30% запас
    per_block = int((chars_to_add / max(1, len(candidates))) * 1.3) + 300
    
    style_label = "высокоакадемический научный" if writing_style == "smart" else "деловой научный"
    style_sys = (
        f"Ты профессионально пишешь академический текст на русском языке в стиле «{style_label}». "
        "СТРОГО: без markdown (никаких **, ##, ---, ```), "
        "без заголовков, без буллетов и нумерованных списков. "
        "Используй ГОСТ-сноски [1, с. 45], [2], [3] в тексте. "
        "Только сплошной академический текст, абзацами по 5–8 развёрнутых предложений."
    )

    for i in candidates:
        title = blocks[i][0]
        text = blocks[i][2] or ""
        orig_field2 = text
        need = per_block
        
        if prog:
            await prog.update(label=f"➕ Дописываю «{title[:30]}…»", force=True)

        attempts = 0
        max_attempts = 3
        while need > 100 and attempts < max_attempts:
            attempts += 1
            tail = text[-800:] if text else ""
            user_prompt = (
                f"Тема всей работы: «{topic}».\n"
                f"Текущий раздел: «{title}».\n\n"
                f"ЗАДАЧА: продолжи и РАЗВЕРНИ этот раздел, "
                f"добавь МИНИМУМ {need} знаков с пробелами.\n\n"
                f"⚠️ ВНИМАНИЕ: Пользователь просил написать строго заданный объем страниц. "
                f"Ты должен четко понимать лимит страниц и НЕ ОБРЫВАТЬ мысль на полуслове! "
                f"Заканчивай свои мысли логично, последовательно, но при этом кратко и лаконично. "
                f"Каждое предложение должно быть полностью завершено.\n\n"
                f"Правила:\n"
                f"— Не повторяй то, что уже сказано.\n"
                f"— Развивай мысль: новые аргументы, примеры.\n"
                f"— Минимум 3 полноценных абзаца.\n"
                f"— Никаких заголовков, списков.\n\n"
                f"Хвост уже написанного раздела:\n...{tail}"
            )
            
            try:
                extra, _used = await chat_with_fallback(
                    model_key,
                    [{"role": "system", "content": style_sys},
                     {"role": "user",   "content": user_prompt}],
                    tokens_for_chars(int(need * 1.5)),
                )
            except Exception as e:
                print(f"[EXPAND] Ошибка ИИ: {e}")
                break

            if not extra or len(extra.strip()) < 100:
                continue

            extra = sanitize_llm_text(extra.strip())
            text = (text.rstrip() + "\n\n" + extra) if text else extra
            need -= len(extra)

        # Обновляем блок
        if blocks[i][3]:
            last_idx = len(blocks[i][3]) - 1
            stitle, stext = blocks[i][3][last_idx]
            orig_text = orig_field2 or ""
            new_part = text[len(orig_text):].strip() if text.startswith(orig_text) else text
            if new_part:
                blocks[i][3][last_idx] = (stitle, (stext + "\n\n" + new_part) if stext else new_part)
            blocks[i][2] = "\n\n".join(st for _, st in blocks[i][3] if st)
        else:
            blocks[i][2] = text

    return [tuple(b) for b in blocks]


# ═══════════════════════════════════════════════════════════════
#  ФУНКЦИЯ ДЛЯ ТОЧНОЙ ПОДГОНКИ СТРАНИЦ (НОВАЯ)
# ═══════════════════════════════════════════════════════════════


async def precise_page_adjustment(
    tmp_in: str,
    blocks: list[tuple],
    target_pages: int,
    topic: str,
    model_key: str,
    writing_style: str,
    data: dict,
    gost: dict,
    prog: Optional["Progress"] = None,
    work_dir: str = None,
) -> tuple[list[tuple], int]:
    """
    Точная подгонка количества страниц.
    Возвращает (блоки, финальное_количество_страниц)
    """
    if work_dir is None:
        work_dir = os.path.dirname(tmp_in)
    
    measure_dir = os.path.join(work_dir, "_measure")
    os.makedirs(measure_dir, exist_ok=True)
    
    max_iters = 10
    
    for it in range(max_iters):
        # Измеряем текущее количество страниц
        real_pages = await measure_pages_async(tmp_in, measure_dir)
        
        if real_pages is None:
            total_chars = _blocks_text_total(blocks)
            real_pages = target_pages_from_chars(total_chars)
        
        diff = real_pages - target_pages
        print(f"[ADJUST] Итерация {it+1}: цель={target_pages}, факт={real_pages}, разница={diff:+d}")
        
        if prog:
            await prog.update(
                label=f"📏 Подгонка: {real_pages}/{target_pages} стр (попытка {it+1})",
                force=True,
            )
        
        # Достигли цели?
        if diff == 0:
            print(f"[ADJUST] ✅ Цель достигнута!")
            break
        elif it >= 7 and abs(diff) <= 1:
            print(f"[ADJUST] ✅ Приемлемая цель достигнута на поздней итерации: {real_pages} стр.")
            break
        
        # Рассчитываем сколько символов нужно добавить/убрать
        chars_per_real_page = CHARS_PER_PAGE
        
        if diff > 0:
            # Слишком много страниц — обрезаем
            chars_to_remove = int(diff * chars_per_real_page)
            print(f"[ADJUST] ✂️ Обрезаю {chars_to_remove} знаков")
            blocks = _trim_blocks_by_chars(blocks, chars_to_remove)
        else:
            # Слишком мало страниц — добавляем
            chars_to_add = int(abs(diff) * chars_per_real_page)
            print(f"[ADJUST] ➕ Добавляю {chars_to_add} знаков")
            blocks = await _expand_blocks_by_chars(
                blocks, chars_to_add, topic, model_key, writing_style, prog,
            )
        
        # Пересобираем DOCX
        docx_raw = build_docx_bytes(data, blocks, gost)
        with open(tmp_in, "wb") as f:
            f.write(docx_raw)
    
    # Финальный замер
    final_pages = await measure_pages_async(tmp_in, measure_dir)
    if final_pages is None:
        final_pages = target_pages
    
    print(f"[ADJUST] 🎯 Финальный результат: {final_pages} страниц")
    return blocks, final_pages
# ═══════════════════════════════════════════════════════════════
#  «СВОЙ ГОСТ» — ПАРСИНГ ТРЕБОВАНИЙ ЧЕРЕЗ ИИ
# ═══════════════════════════════════════════════════════════════

async def parse_custom_gost_via_ai(model_key: str, gost_text: str) -> dict:
    schema_hint = {
        "font_name":            "Times New Roman",
        "font_size":            14,
        "line_spacing":         1.5,
        "first_line_indent_cm": 1.25,
        "left_margin_mm":       30,
        "right_margin_mm":      15,
        "top_margin_mm":        20,
        "bottom_margin_mm":     20,
        "page_number_position": "bottom_center",
    }

    system = (
        "Ты извлекаешь параметры оформления из описания требований пользователя. "
        "Ответ СТРОГО в JSON без объяснений и без markdown. "
        "Если параметр не указан — используй значение по умолчанию из schema."
    )

    user = (
        "schema (defaults):\n"
        + json.dumps(schema_hint, ensure_ascii=False)
        + "\n\nТребования пользователя:\n"
        + gost_text
        + "\n\nВерни ТОЛЬКО JSON."
    )

    raw, _ = await chat_with_fallback(
        model_key,
        [{"role": "system", "content": system}, {"role": "user", "content": user}],
        max_tokens=800,
    )
    raw = raw.strip()

    m = re.search(r"\{.*\}", raw, flags=re.S)
    if m:
        raw = m.group(0)

    try:
        data = json.loads(raw)
    except Exception:
        return schema_hint

    out = dict(schema_hint)
    for k in out:
        if k in data and data[k] not in (None, ""):
            out[k] = data[k]

    # Приведение типов
    try:
        out["font_size"] = int(out["font_size"])
    except Exception:
        out["font_size"] = 14

    for fkey in ("line_spacing", "first_line_indent_cm"):
        try:
            out[fkey] = float(out[fkey])
        except Exception:
            out[fkey] = schema_hint[fkey]

    for ikey in ("left_margin_mm", "right_margin_mm", "top_margin_mm", "bottom_margin_mm"):
        try:
            out[ikey] = int(out[ikey])
        except Exception:
            out[ikey] = schema_hint[ikey]

    if out.get("page_number_position") not in ("bottom_center", "top_center"):
        out["page_number_position"] = "bottom_center"

    return out


# ═══════════════════════════════════════════════════════════════
#  КЛАВИАТУРЫ — КРАСИВЫЕ
# ═══════════════════════════════════════════════════════════════

def kb_doc_type() -> InlineKeyboardMarkup:
    b = InlineKeyboardBuilder()
    for k, v in DOC_TYPES.items():
        label = f"{v['name']}  {v['min_pages']}–{v['max_pages']} стр."
        b.button(text=label, callback_data=f"dtype_{k}")
    b.adjust(1)
    return b.as_markup()


def kb_mode() -> InlineKeyboardMarkup:
    b = InlineKeyboardBuilder()
    b.button(text="🆓 Бесплатно  (быстро)", callback_data="mode_free")
    b.button(text="⭐ Платный режим  (безлимит)", callback_data="mode_paid")
    b.adjust(1)
    return b.as_markup()


def kb_writing_style() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(
                text="🎓 Умно — сложный академический язык, термины, глубокий анализ",
                callback_data="style_smart",
            )],
            [InlineKeyboardButton(
                text="📝 Классически — чёткий деловой стиль, понятные формулировки",
                callback_data="style_classic",
            )],
        ]
    )


def kb_source_choice() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="✅ Да — добавлю свои материалы", callback_data="source_yes")],
            [InlineKeyboardButton(text="❌ Нет — только по теме",         callback_data="source_no")],
        ]
    )


def kb_institution() -> InlineKeyboardMarkup:
    b = InlineKeyboardBuilder()
    for k, v in INSTITUTION_TYPES.items():
        b.button(text=v["name"], callback_data=f"inst_{k}")
    b.adjust(2)
    return b.as_markup()


def kb_subject() -> InlineKeyboardMarkup:
    b = InlineKeyboardBuilder()
    for s in SUBJECTS:
        b.button(text=s, callback_data=f"subj_{s}")
    b.button(text="✏️ Другой предмет", callback_data="subj_other")
    b.adjust(3)
    return b.as_markup()


def kb_city() -> InlineKeyboardMarkup:
    b = InlineKeyboardBuilder()
    for c in CITIES:
        b.button(text=c, callback_data=f"city_{c}")
    b.button(text="✏️ Другой город", callback_data="city_other")
    b.adjust(3)
    return b.as_markup()


def kb_page_number() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="⬇️ Внизу по центру (стандарт)", callback_data="pagepos_bottom")],
            [InlineKeyboardButton(text="⬆️ Вверху по центру",           callback_data="pagepos_top")],
        ]
    )


def kb_final() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="🔁 Создать новую работу", callback_data="final_new")],
            [InlineKeyboardButton(text="⚙️ Настроить ГОСТ",       callback_data="custom_gost")],
            [InlineKeyboardButton(text="📊 Мой баланс генераций", callback_data="show_limits")],
        ]
    )


def kb_models() -> InlineKeyboardMarkup:
    b = InlineKeyboardBuilder()
    for k, info in AI_MODELS.items():
        if not info.get("api_key") or info.get("_fatal"):
            continue
        status = info.get("status", ModelStatus.UNKNOWN)
        b.button(
            text=f"{info['name']}  {info['price_per_page']}⭐/стр  {status}",
            callback_data=f"model_{k}",
        )
    b.adjust(1)
    return b.as_markup()


def kb_cancel() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[[InlineKeyboardButton(text="❌ Отмена", callback_data="cancel_flow")]]
    )


# ═══════════════════════════════════════════════════════════════
#  FSM СОСТОЯНИЯ
# ═══════════════════════════════════════════════════════════════

class WorkState(StatesGroup):
    doc_type           = State()
    custom_doc_name    = State()
    mode               = State()
    writing_style      = State()
    topic              = State()
    source_choice      = State()
    source_content     = State()
    institution_type   = State()
    org_type           = State()
    institution        = State()
    group              = State()
    author             = State()
    teacher            = State()
    subject            = State()
    city               = State()
    pages              = State()
    page_number_position = State()
    model              = State()
    payment            = State()
    gost_free_text     = State()


# ═══════════════════════════════════════════════════════════════
#  БОТ И ДИСПЕТЧЕР
# ═══════════════════════════════════════════════════════════════

bot = Bot(token=BOT_TOKEN)
dp  = Dispatcher(storage=MemoryStorage())


# ═══════════════════════════════════════════════════════════════
#  HTTP СЕРВЕР ДЛЯ RENDER HEALTHCHECK
# ═══════════════════════════════════════════════════════════════

async def _start_http_server() -> None:
    port = int(os.getenv("PORT", "10000"))

    async def health(_: web.Request) -> web.Response:
        return web.Response(
            text=json.dumps({"status": "ok", "bot": "GOST Assistant"}, ensure_ascii=False),
            content_type="application/json",
        )

    app = web.Application()
    app.add_routes([web.get("/health", health), web.get("/", health)])

    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, host="0.0.0.0", port=port)
    await site.start()
    print(f"[HTTP] Healthcheck на порту {port}")


# ═══════════════════════════════════════════════════════════════
#  ВСПОМОГАТЕЛЬНЫЕ ТЕКСТЫ СООБЩЕНИЙ
# ═══════════════════════════════════════════════════════════════

def _welcome_text(user_id: int, first_name: str) -> str:
    limits = get_user_limits_info(user_id)
    vip_mark = " 👑" if is_vip(user_id) else ""
    return (
        f"👋 Привет, <b>{first_name}</b>{vip_mark}!\n\n"
        f"📚 <b>ГОСТ-АССИСТЕНТ</b> — создаёт академические работы\n"
        f"в формате DOCX строго по ГОСТ 7.32-2017.\n\n"
        f"<b>Что умею:</b>\n"
        f"• 📄 Рефераты, курсовые, эссе, доклады\n"
        f"• 📑 Развёрнутые названия глав и подглав\n"
        f"• 📐 Точное соблюдение объёма (±10%)\n"
        f"• 🗂 Автоматическое содержание\n"
        f"• ⚙️ Настройка ГОСТ под ваш вуз\n\n"
        f"{limits}\n\n"
        f"👇 <b>Выберите тип работы:</b>"
    )


def _doc_type_card(doc_type: str, user_id: int) -> str:
    dt     = DOC_TYPES[doc_type]
    limits = get_user_limits_info(user_id)
    return (
        f"✅ <b>{dt['name']}</b>\n\n"
        f"📖 <i>{dt['desc']}</i>\n\n"
        f"📄 Объём: <b>{dt['min_pages']}–{dt['max_pages']} страниц</b>\n"
        f"🗂 Структура: {dt['structure']}\n\n"
        f"{limits}\n\n"
        f"💳 <b>Выберите режим генерации:</b>"
    )


def _mode_free_text() -> str:
    period_h = FREE_COOLDOWN // 3600
    period_s = f"{period_h // 24} дн" if period_h >= 24 else f"{period_h} ч"
    return (
        "🆓 <b>Бесплатный режим</b>\n\n"
        f"• Модель: {AI_MODELS.get(FREE_MODEL_KEY, {}).get('name', 'DeepSeek')}\n"
        f"• Максимум: <b>{FREE_MAX_PAGES} страниц</b>\n"
        f"• Лимит: <b>1 генерация раз в {period_s}</b>\n\n"
        "✍️ Введите <b>тему работы</b> одной строкой:"
    )


def _mode_paid_text() -> str:
    paid_str = "♾ безлимитно" if PAID_DAILY_LIMIT == 0 else f"{PAID_DAILY_LIMIT} в день"
    return (
        "⭐ <b>Платный режим</b>\n\n"
        f"• Любая доступная модель ИИ\n"
        f"• Любое количество страниц\n"
        f"• Генерации: <b>{paid_str}</b>\n\n"
        "✍️ Введите <b>тему работы</b> одной строкой:"
    )


# ═══════════════════════════════════════════════════════════════
#  ХЭНДЛЕРЫ — /start и общие команды
# ═══════════════════════════════════════════════════════════════

@dp.message(Command("start"))
async def cmd_start(message: Message, state: FSMContext) -> None:
    await state.clear()
    first = message.from_user.first_name or "пользователь"
    await message.answer(
        _welcome_text(message.from_user.id, first),
        reply_markup=kb_doc_type(),
        parse_mode="HTML",
    )
    await state.set_state(WorkState.doc_type)


@dp.message(Command("help"))
async def cmd_help(message: Message, state: FSMContext) -> None:
    await message.answer(
        "📖 <b>Помощь по ГОСТ-ассистенту</b>\n\n"
        "<b>Команды:</b>\n"
        "/start — начать новую работу\n"
        "/help — эта справка\n"
        "/limits — мои лимиты\n"
        "/cancel — отменить текущий ввод\n\n"
        "<b>Как это работает:</b>\n"
        "1️⃣ Выбираете тип документа\n"
        "2️⃣ Указываете тему, учреждение, данные\n"
        "3️⃣ ИИ генерирует текст и формирует DOCX\n"
        "4️⃣ Получаете готовый файл по ГОСТ\n\n"
        "<b>Платный режим:</b> выбор ИИ-модели, любой объём, без ежедневного лимита.\n\n"
        "<b>Проблемы?</b> Обратитесь к администратору бота.",
        parse_mode="HTML",
    )


@dp.message(Command("limits"))
async def cmd_limits(message: Message) -> None:
    await message.answer(
        f"📊 <b>Ваши генерации</b>\n\n{get_user_limits_info(message.from_user.id)}",
        parse_mode="HTML",
    )


@dp.message(Command("cancel"))
async def cmd_cancel(message: Message, state: FSMContext) -> None:
    await state.clear()
    await message.answer(
        "❌ <b>Отменено.</b>\n\nНачните заново — /start",
        parse_mode="HTML",
    )


@dp.callback_query(F.data == "cancel_flow")
async def h_cancel_flow(cb: CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    await cb.message.edit_text(
        "❌ <b>Генерация отменена.</b>\n\nНачните заново — /start",
        parse_mode="HTML",
    )
    await cb.answer()


@dp.callback_query(F.data == "show_limits")
async def h_show_limits(cb: CallbackQuery) -> None:
    await cb.answer(
        get_user_limits_info(cb.from_user.id).replace("<b>", "").replace("</b>", ""),
        show_alert=True,
    )


# ═══════════════════════════════════════════════════════════════
#  ХЭНДЛЕРЫ — НАСТРОЙКА ГОСТ
# ═══════════════════════════════════════════════════════════════

@dp.callback_query(F.data == "custom_gost")
async def h_custom_gost(cb: CallbackQuery, state: FSMContext) -> None:
    await cb.message.answer(
        "⚙️ <b>Настройка ГОСТ под ваш вуз</b>\n\n"
        "Вставьте требования к оформлению обычным текстом.\n\n"
        "<b>Пример:</b>\n"
        "<i>Шрифт Arial 12pt, поля: лево 20 мм, право 10 мм, верх 15 мм, низ 15 мм. "
        "Межстрочный интервал 1.0. Отступ первой строки 1.25 см. Нумерация снизу.</i>\n\n"
        "ИИ автоматически извлечёт параметры и сохранит для вашего аккаунта.",
        parse_mode="HTML",
        reply_markup=kb_cancel(),
    )
    await state.set_state(WorkState.gost_free_text)
    await cb.answer()


@dp.message(WorkState.gost_free_text)
async def h_save_custom_gost(message: Message, state: FSMContext) -> None:
    data = await state.get_data()
    doc_type = data.get("doc_type", "referat")

    gost_text = (message.text or "").strip()
    if len(gost_text) < 10:
        await message.answer(
            "❌ Слишком коротко. Вставьте полные требования к оформлению.",
            reply_markup=kb_cancel(),
        )
        return

    wait_msg = await message.answer(
        "⏳ <b>Анализирую требования...</b>",
        parse_mode="HTML",
    )

    parsed = await parse_custom_gost_via_ai(FREE_MODEL_KEY, gost_text)
    save_user_gost_config(message.from_user.id, doc_type, parsed)

    await wait_msg.delete()

    pn = "снизу" if parsed["page_number_position"] == "bottom_center" else "сверху"
    await message.answer(
        "✅ <b>ГОСТ сохранён!</b>\n\n"
        "┌─────────────────────────\n"
        f"│ 🔤 Шрифт: <b>{parsed['font_name']} {parsed['font_size']}pt</b>\n"
        f"│ ↕️ Интервал: <b>{parsed['line_spacing']}</b>\n"
        f"│ ↩️ Отступ: <b>{parsed['first_line_indent_cm']} см</b>\n"
        f"│ 📐 Поля (мм): Л{parsed['left_margin_mm']} П{parsed['right_margin_mm']} "
        f"В{parsed['top_margin_mm']} Н{parsed['bottom_margin_mm']}\n"
        f"│ 🔢 Нумерация: <b>{pn} по центру</b>\n"
        "└─────────────────────────\n\n"
        "Эти настройки применятся к вашим следующим работам.",
        parse_mode="HTML",
    )

    await state.set_state(WorkState.doc_type)
    await message.answer(
        "👇 Выберите тип работы для продолжения:",
        reply_markup=kb_doc_type(),
    )


# ═══════════════════════════════════════════════════════════════
#  ХЭНДЛЕРЫ — ВЫБОР ТИПА ДОКУМЕНТА
# ═══════════════════════════════════════════════════════════════

@dp.callback_query(F.data.startswith("dtype_"))
async def h_doc_type(cb: CallbackQuery, state: FSMContext) -> None:
    doc_type = cb.data.replace("dtype_", "", 1)
    if doc_type not in DOC_TYPES:
        await cb.answer("Неизвестный тип документа", show_alert=True)
        return

    await state.update_data(doc_type=doc_type)

    if doc_type == "custom":
        await cb.message.edit_text(
            "🧩 <b>Свой тип документа</b>\n\n"
            "Введите название документа.\n"
            "<i>Примеры: Отчёт по практике, Лабораторная работа, Реферат по материалам</i>",
            parse_mode="HTML",
            reply_markup=kb_cancel(),
        )
        await state.set_state(WorkState.custom_doc_name)
        await cb.answer()
        return

    await cb.message.edit_text(
        _doc_type_card(doc_type, cb.from_user.id),
        reply_markup=kb_mode(),
        parse_mode="HTML",
    )
    await state.set_state(WorkState.mode)
    await cb.answer()


@dp.message(WorkState.custom_doc_name)
async def h_custom_doc_name(message: Message, state: FSMContext) -> None:
    name = (message.text or "").strip()
    if len(name) < 2:
        await message.answer(
            "❌ Название слишком короткое. Попробуйте ещё раз.",
            reply_markup=kb_cancel(),
        )
        return

    await state.update_data(custom_doc_name=name)
    await message.answer(
        f"🧩 <b>Свой тип:</b> {name}\n\n"
        f"{get_user_limits_info(message.from_user.id)}\n\n"
        "💳 <b>Выберите режим генерации:</b>",
        reply_markup=kb_mode(),
        parse_mode="HTML",
    )
    await state.set_state(WorkState.mode)


# ═══════════════════════════════════════════════════════════════
#  ХЭНДЛЕРЫ — РЕЖИМ И ТЕМА
# ═══════════════════════════════════════════════════════════════

@dp.callback_query(F.data.startswith("mode_"))
async def h_mode(cb: CallbackQuery, state: FSMContext) -> None:
    mode = cb.data.replace("mode_", "", 1)
    await state.update_data(mode=mode)

    await cb.message.edit_text(
        "✍️ <b>Выберите стиль написания работы</b>\n\n"
        "Это влияет на язык, структуру предложений и глубину изложения:",
        parse_mode="HTML",
        reply_markup=kb_writing_style(),
    )
    await state.set_state(WorkState.writing_style)
    await cb.answer()


@dp.callback_query(F.data.in_(["style_smart", "style_classic"]))
async def h_writing_style(cb: CallbackQuery, state: FSMContext) -> None:
    style = cb.data.replace("style_", "", 1)
    await state.update_data(writing_style=style)

    data = await state.get_data()
    mode = data.get("mode", "free")
    text = _mode_free_text() if mode == "free" else _mode_paid_text()
    style_label = "🎓 Умный стиль выбран" if style == "smart" else "📝 Классический стиль выбран"

    await cb.message.edit_text(
        f"✅ {style_label}\n\n{text}",
        parse_mode="HTML",
        reply_markup=kb_cancel(),
    )
    await state.set_state(WorkState.topic)
    await cb.answer()


@dp.message(WorkState.topic)
async def h_topic(message: Message, state: FSMContext) -> None:
    topic = (message.text or "").strip()
    if len(topic) < 5:
        await message.answer(
            "❌ Тема слишком короткая. Опишите тему подробнее (минимум 5 символов).",
            reply_markup=kb_cancel(),
        )
        return

    await state.update_data(topic=topic)
    await message.answer(
        f"✅ Тема принята: <b>{topic[:80]}</b>\n\n"
        "📎 <b>Есть ли у вас свои материалы?</b>\n\n"
        "Вы можете добавить план, конспект или тезисы — ИИ использует их как основу.\n"
        "Или выбрать генерацию только по теме.",
        reply_markup=kb_source_choice(),
        parse_mode="HTML",
    )
    await state.set_state(WorkState.source_choice)


@dp.callback_query(F.data == "source_yes")
async def h_source_yes(cb: CallbackQuery, state: FSMContext) -> None:
    """Пользователь выбрал 'Да — добавлю свои материалы'"""
    await cb.message.edit_text(
        "📎 <b>Отправьте ваши материалы</b>\n\n"
        "Вставьте план, тезисы, конспект или любой текст — ИИ использует его как основу.\n"
        "<i>Максимум 12 000 символов.</i>",
        parse_mode="HTML",
        reply_markup=kb_cancel(),  # <-- ИСПРАВЛЕНО
    )
    await state.set_state(WorkState.source_content)
    await cb.answer()

@dp.callback_query(F.data == "source_no")
async def h_source_no(cb: CallbackQuery, state: FSMContext) -> None:
    """Пользователь выбрал 'Нет — только по теме'"""
    await state.update_data(source_content="")
    await cb.message.edit_text(
        "🏛 <b>Тип учебного заведения</b>\n\nВыберите из списка:",
        reply_markup=kb_institution(),  # <-- ИСПРАВЛЕНО: используем правильную клавиатуру
        parse_mode="HTML",
    )
    await state.set_state(WorkState.institution_type)
    await cb.answer()

@dp.message(WorkState.source_content)
async def h_source_content(message: Message, state: FSMContext) -> None:
    content = (message.text or "").strip()
    await state.update_data(source_content=content)
    await message.answer(
        f"✅ Материалы приняты ({len(content)} символов).\n\n"
        "🏛 <b>Тип учебного заведения:</b>",
        reply_markup=kb_institution(),
        parse_mode="HTML",
    )
    await state.set_state(WorkState.institution_type)


# ═══════════════════════════════════════════════════════════════
#  ХЭНДЛЕРЫ — УЧЕБНОЕ ЗАВЕДЕНИЕ
# ═══════════════════════════════════════════════════════════════

@dp.callback_query(F.data.startswith("inst_"))
async def h_inst(cb: CallbackQuery, state: FSMContext) -> None:
    inst = cb.data.replace("inst_", "", 1)
    await state.update_data(institution_type=inst)

    if inst == "custom":
        await state.update_data(org_type="")
        await cb.message.edit_text(
            "✏️ <b>Введите полное название учебного заведения</b>\n\n"
            "<i>Например: Новосибирский государственный технический университет</i>",
            parse_mode="HTML",
            reply_markup=kb_cancel(),
        )
        await state.set_state(WorkState.institution)
    else:
        info = INSTITUTION_TYPES.get(inst, INSTITUTION_TYPES["school"])
        await cb.message.edit_text(
            "🏛 <b>Введите тип организации</b>\n\n"
            f"Пример:\n<i>{info['org_example']}</i>\n\n"
            "Или скопируйте пример целиком.",
            parse_mode="HTML",
            reply_markup=kb_cancel(),
        )
        await state.set_state(WorkState.org_type)

    await cb.answer()


@dp.message(WorkState.org_type)
async def h_org_type(message: Message, state: FSMContext) -> None:
    await state.update_data(org_type=(message.text or "").strip())
    data = await state.get_data()
    info = INSTITUTION_TYPES.get(data.get("institution_type", "school"), INSTITUTION_TYPES["school"])
    await message.answer(
        "🏫 <b>Введите название учебного заведения</b>\n\n"
        f"Пример:\n<i>{info['name_example']}</i>",
        parse_mode="HTML",
        reply_markup=kb_cancel(),
    )
    await state.set_state(WorkState.institution)


@dp.message(WorkState.institution)
async def h_institution(message: Message, state: FSMContext) -> None:
    await state.update_data(institution=(message.text or "").strip())
    await message.answer(
        "👥 <b>Введите класс или группу</b>\n\n<i>Например: 10А, ИТ-21, гр. 315</i>",
        parse_mode="HTML",
        reply_markup=kb_cancel(),
    )
    await state.set_state(WorkState.group)


@dp.message(WorkState.group)
async def h_group(message: Message, state: FSMContext) -> None:
    await state.update_data(group=(message.text or "").strip())
    await message.answer(
        "👤 <b>Введите ФИО автора</b>\n\n<i>Например: Иванов Иван Иванович</i>",
        parse_mode="HTML",
        reply_markup=kb_cancel(),
    )
    await state.set_state(WorkState.author)


@dp.message(WorkState.author)
async def h_author(message: Message, state: FSMContext) -> None:
    author = (message.text or "").strip()
    if len(author.split()) < 2:
        await message.answer(
            "❌ Введите ФИО полностью (минимум имя и фамилия).",
            reply_markup=kb_cancel(),
        )
        return
    await state.update_data(author=author)
    await message.answer(
        "👨‍🏫 <b>Введите ФИО преподавателя</b>\n\n<i>Например: Петров Пётр Петрович</i>",
        parse_mode="HTML",
        reply_markup=kb_cancel(),
    )
    await state.set_state(WorkState.teacher)


@dp.message(WorkState.teacher)
async def h_teacher(message: Message, state: FSMContext) -> None:
    await state.update_data(teacher=(message.text or "").strip())
    await message.answer(
        "📚 <b>Выберите дисциплину (предмет)</b>",
        reply_markup=kb_subject(),
        parse_mode="HTML",
    )
    await state.set_state(WorkState.subject)


# ═══════════════════════════════════════════════════════════════
#  ХЭНДЛЕРЫ — ПРЕДМЕТ, ГОРОД, СТРАНИЦЫ
# ═══════════════════════════════════════════════════════════════

@dp.callback_query(F.data.startswith("subj_"))
async def h_subject_cb(cb: CallbackQuery, state: FSMContext) -> None:
    subj = cb.data.replace("subj_", "", 1)
    if subj == "other":
        await cb.message.edit_text(
            "✏️ <b>Введите название предмета</b>",
            parse_mode="HTML",
            reply_markup=kb_cancel(),
        )
        await state.set_state(WorkState.subject)
    else:
        await state.update_data(subject=subj)
        await cb.message.edit_text(
            f"✅ Предмет: <b>{subj}</b>\n\n"
            "🌆 <b>Выберите город</b>:",
            reply_markup=kb_city(),
            parse_mode="HTML",
        )
        await state.set_state(WorkState.city)
    await cb.answer()


@dp.message(WorkState.subject)
async def h_subject_text(message: Message, state: FSMContext) -> None:
    await state.update_data(subject=(message.text or "").strip())
    await message.answer(
        "🌆 <b>Выберите город:</b>",
        reply_markup=kb_city(),
        parse_mode="HTML",
    )
    await state.set_state(WorkState.city)


@dp.callback_query(F.data.startswith("city_"))
async def h_city_cb(cb: CallbackQuery, state: FSMContext) -> None:
    city = cb.data.replace("city_", "", 1)
    if city == "other":
        await cb.message.edit_text(
            "✏️ <b>Введите название города</b>",
            parse_mode="HTML",
            reply_markup=kb_cancel(),
        )
        await state.set_state(WorkState.city)
    else:
        await state.update_data(city=city)
        await _ask_pages(cb.message, state, cb.from_user.id)
        await state.set_state(WorkState.pages)
    await cb.answer()


@dp.message(WorkState.city)
async def h_city_text(message: Message, state: FSMContext) -> None:
    await state.update_data(city=(message.text or "").strip())
    await _ask_pages(message, state, message.from_user.id)
    await state.set_state(WorkState.pages)


async def _ask_pages(event: Message, state: FSMContext, user_id: int) -> None:
    data     = await state.get_data()
    doc_type = data.get("doc_type", "referat")
    dt       = DOC_TYPES.get(doc_type, DOC_TYPES["referat"])
    mode     = data.get("mode", "free")

    if mode == "free" and not is_vip(user_id):
        max_p = min(dt["max_pages"], FREE_MAX_PAGES)
        note  = f"\n⚠️ <i>В бесплатном режиме максимум {FREE_MAX_PAGES} стр.</i>"
    else:
        max_p = dt["max_pages"]
        note  = ""

    await event.answer(
        f"📄 <b>Количество страниц</b>\n\n"
        f"Допустимо: <b>{dt['min_pages']}–{max_p}</b> страниц{note}\n\n"
        f"Введите число:",
        parse_mode="HTML",
        reply_markup=kb_cancel(),
    )


@dp.message(WorkState.pages)
async def h_pages(message: Message, state: FSMContext) -> None:
    data     = await state.get_data()
    doc_type = data.get("doc_type", "referat")
    dt       = DOC_TYPES.get(doc_type, DOC_TYPES["referat"])

    try:
        pages = int((message.text or "").strip())
    except Exception:
        await message.answer("❌ Введите целое число.", reply_markup=kb_cancel())
        return

    mode = data.get("mode", "free")

    # Клипируем до допустимого диапазона
    min_p = dt["min_pages"]
    if mode == "free" and not is_vip(message.from_user.id):
        max_p = min(dt["max_pages"], FREE_MAX_PAGES)
    else:
        max_p = dt["max_pages"]

    if pages < min_p or pages > max_p:
        await message.answer(
            f"❌ Допустимое количество страниц: <b>{min_p}–{max_p}</b>.\n"
            f"Вы ввели: {pages}. Попробуйте ещё раз.",
            parse_mode="HTML",
            reply_markup=kb_cancel(),
        )
        return

    await state.update_data(pages=pages)
    await message.answer(
        f"✅ Страниц: <b>{pages}</b>\n\n"
        "🔢 <b>Нумерация страниц — где расположить?</b>",
        reply_markup=kb_page_number(),
        parse_mode="HTML",
    )
    await state.set_state(WorkState.page_number_position)


# ═══════════════════════════════════════════════════════════════
#  ХЭНДЛЕРЫ — НУМЕРАЦИЯ, МОДЕЛЬ, ОПЛАТА
# ═══════════════════════════════════════════════════════════════

@dp.callback_query(F.data.in_(["pagepos_bottom", "pagepos_top"]))
async def h_pagepos(cb: CallbackQuery, state: FSMContext) -> None:
    pos = "top_center" if cb.data == "pagepos_top" else "bottom_center"
    await state.update_data(page_number_position=pos)

    data  = await state.get_data()
    pages = int(data.get("pages", 10))
    mode  = data.get("mode", "free")

    # ── Бесплатный режим ──
    if mode == "free":
        if pages > FREE_MAX_PAGES and not is_vip(cb.from_user.id):
            await cb.message.edit_text(
                f"🚫 <b>В бесплатном режиме максимум {FREE_MAX_PAGES} страниц.</b>\n\n"
                "Используйте платный режим для большего объёма.",
                parse_mode="HTML",
            )
            await state.clear()
            await cb.answer()
            return

        ok, reason = check_user_limit(cb.from_user.id, "free")
        if not ok:
            await cb.message.edit_text(reason, parse_mode="HTML")
            await state.clear()
            await cb.answer()
            return

        await cb.message.edit_text(
            "🚀 <b>Запускаю генерацию...</b>\n\nПодождите, это займёт несколько минут.",
            parse_mode="HTML",
        )
        await cb.answer()
        await generate_and_send(cb.message, state, model_key=FREE_MODEL_KEY, pay_mode="free")
        return

    # ── Платный режим ──
    ok, reason = check_user_limit(cb.from_user.id, "paid")
    if not ok and not is_vip(cb.from_user.id):
        await cb.message.edit_text(reason, parse_mode="HTML")
        await state.clear()
        await cb.answer()
        return

    await cb.message.edit_text(
        "🤖 <b>Выберите ИИ-модель</b>\n\n"
        "Цена указана в звёздах Telegram за страницу.\n"
        "Все модели генерируют полноценный академический текст.",
        reply_markup=kb_models(),
        parse_mode="HTML",
    )
    await state.set_state(WorkState.model)
    await cb.answer()


@dp.callback_query(F.data.startswith("model_"))
async def h_model(cb: CallbackQuery, state: FSMContext) -> None:
    model_key = cb.data.replace("model_", "", 1)
    if model_key not in AI_MODELS or not AI_MODELS[model_key].get("api_key"):
        await cb.answer("⚠️ Модель временно недоступна", show_alert=True)
        return

    data  = await state.get_data()
    pages = int(data.get("pages", 10))
    model = AI_MODELS[model_key]
    total = int(model["price_per_page"]) * pages

    await state.update_data(model_key=model_key)

    if is_vip(cb.from_user.id):
        await cb.message.edit_text(
            f"👑 <b>VIP — оплата не требуется</b>\n\n"
            f"Модель: {model['name']}\n"
            f"Страниц: {pages}\n\n"
            f"🚀 Запускаю генерацию...",
            parse_mode="HTML",
        )
        await cb.answer()
        await generate_and_send(cb.message, state, model_key=model_key, pay_mode="paid")
        return

    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text=f"⭐ Оплатить {total} звёзд", callback_data=f"pay_{total}")],
            [InlineKeyboardButton(text="← Выбрать другую модель",   callback_data="back_to_models")],
        ]
    )
    await cb.message.edit_text(
        f"💳 <b>Подтверждение оплаты</b>\n\n"
        f"┌─────────────────────────\n"
        f"│ 🤖 Модель:  {model['name']}\n"
        f"│ 📄 Страниц: {pages}\n"
        f"│ 💰 Цена:    {model['price_per_page']}⭐ × {pages} = <b>{total}⭐</b>\n"
        f"└─────────────────────────\n\n"
        f"Нажмите кнопку для оплаты через Telegram Stars:",
        reply_markup=kb,
        parse_mode="HTML",
    )
    await state.set_state(WorkState.payment)
    await cb.answer()


@dp.callback_query(F.data == "back_to_models")
async def h_back_to_models(cb: CallbackQuery, state: FSMContext) -> None:
    await cb.message.edit_text(
        "🤖 <b>Выберите ИИ-модель</b>",
        reply_markup=kb_models(),
        parse_mode="HTML",
    )
    await state.set_state(WorkState.model)
    await cb.answer()


@dp.callback_query(F.data.startswith("pay_"))
async def h_pay(cb: CallbackQuery, state: FSMContext) -> None:
    stars     = int(cb.data.replace("pay_", "", 1))
    data      = await state.get_data()
    model_key = data.get("model_key", FREE_MODEL_KEY)
    payload   = json.dumps(
        {"model_key": model_key, "pages": data.get("pages", 10)},
        ensure_ascii=False,
    )

    await cb.bot.send_invoice(
        chat_id=cb.from_user.id,
        title="Академическая работа по ГОСТ",
        description=f"{AI_MODELS[model_key]['name']} · {data.get('pages', 10)} стр. · {data.get('topic','')[:40]}",
        payload=payload,
        provider_token="",
        currency="XTR",
        prices=[LabeledPrice(label="Работа по ГОСТ", amount=stars)],
        reply_markup=InlineKeyboardMarkup(
            inline_keyboard=[[InlineKeyboardButton(text=f"⭐ Оплатить {stars} звёзд", pay=True)]]
        ),
    )
    await cb.answer()


@dp.pre_checkout_query()
async def h_pre_checkout(q: PreCheckoutQuery) -> None:
    await q.answer(ok=True)


@dp.message(F.successful_payment)
async def h_payment_ok(message: Message, state: FSMContext) -> None:
    try:
        payload   = json.loads(message.successful_payment.invoice_payload)
        model_key = payload.get("model_key", FREE_MODEL_KEY)
    except Exception:
        model_key = FREE_MODEL_KEY

    await message.answer(
        "✅ <b>Оплата прошла успешно!</b>\n\n"
        "🚀 Начинаю генерацию работы...",
        parse_mode="HTML",
    )
    await generate_and_send(message, state, model_key=model_key, pay_mode="paid")


@dp.callback_query(F.data == "final_new")
async def h_final_new(cb: CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    first = cb.from_user.first_name or "пользователь"
    await cb.message.edit_text(
        _welcome_text(cb.from_user.id, first),
        reply_markup=kb_doc_type(),
        parse_mode="HTML",
    )
    await state.set_state(WorkState.doc_type)
    await cb.answer()


# ═══════════════════════════════════════════════════════════════
#  ГЛАВНАЯ ГЕНЕРАЦИЯ
# ═══════════════════════════════════════════════════════════════

async def generate_and_send(
    event: Message,
    state: FSMContext,
    model_key: str,
    pay_mode: str,
) -> None:
    """
    Основная функция генерации: запрашивает названия глав,
    генерирует текст блоками, собирает DOCX, отправляет.
    """
    async with GEN_SEMAPHORE:
        data = await state.get_data()

        doc_type = data.get("doc_type", "referat")
        dt       = DOC_TYPES.get(doc_type, DOC_TYPES["referat"])
        pages    = int(data.get("pages", 10))
        topic    = data.get("topic", "")
        subject  = data.get("subject", "")

        # Определяем количество глав по типу документа
        if doc_type in ("esse", "doklad", "article"):
            num_chapters = 0  # у них своя структура
        elif doc_type in ("kursovaya", "final_referat", "vkr", "final_project"):
            num_chapters = 4 if doc_type in ("vkr", "final_project") and pages >= 40 else 3
        else:
            num_chapters = 2  # реферат, контрольная, свой

        # Считаем шаги для прогресс-бара (оценка: ~3 подглавы на главу)
        # 1 = названия глав, блоки = intro + подглавы + lit + conclusion, 1 = DOCX, 1 = отправка
        if doc_type in ("esse",):
            extra_blocks = 6
        elif doc_type in ("doklad",):
            extra_blocks = 5
        elif doc_type in ("article",):
            extra_blocks = 7
        else:
            # chapter_titles ещё не готов — оцениваем: num_chapters * 3 подглавы
            num_subs = max(1, num_chapters * 3)
            extra_blocks = 1 + num_subs + 2  # intro + N subs + lit + conclusion
        total_steps  = 1 + extra_blocks + 2

        progress_msg = await event.answer(
            "⏳ <b>Инициализация...</b>",
            parse_mode="HTML",
        )
        prog = Progress(
            msg=progress_msg,
            title=f"Создаю {dt['word'].lower()}",
            total_steps=total_steps,
            model_name=AI_MODELS.get(model_key, {}).get("name", ""),
        )

        # Запускаем фоновый цикл анимации
        stop_anim = asyncio.Event()
        anim_task = asyncio.create_task(prog.animate_loop(stop_anim))

        try:
            record_user_generation(event.chat.id, pay_mode)

            gost = get_gost_config(doc_type, event.chat.id)
            gost["page_number_position"] = data.get(
                "page_number_position",
                gost.get("page_number_position", "bottom_center"),
            )

            # ── Шаг 1: Генерируем названия глав ──
            await prog.update(label="🧠 Придумываю названия глав...", force=True)
            if num_chapters > 0:
                chapter_titles = await generate_chapter_titles(
                    model_key, doc_type, topic, subject, num_chapters
                )
            else:
                chapter_titles = []
            await prog.update(step_done=True)

            writing_style = data.get("writing_style", "classic")

            # ── Шаги 2–N: Генерируем текст ──
            await prog.update(label="✍️ Генерирую текст разделов...")
            parts = await generate_text_blocks(
                topic=topic,
                pages=pages,
                doc_type=doc_type,
                subject=subject,
                model_key=model_key,
                source=data.get("source_content", ""),
                chapter_titles=chapter_titles,
                writing_style=writing_style,
                prog=prog,
            )

            # ── Проверка соответствия дисциплине (приоритет 🔴) ──
            await prog.update(label="🔎 Проверяю соответствие дисциплине...")
            sample_for_check = "\n\n".join(
                str(parts.get(k, "")) for k in ("intro", "main1", "part1", "ch1_s1")
                if parts.get(k)
            )[:2500]
            relevance_ok, relevance_reason = await verify_discipline_relevance(
                model_key, topic, subject, sample_for_check,
            )
            if not relevance_ok:
                print(f"[RELEVANCE] ⚠️ Текст может не соответствовать дисциплине "
                      f"«{subject}»: {relevance_reason}")

            # ── Сборка структуры ──
            blocks = generate_structure(doc_type, parts, chapter_titles)

            # ── DOCX ──
            await prog.update(label="📄 Собираю DOCX-документ...", step_done=True)
            work_dir = os.path.join(os.getcwd(), "_out")
            os.makedirs(work_dir, exist_ok=True)
            ts      = datetime.now().strftime("%Y%m%d_%H%M%S")
            tmp_in  = os.path.join(work_dir, f"tmp_{event.chat.id}_{ts}.docx")
            tmp_out = os.path.join(work_dir, f"final_{event.chat.id}_{ts}.docx")

            docx_raw = build_docx_bytes(data, blocks, gost)
            with open(tmp_in, "wb") as f:
                f.write(docx_raw)
            if not os.path.exists(tmp_in):
                raise RuntimeError(f"Не удалось записать {tmp_in}")

            # ── ПОДГОНКА СТРАНИЦ (исправленная) ──
            blocks, final_pages = await precise_page_adjustment(
                tmp_in=tmp_in,
                blocks=blocks,
                target_pages=pages,
                topic=topic,
                model_key=model_key,
                writing_style=writing_style,
                data=data,
                gost=gost,
                prog=prog,
                work_dir=work_dir,
            )

            # ── LibreOffice (финальная конвертация — обновит TOC и поля PAGE) ──
            await prog.update(label="🔄 Обновляю содержание (LibreOffice)...", step_done=True)
            updated    = libreoffice_update_docx(tmp_in, tmp_out)
            final_path = tmp_out if updated else tmp_in

            # Если после LibreOffice количество страниц изменилось, подгоняем финальный файл еще раз
            post_lo_pages = await measure_pages_async(final_path, work_dir)
            if post_lo_pages is not None and post_lo_pages != pages:
                print(f"[PAGES] LO сдвинул страницы ({post_lo_pages} вместо {pages}). Запуск финальной коррекции...")
                blocks, final_pages = await precise_page_adjustment(
                    tmp_in=final_path,
                    blocks=blocks,
                    target_pages=pages,
                    topic=topic,
                    model_key=model_key,
                    writing_style=writing_style,
                    data=data,
                    gost=gost,
                    prog=prog,
                    work_dir=work_dir,
                )
                final_pages = await measure_pages_async(final_path, work_dir) or pages
            else:
                final_pages = post_lo_pages or pages

            print(f"[PAGES] 📤 Итог в caption: {final_pages} страниц")

            # ── Имя файла ──
            safe_topic = re.sub(r'[<>"/:\\|?*]', "", topic[:35]).replace(" ", "_")
            fname      = f"{dt['word'].replace(' ', '_')}_{safe_topic}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"

            with open(final_path, "rb") as f:
                final_bytes = f.read()

            # ── Отправляем ──
            await prog.update(label="📤 Отправляю файл...", step_done=True)

            toc_status = "✅ обновлено автоматически" if updated else "⚠️ обновите вручную в Word"
            style_label = "🎓 Умный" if data.get("writing_style") == "smart" else "📝 Классический"
            pages_line = (
                f"📄 Страниц: <b>{final_pages}</b> (заказано {pages})"
                if final_pages != pages
                else f"📄 Страниц: <b>{final_pages}</b> ✅"
            )
            relevance_status = (
                f"✅ соответствует «{subject}»"
                if relevance_ok
                else f"⚠️ проверьте тему: {relevance_reason[:60]}"
            )
            caption = (
                f"🎉 <b>{dt['word']} ГОТОВ!</b>\n\n"
                f"┌─────────────────────────\n"
                f"│ 📖 Тема: {topic[:60]}\n"
                f"│ {pages_line}\n"
                f"│ 🤖 ИИ: {AI_MODELS.get(model_key, {}).get('name', model_key)}\n"
                f"│ ✍️ Стиль: {style_label}\n"
                f"│ 📐 Шрифт: {gost.get('font_name')} {gost.get('font_size')}pt\n"
                f"│ ↕️ Интервал: {gost.get('line_spacing')}\n"
                f"│ 🎓 Дисциплина: {relevance_status}\n"
                f"│ 📑 Содержание: {toc_status}\n"
                f"└─────────────────────────"
            )

            await event.answer_document(
                BufferedInputFile(final_bytes, filename=fname),
                caption=caption,
                parse_mode="HTML",
            )

        except Exception as e:
            print(f"[GEN ERROR] {e}")
            await prog.finish(
                "❌ <b>Ошибка генерации</b>\n\n"
                f"Причина: {str(e)[:200]}\n\n"
                "Попробуйте ещё раз или выберите другую модель (/start)."
            )
            await state.clear()
            return

        finally:
            stop_anim.set()
            anim_task.cancel()
            try:
                await anim_task
            except asyncio.CancelledError:
                pass

        await prog.delete()

        await event.answer(
            "✅ <b>Работа готова!</b>\n\n"
            "Если содержание не обновилось — откройте файл в Word и нажмите:\n"
            "<code>Ctrl+A → F9 → Обновить всё поле</code>",
            reply_markup=kb_final(),
            parse_mode="HTML",
        )
        await state.clear()

        # Чистим временные файлы
        for tmp in (tmp_in, tmp_out):
            try:
                if os.path.exists(tmp):
                    os.remove(tmp)
            except Exception:
                pass


# ═══════════════════════════════════════════════════════════════
#  ЗАПУСК
# ═══════════════════════════════════════════════════════════════

async def main() -> None:
    print("═" * 62)
    print("  🤖  ГОСТ-АССИСТЕНТ v2.7")
    print("═" * 62)
    print(f"  LibreOffice : {shutil.which('soffice') or '❌ не найден'}")
    print(f"  DeepSeek    : {'✅' if DEEPSEEK_KEY else '❌ нет ключа'}")
    print(f"  OpenRouter  : {'✅' if OPENROUTER_KEY else '❌ нет ключа'}")
    print(f"  Groq        : {'✅' if GROQ_KEY else '❌ нет ключа'}")
    print(f"  VIP users   : {VIP_USERS or 'нет'}")
    _fcd_days = FREE_COOLDOWN / 86400
    print(f"  Free limit  : 1 ген / {_fcd_days:.1f} дн, max {FREE_MAX_PAGES} стр.")
    print(f"  Paid limit  : {'безлимит' if PAID_DAILY_LIMIT == 0 else str(PAID_DAILY_LIMIT)+'/день'}")
    print("═" * 62)

    await _start_http_server()

    while True:
        try:
            print("[BOT] Запускаю polling...")
            await dp.start_polling(bot)
        except Exception as e:
            print(f"[BOT] Ошибка поллинга: {e}")
            await asyncio.sleep(5)


if __name__ == "__main__":
    asyncio.run(main())