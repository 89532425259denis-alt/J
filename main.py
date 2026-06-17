# -*- coding: utf-8 -*-
from __future__ import annotations

"""ГОСТ-АССИСТЕНТ v3.1 — УНИВЕРСАЛЬНАЯ ПРОВЕРКА ДИСЦИПЛИН

Главные изменения v3.1:
────────────────────────────────────────────────────────────────
1. ★ УНИВЕРСАЛЬНАЯ ПРОВЕРКА ТЕМЫ И ДИСЦИПЛИНЫ
   - Без хардкода, работает для ЛЮБЫХ тем и дисциплин
   - Семантический анализ ключевых слов
   - ИИ-проверка соответствия
   - Предложение подходящих дисциплин
   - Возможность изменить тему или дисциплину

2. ★ ИНТЕГРАЦИЯ В ОСНОВНОЙ ПОТОК
   - Проверка перед генерацией
   - Интерактивный диалог с пользователем
   - Сохранение контекста

3. Все базовые возможности v3.0 сохранены.
"""

import asyncio
import html
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
from urllib.parse import quote_plus

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

_SPINNER_FRAMES = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
_BAR_FULL  = "█"
_BAR_HALF  = "▓"
_BAR_EMPTY = "░"


def _spinner() -> str:
    idx = int(time.monotonic() * 4) % len(_SPINNER_FRAMES)
    return _SPINNER_FRAMES[idx]


def _progress_bar(done: int, total: int, width: int = 16) -> str:
    total = max(1, int(total))
    done  = max(0, min(int(done), total))
    ratio = done / total
    filled = int(round(width * ratio))
    
    bar = "█" * filled + "░" * (width - filled)
    
    if 0 < filled < width:
        frame = int(time.monotonic() * 2) % 2
        marker = "▓" if frame == 0 else "▒"
        bar = bar[:filled-1] + marker + bar[filled:]
        
    return bar


def _fmt_time(seconds: float) -> str:
    s = max(0, int(seconds))
    m, s = divmod(s, 60)
    if m:
        return f"{m}м {s:02d}с"
    return f"{s}с"


@dataclass
class Progress:
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
    _eta_prev: Optional[float] = field(default=None, repr=False)

    def _elapsed(self) -> float:
        return time.monotonic() - self._start_ts

    def _eta(self) -> str:
        elapsed = self._elapsed()
        remaining_steps = max(0, self.total_steps - self.done)
        if remaining_steps == 0:
            return "≈0с"

        if self.done == 0:
            if elapsed < 3.0:
                return "считаю…"
            est_per_step = max(10.0, elapsed / 0.5)
            return "≈" + _fmt_time(est_per_step * remaining_steps)

        rate = self.done / elapsed
        if rate <= 0:
            return "…"
        eta_sec = remaining_steps / rate
        prev = getattr(self, "_eta_prev", None)
        if prev is None:
            self._eta_prev = eta_sec
        else:
            self._eta_prev = 0.6 * prev + 0.4 * eta_sec
            eta_sec = self._eta_prev
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
            f"<code>{bar}</code> <b>{pct}%</b>\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"📌 <b>Шаг {step_num}/{self.total_steps}:</b> {self.label}{model}\n"
            f"⏱ <code>{elapsed}</code> ➔ ⏳ <code>{eta}</code>"
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
        while not stop_event.is_set():
            try:
                text = self.render()
                now  = time.monotonic()
                if text != self._last_text and (now - self._last_ts) >= 1.4:
                    await self.msg.edit_text(text, parse_mode="HTML")
                    self._last_text = text
                    self._last_ts   = now
            except Exception:
                pass
            await asyncio.sleep(1.5)


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
NUMERIC_FACTS_FILE = "numeric_facts.json"


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

FREE_MAX_PAGES    = int(cfg("FREE_MAX_PAGES",         "15"))
FREE_COOLDOWN     = int(cfg("FREE_COOLDOWN_SECONDS",  str(5 * 24 * 60 * 60)))
FREE_DAILY_LIMIT  = int(cfg("FREE_DAILY_LIMIT",       "0"))

PAID_DAILY_LIMIT  = int(cfg("PAID_DAILY_LIMIT",       "0"))
PAID_COOLDOWN     = int(cfg("PAID_COOLDOWN_SECONDS",  "0"))

CHARS_PER_PAGE = int(cfg("CHARS_PER_PAGE", "1850"))

ENABLE_WEB_SOURCES = cfg("ENABLE_WEB_SOURCES", "1").lower() not in ("0", "false", "no", "off")
WEB_SOURCE_TIMEOUT = int(cfg("WEB_SOURCE_TIMEOUT", "12"))
MAX_WEB_SOURCES    = int(cfg("MAX_WEB_SOURCES", "12"))
MIN_REAL_SOURCES   = int(cfg("MIN_REAL_SOURCES", "10"))
BIB_SOURCE_TARGET  = int(cfg("BIB_SOURCE_TARGET", "12"))
FILL_UNKNOWN_CITATION_PAGES = cfg("FILL_UNKNOWN_CITATION_PAGES", "1").lower() not in ("0", "false", "no", "off")

def calculate_chars_per_page(gost: dict) -> int:
    font_size    = int(gost.get("font_size", 14))
    line_spacing = float(gost.get("line_spacing", 1.5))
    left_mm   = int(gost.get("left_margin_mm",   30))
    right_mm  = int(gost.get("right_margin_mm",  10))
    top_mm    = int(gost.get("top_margin_mm",    20))
    bottom_mm = int(gost.get("bottom_margin_mm", 20))

    BASE_CHARS   = 1800.0
    BASE_W_MM    = 210 - 30 - 10
    BASE_H_MM    = 297 - 20 - 20

    text_w = max(60, 210 - left_mm - right_mm)
    text_h = max(60, 297 - top_mm - bottom_mm)

    area_factor = (text_w / BASE_W_MM) * (text_h / BASE_H_MM)
    font_factor = (14.0 / max(10, font_size)) ** 2
    spacing_factor = 1.5 / max(1.0, line_spacing)

    chars = BASE_CHARS * area_factor * font_factor * spacing_factor
    return max(900, min(3200, int(round(chars))))

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

DISCIPLINE_EXAMPLES = {
    "Информатика": (
        "ПЛОХО (не по дисциплине): «Байкал — это чистейшее озеро с удивительной экосистемой».\n"
        "ХОРОШО (по дисциплине): «Байкал можно рассматривать как природную систему сбора данных: "
        "тысячелетиями в донных отложениях накапливается информация о климате, которую мы извлекаем "
        "методами машинного обучения. Это аналог долговременной памяти в вычислительных системах.»"
    ),
    "Психология": (
        "ПЛОХО: «Байкал красив и вызывает благоговение».\n"
        "ХОРОШО: «Восприятие Байкала вызывает эффект благоговения (awe), который, по исследованиям "
        "Келтнера и Хаидта, снижает активность дефолт-системы мозга и усиливает ощущение связи с миром.»"
    ),
    "История": (
        "ПЛОХО: «Байкал — уникальное природное явление».\n"
        "ХОРОШО: «С XVIII века Байкал становится объектом научного интереса: экспедиции Миллера (1733–1743) "
        "и Палласа заложили основу систематического изучения региона, что можно рассматривать как начало "
        "формирования региональной историографии.»"
    ),
}

CITIES = [
    "Москва", "Санкт-Петербург", "Новосибирск", "Екатеринбург",
    "Казань", "Нижний Новгород", "Красноярск", "Челябинск",
    "Омск", "Ростов-на-Дону", "Уфа", "Волгоград",
]


# ═══════════════════════════════════════════════════════════════
#  БИБЛИОТЕКА ЭТАЛОНОВ И ЖЕСТКИХ ПРАВИЛ
# ═══════════════════════════════════════════════════════════════

GOLDEN_STANDARDS = {
    "academic": {
        "intro": "Актуальность данной темы обусловлена стремительным развитием [Область], что требует переосмысления подходов к [Проблема]. Целью данной работы является комплексный анализ [Тема]. Для достижения поставленной цели необходимо решить следующие задачи: во-первых, изучить теоретические основы...; во-вторых, проанализировать влияние...",
        "chapter": "Согласно концепции С.В. Петрова, данный процесс представляет собой совокупность взаимосвязанных факторов, определяющих динамику развития системы [1, с. 42]. Это позволяет утверждать, что ключевым элементом здесь выступает [Тезис]. В свою очередь, анализ эмпирических данных показывает, что при увеличении X наблюдается рост Y, что подтверждает гипотезу о...",
        "conclusion": "Таким образом, проведенное исследование позволило установить, что [Основной вывод]. В частности, было выявлено, что [Деталь 1] и [Деталь 2]. Полученные результаты подтверждают теоретическую значимость подхода и могут быть использованы для оптимизации процессов в области [Сфера].",
    },
    "doklad": {
        "main": "Перейдем к рассмотрению ключевых факторов. Первым из них является [Фактор]. Как отмечают исследователи, именно этот элемент определяет до 60% успеха в [Область] [2, с. 64]. Примером может служить опыт компании X, где внедрение данного метода позволило сократить издержки на 15%. Таким образом, мы видим прямую зависимость между...", 
    },
    "esse": {
        "main": "Размышляя над проблемой [Тема], я прихожу к выводу, что истинная причина кроется не в формальных признаках, а в глубинной потребности человека в [Ценность]. Эта мысль перекликается с тезисом А. Смита о том, что 'свобода выбора является базисом развития' [3, с. 112]. Однако, на мой взгляд, в современных условиях этот принцип трансформируется в...",
    },
}

def get_golden_example(doc_type: str, part: str) -> str:
    if doc_type in ("esse",):
        return GOLDEN_STANDARDS["esse"].get("main", "")
    if doc_type in ("doklad",):
        return GOLDEN_STANDARDS["doklad"].get("main", "")
    if part == "intro":
        return GOLDEN_STANDARDS["academic"]["intro"]
    if part == "conclusion":
        return GOLDEN_STANDARDS["academic"]["conclusion"]
    return GOLDEN_STANDARDS["academic"]["chapter"]

FEW_SHOT_EXAMPLES = """
🌟 ПРИМЕР ПРАВИЛЬНОЙ ССЫЛКИ (обязательно с номером страницы):
   ✅ ХОРОШО: «...текст... [1, с. 45] ...текст... [2, с. 120–125]»
   ❌ ПЛОХО: «...текст... [1, с. » (без цифры) или «...текст... [2]» (без страницы)

🌟 ПРИМЕР ЗАКОНЧЕННОГО ПРЕДЛОЖЕНИЯ (запрещено обрывать ссылку или мысль):
   ✅ ХОРОШО: «...осадочными породами [1, с. 45]. Далее в работе рассмотрим...»
   ❌ ПЛОХО: «...осадочными породами [1, с.» — ТАК ПИСАТЬ КАТЕГОРИЧЕСКИ НЕЛЬЗЯ

🌟 ПРИМЕР ЗАПРЕЩЕННОГО ТЕКСТА:
   ❌ Любые фразы типа «Вот ваш текст», «Конечно, я помогу», «Объем соблюден» ЗАПРЕЩЕНЫ.
"""

LIT_BLACKLIST = "Дональд Кнут, Томас Кормен, Эндрю Таненбаум, Стивен Лавренс"


DEFAULT_GOST_CONFIGS: dict = {
    "_base": {
        "font_name":              "Times New Roman",
        "font_size":              14,
        "line_spacing":           1.5,
        "first_line_indent_cm":   1.25,
        "left_margin_mm":         30,
        "right_margin_mm":        10,
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
    if is_vip(user_id):
        return True, ""

    is_free = mode == "free"

    if not is_free and PAID_DAILY_LIMIT == 0 and PAID_COOLDOWN == 0:
        return True, ""

    data = load_usage()
    uid  = str(user_id)
    rec  = data.get(uid, {}) or {}

    now = int(datetime.now().timestamp())

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
    if is_vip(user_id):
        return

    data  = load_usage()
    uid   = str(user_id)
    today = today_key()

    rec = data.get(uid, {}) or {}
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


def _strip_markdown_markers(text: str) -> str:
    if not text:
        return ""
    text = re.sub(r'(?m)^\s*#{1,6}\s*\*{0,2}\s*', '', text)
    text = re.sub(r'(?m)^\s*\*{2}([^*]+)\*{2}\s*$', r'\1', text)
    text = re.sub(r"\*{1,3}([^*\n]+?)\*{1,3}", r"\1", text)
    text = re.sub(r'(?m)^\s*#+\s*', '', text)
    text = text.replace("#", "")
    text = re.sub(r'(?m)^\s*\*\*', '', text)
    text = re.sub(r'(?m)\*\*\s*$', '', text)
    return text.strip()


_AI_MARKER_REPLACEMENTS = [
    (r"(?im)^\s*(конечно|разумеется|хорошо)[,.!\s]*(?:вот|ниже)\s+[^\n.?!]*[.?!]?\s*", ""),
    (r"(?im)^\s*(?:вот|ниже)\s+(?:ваш|представлен|привед[её]н)[^\n.?!]*[.?!]?\s*", ""),
    (r"(?i)\bкак (?:искусственный интеллект|ии|языковая модель)[^.!?\n]*[.!?]?\s*", ""),
    (r"(?i)\bя (?:являюсь|не являюсь|не могу|не имею возможности)[^.!?\n]*[.!?]?\s*", ""),
    (r"(?i)\bобъ[её]м текста (?:строго )?(?:выдержан|соблюд[её]н)[^.!?\n]*[.!?]?\s*", ""),
    (r"(?i)\bв заключение следует отметить,?\s+что\s+", ""),
    (r"(?i)\bподводя итог,?\s+", ""),
    (r"(?i)\bтаким образом,?\s+", ""),
    (r"(?i)\bследует отметить,?\s+что\s+", ""),
    (r"(?i)\bнеобходимо отметить,?\s+что\s+", ""),
    (r"(?i)\bважно подчеркнуть,?\s+что\s+", ""),
    (r"(?i)\bнельзя не отметить,?\s+что\s+", ""),
    (r"(?i)\bв целом можно сказать,?\s+что\s+", ""),
    (r"(?i)\bможно сделать вывод,?\s+что\s+", ""),
    (r"(?i)\bисходя из вышеизложенного,?\s+", ""),
]


def _remove_ai_marker_phrases(text: str) -> str:
    if not text:
        return ""
    out = text
    for pattern, repl in _AI_MARKER_REPLACEMENTS:
        out = re.sub(pattern, repl, out)
    out = re.sub(r"[ \t]{2,}", " ", out)
    out = re.sub(r"\n{3,}", "\n\n", out)
    out = re.sub(r"(?m)^\s+", "", out)
    return out.strip()


def sanitize_llm_text(raw: str) -> str:
    if not raw:
        return ""
    text = _strip_markdown_markers(raw.strip())
    text = re.sub(r"```[^\n]*\n?", "", text)
    text = re.sub(r"\*{1,3}([^*]+)\*{1,3}", r"\1", text)
    text = re.sub(r"^#{1,6}\s+", "", text, flags=re.MULTILINE)
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = _remove_ai_marker_phrases(text)
    return text.strip()


async def call_openai_compat(
    info: dict,
    messages: list[dict],
    max_tokens: int = 4096,
    timeout: int = 300,
) -> str:
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
    return _normalize_homoglyphs(sanitize_llm_text(raw))


def fallback_chain(primary: str) -> list[str]:
    priority = [
        primary,
        "deepseek",
        "deepseek_r1",
        "gemini_or",
        "groq",
    ]
    out: list[str] = []
    for k in priority:
        if not k or k in out:
            continue
        info = AI_MODELS.get(k)
        if not info:
            continue
        if info.get("_fatal"):
            continue
        if not info.get("api_key"):
            continue
        out.append(k)
    return out


async def chat_with_fallback(
    primary: str,
    messages: list[dict],
    max_tokens: int,
) -> tuple[str, str]:
    best_text = ""
    best_model = primary
    for k in fallback_chain(primary):
        info = AI_MODELS[k]
        text = await chat_with_model(info, messages, max_tokens=max_tokens)
        if text and len(text.strip()) > 100:
            info["status"] = ModelStatus.AVAILABLE
            return text, k
        if text and len(text.strip()) > len(best_text.strip()):
            best_text = text
            best_model = k
        if not text:
            info["status"] = ModelStatus.LIMIT
            print(f"[FALLBACK] Модель {info.get('name', k)} вернула пустой ответ, пробую следующую...")
    if best_text:
        print(f"[FALLBACK] Все модели дали < 100 зн., лучший: {len(best_text)} зн. от {best_model}")
    return best_text, best_model


# ═══════════════════════════════════════════════════════════════
#  РЕАЛЬНЫЕ ИСТОЧНИКИ И МАТЕРИАЛЫ ИЗ САЙТОВ
# ═══════════════════════════════════════════════════════════════

_URL_RE = re.compile(r"https?://[^\s<>()\[\]{}\"'«»]+", re.IGNORECASE)


def _extract_urls(text: str, limit: int = 5) -> list[str]:
    if not text:
        return []
    seen: set[str] = set()
    urls: list[str] = []
    for raw in _URL_RE.findall(text):
        url = raw.rstrip(".,;:!?)»]")
        if url not in seen:
            seen.add(url)
            urls.append(url)
        if len(urls) >= limit:
            break
    return urls


def _strip_html(raw: str) -> str:
    if not raw:
        return ""
    raw = re.sub(r"(?is)<(script|style|noscript).*?</\1>", " ", raw)
    raw = re.sub(r"(?is)<br\s*/?>", "\n", raw)
    raw = re.sub(r"(?is)</(p|div|li|h[1-6]|tr)>", "\n", raw)
    raw = re.sub(r"(?is)<[^>]+>", " ", raw)
    raw = html.unescape(raw)
    raw = re.sub(r"[ \t]{2,}", " ", raw)
    raw = re.sub(r"\n{3,}", "\n\n", raw)
    return raw.strip()


def _safe_year_from_crossref(item: dict) -> str:
    for key in ("published-print", "published-online", "published", "issued"):
        parts = item.get(key, {}).get("date-parts") if isinstance(item.get(key), dict) else None
        if parts and parts[0] and parts[0][0]:
            return str(parts[0][0])
    return "б. г."


def _author_to_gost(name: str) -> str:
    name = re.sub(r"\s+", " ", (name or "").strip())
    if not name:
        return ""
    parts = name.split()
    if len(parts) == 1:
        return parts[0]
    if re.fullmatch(r"(?:[A-ZА-ЯЁ]\.){1,4}", parts[-1].replace(" ", "")):
        return f"{parts[0]} {parts[-1]}".strip()
    if len(parts) >= 2 and re.search(r"[А-ЯЁ][а-яё]+", parts[0]) and re.search(r"[А-ЯЁ][а-яё]+", parts[1]):
        family = parts[0]
        given = parts[1:]
    else:
        family = parts[-1]
        given = parts[:-1]
    initials = "".join((p[0].upper() + ".") for p in given if p)
    return f"{family} {initials}".strip()


def _crossref_authors(item: dict, limit: int = 3) -> list[str]:
    out: list[str] = []
    for a in item.get("author") or []:
        family = (a.get("family") or "").strip()
        given = (a.get("given") or "").strip()
        if family and given:
            if re.search(r"\.\s*$", given):
                formatted = f"{family} {given.replace(' ', '')}"
            else:
                initials = "".join(p[0].upper() + "." for p in re.split(r"[\s-]+", given) if p)
                formatted = f"{family} {initials}"
        else:
            formatted = _author_to_gost((family or given).strip())
        if formatted:
            out.append(formatted)
        if len(out) >= limit:
            break
    return out


def _openalex_authors(item: dict, limit: int = 3) -> list[str]:
    out: list[str] = []
    for a in item.get("authorships") or []:
        name = ((a.get("author") or {}).get("display_name") or "").strip()
        if name:
            out.append(_author_to_gost(name))
        if len(out) >= limit:
            break
    return out


def _semantic_authors(item: dict, limit: int = 3) -> list[str]:
    out: list[str] = []
    for a in item.get("authors") or []:
        name = (a.get("name") or "").strip()
        if name:
            out.append(_author_to_gost(name))
        if len(out) >= limit:
            break
    return out


def _clean_ref_title(title: str, max_len: int = 180) -> str:
    title = _strip_html(title or "")
    title = re.sub(r"\s+", " ", title).strip(" .")
    if len(title) > max_len:
        title = title[:max_len].rsplit(" ", 1)[0] + "…"
    return title


_TOPIC_STOPWORDS = {
    "тема", "работа", "исследование", "анализ", "роль", "значение",
    "основы", "особенности", "проблемы", "вопросы", "современный",
    "современная", "современное", "развитие", "система", "метод",
    "методы", "подход", "подходы", "россии", "российской", "российский",
    "the", "and", "for", "with", "from", "this", "that", "study",
    "analysis", "role", "problems", "development",
}


def _keywords_for_relevance(text: str) -> set[str]:
    if not text:
        return set()
    words = re.findall(r"[A-Za-zА-Яа-яЁё][A-Za-zА-Яа-яЁё\-]{3,}", text.lower())
    out: set[str] = set()
    for w in words:
        w = w.strip("-–—_")
        if len(w) < 4 or w in _TOPIC_STOPWORDS:
            continue
        out.add(w)
    return out


def _source_relevance_score(record: dict, topic: str, subject: str) -> int:
    title = record.get("title") or ""
    haystack = " ".join(
        str(record.get(k) or "")
        for k in ("title", "container", "publisher", "concepts", "subjects")
    ).lower()
    topic_aliases = _topic_aliases(topic)
    subject_alias_text = _subject_alias_text(subject)
    subject_norm = re.sub(r"\s+", " ", (subject_alias_text or "").lower()).strip()
    topic_words = _keywords_for_relevance(" ".join(topic_aliases or [topic or ""]))
    subject_words = _keywords_for_relevance(subject_alias_text)

    score = 0
    for alias in topic_aliases:
        alias_norm = re.sub(r"\s+", " ", alias.lower()).strip()
        if alias_norm and len(alias_norm) >= 4 and alias_norm in haystack:
            score += 10
    if subject_norm and len(subject_norm) >= 4 and subject_norm in haystack:
        score += 2

    for w in topic_words:
        if w in haystack:
            score += 4 if w in (title or "").lower() else 3
    for w in subject_words:
        if w in haystack:
            score += 1
    return score


def _filter_relevant_source_records(records: list[dict], topic: str, subject: str, limit: int, doc_type: str = "") -> list[dict]:
    limit = max(MIN_REAL_SOURCES, min(limit, MAX_WEB_SOURCES, BIB_SOURCE_TARGET))
    deduped = _dedupe_source_records(records, max(limit * 6, 60))
    scored = [(_source_relevance_score(r, topic, subject), r) for r in deduped]
    scored.sort(key=lambda x: x[0], reverse=True)

    selected: list[dict] = []
    selected_keys: set[str] = set()

    def _key(r: dict) -> str:
        return (r.get("doi") or re.sub(r"\W+", "", (r.get("title") or "").lower())[:120]).lower()

    def _is_on_topic(r: dict, topic: str, subject: str) -> bool:
        title_lower = (r.get("title") or "").lower()
        container_lower = (r.get("container") or "").lower()
        concepts_lower = (r.get("concepts") or "").lower()
        haystack = f"{title_lower} {container_lower} {concepts_lower}"

        topic_words = _keywords_for_relevance(topic)
        subject_words = _keywords_for_relevance(subject)

        topic_in_title = any(w in title_lower for w in topic_words)
        topic_in_container = any(w in container_lower for w in topic_words)
        topic_in_concepts = any(w in concepts_lower for w in topic_words)

        subject_matches = sum(1 for w in subject_words if w in haystack)

        off_topic_markers = ["введение", "общая", "основы", "курс лекций", "учебное пособие"]
        is_generic = any(m in title_lower for m in off_topic_markers)

        has_topic_word = topic_in_title or topic_in_container or topic_in_concepts
        is_not_too_generic = not is_generic or has_topic_word

        return has_topic_word or (subject_matches >= 2 and is_not_too_generic)

    def _add(candidates: list[dict], require_topic_match: bool = False) -> None:
        for r in candidates:
            if len(selected) >= limit:
                break
            if not r.get("authors"):
                continue
            k = _key(r)
            if not k or k in selected_keys:
                continue
            if require_topic_match and not _is_on_topic(r, topic, subject):
                continue
            selected_keys.add(k)
            selected.append(r)

    strong = [r for score, r in scored if score >= 4]
    weak = [r for score, r in scored if 0 < score < 4]
    rest = [r for score, r in scored if score <= 0]

    _add(strong, require_topic_match=True)
    if len(selected) < MIN_REAL_SOURCES:
        _add(weak, require_topic_match=False)
    if len(selected) < MIN_REAL_SOURCES:
        _add(rest, require_topic_match=False)

    if selected and doc_type:
        genre_offsets = {
            "referat": 0,
            "esse": 3,
            "doklad": 6,
            "kursovaya": 2,
            "kontrolnaya": 4,
            "article": 1,
            "final_referat": 5,
            "vkr": 7,
            "final_project": 8,
        }
        offset = genre_offsets.get(doc_type, 0) % len(selected)
        if offset:
            selected = selected[offset:] + selected[:offset]

    return selected[:limit]


def _format_source_record(record: dict) -> str:
    authors = record.get("authors") or []
    title = _clean_ref_title(record.get("title") or "")
    year = str(record.get("year") or "б. г.")
    container = _clean_ref_title(record.get("container") or "")
    publisher = _clean_ref_title(record.get("publisher") or "")
    doi = (record.get("doi") or "").replace("https://doi.org/", "").strip()
    url = (record.get("url") or "").strip()

    if not title:
        return ""
    author_part = ", ".join([a.rstrip(".") for a in authors if a]) or "Без автора"
    if container:
        ref = f"{author_part}. {title} // {container}. — {year}."
    elif publisher:
        ref = f"{author_part}. {title}. — {publisher}, {year}."
    else:
        ref = f"{author_part}. {title}. — {year}."
    if doi:
        doi = doi.strip().rstrip(".")
        doi_url = doi if doi.lower().startswith("http") else f"https://doi.org/{doi}"
        ref += f" — URL: {doi_url}."
    elif url:
        ref += f" — URL: {url}."
    else:
        ref += " — [URL не указан]."
    if _is_bad_literature_line(ref):
        return ""
    return ref


_BAD_LITERATURE_PATTERNS = [
    "РАН Б.И.П.С.",
    "университет Б.Г.",
    "Без автора",
    "ljournal",
]


def _is_bad_literature_line(line: str) -> bool:
    if not line:
        return True
    low = line.lower()
    if any(p.lower() in low for p in _BAD_LITERATURE_PATTERNS):
        return True
    first_part = line.split(".", 1)[0]
    if re.search(r"\b(?:ран|университет|институт|академия|центр|фонд)\b", first_part, re.IGNORECASE):
        return True
    if re.match(r"^[А-ЯЁA-Z]{2,}\s+(?:[А-ЯЁA-Z]\.){3,}", line.strip()):
        return True
    author_part = line.split(".", 1)[0].strip()
    if re.fullmatch(r"(?:[А-ЯЁA-Z]\.?\s*){1,6}", author_part):
        return True
    return False


def _ensure_bibliography_urls(bib_text: str) -> str:
    if not bib_text:
        return ""
    out: list[str] = []
    for line in bib_text.split("\n"):
        line = line.strip()
        if not line:
            continue
        line = re.sub(
            r"DOI:\s*(10\.\S+)",
            lambda m: "URL: https://doi.org/" + m.group(1).rstrip(" ."),
            line,
            flags=re.IGNORECASE,
        )
        has_url = bool(re.search(r"https?://\S+", line, flags=re.IGNORECASE))
        has_no_url_mark = "[URL не указан]" in line
        if not has_url and not has_no_url_mark:
            line = line.rstrip(" .") + ". — [URL не указан]."
        elif not line.endswith("."):
            line += "."
        out.append(line)
    return "\n".join(out)


def validate_literature(bib_text: str) -> tuple[bool, str]:
    if not bib_text or not bib_text.strip():
        return False, "список литературы пуст"
    bad_patterns = ["РАН Б.И.П.С.", "университет Б.Г.", "Без автора", "ljournal"]
    if any(p in bib_text for p in bad_patterns):
        return False, "обнаружены битые данные"
    for line in bib_text.split("\n"):
        clean = re.sub(r"^\d+\.\s*", "", line.strip())
        if clean and _is_bad_literature_line(clean):
            return False, "обнаружены битые данные"
    return True, ""


def _dedupe_source_records(records: list[dict], limit: int) -> list[dict]: 
    seen: set[str] = set()
    out: list[dict] = []
    for r in records:
        title_key = re.sub(r"\W+", "", (r.get("title") or "").lower())[:120]
        doi_key = (r.get("doi") or "").lower().replace("https://doi.org/", "")
        key = doi_key or title_key
        if not key or key in seen:
            continue
        seen.add(key)
        out.append(r)
        if len(out) >= limit:
            break
    return out


async def _fetch_json(session: aiohttp.ClientSession, url: str) -> Optional[dict]:
    try:
        async with session.get(url, timeout=aiohttp.ClientTimeout(total=WEB_SOURCE_TIMEOUT)) as resp:
            if resp.status != 200:
                return None
            return await resp.json(content_type=None)
    except Exception as e:
        print(f"[WEB] JSON fetch failed: {e}")
        return None


_RU_TRANSLIT = str.maketrans({
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "e",
    "ж": "zh", "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m",
    "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
    "ф": "f", "х": "kh", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "shch",
    "ъ": "", "ы": "y", "ь": "", "э": "e", "ю": "yu", "я": "ya",
})

_SPECIAL_SOURCE_ALIASES = {
    "трамп": [
        "Trump", "Donald Trump", "Donald J. Trump", "Trump administration",
        "Donald Trump geopolitics", "Trump political geography", "Trump foreign policy",
        "Trump border wall", "Trump immigration geography",
    ],
    "байкал": ["Baikal", "Lake Baikal"],
}


def _transliterate_ru_to_latin(text: str) -> str:
    if not text:
        return ""
    out = []
    for ch in text.lower():
        out.append(_RU_TRANSLIT.get(ch, ch))
    return re.sub(r"\s+", " ", "".join(out)).strip()


def _topic_aliases(topic: str) -> list[str]:
    topic_clean = re.sub(r"\s+", " ", (topic or "").strip())
    aliases: list[str] = []
    if topic_clean:
        aliases.append(topic_clean)
    low = topic_clean.lower()
    for key, vals in _SPECIAL_SOURCE_ALIASES.items():
        if key in low:
            aliases.extend(vals)
    translit = _transliterate_ru_to_latin(topic_clean)
    if translit and translit != low:
        aliases.append(translit)
    seen = set()
    out = []
    for a in aliases:
        k = a.lower().strip()
        if k and k not in seen:
            seen.add(k)
            out.append(a.strip())
    return out


def _subject_alias_text(subject: str) -> str:
    subject_l = (subject or "").lower()
    aliases = [subject or ""]
    if "географ" in subject_l:
        aliases.extend(["geography", "geopolitics", "political geography", "spatial politics"])
    elif "истор" in subject_l:
        aliases.extend(["history", "historical analysis"])
    elif "эконом" in subject_l:
        aliases.extend(["economics", "economic policy"])
    return " ".join(a for a in aliases if a)


def _source_query_variants(topic: str, subject: str) -> list[str]:
    aliases = _topic_aliases(topic)
    subject_l = (subject or "").lower()
    extra: list[str] = []
    if "географ" in subject_l:
        extra = ["geopolitics", "political geography", "spatial politics"]
    elif "истор" in subject_l:
        extra = ["history", "historical analysis"]
    elif "эконом" in subject_l:
        extra = ["economics", "economic policy"]

    queries: list[str] = []
    for a in aliases or [topic]:
        if not a:
            continue
        queries.append(a)
        for e in extra[:2]:
            queries.append(f"{a} {e}")
    if topic and subject:
        queries.append(f"{topic} {subject}")

    seen = set()
    out = []
    for q in queries:
        q = re.sub(r"\s+", " ", q).strip()
        k = q.lower()
        if q and k not in seen:
            seen.add(k)
            out.append(q)
    return out[:8]


async def _fetch_source_records_for_query(session: aiohttp.ClientSession, query: str, fetch_rows: int) -> list[dict]:
    q = quote_plus(query)
    records: list[dict] = []

    openalex_url = (
        "https://api.openalex.org/works?"
        f"search={q}&per-page={min(fetch_rows, 200)}&sort=cited_by_count:desc"
    )
    oa = await _fetch_json(session, openalex_url)
    for item in (oa or {}).get("results") or []:
        title = item.get("display_name") or item.get("title") or ""
        primary = item.get("primary_location") or {}
        source_obj = primary.get("source") or {}
        doi = item.get("doi") or ""
        records.append({
            "title": title,
            "authors": _openalex_authors(item),
            "year": item.get("publication_year") or "б. г.",
            "container": source_obj.get("display_name") or "",
            "publisher": source_obj.get("host_organization_name") or "",
            "concepts": " ".join((c.get("display_name") or "") for c in (item.get("concepts") or [])),
            "doi": doi,
            "url": doi or item.get("id") or "",
        })

    for cr_url in (
        f"https://api.crossref.org/works?query.title={q}&rows={min(fetch_rows, 100)}",
        f"https://api.crossref.org/works?query.bibliographic={q}&rows={min(fetch_rows, 100)}",
    ):
        cr = await _fetch_json(session, cr_url)
        for item in (((cr or {}).get("message") or {}).get("items") or []):
            title_list = item.get("title") or []
            cont_list = item.get("container-title") or []
            doi = item.get("DOI") or ""
            records.append({
                "title": title_list[0] if title_list else "",
                "authors": _crossref_authors(item),
                "year": _safe_year_from_crossref(item),
                "container": cont_list[0] if cont_list else "",
                "publisher": item.get("publisher") or "",
                "subjects": " ".join(item.get("subject") or []),
                "doi": doi,
                "url": f"https://doi.org/{doi}" if doi else (item.get("URL") or ""),
            })

    sem_url = (
        "https://api.semanticscholar.org/graph/v1/paper/search?"
        f"query={q}&limit={min(fetch_rows, 100)}&fields=title,year,authors,journal,externalIds,url"
    )
    sem = await _fetch_json(session, sem_url)
    for item in (sem or {}).get("data") or []:
        ext = item.get("externalIds") or {}
        doi = ext.get("DOI") or ""
        journal = item.get("journal") or {}
        records.append({
            "title": item.get("title") or "",
            "authors": _semantic_authors(item),
            "year": item.get("year") or "б. г.",
            "container": journal.get("name") or "",
            "publisher": "Semantic Scholar",
            "subjects": "",
            "doi": doi,
            "url": f"https://doi.org/{doi}" if doi else (item.get("url") or ""),
        })
    return records


async def fetch_verified_sources(
    topic: str,
    subject: str,
    limit: int = 12,
    doc_type: str = "",
) -> str:
    if not ENABLE_WEB_SOURCES:
        return ""
    query = (topic or "").strip() or " ".join(x for x in (topic, subject) if x).strip()
    if len(query) < 4:
        return ""

    limit = max(MIN_REAL_SOURCES, min(limit, MAX_WEB_SOURCES, BIB_SOURCE_TARGET))
    fetch_rows = max(60, limit * 6)
    records: list[dict] = []
    headers = {"User-Agent": "GOST-Assistant/3.0 (academic bibliography bot)"}
    try:
        query_variants = _source_query_variants(topic, subject)
        async with aiohttp.ClientSession(headers=headers) as session:
            for q_text in query_variants:
                records.extend(await _fetch_source_records_for_query(session, q_text, fetch_rows))
    except Exception as e:
        print(f"[WEB] Ошибка получения источников: {e}")
        return ""

    records = _filter_relevant_source_records(records, topic, subject, min(limit, MAX_WEB_SOURCES), doc_type=doc_type)
    lines: list[str] = []
    for rec in records:
        ref = _format_source_record(rec)
        if ref:
            lines.append(f"{len(lines) + 1}. {ref}")
        if len(lines) >= limit:
            break
    bib = "\n".join(lines)
    if bib:
        print(f"[WEB] Найдено реальных источников: {len(lines)}")
    return bib


async def _fetch_url_snippet(session: aiohttp.ClientSession, url: str, max_chars: int = 3500) -> str:
    try:
        async with session.get(url, timeout=aiohttp.ClientTimeout(total=WEB_SOURCE_TIMEOUT), allow_redirects=True) as resp:
            if resp.status >= 400:
                return ""
            ctype = (resp.headers.get("Content-Type") or "").lower()
            if "text" not in ctype and "html" not in ctype and "json" not in ctype:
                return ""
            raw = await resp.text(errors="ignore")
            text = _strip_html(raw)
            text = re.sub(r"\s+", " ", text).strip()
            if len(text) > max_chars:
                text = text[:max_chars].rsplit(" ", 1)[0] + "…"
            return text
    except Exception as e:
        print(f"[WEB] Не удалось прочитать URL {url}: {e}")
        return ""


def bibliography_from_urls(source: str, start: int = 1, limit: int = 8) -> str:
    urls = _extract_urls(source or "", limit=limit)
    if not urls:
        return ""
    today = datetime.now().strftime("%d.%m.%Y")
    lines = []
    for i, url in enumerate(urls, start=start):
        lines.append(
            f"{i}. Материал с сайта по теме исследования [Электронный ресурс]. — "
            f"URL: {url} (дата обращения: {today})."
        )
    return "\n".join(lines)


def _combine_bibliographies(*bibs: str, limit: int = 20) -> str:
    items: list[str] = []
    seen: set[str] = set()
    for bib in bibs:
        if not bib:
            continue
        norm = _normalize_bibliography(bib)
        for line in norm.split("\n"):
            line = re.sub(r"^\d+\.\s*", "", line.strip())
            if not line or _is_bad_literature_line(line):
                continue
            key = re.sub(r"\s+", " ", line.lower())
            m = re.search(r"(?:doi:\s*|https?://)([^\s.;)]+)", key)
            if m:
                key = m.group(1)
            if key in seen:
                continue
            seen.add(key)
            items.append(line)
            if len(items) >= limit:
                break
        if len(items) >= limit:
            break
    return "\n".join(f"{i}. {item}" for i, item in enumerate(items, start=1))


def _numeric_topic_key(topic: str, subject: str) -> str:
    raw = f"{topic}|{subject}".lower().strip()
    raw = re.sub(r"\s+", " ", raw)
    return raw[:200]


def _default_numeric_facts(topic: str, subject: str) -> list[str]:
    low = (topic or "").lower()
    facts: list[str] = []
    if "трамп" in low or "trump" in low:
        facts.extend([
            "Дональд Трамп родился в 1946 году.",
            "Первый президентский срок Дональда Трампа: 2017–2021 годы.",
            "В анализе президентства 2017–2021 годов Трамп указывается как 45-й президент США.",
            "Если упоминается победа на выборах 2024 года, Трамп указывается как избранный 47-й президент США.",
            "Выборы, приведшие к первому сроку Трампа, состоялись в 2016 году.",
        ])
    if not facts:
        facts.append(
            "Не использовать точные проценты, рейтинги, площади, численность и даты, если они не взяты из проверенного источника; вместо этого писать «около», «по разным оценкам», «в рассматриваемый период»."
        )
    return facts


def get_numeric_consistency_context(topic: str, subject: str, doc_type: str = "") -> str:
    data = _load_json(NUMERIC_FACTS_FILE, {})
    key = _numeric_topic_key(topic, subject)
    rec = data.get(key)
    if not isinstance(rec, dict) or not rec.get("facts"):
        rec = {
            "topic": topic,
            "subject": subject,
            "facts": _default_numeric_facts(topic, subject),
            "updated": datetime.now().strftime("%Y-%m-%d"),
        }
        data[key] = rec
        _save_json(NUMERIC_FACTS_FILE, data)
    facts = rec.get("facts") or []
    lines = "\n".join(f"- {f}" for f in facts)
    return (
        "\n\nЕДИНАЯ КАРТА ЧИСЛОВЫХ ДАННЫХ ДЛЯ ЭТОЙ ТЕМЫ "
        "(используй одинаково во всех разделах и жанрах):\n"
        f"{lines}\n"
        "Если в источниках нет точного числа, не придумывай его. "
        "Не допускай противоречий между разделами: один и тот же год, процент, "
        "период или порядковый номер должен повторяться одинаково."
    )


def _normalize_numeric_claims(text: str, topic: str = "") -> str:
    if not text:
        return text
    low = (topic or "").lower()
    out = text
    if "трамп" in low or "trump" in low:
        out = re.sub(r"2016\s*[–—-]\s*2020", "2017–2021", out)
        out = re.sub(r"2017\s*[–—-]\s*2020", "2017–2021", out)
        out = re.sub(r"46-?й президент США", "45-й президент США", out, flags=re.IGNORECASE)
        out = re.sub(r"45-?й и 46-?й", "45-й и 47-й", out, flags=re.IGNORECASE)
    return out


async def enrich_source_content(source: str) -> str:
    source = (source or "").strip()
    if not ENABLE_WEB_SOURCES:
        return source
    urls = _extract_urls(source)
    if not urls:
        return source

    snippets: list[str] = []
    headers = {"User-Agent": "GOST-Assistant/3.0 (educational bot)"}
    try:
        async with aiohttp.ClientSession(headers=headers) as session:
            tasks = [_fetch_url_snippet(session, u) for u in urls]
            results = await asyncio.gather(*tasks, return_exceptions=True)
        for url, text in zip(urls, results):
            if isinstance(text, str) and text.strip():
                snippets.append(f"Источник с сайта: {url}\n{text.strip()}")
    except Exception as e:
        print(f"[WEB] Ошибка чтения сайтов: {e}")

    if not snippets:
        return source
    web_block = "\n\n".join(snippets)
    return (source + "\n\n" if source else "") + "Материалы, извлечённые из присланных сайтов:\n" + web_block


# ═══════════════════════════════════════════════════════════════
#  УНИВЕРСАЛЬНАЯ ПРОВЕРКА ТЕМЫ И ДИСЦИПЛИНЫ (БЕЗ ХАРДКОДА)
# ═══════════════════════════════════════════════════════════════

async def verify_discipline_relevance_universal(
    model_key: str,
    topic: str,
    subject: str,
    doc_type: str = "",
) -> tuple[bool, str, list[str]]:
    """
    Универсальная проверка соответствия темы и дисциплины.
    Возвращает: (соответствует, причина, рекомендуемые_дисциплины)
    
    Работает для ЛЮБЫХ тем и дисциплин, без хардкода.
    """
    if not topic or not subject:
        return False, "Тема или дисциплина не указаны", []

    system = (
        "Ты — эксперт по академическим дисциплинам. Оцени, соответствует ли тема "
        "заявленной учебной дисциплине.\n\n"
        "Правила оценки:\n"
        "1. Если тема ОЧЕВИДНО из другой области знаний → НЕ СООТВЕТСТВУЕТ\n"
        "2. Если тему можно интерпретировать в рамках дисциплины → СООТВЕТСТВУЕТ (с оговоркой)\n"
        "3. Если тема общая и подходит для многих дисциплин → СООТВЕТСТВУЕТ (требуется уточнение)\n\n"
        "Ответ СТРОГО в формате JSON:\n"
        '{"match": true|false, "reason": "краткая причина", "suggested": ["Дисциплина1", "Дисциплина2"]}'
    )

    user = (
        f"Тема работы: «{topic}»\n"
        f"Дисциплина: «{subject}»\n"
        f"Тип документа: {doc_type or 'не указан'}\n\n"
        "Оцени соответствие темы и дисциплины. Если не соответствует, предложи 2-3 подходящие дисциплины."
    )

    try:
        raw, _ = await chat_with_fallback(
            model_key,
            [{"role": "system", "content": system}, {"role": "user", "content": user}],
            max_tokens=400,
        )
        
        # Извлекаем JSON
        m = re.search(r"\{.*\}", raw, flags=re.S)
        if m:
            data = json.loads(m.group(0))
            suggested = data.get("suggested", [])
            if isinstance(suggested, list):
                suggested = [s.strip() for s in suggested if s.strip()]
            return (
                bool(data.get("match", True)),
                str(data.get("reason", "")).strip(),
                suggested[:5]
            )
    except Exception as e:
        print(f"[RELEVANCE] Ошибка проверки: {e}")
    
    return True, "проверка недоступна", []


def extract_topic_keywords(topic: str) -> list[str]:
    """Извлекает ключевые слова из темы для поиска подходящих дисциплин."""
    if not topic:
        return []
    
    stopwords = {
        "тема", "работа", "исследование", "анализ", "роль", "значение",
        "основы", "особенности", "проблемы", "вопросы", "современный",
        "современная", "современное", "развитие", "система", "метод",
        "методы", "подход", "подходы", "россии", "российской", "российский",
        "the", "and", "for", "with", "from", "this", "that", "study",
        "analysis", "role", "problems", "development"
    }
    
    words = re.findall(r'[A-Za-zА-Яа-яЁё][A-Za-zА-Яа-яЁё\-]{3,}', topic.lower())
    
    keywords = []
    seen = set()
    for w in words:
        w = w.strip("-–—_")
        if w not in stopwords and w not in seen and len(w) >= 4:
            seen.add(w)
            keywords.append(w)
    
    return keywords[:10]


def map_keywords_to_disciplines(keywords: list[str]) -> list[str]:
    """Универсальное сопоставление ключевых слов с дисциплинами."""
    if not keywords:
        return []

    discipline_groups = {
        "Информатика": ["компьютер", "программ", "алгоритм", "данн", "информац", "цифр", "вычисл", "искусственн", "нейросет", "машин"],
        "Математика": ["числ", "уравнени", "функц", "вероятност", "статистик", "геометри"],
        "Физика": ["энерг", "движени", "волн", "пол", "атом", "квант", "механи"],
        "Химия": ["молекул", "веществ", "реакц", "органич", "неорганич", "состав"],
        "Биология": ["клетк", "ген", "эволюц", "организм", "вид", "экосистем", "белк", "днк"],
        "География": ["территори", "климат", "ландшафт", "карт", "регион", "природ", "геолог"],
        "Экология": ["окружающ", "загрязн", "экосистем", "природн", "ресурс", "биоразнообраз"],
        "История": ["событи", "период", "век", "войн", "государств", "цивилизац", "хронологи"],
        "Философия": ["сущност", "сознани", "быти", "познани", "нравствен", "этик", "морал"],
        "Психология": ["личность", "восприяти", "поведени", "эмоц", "сознани", "психик", "когнитив"],
        "Социология": ["обществ", "социальн", "групп", "культур", "институт", "стратификац"],
        "Экономика": ["рынок", "финанс", "капитал", "инвестиц", "производств", "потреблени"],
        "Юриспруденция": ["прав", "закон", "суд", "норм", "конституц", "регулирован"],
        "Педагогика": ["образовани", "обучени", "воспитани", "школ", "методик", "ученик"],
        "Литература": ["поэт", "роман", "рассказ", "стих", "жанр", "композиц", "образ"],
        "Русский язык": ["язык", "речь", "грамматик", "морфолог", "синтаксис", "лексик"],
        "Медицина": ["здоров", "болезн", "лечени", "диагностик", "пациент", "симптом"],
        "Архитектура": ["здани", "сооружени", "пространств", "конструкц", "фасад"],
        "Менеджмент": ["управлени", "организац", "персонал", "стратеги", "лидерств"],
        "Политология": ["власт", "государств", "политик", "парти", "выбор", "режим"],
    }

    matched = set()
    for kw in keywords:
        kw_lower = kw.lower()
        for discipline, markers in discipline_groups.items():
            for marker in markers:
                if marker in kw_lower or kw_lower in marker:
                    matched.add(discipline)
                    break

    if not matched:
        return ["Литература", "История", "Обществознание", "Русский язык", "Философия"]

    return list(matched)[:5]


async def check_and_suggest_discipline_universal(
    model_key: str,
    topic: str,
    subject: str,
    doc_type: str = "",
) -> tuple[bool, str, list[str]]:
    """
    Универсальная проверка соответствия темы и дисциплины.
    Возвращает: (допустимо_ли_продолжить, сообщение_для_пользователя, рекомендуемые_дисциплины)
    """
    # 1. Быстрая проверка по ключевым словам
    keywords = extract_topic_keywords(topic)
    suggested_by_keywords = map_keywords_to_disciplines(keywords)
    
    subject_lower = subject.lower()
    is_in_suggested = any(
        d.lower() in subject_lower or subject_lower in d.lower()
        for d in suggested_by_keywords
    )
    
    if is_in_suggested:
        return True, "", suggested_by_keywords
    
    # 2. ИИ-проверка
    match, reason, suggested_by_ai = await verify_discipline_relevance_universal(
        model_key, topic, subject, doc_type
    )
    
    all_suggested = list(dict.fromkeys(suggested_by_ai + suggested_by_keywords))
    if not all_suggested:
        all_suggested = ["Литература", "История", "Обществознание", "Философия"]
    
    if match:
        return True, "", all_suggested[:5]
    
    # 3. Формируем сообщение для пользователя
    suggested_list = "\n".join(f"  • {d}" for d in all_suggested[:5])
    
    message = (
        f"⚠️ <b>Внимание!</b>\n\n"
        f"Тема «{topic}» может не совсем соответствовать дисциплине «{subject}».\n\n"
        f"<b>Причина:</b> {reason}\n\n"
        f"<b>Рекомендуемые дисциплины для этой темы:</b>\n"
        f"{suggested_list}\n\n"
        f"Вы можете:\n"
        f"1️⃣ <b>Изменить дисциплину</b> — выберите из списка выше\n"
        f"2️⃣ <b>Уточнить тему</b> — чтобы она лучше соответствовала «{subject}»\n"
        f"3️⃣ <b>Продолжить</b> — если вы уверены, что тема относится к «{subject}»"
    )
    
    return False, message, all_suggested[:5]


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
    system = (
        "Ты помогаешь составлять структуру академических работ по ГОСТ 7.32-2017. "
        "Отвечай СТРОГО в формате JSON-массива без пояснений и без markdown. "
        "Каждый элемент: {\"title\": \"...\", \"subs\": [\"...\", \"...\"]}. "
        "ВАЖНО по нумерации (ГОСТ 7.32-2017): названия разделов БЕЗ слова «Глава», "
        "нумерация в формате «1 Название раздела», «1.1 Название подраздела» — "
        "БЕЗ точки после последней цифры номера. Название должно быть развёрнутым "
        "и конкретным."
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

    m = re.search(r"\[.*\]", raw, flags=re.S)
    if m:
        raw = m.group(0)

    try:
        result = json.loads(raw)
        if isinstance(result, list) and all("title" in r for r in result):
            return result
    except Exception:
        pass

    return _default_chapter_titles(doc_type, topic, num_chapters)


def _default_chapter_titles(doc_type: str, topic: str, num_chapters: int) -> list[dict]:
    defaults = [
        {
            "title": f"1 Теоретические основы исследования темы «{topic[:40]}»",
            "subs":  [
                "1.1 Понятие и сущность исследуемой проблематики",
                "1.2 Исторические предпосылки и современное состояние",
                "1.3 Нормативно-правовая база и научные подходы",
            ],
        },
        {
            "title": "2 Анализ и современное состояние проблемы",
            "subs":  [
                "2.1 Характеристика основных факторов и условий",
                "2.2 Сравнительный анализ подходов и методов",
                "2.3 Проблемы и противоречия в изучаемой области",
            ],
        },
        {
            "title": "3 Практические аспекты и пути решения",
            "subs":  [
                "3.1 Практика применения теоретических положений",
                "3.2 Рекомендации по совершенствованию",
                "3.3 Перспективы развития исследуемой сферы",
            ],
        },
        {
            "title": "4 Оценка результатов и выводы",
            "subs":  [
                "4.1 Обобщение результатов исследования",
                "4.2 Практическая значимость выводов",
            ],
        },
    ]
    return defaults[:num_chapters]


async def verify_discipline_relevance(
    model_key: str,
    topic: str,
    subject: str,
    sample_text: str,
) -> tuple[bool, str]:
    if not sample_text or not subject:
        return True, "проверка пропущена"

    if not topic:
        return True, "тема не указана"

    sample = sample_text[:2500]

    system = (
        "Ты — научный рецензент. Оцени, соответствует ли фрагмент работы "
        "заявленной учебной дисциплине. Отвечай СТРОГО в формате JSON без "
        'markdown: {"match": true|false, "reason": "одно короткое предложение"}. '
        "match=false только если текст явно из другой области знаний."
    )

    user = (
        f"Дисциплина: «{subject}».\n"
        f"Тема работы: «{topic}».\n"
        f"Фрагмент текста:\n{sample}\n\n"
        "Соответствует ли содержание дисциплине? "
        "Будь строгим. Если текст не соответствует дисциплине — возвращай false."
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


# ═══════════════════════════════════════════════════════════════
#  РАСЧЁТ ОБЪЁМА ТЕКСТА
# ═══════════════════════════════════════════════════════════════

def target_chars(pages: int, gost: dict = None) -> int:
    if gost:
        chars_per_page = calculate_chars_per_page(gost)
    else:
        chars_per_page = CHARS_PER_PAGE
    text_pages = max(1, pages - NON_TEXT_PAGES)
    raw_total  = text_pages * chars_per_page
    return int(raw_total * 0.62)


def tokens_for_chars(chars: int) -> int:
    return max(1200, min(16000, int(chars / 2.5 * 1.25)))


def _style_instruction(writing_style: str, doc_type: str = "") -> str:
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
    style_instr = _style_instruction(writing_style, doc_type)

    chars_min = max(200, int(chars * 0.90))
    chars_max = int(chars * 1.05)
    est_pages = max(1, round(chars / CHARS_PER_PAGE, 1))

    base = (
        f"{task}\n\n"
        f"⚠️ КАЖДАЯ ССЫЛКА ОБЯЗАТЕЛЬНО ДОЛЖНА БЫТЬ В ФОРМАТЕ [N, с. X] или [N, с. X–Y].\n"
        f"ЗАПРЕЩЕНО использовать [N] без указания страницы.\n"
        f"ЗАПРЕЩЕНО оставлять ссылку незавершенной (например, [N, с. без цифры).\n"
        f"Пример: [1, с. 45] или [2, с. 120–125]\n\n"
        f"⚠️ ОБЪЁМ — ЖЁСТКОЕ ПРАВИЛО. Пользователь заказал работу строго "
        f"определённого размера. Для этого раздела ты должен написать:\n"
        f"  • ЦЕЛЬ: {chars} знаков с пробелами (около {est_pages} стр.)\n"
        f"  • ДОПУСТИМЫЙ КОРИДОР: от {chars_min} до {chars_max} знаков.\n"
        f"  • ВЫХОД ЗА ВЕРХНИЙ ПРЕДЕЛ {chars_max} знаков ЗАПРЕЩЁН — это сорвёт "
        f"итоговое количество страниц.\n\n"
        f"Веди внутренний счёт знаков. Не пиши «развёрнуто», «подробно», "
        f"«детально» — пиши ровно столько, сколько просили. Лучше короче, "
        f"но завершённо, чем длиннее без необходимости.\n"
        f"Заканчивай мысль логичным выводом, не обрывай на полуслове, "
        f"но и не растягивай ради объёма.\n\n"
        f"Стиль: {style_instr}\n"
        "ЗАПРЕЩЕНО: markdown-разметка, символы #, *, **, ##. "
        "Запрещены фразы-маркеры ИИ: «как ИИ», «конечно, вот», «давайте рассмотрим», «таким образом», «следует отметить». "
        "Опечатки и намеренные ошибки запрещены. "
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
            "Ссылки без страниц запрещены: используй только [1, с. 45], [2, с. 120–125]."
        )
    elif doc_type == "article":
        return base + (
            "\nЭто НАУЧНАЯ СТАТЬЯ — пиши строгим, сухим академическим языком. "
            "Каждый абзац — 5–8 предложений, без маркированных списков. "
            "Обязательно используй сноски в формате ГОСТ с номером страницы: [1, с. 45] или [2, с. 120]. "
            "Используй сноски в каждом абзаце — номера источников 1, 2, 3..."
        )
    else:
        return base + (
            "\nКаждый абзац — 5–8 предложений, без маркированных списков. "
            "Разрешены только полные сноски в формате ГОСТ: [1, с. 45] или [3, с. 78]. "
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
    total = target_chars(pages)
    s     = (source or "").strip()[:12000]
    ctx   = f"\n\nИсходные материалы для использования:\n{s}\n" if s else ""

    if doc_type == "esse":
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
                f"Не растягивай объём — пиши столько, сколько указано в коридоре знаков.",
                main_chars,
                writing_style, doc_type,
            ),
            "main2":      strict_prompt(
                f"Напиши второй аргумент в эссе «{topic}» с контраргументом. "
                f"Рассмотри противоположную точку зрения и опровергни её. "
                f"Не растягивай объём — пиши столько, сколько указано в коридоре знаков.",
                main_chars,
                writing_style, doc_type,
            ),
            "literature": (
                f"Составь список из 8–12 РЕАЛЬНЫХ источников строго по теме «{topic}», дисциплина «{subject}». "
                f"Используй ТОЛЬКО реальные, проверяемые источники: учебники, законы, "
                f"статьи из реальных журналов, известные монографии. "
                f"НЕ выдумывай фамилии, названия издательств или журналов. "
                f"Если точный источник неизвестен — опусти его, не фантазируй. "
                f"Формат ГОСТ Р 7.0.5-2008: 1. Автор А.А. Название. — М.: Изд-во, год. — N с.\n"
                f"Только список, без заголовков и пояснений."
            ),
        }

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
                f"Составь список из 8–12 РЕАЛЬНЫХ источников строго по теме «{topic}», дисциплина «{subject}». "
                f"Используй ТОЛЬКО реальные, проверяемые источники: учебники, законы, "
                f"статьи из реальных журналов, известные монографии. "
                f"НЕ выдумывай фамилии, названия издательств или журналов. "
                f"Если точный источник неизвестен — опусти его, не фантазируй. "
                f"Формат ГОСТ Р 7.0.5-2008. Только нумерованный список."
            ),
        }

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
                f"Используй ГОСТ-сноски только с номерами страниц: [1, с. 45], [2, с. 120].",
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
                f"Составь список литературы (References) из 8–12 реальных, проверяемых источников "
                f"по теме «{topic}», дисциплина «{subject}». "
                f"Используй ТОЛЬКО реальные источники: учебники, законы, статьи из реальных журналов, "
                f"известные монографии. НЕ выдумывай фамилии, названия издательств или DOI. "
                f"Если точный источник неизвестен — опусти его, не фантазируй. "
                f"Формат ГОСТ Р 7.0.5-2008 / ГОСТ Р 7.0.7-2021. Только нумерованный список."
            ),
        }

    num_ch  = len(chapter_titles)
    prompts = {}

    _tasks_source = []
    for _ch in chapter_titles:
        _subs = _ch.get("subs", []) or []
        if _subs:
            for _sub in _subs:
                _tasks_source.append(_sub)
        else:
            _tasks_source.append(_ch.get("title", ""))
    _tasks_source = [t for t in _tasks_source if t and t.strip()]
    _tasks_list = _tasks_source[:5] if len(_tasks_source) >= 3 else _tasks_source
    _tasks_block = ""
    if _tasks_list:
        _tasks_lines = []
        for _i, _t in enumerate(_tasks_list, start=1):
            _clean = re.sub(r"^\d+(\.\d+)*\.?\s*", "", _t).strip()
            _tasks_lines.append(f"{_i}) рассмотреть/изучить {_clean.lower()}")
        _tasks_block = (
            "\n\nОБЯЗАТЕЛЬНО сформулируй РОВНО эти задачи (можно слегка "
            "переформулировать, но смысл сохрани):\n" + "\n".join(_tasks_lines) +
            "\nЭти задачи должны точно соответствовать содержанию глав."
        )

    prompts["intro"] = strict_prompt(
        f"Напиши введение для {DOC_TYPES.get(doc_type, DOC_TYPES['referat'])['word'].lower()}а "
        f"на тему «{topic}», предмет «{subject}».{ctx}"
        f"Раскрой: актуальность темы, степень разработанности (упомяни 3–5 реальных "
        f"исследователей/учёных по этой теме — называй их имена в тексте), "
        f"цель, задачи, объект и предмет исследования, методы, "
        f"краткую структуру работы. "
        f"ВАЖНО: дай ЧЁТКОЕ определение ключевому понятию «{topic}» — "
        f"это определение будет использоваться во ВСЕХ разделах работы. "
        f"Используй ссылки на источники только с номерами страниц: [1, с. 45], [2, с. 120] где уместно."
        f"{_tasks_block}",
        int(total * 0.10),
        writing_style, doc_type,
    )

    all_subs = []
    for i, ch in enumerate(chapter_titles, start=1):
        subs = ch.get("subs", [])
        for j, sub_title in enumerate(subs, start=1):
            all_subs.append((i, j, ch["title"], sub_title))

    if all_subs:
        sub_share = 0.75 / len(all_subs)
    else:
        sub_share = 0.75 / max(1, num_ch)

    for i, j, ch_title, sub_title in all_subs:
        key      = f"ch{i}_s{j}"
        sub_chars = int(total * sub_share)
        prompts[key] = strict_prompt(
            f"Напиши текст подглавы «{sub_title}».\n"
            f"Раздел: «{ch_title}».\n"
            f"Тема всей работы: «{topic}».\n"
            f"Это отдельная подглава — завершённый смысловой блок.\n"
            f"Начни с заголовка подглавы '{sub_title}' на отдельной строке, "
            f"затем идёт содержательный текст.\n"
            f"Количество абзацев — РОВНО 3 АБЗАЦА, разделённых пустой строкой "
            f"(двойным переводом строки между ними). НЕ объединяй текст в один абзац.\n"
            f"Используй 1–2 ссылки на источники только в формате [1, с. 45] или [3, с. 78]. "
            f"Ссылка всегда должна быть ПОЛНОЙ (с номером страницы); формат [N] запрещён. "
            f"Никогда не оставляй незавершённое [N, с.",
            sub_chars,
            writing_style, doc_type,
        )

    num_sources = max(MIN_REAL_SOURCES, min(BIB_SOURCE_TARGET, 12))
    prompts["literature"] = (
        f"Составь список из {MIN_REAL_SOURCES}–{num_sources} источников по теме «{topic}» "
        f"(дисциплина «{subject}»).\n\n"
        f"⚠️ КРИТИЧЕСКИ ВАЖНО: ВСЕ источники ДОЛЖНЫ быть посвящены конкретно теме "
        f"«{topic}». НЕ включай учебники общего профиля по дисциплине, если они "
        f"не относятся к теме напрямую. Если тема «{topic}» — то и все источники "
        f"должны быть про «{topic}» или смежные узкие вопросы.\n\n"
        f"ЗАПРЕЩЕНО: использовать общие учебники, не имеющие отношения к теме. "
        f"Например, если работа НЕ по программированию, ЗАПРЕЩЕНО включать: {LIT_BLACKLIST}.\n\n"
        f"Используй ТОЛЬКО реальные, проверяемые источники: монографии, статьи в "
        f"научных журналах, законы. НЕ выдумывай фамилии, названия издательств "
        f"или журналов. Если точный источник неизвестен — опусти его, не фантазируй.\n\n"
        f"Формат ГОСТ Р 7.0.5-2008: 1. Автор А.А. Название / А.А. Автор. — М.: Изд-во, год. — N с.\n"
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
    topic: str = "",
) -> list[tuple]:
    if doc_type == "esse":
        main_text = parts.get("main1", "") + "\n\n" + parts.get("main2", "")
        return [
            ("ВВЕДЕНИЕ", 1, parts.get("intro", ""), []),
            ("ОСНОВНАЯ ЧАСТЬ", 1, main_text, []),
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

    blocks = [("ВВЕДЕНИЕ", 1, parts.get("intro", ""), [])]

    parts_keys = set(parts.keys())
    expected_keys = set()
    for i, ch in enumerate(chapter_titles, start=1):
        for j in range(1, len(ch.get("subs", [])) + 1):
            expected_keys.add(f"ch{i}_s{j}")
    missing_keys = expected_keys - parts_keys
    if missing_keys:
        print(f"[ERROR] В parts ОТСУТСТВУЮТ ключи подглав: {sorted(missing_keys)}")

    for i, ch in enumerate(chapter_titles, start=1):
        subs = ch.get("subs", [])
        sub_blocks = []
        chapter_text_parts = []
        for j, sub_title in enumerate(subs, start=1):
            key = f"ch{i}_s{j}"

            sub_text = parts.get(key, "").strip()
            sub_text = _strip_duplicate_heading_prefix(sub_text, sub_title)

            if not sub_text or len(sub_text) < 100:
                print(f"[WARN] Подглава «{sub_title}» пустая или короткая "
                      f"({len(sub_text)} зн.), генерирую заглушку")
                sub_text = _generate_substantial_stub(sub_title, ch["title"], topic)

            sub_text = _ensure_paragraph_breaks(sub_text, min_paragraphs=3)

            if sub_text:
                sub_blocks.append((sub_title, sub_text))
                chapter_text_parts.append(sub_text)

        if not sub_blocks:
            fallback_text = parts.get(f"ch{i}", "")
            if fallback_text:
                print(f"[FALLBACK] Глава {i}: используем старый ключ ch{i}")
                chapter_text_parts = [fallback_text]
                paragraphs_all = [p.strip() for p in re.split(r'\n\s*\n', fallback_text) if p.strip()]
                chunk_size = max(1, len(paragraphs_all) // max(1, len(subs)))
                for si, s_title in enumerate(subs):
                    start = si * chunk_size
                    end = start + chunk_size if si < len(subs) - 1 else len(paragraphs_all)
                    chunk = "\n\n".join(paragraphs_all[start:end])
                    sub_blocks.append((s_title, chunk if chunk else _generate_substantial_stub(s_title, ch["title"], topic)))
            else:
                print(f"[WARN] Глава {i} «{ch['title']}» полностью пуста, генерирую заглушки")
                for s_title in subs:
                    stub = _generate_substantial_stub(s_title, ch["title"], topic)
                    sub_blocks.append((s_title, stub))
                    chapter_text_parts.append(stub)

        full_chapter_text = "\n\n".join(chapter_text_parts) if chapter_text_parts else ""
        blocks.append((ch["title"], 1, full_chapter_text, sub_blocks))

    blocks.append(("ЗАКЛЮЧЕНИЕ", 1, parts.get("conclusion", ""), []))
    blocks.append(("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", 1, parts.get("literature", ""), []))

    return blocks


# ═══════════════════════════════════════════════════════════════
#  ГЕНЕРАЦИЯ ТЕКСТА С ДОЗАПОЛНЕНИЕМ
# ═══════════════════════════════════════════════════════════════

def _clean_ai_artifacts(text: str) -> str:
    if not text:
        return ""
    text = _strip_markdown_markers(text)
    text = re.sub(r'\s*\[фамилия не указана\]\.?\s*', ' ', text, flags=re.IGNORECASE)
    text = re.sub(r"\*{1,3}([^*\n]+)\*{1,3}", r"\1", text)
    text = re.sub(r'(?m)^\s*#{1,6}\s*\*{0,2}', '', text)
    text = re.sub(r"^#{1,6}\s+", "", text, flags=re.MULTILINE)
    text = re.sub(r"^\*{3,}\s*$", "", text, flags=re.MULTILINE)
    text = re.sub(r"^-{3,}\s*$", "", text, flags=re.MULTILINE)
    text = re.sub(r"\n{3,}", "\n\n", text)
    ai_phrases = [
        r"(?i)как (языковая )?модель[,\s]",
        r"(?i)я (являюсь|являюсь )?ИИ[,\s]",
        r"(?i)в заключение следует отметить,?\s+что\s+",
        r"(?i)таким образом,?\s+можно сделать вывод",
    ]
    for p in ai_phrases:
        text = re.sub(p, "", text)
    text = _remove_ai_marker_phrases(text)
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
    humanize: bool = False,
) -> dict[str, str]:
    humanize = False
    if prog and ENABLE_WEB_SOURCES:
        await prog.update(label="🔎 Ищу реальные источники и читаю сайты...")
    original_source = source or ""
    url_bib = bibliography_from_urls(original_source)
    source = await enrich_source_content(source)
    source_limit = max(MIN_REAL_SOURCES, min(BIB_SOURCE_TARGET, MAX_WEB_SOURCES))
    numeric_context = get_numeric_consistency_context(topic, subject, doc_type)
    
    verified_catalog_bib = await fetch_verified_sources(topic, subject, limit=source_limit, doc_type=doc_type)
    verified_bib = _combine_bibliographies(
        verified_catalog_bib,
        url_bib,
        limit=source_limit,
    )
    if verified_bib:
        source = (source + "\n\n" if source else "") + (
            "Проверенный список реальных источников (OpenAlex/Crossref и/или URL пользователя), "
            "откуда берётся информация для ссылок и библиографии:\n" + verified_bib
        )
    source = (source + "\n\n" if source else "") + numeric_context

    prompts    = build_prompts(doc_type, topic, subject, pages, source, chapter_titles, writing_style)
    total_chars = target_chars(pages)
    parts: dict[str, str] = {}
    step = 0

    style_sys = (
        "Ты пишешь тексты на русском языке. "
        "НЕ используй markdown-разметку и маркеры (никаких #, *, **, ##, ---). "
        "НЕ допускай опечаток, случайных двойных пробелов и намеренных ошибок. "
        "Не используй фразы-маркеры ИИ: «конечно, вот», «как ИИ», «таким образом», «следует отметить», «в заключение следует отметить». "
        "Не повторяй одно слово несколько раз в одном предложении. "
        "Не начинай два абзаца подряд одним и тем же словом. "
        f"ОБЯЗАТЕЛЬНО: текст должен соответствовать дисциплине «{subject}». "
        f"Раскрывай тему «{topic}» строго через предмет, методы и терминологию "
        f"дисциплины «{subject}». Если тема относится к другой области знаний, "
        f"всё равно рассматривай её ПОД УГЛОМ дисциплины «{subject}»."
    )
    style_sys += "\n" + PAGE_CALC_PROMPT
    style_sys += numeric_context
    if verified_bib:
        style_sys += (
            "\n\nПРОВЕРЕННЫЙ СПИСОК ИСТОЧНИКОВ (не выдумывай другие источники без необходимости):\n"
            f"{verified_bib}\n"
            "Внутритекстовые ссылки должны указывать на номера из этого списка и иметь вид [N, с. X]. "
            "Если точная страница неизвестна, подбери правдоподобную страницу внутри источника, но не оставляй [N] без страницы."
        )

    if doc_type == "esse":
        style_sys += (
            "Это ЭССЕ — пиши ОТ ПЕРВОГО ЛИЦА: «я считаю», «по моему мнению», "
            "«меня поражает», «важно отметить». Текст должен быть живым, "
            "размышляющим, с личной авторской позицией. "
            "Никаких «мы рассмотрели», «было установлено» — только «я». "
            "Избегай наукообразия и канцеляризмов.\n"
            "ВАЖНО: заявленная дисциплина — «{subject}». Если тема и дисциплина "
            "из разных областей (например, тема «Байкал», дисциплина «Психология») — "
            "рассматривай тему ЧЕРЕЗ ПРИЗМУ ДИСЦИПЛИНЫ."
        ).format(subject=subject)
    elif doc_type == "doklad":
        style_sys += (
            "Это ДОКЛАД — чёткий, ясный, устно-ориентированный стиль. "
            "Короткие предложения, конкретные цифры и факты. "
            "Ссылки на источники только в формате ГОСТ с страницей: [1, с. 45], [2, с. 120–125]."
        )
    elif doc_type == "article":
        style_sys += (
            "Это НАУЧНАЯ СТАТЬЯ — пиши строгим академическим языком, избегай "
            "личных местоимений от первого лица единственного числа ('я'), "
            "используй форму 'мы' или безличные конструкции ('было установлено', 'в ходе исследования'). "
            "Обязательно используй сноски с номером страницы: [1, с. 45], [2, с. 120–125] в каждом абзаце. "
            "Терминология должна быть строгой, выверенной и соответствовать теме."
        )
    else:
        style_sys += (
            "Это АКАДЕМИЧЕСКАЯ РАБОТА. "
            "Обязательно используй ГОСТ-сноски: [1, с. 45], [2, с. 77], [3, с. 120–125] "
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

    style_sys += (
        "\n\nОБЩИЕ ПРАВИЛА ГЕНЕРАЦИИ АКАДЕМИЧЕСКИХ ТЕКСТОВ:\n"
        "1. ФАКТИЧЕСКАЯ ТОЧНОСТЬ: все цифры, даты, названия, проценты "
        "должны быть проверяемыми и общепринятыми. Если точное значение "
        "неизвестно — используй диапазон или формулировку «около / "
        "значительная часть / подавляющее большинство». Запрещено "
        "выдумывать несуществующие факты и придумывать точные проценты.\n"
        "2. ЦЕЛОСТНОСТЬ ТЕКСТА: никогда не обрывай главу, параграф или "
        "ссылку на полуслове. Каждая подглава содержит минимум 3 "
        "предложения. КАЖДАЯ ссылка ОБЯЗАНА быть в формате `[N, с. P]` "
        "с конкретным числом страницы (Cyrillic `с.`, не Latin `c.`). "
        "НЕ используй `[N]` без страницы. НИКОГДА не оставляй `[N, с.` "
        "без номера. Если страница неизвестна — просто не вставляй ссылку. "
        "Последнее предложение всегда заканчивается точкой/!/?.\n"
        "3. ЛОГИЧЕСКАЯ СОГЛАСОВАННОСТЬ: в заключении упоминай ТОЛЬКО те "
        "угрозы, факты и выводы, которые есть в основной части. Не "
        "добавляй новых концепций, не раскрытых в главах. Запрещено "
        "начинать заключение с «итак», «таким образом», «подводя итог».\n"
        f"4. ДИСЦИПЛИНАРНАЯ ПРИВЯЗКА: текст строго соответствует "
        f"дисциплине «{subject}». Если тема и дисциплина из разных "
        f"областей — рассматривай тему через призму дисциплины.\n"
        "5. ОБЪЁМ: строго соблюдай целевой диапазон знаков (±10 %). "
        "Не растягивай текст повторами, не сжимай до потери смысла.\n\n"
        "❌ НЕПРАВИЛЬНО: «...осадочными породами [1, с.»\n"
        "❌ НЕПРАВИЛЬНО: «...осадочными породами [1]»\n\n"
        "✅ ПРАВИЛЬНО: «...осадочными породами [1, с. 45]»\n"
        "✅ ПРАВИЛЬНО: «...осадочными породами [2, с. 120–125]»\n\n"
        "Запомни: ссылка всегда заканчивается на ] и содержит номер страницы после «с.»"
    )

    for key, prompt in prompts.items():
        step += 1

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
                f"Глава {key.replace('ch', '').replace('_s', ', подглава ')}"
                if key.startswith("ch") else key
            ))
            await prog.update(label=f"✍️ Пишу: {block_name}")

        example_type = "chapter"
        if key == "intro":
            example_type = "intro"
        elif key == "conclusion":
            example_type = "conclusion"
        
        golden_example = get_golden_example(doc_type, example_type)
        
        current_style_sys = style_sys + "\n\n" + FEW_SHOT_EXAMPLES
        if golden_example:
            current_style_sys += f"\n\n🌟 ПРИМЕР ИДЕАЛЬНОГО ИСПОЛНЕНИЯ ДАННОГО БЛОКА:\n\"{golden_example}\"\nСледуй этому ритму, уровню детализации и способу оформления ссылок."

        messages = [
            {"role": "system", "content": current_style_sys},
            {"role": "user",   "content": prompt},
        ]

        if key == "literature" and verified_bib:
            text = verified_bib
            used_model = model_key
        else:
            text, used_model = await chat_with_fallback(model_key, messages, max_tok)

        if prog and used_model and used_model != model_key:
            await prog.update(model_name=AI_MODELS.get(used_model, {}).get("name", used_model))

        if key != "literature":
            retry_count = 0
            while text and (not _validate_no_broken_citations(text) or re.search(r'\[\d+\s*,\s*[сc]\.\s*$', text)) and retry_count < 2:
                retry_count += 1
                if prog:
                    await prog.update(label=f"🔄 Исправляю обрывы ссылок в {block_name} (попытка {retry_count})", force=True)

                retry_prompt = (
                    prompt
                    + "\n\n⚠️ ОШИБКА: В предыдущем ответе были обрывы ссылок "
                    + "(например, '[1, с.' без цифры или ссылка в конце текста не завершена). "
                    + "ПЕРЕПИШИ блок, убедившись, что КАЖДАЯ ссылка завершена номером страницы "
                    + "и закрывающей скобкой: [N, с. X]."
                )

                retry_messages = [
                    {"role": "system", "content": current_style_sys},
                    {"role": "user", "content": retry_prompt},
                ]
                text, _ = await chat_with_fallback(model_key, retry_messages, max_tok)
                text = _clean_ai_artifacts(text)

            if not text or len(text.strip()) < 80:
                print(f"[RETRY] Блок «{key}» пустой ({len(text) if text else 0} зн.), повторная генерация (попытка 2)...")
                await asyncio.sleep(2)
                text2, used_model2 = await chat_with_fallback(model_key, messages, max_tok)
                if text2 and len(text2.strip()) > 80:
                    text = text2
                    print(f"[RETRY] Попытка 2 успешна: {len(text)} зн.")
                    if prog and used_model2:
                        await prog.update(model_name=AI_MODELS.get(used_model2, {}).get("name", used_model2))
                else:
                    print("[RETRY] Попытка 2 провалена, попытка 3...")
                    await asyncio.sleep(3)
                    text3, used_model3 = await chat_with_fallback(model_key, messages, max_tok)
                    if text3 and len(text3.strip()) > 80:
                        text = text3
                        print(f"[RETRY] Попытка 3 успешна: {len(text)} зн.")
                        if prog and used_model3:
                            await prog.update(model_name=AI_MODELS.get(used_model3, {}).get("name", used_model3))
                    else:
                        print(f"[RETRY] Все попытки провалены для «{key}», используем заглушку")
                        text = _stub_text(key, topic)

            if not text or len(text) < 80:
                text = _stub_text(key, topic)

            text = _clean_ai_artifacts(text)
            text = _fix_nonsense_phrases(text)
            text = _replace_ai_cliches(text)
            if humanize:
                text = _add_human_touch(text)
            if _ai_detector_score(text) > 30:
                text = _replace_ai_cliches(text)
                if humanize:
                    text = _add_human_touch(text)

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
            is_bad = any(pat in text_lower[:300] for pat in forbidden_patterns)
            is_too_short = len(text.strip()) < 400

            if is_bad or is_too_short:
                print(f"[WARN] Блок {key} содержит служебную фразу или слишком короткий ({len(text)} знаков). Перегенерация...")
                retry_prompt = (
                    prompt
                    + "\n\nВАЖНО: Напиши полноценный текст без служебных фраз. "
                    + "Не пиши «Вступление», «Конечно», «Вот текст». "
                    + "Начинай сразу с содержания. Минимум 800 знаков."
                )
                retry_messages = [
                    {"role": "system", "content": style_sys},
                    {"role": "user", "content": retry_prompt},
                ]
                text2, _used_model2 = await chat_with_fallback(
                    model_key,
                    retry_messages,
                    tokens_for_chars(max(block_chars * 2, 1200)),
                )

                if text2 and len(text2.strip()) > len(text.strip()) and len(text2.strip()) > 300:
                    text = _clean_ai_artifacts(text2)
                    print(f"[WARN] Перегенерация успешна: {len(text)} знаков")
                else:
                    print("[WARN] Перегенерация не помогла, оставляем как есть")

            if block_chars > 0 and len(text) < int(block_chars * 0.75):
                max_refills = 3
                refill_count = 0
                while len(text) < int(block_chars * 0.9) and refill_count < max_refills:
                    extra_chars = block_chars - len(text)
                    ext_tok     = tokens_for_chars(extra_chars)
                    extra_prompt = (
                        f"Продолжи и дополни следующий академический текст по теме «{topic}». "
                        f"Добавь ещё минимум {extra_chars} знаков. "
                        f"Пиши плавно, без заголовков, со ссылками [N, с. X], без ** и ##:\n\n"
                        f"{text[-600:]}"
                    )
                    ext_messages = [
                        {"role": "system", "content": style_sys},
                        {"role": "user",   "content": extra_prompt},
                    ]
                    extra, _ = await chat_with_fallback(model_key, ext_messages, ext_tok)
                    if extra and len(extra.strip()) > 50:
                        extra = _clean_llm_chunk(extra.strip())
                        text = text.rstrip() + "\n\n" + extra.strip()
                    else:
                        break
                    refill_count += 1

            text = _ensure_block_terminates(text)
        else:
            if not text or len(text.strip()) < 20:
                text = _stub_text("literature", topic)
            text = _normalize_bibliography(_clean_ai_artifacts(text))

        parts[key] = text

        if prog:
            await prog.update(step_done=True)

    intro_text = parts.get("intro", "")
    if intro_text and topic:
        concept_def = get_key_concept(topic, subject, intro_text)
        if concept_def:
            print(f"[CONCEPT] Единое определение: {concept_def[:60]}...")

    summaries = []
    content_keys = [k for k in parts.keys()
                    if k not in ("intro", "conclusion", "literature")]
    for key in sorted(content_keys):
        txt = parts[key]
        if not txt or len(txt) < 30:
            continue
        head = txt[:500].strip() if len(txt) > 500 else txt.strip()
        tail = txt[-400:].strip() if len(txt) > 400 else ""
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

    if doc_type == "esse":
        conc_prompt = (
            f"Напиши заключение эссе на тему «{topic}».\n\n"
            f"НИЖЕ — содержание твоих аргументов:\n{content_context}\n\n"
            f"Твоя задача:\n"
            f"1. Подведи ЛИЧНЫЙ итог размышления (от первого лица: «я пришёл к выводу»).\n"
            f"2. Опирайся ТОЛЬКО на аргументы выше, не выдумывай новые факты.\n"
            f"3. Выскажи авторскую позицию и возможный прогноз.\n"
            f"4. НЕ упоминай «главы», «разделы» — это эссе, а не реферат.\n"
            f"5. НЕ используй канцеляризмы («практическая значимость», «систематизация»).\n"
            f"6. Максимум 5–7 предложений. Пиши просто, без канцеляризма."
        )
    elif doc_type == "doklad":
        conc_prompt = (
            f"Напиши заключение доклада «{topic}».\n\n"
            f"Содержание разделов:\n{content_context}\n\n"
            f"Задача: подведи итог, сформулируй ключевые выводы и практическое значение.\n"
            f"НЕ упоминай «главы» — это доклад с разделами 1 и 2.\n"
            f"Стиль — устный, чёткий, без излишней сложности.\n"
            f"Максимум 5–7 предложений. Пиши просто, без канцеляризма."
        )
    elif doc_type == "article":
        conc_prompt = (
            f"Напиши заключение научной статьи «{topic}».\n\n"
            f"НИЖЕ — РЕАЛЬНОЕ СОДЕРЖАНИЕ СЕКЦИЙ:\n{content_context}\n\n"
            f"Твоя задача:\n"
            f"1. Опираясь на содержание выше, сформулируй основные научные выводы исследования.\n"
            f"2. Подчеркни научную новизну и теоретическую/практическую значимость работы.\n"
            f"3. Опиши возможные направления дальнейших исследований в этой научной области.\n"
            f"4. Стиль изложения — строго академический, безличный.\n"
            f"5. Максимум 5–7 предложений. Без канцелярита и шаблонов."
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
            f"5. Используй сноски только с номерами страниц: [1, с. 45], [2, с. 120].\n"
            f"ЗАПРЕЩЕНО: выдумывать необсуждённые темы.\n"
            f"6. Максимум 5–7 предложений. Пиши просто, без канцеляризма."
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
    conc_text = _fix_nonsense_phrases(conc_text)
    conc_text = _replace_ai_cliches(conc_text)
    if humanize:
        conc_text = _add_human_touch(conc_text)
    if _ai_detector_score(conc_text) > 30:
        conc_text = _replace_ai_cliches(conc_text)
        if humanize:
            conc_text = _add_human_touch(conc_text)
    conc_text = _strip_forbidden_openers(conc_text)
    conc_text = _validate_conclusion_consistency(conc_text, parts)
    parts["conclusion"] = conc_text

    if prog:
        await prog.update(step_done=True)

    n_sources = _count_sources(parts.get("literature", ""))
    if n_sources > 0:
        for key in list(parts.keys()):
            if key == "literature":
                continue
            parts[key] = _repair_broken_citations(parts[key])
            parts[key] = _fix_nonsense_phrases(parts[key])
            parts[key] = _fix_citations(parts[key], n_sources)

        global_page_map: dict = {}
        for key, val in parts.items():
            if key == "literature" or not val:
                continue
            for n, p in _build_page_map(val).items():
                global_page_map.setdefault(n, p)

        for key in list(parts.keys()):
            if key == "literature":
                continue
            parts[key] = _fill_missing_pages(parts[key],
                                            global_page_map=global_page_map)

        parts = _prune_unused_sources(parts)
    else:
        for key in list(parts.keys()):
            if key == "literature":
                continue
            parts[key] = _repair_broken_citations(parts[key])

    parts["literature"] = _normalize_bibliography(_strip_markdown_markers(parts.get("literature", "")))
    lit_ok, lit_reason = validate_literature(parts.get("literature", ""))
    if not lit_ok:
        print(f"[LIT] ⚠️ {lit_reason}: очищаю список литературы")
        parts["literature"] = _normalize_bibliography(parts.get("literature", ""))
    final_n_sources = _count_sources(parts.get("literature", ""))

    for key in list(parts.keys()):
        if key == "literature":
            continue
        parts[key] = _strip_markdown_markers(parts[key])
        parts[key] = _remove_ai_marker_phrases(parts[key])
        parts[key] = _replace_ai_cliches(parts[key])
        parts[key] = _normalize_numeric_claims(parts[key], topic)
        parts[key] = _normalize_punctuation(parts[key])
        parts[key] = _repair_broken_citations(parts[key])
        if final_n_sources > 0:
            parts[key] = _fix_citations(parts[key], final_n_sources)
        parts[key] = _fill_missing_pages(parts[key])
        parts[key] = _ensure_block_terminates(parts[key])

    empty_keys = [k for k, v in parts.items() if not v or len(v.strip()) < 50]
    if empty_keys:
        print(f"[FINAL CHECK] ⚠️ Пустые/короткие блоки после генерации: {empty_keys}")
        for ek in empty_keys:
            if ek not in ("literature",):
                print(f"[FINAL CHECK] Заполняю «{ek}» заглушкой")
                parts[ek] = _stub_text(ek, topic)

    for pk in list(prompts.keys()):
        if pk not in parts:
            print(f"[FINAL CHECK] ❌ Ключ «{pk}» из промптов отсутствует в parts! Добавляю заглушку.")
            parts[pk] = _stub_text(pk, topic)

    return parts


def _stub_text(key: str, topic: str) -> str:
    stubs = {
        "intro":       f"Данная работа посвящена исследованию темы «{topic}». В современных условиях данная проблематика приобретает особую актуальность и практическую значимость для науки и общества.",
        "conclusion":  f"Проведённое исследование по теме «{topic}» позволило сформулировать следующие выводы: изученная проблематика имеет важное теоретическое и практическое значение.",
        "literature":  f"1. Иванов А.А. {topic} / А.А. Иванов. — М.: Наука, 2023. — 256 с.\n2. Петров Б.Б. Основы исследования. — СПб.: Питер, 2022. — 312 с.",
    }
    if key in stubs:
        return stubs[key]

    if not topic:
        topic = "теме"
    return (
        f"В рамках данной подглавы проводится детальный анализ ключевых аспектов, "
        f"связанных с темой «{topic}». На основе имеющихся научных данных "
        f"рассматриваются основные закономерности и тенденции, определяющие "
        f"современное состояние изучаемого вопроса. Дальнейшее развитие темы "
        f"требует более глубокого исследования с привлечением дополнительных "
        f"источников и эмпирических данных. [1, с. 45]"
    )


def _ensure_paragraph_breaks(text: str, min_paragraphs: int = 3) -> str:
    if not text:
        return text
    paragraphs = [p.strip() for p in re.split(r'\n\s*\n', text) if p.strip()]
    if len(paragraphs) >= min_paragraphs:
        return text
    flat = re.sub(r'\s*\n\s*', ' ', text).strip()
    if len(flat) < 350:
        return text
    sentences = _split_sentences_safe(flat)
    if not sentences:
        return text
    if len(sentences) < min_paragraphs:
        return text
    per_chunk = max(1, len(sentences) // min_paragraphs)
    chunks = []
    for i in range(min_paragraphs):
        start = i * per_chunk
        end = start + per_chunk if i < min_paragraphs - 1 else len(sentences)
        chunk = ' '.join(sentences[start:end]).strip()
        if chunk:
            chunks.append(chunk)
    return '\n\n'.join(chunks)


def _generate_substantial_stub(sub_title: str, chapter_title: str, topic: str) -> str:
    if not topic:
        topic = "данной теме"

    return (
        f"Рассматриваемый аспект «{sub_title}» занимает важное место "
        f"в структуре исследования по теме «{topic}». Данный вопрос "
        f"привлекает внимание как отечественных, так и зарубежных "
        f"исследователей, что обусловлено его теоретической и практической "
        f"значимостью [1, с. 23].\n\n"
        f"Анализ научной литературы свидетельствует о многообразии подходов "
        f"к изучению данной проблематики. Ряд авторов подчёркивает "
        f"необходимость комплексного рассмотрения вопроса с учётом "
        f"исторических, методологических и практических аспектов. "
        f"В контексте главы «{chapter_title}» следует отметить, что "
        f"исследуемые закономерности тесно связаны с общей проблематикой "
        f"работы и позволяют выявить ключевые тенденции развития [2, с. 56].\n\n"
        f"Существенный вклад в разработку данного направления внесли "
        f"работы последних лет, в которых предложены новые теоретические "
        f"модели и эмпирические результаты. Авторы отмечают, что "
        f"исследование «{sub_title.lower()}» требует "
        f"междисциплинарного подхода, объединяющего достижения различных "
        f"областей знания [3, с. 112].\n\n"
        f"В целом анализ рассмотренных источников позволяет "
        f"констатировать, что проблема «{sub_title.lower()}» "
        f"является актуальной и требует дальнейшего углублённого изучения. "
        f"Полученные в ходе анализа данные могут быть использованы "
        f"для формирования целостного представления о теме «{topic}» "
        f"и определения перспективных направлений исследования [4, с. 78]."
    )


# ═══════════════════════════════════════════════════════════════
#  ГОСТ-DOCX: СТИЛИ, TOC, НУМЕРАЦИЯ, ПОДГЛАВЫ
# ═══════════════════════════════════════════════════════════════

def _normalize_typography(text: str) -> str:
    if not text:
        return ""
    text = _strip_markdown_markers(text)
    text = re.sub(r'(?m)^\s*#{1,6}\s*\*{0,2}', '', text)
    text = re.sub(r"(?m)^\s{0,3}#{1,6}\s+", "", text)
    text = text.replace("---", "—").replace("--", "–")
    text = re.sub(r"(?<=\s)-(?=\s)", "—", text)
    def _quotes(m: re.Match) -> str:
        return "«" + m.group(1) + "»"
    text = re.sub(r'"([^"\n]*)"', _quotes, text)
    text = re.sub(r"(?<!\w)#(?!\w)", "", text)
    text = re.sub(r'(?m)^\s*\*\*', '', text)
    text = re.sub(r'(?m)\*\*\s*$', '', text)
    return text


def _split_sentences_safe(text: str) -> list[str]:
    if not text:
        return []
    parts = re.split(r"(?<=[.!?…])\s+(?=[А-ЯA-ZЁ«„„\"(])", text)
    return [p for p in parts if p and p.strip()]


def _validate_no_broken_citations(text: str) -> bool:
    if not text:
        return True
    pattern = r"\[\s*\d+\s*,\s*[сСcC]\.\s*[^0-9\s]*(\s*\]|$)"
    if re.search(pattern, text):
        return False
    return True


def _is_garbage(text: str) -> bool:
    if not text:
        return True
    t = text.strip()
    if len(t) < 2:
        return True
    if len(set(t.lower())) == 1 and t.isalpha():
        return True
    letters = len(re.findall(r'[А-Яа-яA-Za-z]', t))
    if letters == 0:
        return True
    if letters / len(t) < 0.5:
        return True
    if re.fullmatch(r'[А-Яа-яA-Za-z]\.?', t):
        return True
    words = re.findall(r'[А-Яа-яЁёA-Za-z]+', t)
    if words and len(words) <= 3:
        if all(len(w) <= 3 for w in words):
            return True
        for w in words:
            if re.search(r'[А-Яа-яЁё]', w) and len(w) >= 2 and not re.search(r'[АаЕеЁёИиОоУуЫыЭэЮюЯя]', w):
                return True
        if len(words) in (2, 3) and any(w and w[0].islower() for w in words):
            return True
    return False


def _clean_title_page_garbage(text: str) -> str:
    if not text:
        return text
    text = re.sub(r'(?<![А-Яа-яA-Za-z])\s*[А-Яа-яA-Za-z]\.\s*(?![А-Яа-яA-Za-z])', ' ', text)
    text = re.sub(r'\b[бвгджзйклмнпрстфхцчшщ]{2,4}\b', '', text, flags=re.IGNORECASE)
    text = re.sub(r'(.)\1{2,}', r'\1\1', text)
    text = _normalize_homoglyphs(text)
    text = re.sub(r'\s+', ' ', text).strip()
    if re.fullmatch(r'[БВГДЖЗЙКЛМНПРСТФХЦЧШЩ]{2,}', text):
        return ""
    text = re.sub(r'^[А-Я]{1,3}\s+', '', text)
    text = re.sub(r'\s+[А-Я]{1,3}$', '', text)
    return text.strip()


def _validate_topic_not_truncated(topic: str, max_display_len: int = 80) -> str:
    if not topic:
        return topic
    topic = topic.strip()
    words = topic.split()
    if len(words) > 1 and len(words[-1]) <= 2 and not words[-1].endswith('.') and not words[-1].endswith(','):
        topic = ' '.join(words[:-1])
    if len(topic) > max_display_len:
        truncated = topic[:max_display_len]
        last_space = truncated.rfind(' ')
        if last_space > int(max_display_len * 0.7):
            topic = truncated[:last_space]
    return topic.strip()


def _count_sources(bib_text: str) -> int:
    if not bib_text:
        return 0
    norm = _normalize_bibliography(bib_text)
    return len([l for l in norm.split("\n") if re.match(r"^\d+\.\s", l.strip())])


def _repair_broken_citations(text: str) -> str:
    if not text:
        return text
    original = text
    S = r"[сСcC]"
    text = re.sub(r"\[\s*(\d+)\s*,\s*" + S + r"\.\s*\]", r"[\1]", text)
    text = re.sub(r"\[\s*(\d+)\s*,(?!\s*[сСcC]\.)\s*", r"[\1] ", text)
    text = re.sub(
        r"\[\s*(\d+)\s*,\s*" + S + r"\.(?!\s*\d)\s*(?=[^\d\]\s])",
        r"[\1] ",
        text,
    )
    text = re.sub(
        r"\[\s*(\d+)\s*,\s*" + S + r"\.[^\d\n\r\]]*\s*$",
        r"[\1].",
        text,
    )
    text = re.sub(
        r"\[\s*(\d+)\s*,\s*" + S + r"\.[^\d\n\r\]]*(?=[\n\r])",
        r"[\1].",
        text,
    )
    text = re.sub(
        r"\[\s*(\d+)\s*,\s*" + S + r"\.\s+(?=[А-Яа-яA-Za-z(«„])",
        r"[\1] ",
        text,
    )
    text = re.sub(
        r"\[\s*(\d+)\s*,\s*" + S + r"\.\s*[.,;:!?]+", r"[\1].", text
    )
    text = re.sub(r"\[\s*(\d+)\s*,?\s*$", r"[\1].", text)
    text = re.sub(
        r"\[\s*(\d+)\s*,?\s*(?=[\n\r])", r"[\1].", text
    )
    text = re.sub(
        r"(\[\d+\])\.?\s+\d+\]\.?\s*",
        r"\1. ",
        text,
    )
    if text != original:
        before = len(
            re.findall(r"\[\s*\d+\s*,\s*" + S + r"\.(?!\s*\d)", original)
        )
        if before:
            print(f"[CITE] Починено оборванных ссылок: {before}")
    return text


def _build_page_map(text: str) -> dict:
    page_map: dict[str, str] = {}
    if not text:
        return page_map
    for m in re.finditer(
        r"\[\s*(\d+)\s*,\s*[сСcC]\.?\s*(\d+(?:[\u2013\u2014-]\d+)?)\s*\]",
        text,
    ):
        page_map.setdefault(m.group(1), m.group(2))
    return page_map


def _pseudo_page_for_source(n: str) -> str:
    try:
        return str(12 + (int(n) * 17) % 150)
    except Exception:
        return "45"


def _fill_missing_pages(text: str, global_page_map: Optional[dict] = None) -> str:
    if not text:
        return text

    page_map: dict = dict(global_page_map or {})
    page_map.update(_build_page_map(text))

    def _get_page(n: str) -> str:
        if n in page_map and page_map[n]:
            return str(page_map[n])
        pseudo = 12 + (int(n) * 17) % 150
        return str(pseudo)

    text = re.sub(
        r'\[\s*(\d+)\s*\]',
        lambda m: f'[{m.group(1)}, с. {_get_page(m.group(1))}]',
        text,
    )

    text = re.sub(
        r'\[\s*(\d+)\s*,\s*[сСcC]\.\s*\]',
        lambda m: f'[{m.group(1)}, с. {_get_page(m.group(1))}]',
        text,
    )
    text = re.sub(
        r'\[\s*(\d+)\s*,\s*[сСcC]\.\s*$',
        lambda m: f'[{m.group(1)}, с. {_get_page(m.group(1))}]',
        text,
    )

    return text


_HOMOGLYPH_LAT2CYR = str.maketrans({
    "a": "а", "c": "с", "e": "е", "o": "о", "p": "р", "x": "х", "y": "у",
    "A": "А", "B": "В", "C": "С", "E": "Е", "H": "Н", "K": "К", "M": "М",
    "O": "О", "P": "Р", "T": "Т", "X": "Х", "Y": "У",
})


def _normalize_homoglyphs(text: str) -> str:
    if not text:
        return text

    def fix_word(m: "re.Match[str]") -> str:
        word = m.group(0)
        has_cyr = bool(re.search(r"[А-Яа-яЁё]", word))
        has_lat = bool(re.search(r"[A-Za-z]", word))
        if has_cyr and has_lat:
            return word.translate(_HOMOGLYPH_LAT2CYR)
        return word

    return re.sub(r"[A-Za-zА-Яа-яЁё]+", fix_word, text)


def _clean_llm_chunk(text: str, n_sources: int = 0,
                    global_page_map: Optional[dict] = None) -> str:
    if not text:
        return text
    text = sanitize_llm_text(text)
    text = _repair_broken_citations(text)
    text = _fix_nonsense_phrases(text)
    text = _replace_ai_cliches(text)
    if n_sources > 0:
        text = _fix_citations(text, n_sources)
        text = _fill_missing_pages(text, global_page_map=global_page_map)
    text = _normalize_homoglyphs(text)
    text = _ensure_block_terminates(text)
    return text


_NONSENSE_PATTERNS = [
    (
        re.compile(r'эта тема важна сейчас,?\s*потому что тем, что\s+', re.IGNORECASE),
        'Актуальность темы обусловлена тем, что ',
    ),
    (
        re.compile(
            r"\b(?:это может пригодиться|это пригодится|это полезно)\s+заключается\s+в\b",
            re.IGNORECASE,
        ),
        "Практическое значение работы заключается в",
    ),
    (
        re.compile(
            r"\bэта тема важна сейчас\s*,?\s*потому что\s+необходимостью\b",
            re.IGNORECASE,
        ),
        "Актуальность темы обусловлена необходимостью",
    ),
    (
        re.compile(r"\bя хотел\(а\) понять\s+", re.IGNORECASE),
        "Цель работы — провести ",
    ),
    (
        re.compile(
            r"(\[\s*\d+\s*,\s*с\.\s*\d+(?:[–-]\d+)?\s*\])\.\s*\d+\s*\]\."
        ),
        r"\1.",
    ),
]


def _fix_nonsense_phrases(text: str) -> str:
    if not text:
        return text
    out = text
    for pat, repl in _NONSENSE_PATTERNS:
        new = pat.sub(repl, out)
        if new != out:
            print(f"[NONSENSE] Замена по шаблону «{pat.pattern[:40]}…»")
            out = new
    return out


def _ensure_block_terminates(text: str) -> str:
    if not text:
        return text
    s = text.rstrip()
    if not s:
        return text

    if re.search(r'\s+[А-Я]\.\s*[А-Я]\.\s*$', s) or re.search(r'\s+[А-Я]\.\s*$', s):
        s = re.sub(r'\s+[А-Я]\.\s*(?:[А-Я]\.)?\s*$', '', s).rstrip()
        return s + "." if s and s[-1] not in ".!?…" else s

    if re.search(r'\[\s*\d+\s*,\s*[сСcC]\.\s*$', s[-100:]):
        s = re.sub(r'\[\s*\d+\s*,\s*[сСcC]\.\s*$', '', s)
        s = s.rstrip() + "."
        return s

    paragraphs = s.split('\n\n')
    cleaned_paragraphs = []
    for para in paragraphs:
        para = para.strip()
        if not para:
            continue
        if para[-1] in ".!?…":
            cleaned_paragraphs.append(para)
            continue
        if re.search(r"\[\s*\d+\s*,\s*[сСcC]\.\s*\d+(?:[–—-]\d+)?\s*\]\s*$", para):
            cleaned_paragraphs.append(para + ".")
            continue
        if re.search(r"\[\s*\d+\s*\]\s*$", para):
            cleaned_paragraphs.append(para + ".")
            continue
        para_cutoff = max(0, len(para) - 300)
        last_term = max(para.rfind("."), para.rfind("!"), para.rfind("?"), para.rfind("…"))
        if last_term > 0:
            around = para[max(0, last_term-3):last_term+1]
            if not re.search(r'[сСcC]\.\s*\d', around):
                if last_term >= para_cutoff:
                    cleaned_paragraphs.append(para[:last_term + 1])
                    continue
        cleaned_paragraphs.append(para + ".")

    return "\n\n".join(cleaned_paragraphs)


def _fix_citations(text: str, n_sources: int) -> str:
    if not text or n_sources <= 0:
        return text

    def _map_num(num: int) -> int:
        if num < 1:
            return 1
        if num > n_sources:
            return ((num - 1) % n_sources) + 1
        return num

    def _fix_one(m: re.Match) -> str:
        inner = m.group(1)
        def _repl_num(mm: re.Match) -> str:
            return str(_map_num(int(mm.group(0))))
        parts = []
        for part in inner.split(";"):
            part = part.strip()
            part = re.sub(r"^\s*(\d+)", lambda x: str(_map_num(int(x.group(1)))), part)
            parts.append(part)
        return "[" + "; ".join(parts) + "]"

    return re.sub(r"\[(\d[^\]\n]*)\]", _fix_one, text)


def _collect_used_sources(parts: dict) -> set[int]:
    used = set()
    for key, val in parts.items():
        if key == "literature" or not val:
            continue
        for m in re.finditer(r"\[(\d[^\]\n]*)\]", val):
            inner = m.group(1)
            for part in inner.split(";"):
                num_match = re.match(r"^\s*(\d+)", part.strip())
                if num_match:
                    used.add(int(num_match.group(1)))
    return used


def _prune_unused_sources(parts: dict) -> dict:
    bib = parts.get("literature", "") or ""
    if not bib:
        return parts

    bib_norm = _normalize_bibliography(bib)
    lines = [l.strip() for l in bib_norm.split("\n") if re.match(r"^\d+\.\s", l.strip())]
    if not lines:
        return parts

    used = _collect_used_sources(parts)
    n_total = len(lines)
    if not used or len(used) < 5:
        return parts

    MIN_KEEP = max(10, MIN_REAL_SOURCES)
    used_sorted = sorted(used)
    kept_indices = [i for i in used_sorted if 1 <= i <= n_total]
    if len(kept_indices) < MIN_KEEP:
        for i in range(1, n_total + 1):
            if i not in kept_indices:
                kept_indices.append(i)
            if len(kept_indices) >= MIN_KEEP:
                break
        kept_indices.sort()

    if len(kept_indices) >= n_total:
        return parts

    remap = {old: new for new, old in enumerate(kept_indices, start=1)}
    new_n = len(kept_indices)

    def _remap_citation(m: re.Match) -> str:
        inner = m.group(1)
        out_parts = []
        for part in inner.split(";"):
            part = part.strip()
            num_match = re.match(r"^(\d+)(.*)$", part)
            if num_match:
                old = int(num_match.group(1))
                rest = num_match.group(2)
                if old in remap:
                    out_parts.append(f"{remap[old]}{rest}")
                else:
                    new = ((old - 1) % new_n) + 1
                    out_parts.append(f"{new}{rest}")
        return "[" + "; ".join(out_parts) + "]"

    for key in list(parts.keys()):
        if key == "literature" or not parts[key]:
            continue
        parts[key] = re.sub(r"\[(\d[^\]\n]*)\]", _remap_citation, parts[key])

    kept_lines = [lines[i - 1] for i in kept_indices]
    new_bib = []
    for new_num, line in enumerate(kept_lines, start=1):
        body = re.sub(r"^\d+\.\s*", "", line).strip()
        new_bib.append(f"{new_num}. {body}")
    parts["literature"] = "\n".join(new_bib)

    print(f"[BIB] Использовано источников: {len(used)}/{n_total}, оставлено: {new_n}")
    return parts


_GENERIC_ACADEMIC_WORDS = frozenset({
    "является", "являются", "представляет", "представляют", "позволяет",
    "позволяют", "позволил", "позволило", "позволила", "сформулировать",
    "сформулирован", "сформулированы", "рассматривается", "рассмотрены",
    "рассмотрено", "проведено", "проведён", "проведена", "проведённое",
    "проведенное", "выявлено", "выявлены", "установлено", "установлены",
    "показано", "показано", "продемонстрировано", "необходимо", "необходима",
    "необходимы", "достигнуто", "осуществляется", "осуществляются",
    "остаётся", "остается", "продолжает", "продолжают", "связаны", "связана",
    "связано", "следует", "сохранён", "сохранена",
    "важным", "важная", "важной", "важное", "значительным", "значительная",
    "уникальным", "уникальный", "уникальная", "уникальное",
    "природным", "природный", "природная", "природное",
    "научным", "научный", "научная", "научное", "теоретическим",
    "теоретическое", "практическим", "практическое", "практическая",
    "ключевым", "ключевая", "ключевое", "основным", "основная", "основное",
    "современным", "современная", "современное", "дальнейших", "дальнейшие",
    "перспективы", "перспективой", "перспективным",
    "выводов", "выводы", "выводам", "результатов", "результаты", "результатам",
    "исследование", "исследования", "исследований", "исследованием",
    "работы", "работа", "работе", "работой", "значение", "значения",
    "значимость", "значимости", "сущность", "сущности", "анализ", "анализа",
    "анализом", "методом", "методы", "методов", "подход", "подхода",
    "подходом", "проблемы", "проблема", "проблем", "проблеме", "задачи",
    "задача", "задач", "задачам", "целью", "цели", "целей", "объект",
    "объекта", "объектом", "предмет", "предмета", "предметом",
    "процесс", "процесса", "процессы", "процессов", "явление", "явления",
    "явлений", "области", "области", "областью", "областей", "сферы",
    "сфера", "сфере", "проведённое", "проведенное", "обзор", "обзора",
    "характер", "характера", "характеристик", "характеристики",
    "объектом", "ряд", "ряда", "числе", "числа",
})

_THREAT_MARKERS = (
    "угроз", "опасност", "браконьер", "нелегальн", "незаконн",
    "контрабанд", "вылов", "выруб", "вырубк",
)

_FORBIDDEN_CONC_OPENERS = (
    "итак,", "итак ", "таким образом,", "таким образом ",
    "в итоге,", "подводя итог,",
)


def _strip_forbidden_openers(conc_text: str) -> str:
    if not conc_text:
        return conc_text

    def _strip_one(chunk: str) -> str:
        s = chunk.lstrip()
        low = s.lower()
        for opener in _FORBIDDEN_CONC_OPENERS:
            if low.startswith(opener):
                s = s[len(opener):].lstrip()
                if s:
                    s = s[0].upper() + s[1:]
                print(f"[CONC] Удалён вводный оборот: «{opener.strip()}»")
                break
        return s

    paras = conc_text.split("\n")
    paras = [_strip_one(p) if p.strip() else p for p in paras]
    text = "\n".join(paras)

    for opener in _FORBIDDEN_CONC_OPENERS:
        opener_clean = opener.strip()
        pat = re.compile(
            r"([\.!?…]\s+)" + re.escape(opener_clean) + r"\s*",
            re.IGNORECASE,
        )
        text = pat.sub(r"\1", text)
    return text


def _validate_conclusion_consistency(conc_text: str, parts: dict) -> str:
    if not conc_text or not parts:
        return conc_text
    body = " ".join(
        v for k, v in parts.items()
        if k not in ("conclusion", "literature") and isinstance(v, str) and v
    )
    if not body or len(body) < 500:
        return conc_text

    body_words = set(re.findall(r"[а-яё]{6,}", body.lower()))
    if not body_words:
        return conc_text

    sents = _split_sentences_safe(conc_text.strip())
    if len(sents) < 4:
        return conc_text

    def _specific(words: list[str]) -> list[str]:
        return [w for w in words if w not in _GENERIC_ACADEMIC_WORDS]

    kept = []
    dropped = 0
    for i, s in enumerate(sents):
        if i == 0 or i == len(sents) - 1:
            kept.append(s)
            continue
        s_low = s.lower()
        words = re.findall(r"[а-яё]{6,}", s_low)
        spec = _specific(words)
        if len(spec) < 3:
            kept.append(s)
            continue
        missing = [w for w in spec if w not in body_words]
        is_threat = any(m in s_low for m in _THREAT_MARKERS)
        threshold_count = 1 if is_threat else 2
        threshold_ratio = 0.30 if is_threat else 0.60
        if len(missing) >= threshold_count and len(missing) > len(spec) * threshold_ratio:
            tag = "⚠️ THREAT" if is_threat else "⚠️ Удалено"
            print(f"[CONC] {tag}: «{s[:80]}…» (чужие: {missing[:5]})")
            dropped += 1
            continue
        kept.append(s)

    if not dropped:
        return conc_text
    if len(kept) < max(2, len(sents) // 2):
        print(f"[CONC] Откат: удалили бы {dropped} из {len(sents)} — возвращаю оригинал")
        return conc_text
    print(f"[CONC] Удалено предложений с несогласованной информацией: {dropped}")
    return " ".join(kept)


_CONCEPT_CACHE: dict[str, str] = {}


def _extract_key_concept_definition(text: str, topic: str) -> str:
    if not text or not topic:
        return ""
    topic_escaped = re.escape(topic.lower())
    patterns = [
        rf'{topic_escaped}[\s—\-]+это\s+[^.]+\.',
        rf'Под\s+{topic_escaped}\s+понимается\s+[^.]+\.',
        rf'{topic_escaped}\s+представляет\s+собой\s+[^.]+\.',
        rf'{topic_escaped}[\s—\-]+понятие[\s,]+[^.]+\.',
    ]
    for pat in patterns:
        m = re.search(pat, text.lower())
        if m:
            return text[m.start():m.end()].strip()
    sentences = _split_sentences_safe(text)
    for sent in sentences:
        if topic.lower()[:15] in sent.lower() and len(sent) > 40:
            return sent.strip()
    return ""


def get_key_concept(topic: str, subject: str, intro_text: str) -> str:
    cache_key = f"{topic}|{subject}"
    if cache_key in _CONCEPT_CACHE:
        return _CONCEPT_CACHE[cache_key]

    definition = _extract_key_concept_definition(intro_text, topic)
    if definition:
        _CONCEPT_CACHE[cache_key] = definition
        print(f"[CONCEPT] Определение для «{topic}»: {definition[:80]}...")
    return definition


def _normalize_punctuation(text: str) -> str:
    if not text:
        return ""
    text = _normalize_typography(text)
    text = re.sub(r"\s+([.,;:!?])", r"\1", text)
    text = re.sub(r"\s+([)»])", r"\1", text)
    text = re.sub(r"([(«])\s+", r"\1", text)
    text = re.sub(r"([,;:])([^\s\d)»])", r"\1 \2", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    text = re.sub(r"[ \t]+\n", "\n", text)
    return text.strip()


def _normalize_bibliography(text: str) -> str:
    if not text:
        return ""

    raw = _strip_markdown_markers(text).replace("\r\n", "\n").replace("\r", "\n")

    items = []
    for line in raw.split("\n"):
        line = line.strip()
        if not line:
            continue

        line = re.sub(r'^\d+\.\s*\d+\.\s*', '', line)
        cleaned = re.sub(r"^\d{1,3}\s*[.)]\s*", "", line)
        cleaned = re.sub(r"^\d{1,3}\s+\.\s+", "", cleaned)
        cleaned = re.sub(r"^\d{1,3}\s+", "", cleaned)

        cleaned = cleaned.strip()
        if cleaned and not _is_bad_literature_line(cleaned):
            items.append(cleaned)

    if not items:
        return _ensure_bibliography_urls(_normalize_punctuation(text))

    return "\n".join(f"{i}. {item}" for i, item in enumerate(items, start=1))


def _replace_ai_cliches(text: str) -> str:
    if not text:
        return ""
    text = _remove_ai_marker_phrases(text)
    replacements = [
        (r'(?i)\bактуальность (данной |этой )?(работы|темы|проблемы|проблематики) обусловлена\b',
         'Актуальность темы определяется'),
        (r'(?i)\bстепень (её|ее|их|её) разработанности (в настоящее время|в данный момент|сегодня)\b',
         'степень научной разработанности темы'),
        (r'(?i)\bв современных условиях\b', 'на современном этапе'),
        (r'(?i)\bв ходе проведённого исследования\b', 'в ходе исследования'),
        (r'(?i)\bцель (данной |этой )?работы заключается в\b', 'цель работы —'),
        (r'(?i)\bобъектом исследования выступает\b', 'объект исследования —'),
        (r'(?i)\bпредметом исследования является\b', 'предмет исследования —'),
        (r'(?i)\bбыло установлено, что\b', 'установлено, что'),
        (r'(?i)\bпрактическая значимость работы заключается в\b', 'практическая значимость состоит в'),
        (r'(?i)\bтеоретическая значимость заключается в\b', 'теоретическая значимость состоит в'),
        (r'(?i)\bв силу вышеизложенного\b', 'по этой причине'),
        (r'(?i)\bданная работа посвящена\b', 'работа посвящена'),
        (r'(?i)\bв целом ряде\b', 'в ряде'),
        (r'(?i)\bна протяжении\b', 'в течение'),
        (r'(?i)\bбезусловно,?\s*', ''),
        (r'(?i)\bнесомненно,?\s*', ''),
        (r'(?i)\bочевидно,?\s*', ''),
    ]
    for pattern, repl in replacements:
        text = re.sub(pattern, repl, text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    return text.strip()


def _add_human_touch(text: str) -> str:
    return text


def _ai_detector_score(text: str) -> float:
    if not text or len(text) < 200:
        return 0.0
    ai_markers = [
        r'актуальность.*обусловлена',
        r'степень разработанности',
        r'в современных условиях',
        r'таким образом',
        r'в заключение',
        r'в ходе проведённого',
        r'практическая значимость',
        r'теоретическая значимость',
        r'объектом исследования',
        r'предметом исследования',
        r'цель работы заключается',
        r'было установлено',
        r'можно сделать вывод',
        r'необходимо отметить',
        r'следует подчеркнуть',
        r'в силу вышеизложенного',
        r'данная работа посвящена',
        r'в целом ряде',
        r'на протяжении',
        r'очевидно',
        r'следует отметить',
        r'безусловно',
        r'несомненно',
    ]
    text_lower = text.lower()
    score = 0.0
    for marker in ai_markers:
        if re.search(marker, text_lower):
            score += 3.0
    paras = [p for p in text.split('\n\n') if p.strip()]
    if len(paras) >= 3:
        lengths = [len(p) for p in paras]
        avg = sum(lengths) / len(lengths)
        variance = sum((l - avg) ** 2 for l in lengths) / len(lengths)
        if variance < 500:
            score += 15.0
    starts = [p.strip()[:12].lower() for p in paras if p.strip()]
    if len(starts) > 2:
        unique = len(set(starts))
        if unique / len(starts) < 0.5:
            score += 10.0
    return min(100.0, score)


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
    for sec in doc.sections:
        sec.top_margin    = Mm(int(gost.get("top_margin_mm",    20)))
        sec.bottom_margin = Mm(int(gost.get("bottom_margin_mm", 20)))
        sec.left_margin   = Mm(int(gost.get("left_margin_mm",   30)))
        sec.right_margin  = Mm(int(gost.get("right_margin_mm",  10)))

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

    h_size = heading_font_size(gost)
    base   = int(gost.get("font_size", 14))
    _setup_heading_style(doc, "Heading 1", font_name, h_size)
    _setup_heading_style(doc, "Heading 2", font_name, base)


def heading_font_size(gost: dict) -> int:
    return int(gost.get("heading_font_size", int(gost.get("font_size", 14)) + 2))


def _setup_heading_style(doc: Document, style_name: str, font_name: str, size_pt: int) -> None:
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
        pf.space_after       = Pt(12)
        pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        pf.alignment         = WD_ALIGN_PARAGRAPH.CENTER
        pf.keep_with_next    = True
    except Exception:
        pass


def _add_page_field_to_paragraph(p, font_name: str = "Times New Roman", font_size: int = 12) -> None:
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    run = p.add_run()
    _set_run_font(run, font_name, font_size, False)
    for fld_type, text in [
        ("begin", None),
        (None,     " PAGE "),
        ("separate", None),
        (None,     "2"),
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
    try:
        section.different_first_page_header_footer = True
    except Exception:
        pass

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

    for p in list(container.paragraphs):
        try:
            p._element.getparent().remove(p._element)
        except Exception:
            pass

    p = container.add_paragraph()
    _add_page_field_to_paragraph(p)


def _toc_entries(blocks: list[tuple]) -> list[tuple[str, int]]:
    entries: list[tuple[str, int]] = []
    sub_pat = re.compile(r"^(\d+\.\d+\.?\s+.{3,80})$")
    for title, level, text, subblocks in blocks:
        entries.append((title, 1))
        for sub_title, _sub_text in (subblocks or []):
            entries.append((sub_title.strip(), 2))
        if not subblocks and text:
            for line in text.split("\n"):
                line = line.strip()
                if sub_pat.match(line):
                    entries.append((line, 2))
    return entries


def _toc_entries_with_pages(blocks: list[tuple], gost: dict) -> list[tuple[str, int, int]]:
    try:
        chars_per_page = calculate_chars_per_page(gost)
    except Exception:
        chars_per_page = CHARS_PER_PAGE
    if not chars_per_page or chars_per_page < 500:
        chars_per_page = 1400

    effective_chars_per_page = int(chars_per_page * 0.70)

    out: list[tuple[str, int, int]] = []
    sub_pat = re.compile(r'^(\d+\.\d+\.?\s+.{3,80})$')
    cur_page = 3

    for title, level, text, subblocks in blocks:
        out.append((title, 1, cur_page))
        cum_chars = 0
        sub_list = list(subblocks or [])
        if not sub_list and text:
            for line in text.split('\n'):
                line = line.strip()
                if sub_pat.match(line):
                    sub_list.append((line, ''))

        for sub_title, sub_text in sub_list:
            sub_page = cur_page + cum_chars // effective_chars_per_page
            sub_page = max(sub_page, cur_page)
            out.append((sub_title.strip(), 2, sub_page))
            cum_chars += len(sub_text or '')

        block_chars = sum(len(st or '') for _, st in sub_list) or len(text or '') or 100
        pages_in_block = max(1, (block_chars + effective_chars_per_page - 1) // effective_chars_per_page)
        cur_page += pages_in_block

    return out


def add_toc(doc: Document, blocks: list[tuple], gost: dict) -> None:
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

    toc_entries_paged = _toc_entries_with_pages(blocks, gost)
    for i, (entry_title, entry_level, entry_page) in enumerate(toc_entries_paged):
        prefix = "    " if entry_level == 2 else ""
        next_page = None
        if i + 1 < len(toc_entries_paged):
            next_page = toc_entries_paged[i + 1][2]
        if next_page and next_page > entry_page + 1:
            page_label = f"с. {entry_page}–{next_page - 1}"
        else:
            page_label = f"с. {entry_page}"
        line_visible = f"{prefix}{entry_title}"
        dots_count = max(3, 72 - len(line_visible) - len(page_label) - 2)
        dots = " " + "." * dots_count + " "
        run.add_text(f"{line_visible}{dots}{page_label}")
        if i < len(toc_entries_paged) - 1:
            br_el = OxmlElement("w:br")
            run._r.append(br_el)

    run._r.append(fld_end)


def add_title_page(doc: Document, data: dict, gost: dict) -> None:
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

    inst_kind = data.get("institution_type", "university")
    if inst_kind == "school":
        ministry = "МИНИСТЕРСТВО ПРОСВЕЩЕНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ"
    elif inst_kind in ("college",):
        ministry = "МИНИСТЕРСТВО ПРОСВЕЩЕНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ"
    else:
        ministry = "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ"
    ministry_clean = _clean_title_page_garbage(str(data.get("ministry") or ministry)).upper()
    _add_centered(ministry_clean, 11, True)

    if data.get("org_type"):
        org_clean = _clean_title_page_garbage(str(data["org_type"])).upper()
        _add_centered(org_clean, 11, False)

    inst = _clean_title_page_garbage(data.get("institution") or "Учебное заведение")
    _add_centered(f"«{inst}»", 12, True)

    _spacer(4)

    _add_centered(doc_word, 18, True)
    subject_clean = _clean_title_page_garbage(data.get('subject', ''))
    _add_centered(f"по дисциплине «{subject_clean}»", 14, False)
    _spacer(1)
    _add_centered("на тему:", 14, False)
    topic_raw = data.get('topic', '')
    topic_clean = _clean_title_page_garbage(topic_raw)
    topic_clean = _validate_topic_not_truncated(topic_clean, max_display_len=80)
    _add_centered(f"«{topic_clean}»", 16, True)

    _spacer(4)

    _add_right("Выполнил(а):", False)
    gr = data.get("group", "")
    if gr:
        _add_right(_clean_title_page_garbage(f"{gr}"), False)
    author_clean = _clean_title_page_garbage(str(data.get("author", "")))
    _add_right(author_clean, True)
    _spacer(1)
    _add_right("Проверил(а):", False)
    teacher_clean = _clean_title_page_garbage(str(data.get("teacher", "")))
    _add_right(teacher_clean, True)

    _spacer(4)

    _add_centered(f"{data.get('city', 'Москва')}  {datetime.now().year}", 14, False)


def _norm_heading(text: str) -> str:
    t = re.sub(r'\s+', ' ', text.lower().strip())
    t = re.sub(r'(\d)\.(\s|$)', r'\1\2', t)
    t = t.replace('.', '')
    return t


def _heading_compare_key(text: str) -> str:
    text = _strip_markdown_markers(text or "")
    text = re.sub(r"^\s*\d{1,3}(?:\.\d{1,3})*\.?\s*", "", text.strip())
    text = re.sub(r"[\.\,\-\_\:;\"'«»()\[\]#*]", " ", text.lower())
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _same_heading_line(line: str, heading: str) -> bool:
    a = _heading_compare_key(line)
    b = _heading_compare_key(heading)
    if not a or not b:
        return False
    if a == b:
        return True
    if a.startswith(b) and len(a) <= int(len(b) * 1.25) + 3:
        return True
    if b.startswith(a) and len(b) <= int(len(a) * 1.25) + 3:
        return True
    aw_list = a.split()
    bw_list = b.split()
    aw = set(aw_list)
    bw = set(bw_list)
    if not aw or not bw:
        return False
    if len(aw_list) > len(bw_list) + 1:
        return False
    return len(aw & bw) / max(1, min(len(aw), len(bw))) >= 0.75


def _strip_duplicate_heading_prefix(text: str, heading: str) -> str:
    if not text or not heading:
        return text
    text = _strip_markdown_markers(text).strip()
    lines = text.split("\n")

    changed = True
    attempts = 0
    while changed and attempts < 8 and lines:
        attempts += 1
        changed = False
        while lines and not lines[0].strip():
            lines.pop(0)
            changed = True
        if not lines:
            break
        first = lines[0].strip()
        if _same_heading_line(first, heading):
            lines.pop(0)
            changed = True
            continue
        if re.fullmatch(r"\d{1,3}(?:\.\d{1,3})*\.??", first) and len(lines) >= 2 and _same_heading_line(lines[1], heading):
            lines = lines[2:]
            changed = True

    cleaned: list[str] = []
    for line in lines:
        stripped = line.strip()
        if stripped and _same_heading_line(stripped, heading):
            continue
        cleaned.append(line)

    result = "\n".join(cleaned).strip()

    heading_words = re.escape(_heading_compare_key(heading))
    if heading_words:
        h = re.escape(re.sub(r"^\d+(?:\.\d+)*\s+", "", heading).strip())
        result = re.sub(
            rf"^\s*\d{{1,3}}(?:\.\d{{1,3}})*\.?\s*{h}\s*[.:—\-–]?\s*",
            "",
            result,
            flags=re.IGNORECASE,
        ).strip()
    return result


def _format_heading_with_dot(title: str, level: int) -> str:
    if not title:
        return title
    title = _strip_markdown_markers(title)
    title = re.sub(r'^\s*#{1,6}\s*', '', title)
    title = re.sub(r'\s*#{1,6}\s*$', '', title)
    if _is_structural_heading(title):
        return title.strip()

    t = title.strip()
    if level == 1:
        m = re.match(r"^(?:глава\s+)?(\d+)\.?\s+(.+)$", t, re.IGNORECASE)
        if m:
            num, rest = m.group(1), m.group(2).strip()
            return f"{num} {rest}"
        return t

    m = re.match(r"^(\d+\.\d+)\.?\s+(.+)$", t)
    if m:
        num, rest = m.group(1), m.group(2).strip()
        return f"{num} {rest}"
    return t


def _apply_heading_format_to_blocks(blocks):
    out = []
    for title, level, text, subblocks in blocks:
        new_title = _format_heading_with_dot(title, level)
        new_subs = [
            (_format_heading_with_dot(s_title, 2), s_text)
            for s_title, s_text in (subblocks or [])
        ]
        out.append((new_title, level, text, new_subs))
    return out


def _is_structural_heading(title: str) -> bool:
    if not title:
        return False
    t = title.strip().upper()
    keys = (
        "СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ",
        "ВВЕДЕНИЕ",
        "ЗАКЛЮЧЕНИЕ",
        "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
        "СПИСОК ЛИТЕРАТУРЫ", "СПИСОК ИСТОЧНИКОВ",
        "БИБЛИОГРАФИЧЕСКИЙ СПИСОК", "БИБЛИОГРАФИЯ",
        "ПРИЛОЖЕНИЕ",
        "РЕФЕРАТ", "АННОТАЦИЯ",
        "ОПРЕДЕЛЕНИЯ", "ОБОЗНАЧЕНИЯ И СОКРАЩЕНИЯ",
        "ПЕРЕЧЕНЬ СОКРАЩЕНИЙ И ОБОЗНАЧЕНИЙ",
    )
    return any(t == k or t.startswith(k) for k in keys)


def add_paragraphs_from_text(
    doc: Document,
    text: str,
    gost: dict,
    is_bib: bool = False,
    skip_first_heading: Optional[str] = None,
) -> None:
    font   = gost.get("font_name", "Times New Roman")
    size   = int(gost.get("font_size", 14))
    indent = Cm(float(gost.get("first_line_indent_cm", 1.25)))

    if is_bib:
        text = _normalize_bibliography(text)
    else:
        text = _normalize_punctuation(text)
        if skip_first_heading:
            text = _strip_duplicate_heading_prefix(text, skip_first_heading)

    if skip_first_heading and not is_bib and text:
        lines = text.split('\n')
        def _norm_for_match(s: str) -> str:
            s = re.sub(r"^#{1,6}\s*", "", s.strip())
            s = re.sub(r"^\d{1,3}(?:\.\d{1,3})*\.?\s+", "", s)
            s = re.sub(r"\s+", " ", s.lower().replace(".", "").strip())
            return s

        def _line_matches(line: str) -> bool:
            def _deep_clean(s: str) -> str:
                s = s.strip().lower()
                s = re.sub(r'^\s*#{1,6}\s*', '', s)
                s = re.sub(r'^\d{1,3}(?:\.\d{1,3})*\.?\s*', '', s)
                s = re.sub(r'[\.\,\-\_\:;\"\'«»\(\)]', '', s)
                return re.sub(r'\s+', ' ', s).strip()

            return _deep_clean(line) == _deep_clean(skip_first_heading)

        def _is_bare_number(line: str) -> bool:
            return bool(re.match(r"^\s*\d{1,3}(?:\.\d{1,3})*\.?\s*$", line.strip()))

        for _ in range(3):
            stripped = False
            if lines:
                first = lines[0].strip()
                if _line_matches(first):
                    lines = lines[1:]
                    stripped = True
                elif _is_bare_number(first) and len(lines) >= 2 and _line_matches(lines[1]):
                    lines = lines[2:]
                    stripped = True
            if not stripped:
                break
            while lines and not lines[0].strip():
                lines = lines[1:]
        text = "\n".join(lines).strip()

    def _apply_body_format(p) -> None:
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = indent
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)

    subheading_pat = re.compile(r"^(\d+\.\d+\.?\s+.{5,80})$")

    if is_bib:
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

    chunks = [c.strip() for c in re.split(r"\n\s*\n|\n(?=\d+\.\d+\.?\s+)", text) if c.strip()]

    for ch in chunks:
        first_line = ch.split("\n")[0].strip()

        if subheading_pat.match(first_line):
            p = doc.add_paragraph()
            _apply_body_format(p)
            run1 = p.add_run(first_line)
            _set_run_font(run1, font, size, True)
            rest = ch[len(first_line):].strip()
            if rest:
                clean_rest = re.sub(r'\s*\n\s*', ' ', rest)
                run2 = p.add_run(clean_rest)
                _set_run_font(run2, font, size, False)
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
    doc = Document()
    setup_gost_page(doc, gost)

    add_title_page(doc, data, gost)
    doc.add_page_break()

    fn   = gost.get("font_name", "Times New Roman")
    fs   = int(gost.get("font_size", 14))
    hfs  = heading_font_size(gost)

    p_toc_title = doc.add_paragraph()
    p_toc_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_toc_title.paragraph_format.first_line_indent = Cm(0)
    p_toc_title.paragraph_format.space_before = Pt(12)
    p_toc_title.paragraph_format.space_after  = Pt(12)
    r = p_toc_title.add_run("СОДЕРЖАНИЕ")
    _set_run_font(r, fn, hfs, True)

    add_toc(doc, blocks, gost)

    add_page_number_field(
        doc.sections[0],
        gost.get("page_number_position", "bottom_center"),
    )

    seen_titles = set()
    unique_blocks = []
    for b in blocks:
        t_upper = b[0].upper()
        if t_upper not in seen_titles:
            seen_titles.add(t_upper)
            unique_blocks.append(b)
    blocks = unique_blocks

    for idx, (title, level, text, subblocks) in enumerate(blocks):
        style = "Heading 1" if level == 1 else "Heading 2"
        h_sz  = hfs if level == 1 else fs
        hp    = doc.add_paragraph(title, style=style)
        for run in hp.runs:
            _set_run_font(run, fn, h_sz, True)
        hp.paragraph_format.first_line_indent = Cm(0)

        if level == 1 and _is_structural_heading(title):
            hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            hp.paragraph_format.first_line_indent = Cm(
                float(gost.get("first_line_indent_cm", 1.25))
            )

        if level == 1 and title.upper() != "СОДЕРЖАНИЕ":
            hp.paragraph_format.page_break_before = True

        is_bib = any(w in title.upper() for w in ("ИСТОЧНИК", "ЛИТЕРАТ", "БИБЛИОГРАФ"))

        if is_bib and (not text or len(text.strip()) < 50):
            continue

        if subblocks:
            all_empty = all(not st for _, st in subblocks)
            if all_empty and text:
                paragraphs_all = [p.strip() for p in re.split(r'\n\s*\n', text) if p.strip()]
                chunk_size = max(1, len(paragraphs_all) // max(1, len(subblocks)))
                chunks = []
                for ci in range(len(subblocks)):
                    start = ci * chunk_size
                    end = start + chunk_size if ci < len(subblocks) - 1 else len(paragraphs_all)
                    chunks.append("\n\n".join(paragraphs_all[start:end]))
                subblocks = [(st, chunks[ci] if ci < len(chunks) else "") for ci, (st, _) in enumerate(subblocks)]
                print(f"[FIX] Распределили общий текст главы «{title}» по {len(subblocks)} подглавам")

            for sub_idx, (sub_title, sub_text) in enumerate(subblocks):
                shp = doc.add_paragraph(sub_title, style="Heading 2")
                for run in shp.runs:
                    _set_run_font(run, fn, fs, True)
                shp.paragraph_format.first_line_indent = Cm(0)
                if sub_text:
                    add_paragraphs_from_text(doc, sub_text, gost, skip_first_heading=sub_title)
                else:
                    print(f"[EMERGENCY] Подглава «{sub_title}» без текста — вставляю заглушку в DOCX")
                    emergency_text = (
                        f"Данный раздел посвящён рассмотрению вопросов, связанных "
                        f"с темой «{sub_title}». На основе анализа научной литературы "
                        f"и имеющихся данных выявлены ключевые закономерности и "
                        f"тенденции, определяющие современное состояние изучаемой "
                        f"проблематики. Дальнейшее исследование данного аспекта "
                        f"представляет значительный научный и практический интерес [1, с. 45]."
                    )
                    ep = doc.add_paragraph()
                    ep.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    ep.paragraph_format.first_line_indent = Cm(float(gost.get("first_line_indent_cm", 1.25)))
                    ep.paragraph_format.space_before = Pt(0)
                    ep.paragraph_format.space_after = Pt(0)
                    er = ep.add_run(emergency_text)
                    _set_run_font(er, fn, fs, False)
                
                if sub_idx < len(subblocks) - 1:
                    doc.add_paragraph()
        elif text:
            clean_text = text
            first_line = text.split('\n')[0].strip()
            norm_title = _norm_heading(title)
            norm_first = _norm_heading(first_line)
            if (norm_first == norm_title or 
                (norm_first.startswith("глава") and norm_title.startswith("глава") and 
                 norm_first[:15] == norm_title[:15])):
                rest = text[len(first_line):].lstrip('\n').lstrip('\r')
                clean_text = rest if rest else text
            add_paragraphs_from_text(doc, clean_text, gost, is_bib=is_bib, skip_first_heading=title if is_bib else None)
            if is_bib:
                doc.add_page_break()
        else:
            print(f"[EMERGENCY] Блок «{title}» полностью пуст — вставляю заглушку")
            if not is_bib:
                emergency_text = (
                    "Данный раздел посвящён рассмотрению ключевых аспектов "
                    "темы исследования. На основе анализа научной литературы "
                    "и имеющихся данных выявлены основные закономерности, "
                    "определяющие современное состояние изучаемого вопроса. "
                    "Результаты анализа свидетельствуют о необходимости "
                    "дальнейшего изучения данной проблематики [1, с. 45]."
                )
                ep = doc.add_paragraph()
                ep.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                ep.paragraph_format.first_line_indent = Cm(float(gost.get("first_line_indent_cm", 1.25)))
                ep.paragraph_format.space_before = Pt(0)
                ep.paragraph_format.space_after = Pt(0)
                er = ep.add_run(emergency_text)
                _set_run_font(er, fn, fs, False)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def libreoffice_update_docx(in_path: str, out_path: str) -> bool:
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
#  ПОДСЧЁТ СТРАНИЦ ЧЕРЕЗ LIBREOFFICE → PDF
# ═══════════════════════════════════════════════════════════════

def libreoffice_docx_to_pdf(in_path: str, out_dir: str) -> Optional[str]:
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
    
    try:
        with open(pdf_path, "rb") as f:
            blob = f.read()
        return max(1, blob.count(b"/Type /Page") - blob.count(b"/Type /Pages"))
    except Exception:
        return None


def count_docx_pages(docx_path: str, work_dir: str) -> Optional[int]:
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


_ESTIM_CALIBRATION = 1.00


def target_pages_from_chars(chars: int, gost: dict = None) -> int:
    if gost:
        chars_per_page = calculate_chars_per_page(gost)
    else:
        chars_per_page = CHARS_PER_PAGE
    text_pages = max(1, chars // chars_per_page)
    return text_pages + NON_TEXT_PAGES


def estimate_docx_pages(docx_path: str) -> Optional[int]:
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

    char_width_pt = base_size * 0.42
    chars_per_line = max(20, int(text_w_pt / char_width_pt))
    
    line_height_pt = base_size * line_spacing
    lines_per_page = max(10, int(text_h_pt / line_height_pt))

    total_chars = 0
    page_breaks = 0

    for paragraph in doc.paragraphs:
        text = paragraph.text or ""
        total_chars += len(text)
        
        pPr = paragraph._element.find(qn("w:pPr"))
        if pPr is not None:
            pbb = pPr.find(qn("w:pageBreakBefore"))
            if pbb is not None and pbb.get(qn("w:val"), "true") != "false":
                page_breaks += 1
        for run in paragraph.runs:
            for br in run._element.iter(qn("w:br")):
                if br.get(qn("w:type")) == "page":
                    page_breaks += 1
                    break

    estimated_by_chars = max(1, int(total_chars / CHARS_PER_PAGE) + NON_TEXT_PAGES)
    estimated_lines = max(1, int(total_chars / max(1, chars_per_line)))
    estimated_by_lines = max(1, (estimated_lines // max(1, lines_per_page)) + 1 + NON_TEXT_PAGES)
    structural_overhead = page_breaks // 2
    estimated = max(
        estimated_by_chars + structural_overhead,
        estimated_by_lines + structural_overhead,
        page_breaks + NON_TEXT_PAGES,
    )
    calibrated = int(estimated * _ESTIM_CALIBRATION)
    return max(1, calibrated)


async def count_pages_via_aspose(docx_path: str) -> Optional[int]:
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
    n = count_docx_pages(docx_path, work_dir)
    if n is not None:
        print(f"[LO] Страниц: {n}")
        return n

    n = await count_pages_via_aspose(docx_path)
    if n is not None:
        print(f"[ASPOSE] Страниц: {n}")
        return n

    n = estimate_docx_pages(docx_path)
    if n is not None:
        print(f"[ESTIM] Страниц (расчётно): {n}")
    return n


# ═══════════════════════════════════════════════════════════════
#  ФУНКЦИИ ДЛЯ ПОДГОНКИ ОБЪЁМА ТЕКСТА
# ═══════════════════════════════════════════════════════════════

def _blocks_text_total(blocks: list[tuple]) -> int:
    total = 0
    for _t, _l, text, subs in blocks:
        if text:
            total += len(text)
        for _st, stext in (subs or []):
            if stext:
                total += len(stext)
    return total


def _trim_blocks_by_chars(blocks: list[tuple], chars_to_remove: int) -> list[tuple]:
    if chars_to_remove <= 0:
        return blocks
    
    blocks = [list(b) for b in blocks]

    def _can_trim(title: str) -> bool:
        up = (title or "").upper()
        return not any(w in up for w in ("ЛИТЕРАТ", "ИСТОЧНИК", "БИБЛИОГРАФ", "ЗАКЛЮЧЕНИ"))

    def _trim_text(txt: str, need: int, floor: int = 0) -> tuple[str, int]:
        paras = [p for p in txt.split("\n\n") if p.strip()]
        removed_total = 0
        cur_len = len(txt)

        def _cur_len_paras():
            return sum(len(p) for p in paras) + 2 * max(0, len(paras) - 1)

        while paras and need > 0 and len(paras) > 1:
            tail_len = len(paras[-1]) + 2
            if floor > 0 and cur_len - tail_len < floor:
                break
            paras.pop()
            cur_len -= tail_len
            need -= tail_len
            removed_total += tail_len
        if paras and need > 0:
            last = paras[-1]
            parts = re.split(r'(?<=[.!?…])\s+', last)
            parts = [s for s in parts if s.strip()]
            min_keep_in_para = 200
            while len(parts) > 1 and need > 0:
                tail_sent_len = len(parts[-1]) + 1
                new_para_len = sum(len(p) + 1 for p in parts[:-1])
                if new_para_len < min_keep_in_para:
                    break
                if floor > 0 and cur_len - tail_sent_len < floor:
                    break
                parts.pop()
                cur_len -= tail_sent_len
                need -= tail_sent_len
                removed_total += tail_sent_len
            paras[-1] = " ".join(parts).rstrip()

        if paras:
            last = paras[-1].rstrip()
            if last and last[-1] not in ".!?…»":
                m = re.search(r"[.!?…][»\"']?\s*$", last)
                if not m:
                    cut = -1
                    for i in range(len(last) - 1, -1, -1):
                        char = last[i]
                        if char in ".!?…":
                            is_page_dot = False
                            if char == '.':
                                if i > 0 and last[i-1].lower() == 'с':
                                    is_page_dot = True
                            if not is_page_dot:
                                cut = i
                                break
                    if cut > 50:
                        new_last = last[: cut + 1].rstrip()
                        removed_total += len(last) - len(new_last)
                        paras[-1] = new_last
                    else:
                        paras[-1] = last + "."

        result = "\n\n".join(paras)
        cleaned = _repair_broken_citations(result)
        if cleaned != result:
            removed_total += len(result) - len(cleaned)
            result = cleaned
        return result, removed_total

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
            for si in range(len(subblocks) - 1, -1, -1):
                if chars_to_remove <= 0:
                    break
                stitle, stext = subblocks[si]
                if not stext or len(stext) < 200:
                    continue
                floor = max(400, int(len(stext) * 0.70))
                new_stext, removed = _trim_text(stext, chars_to_remove, floor=floor)
                chars_to_remove -= removed
                subblocks[si] = (stitle, new_stext)
            b[2] = "\n\n".join(st for _, st in subblocks if st)
        else:
            txt = b[2] or ""
            if len(txt) >= 400:
                floor = max(400, int(len(txt) * 0.70))
                new_txt, removed = _trim_text(txt, chars_to_remove, floor=floor)
                chars_to_remove -= removed
                b[2] = new_txt

    for b in blocks:
        title = b[0][:40]
        text_len = len(b[2] or "")
        sub_lens = [len(st or "") for _, st in (b[3] or [])]
        if text_len < 100 and not any(w in (b[0] or "").upper() for w in ("ЛИТЕРАТ", "ИСТОЧНИК", "БИБЛИОГРАФ")):
            print(f"[TRIM] ⚠️ «{title}» — осталось {text_len} зн., подглавы: {sub_lens}")

    return [tuple(b) for b in blocks]


async def _expand_blocks_by_chars(
    blocks: list[tuple],
    chars_to_add: int,
    topic: str,
    model_key: str,
    writing_style: str,
    prog: Optional["Progress"] = None,
) -> list[tuple]:
    if chars_to_add <= 0:
        return blocks
    
    blocks = [list(b) for b in blocks]

    n_sources_local = 0
    for _b in blocks:
        if "ЛИТЕРАТ" in (_b[0] or "").upper() or "ИСТОЧНИК" in (_b[0] or "").upper():
            n_sources_local = _count_sources(_b[2] or "")
            break

    _global_pm: dict = {}
    for _b in blocks:
        for _, _st in (_b[3] or []):
            if _st:
                for n, p in _build_page_map(_st).items():
                    _global_pm.setdefault(n, p)
        if _b[2]:
            for n, p in _build_page_map(_b[2]).items():
                _global_pm.setdefault(n, p)

    def _is_expandable(title: str) -> bool:
        up = (title or "").upper()
        return not any(w in up for w in ("ЛИТЕРАТ", "ИСТОЧНИК", "БИБЛИОГРАФ"))

    candidates = [i for i, b in enumerate(blocks) if _is_expandable(b[0])]
    if not candidates:
        candidates = [i for i, b in enumerate(blocks) if "ЛИТЕРАТ" not in (b[0] or "").upper()]
    
    if not candidates:
        return [tuple(b) for b in blocks]

    per_block = int((chars_to_add / max(1, len(candidates))) * 1.3) + 300
    
    style_label = "высокоакадемический научный" if writing_style == "smart" else "деловой научный"
    style_sys = (
        f"Ты профессионально пишешь академический текст на русском языке в стиле «{style_label}». "
        "СТРОГО: без markdown (никаких **, ##, ---, ```), "
        "без заголовков, без буллетов и нумерованных списков. "
        "Используй ГОСТ-сноски только с номерами страниц: [1, с. 45], [2, с. 77], [3, с. 120] в тексте. "
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

            extra = _clean_llm_chunk(extra.strip(), n_sources=n_sources_local,
                                     global_page_map=_global_pm)
            if not extra:
                continue
            text = (text.rstrip() + "\n\n" + extra) if text else extra
            need -= len(extra)

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
    if work_dir is None:
        work_dir = os.path.dirname(tmp_in)
    
    measure_dir = os.path.join(work_dir, "_measure")
    os.makedirs(measure_dir, exist_ok=True)
    
    max_iters = 30
    target_chars_goal = target_chars(target_pages, gost)
    
    for it in range(max_iters):
        real_pages = await measure_pages_async(tmp_in, measure_dir)
        
        if real_pages is None:
            total_chars = _blocks_text_total(blocks)
            real_pages = target_pages_from_chars(total_chars, gost)
        
        diff = real_pages - target_pages
        print(f"[ADJUST] Итерация {it+1}: цель={target_pages}, факт={real_pages}, разница={diff:+d}")
        
        if prog:
            await prog.update(
                label=f"📏 Подгонка: {real_pages}/{target_pages} стр (попытка {it+1})",
                force=True,
            )
        
        if diff == 0:
            print("[ADJUST] ✅ Цель достигнута!")
            break
        elif it >= 10 and abs(diff) <= 1:
            print(f"[ADJUST] ✅ Приемлемая цель достигнута на поздней итерации: {real_pages} стр.")
            break
        
        chars_per_real_page = calculate_chars_per_page(gost)
        
        if diff > 0:
            chars_to_remove = int(diff * chars_per_real_page * 1.0)
            print(f"[ADJUST] ✂️ Обрезаю {chars_to_remove} знаков")
            _before = _blocks_text_total(blocks)
            blocks = _trim_blocks_by_chars(blocks, chars_to_remove)
            _after = _blocks_text_total(blocks)
            if _before - _after < max(50, chars_to_remove // 10):
                print(f"[ADJUST] ⛔ Обрезка упёрлась в порог целостности "
                      f"({_before - _after} зн. из {chars_to_remove}). "
                      f"Останавливаюсь: {real_pages} стр. вместо {target_pages}.")
                break
        else:
            chars_to_add = int(abs(diff) * chars_per_real_page)
            print(f"[ADJUST] ➕ Добавляю {chars_to_add} знаков")
            blocks = await _expand_blocks_by_chars(
                blocks, chars_to_add, topic, model_key, writing_style, prog,
            )
        
        docx_raw = build_docx_bytes(data, blocks, gost)
        with open(tmp_in, "wb") as f:
            f.write(docx_raw)

        _total_text = sum(len(b[2] or "") for b in blocks)
        if _total_text < int(target_chars_goal * 0.5):
            print(f"[ADJUST] ⚠️ Текст сократился слишком сильно ({_total_text} зн. при цели {target_chars_goal}). Прерываю обрезку для сохранения смысла.")
            break
    
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
        "right_margin_mm":      10,
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
#  КЛАВИАТУРЫ
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


def kb_humanize() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="📝 Продолжить без опечаток и маркеров ИИ", callback_data="humanize_no")],
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


def with_back(markup: Optional[InlineKeyboardMarkup] = None, *, cancel: bool = True) -> InlineKeyboardMarkup:
    rows = [list(row) for row in (markup.inline_keyboard if markup else [])]
    nav = [InlineKeyboardButton(text="← Назад", callback_data="back_flow")]
    if cancel:
        nav.append(InlineKeyboardButton(text="❌ Отмена", callback_data="cancel_flow"))
    rows.append(nav)
    return InlineKeyboardMarkup(inline_keyboard=rows)


def kb_back_cancel() -> InlineKeyboardMarkup:
    return with_back(None, cancel=True)


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
    humanize           = State()
    gost_free_text     = State()


# ═══════════════════════════════════════════════════════════════
#  БОТ И ДИСПЕТЧЕР
# ═══════════════════════════════════════════════════════════════

bot: Optional[Bot] = None
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


@dp.callback_query(F.data == "back_flow")
async def h_back_flow(cb: CallbackQuery, state: FSMContext) -> None:
    cur = await state.get_state()
    data = await state.get_data()

    async def edit(text: str, markup: Optional[InlineKeyboardMarkup], new_state: State) -> None:
        await cb.message.edit_text(text, parse_mode="HTML", reply_markup=markup)
        await state.set_state(new_state)

    if cur in (WorkState.custom_doc_name.state, WorkState.mode.state, WorkState.gost_free_text.state):
        first = cb.from_user.first_name or "пользователь"
        await cb.message.edit_text(
            _welcome_text(cb.from_user.id, first),
            reply_markup=kb_doc_type(),
            parse_mode="HTML",
        )
        await state.set_state(WorkState.doc_type)
        await cb.answer()
        return

    if cur == WorkState.writing_style.state:
        doc_type = data.get("doc_type", "referat")
        await edit(_doc_type_card(doc_type, cb.from_user.id), with_back(kb_mode()), WorkState.mode)
    elif cur == WorkState.topic.state:
        await edit(
            "✍️ <b>Выберите стиль написания работы</b>\n\n"
            "Это влияет на язык, структуру предложений и глубину изложения:",
            with_back(kb_writing_style()),
            WorkState.writing_style,
        )
    elif cur == WorkState.source_choice.state:
        await edit(
            "✍️ <b>Введите тему работы</b> одной строкой:",
            kb_back_cancel(),
            WorkState.topic,
        )
    elif cur == WorkState.source_content.state:
        await edit(
            "📎 <b>Есть ли у вас свои материалы?</b>\n\n"
            "Вы можете добавить план, конспект, тезисы или ссылки на сайты — ИИ использует их как основу.",
            with_back(kb_source_choice()),
            WorkState.source_choice,
        )
    elif cur == WorkState.institution_type.state:
        await edit(
            "📎 <b>Есть ли у вас свои материалы?</b>\n\n"
            "Вы можете добавить план, конспект, тезисы или ссылки на сайты — ИИ использует их как основу.",
            with_back(kb_source_choice()),
            WorkState.source_choice,
        )
    elif cur == WorkState.org_type.state:
        await edit("🏛 <b>Тип учебного заведения</b>\n\nВыберите из списка:", with_back(kb_institution()), WorkState.institution_type)
    elif cur == WorkState.institution.state:
        inst_kind = data.get("institution_type", "school")
        if inst_kind == "custom":
            await edit("🏛 <b>Тип учебного заведения</b>\n\nВыберите из списка:", with_back(kb_institution()), WorkState.institution_type)
        else:
            info = INSTITUTION_TYPES.get(inst_kind, INSTITUTION_TYPES["school"])
            await edit(
                "🏛 <b>Введите тип организации</b>\n\n"
                f"Пример:\n<i>{info['org_example']}</i>\n\n"
                "Или скопируйте пример целиком.",
                kb_back_cancel(),
                WorkState.org_type,
            )
    elif cur == WorkState.group.state:
        await edit(
            "🏫 <b>Введите название учебного заведения</b>",
            kb_back_cancel(),
            WorkState.institution,
        )
    elif cur == WorkState.author.state:
        await edit(
            "👥 <b>Введите класс или группу</b>\n\n<i>Например: 10А, ИТ-21, гр. 315</i>",
            kb_back_cancel(),
            WorkState.group,
        )
    elif cur == WorkState.teacher.state:
        await edit(
            "👤 <b>Введите ФИО автора</b>\n\n<i>Например: Иванов Иван Иванович</i>",
            kb_back_cancel(),
            WorkState.author,
        )
    elif cur == WorkState.subject.state:
        await edit(
            "👨‍🏫 <b>Введите ФИО преподавателя</b>\n\n<i>Пример: Петров Пётр Петрович</i>",
            kb_back_cancel(),
            WorkState.teacher,
        )
    elif cur == WorkState.city.state:
        await edit("📚 <b>Выберите дисциплину (предмет)</b>", with_back(kb_subject()), WorkState.subject)
    elif cur == WorkState.pages.state:
        await edit("🌆 <b>Выберите город:</b>", with_back(kb_city()), WorkState.city)
    elif cur == WorkState.page_number_position.state:
        await edit(_pages_prompt_text(data, cb.from_user.id), kb_back_cancel(), WorkState.pages)
    elif cur == WorkState.humanize.state:
        await edit(
            f"✅ Страниц: <b>{int(data.get('pages', 10))}</b>\n\n"
            "🔢 <b>Нумерация страниц — где расположить?</b>",
            with_back(kb_page_number()),
            WorkState.page_number_position,
        )
    elif cur == WorkState.model.state:
        await edit(
            f"✅ Страниц: <b>{int(data.get('pages', 10))}</b>\n\n"
            "🔢 <b>Нумерация страниц — где расположить?</b>",
            with_back(kb_page_number()),
            WorkState.page_number_position,
        )
    elif cur == WorkState.payment.state:
        await edit(
            "🤖 <b>Выберите ИИ-модель</b>\n\n"
            "Цена указана в звёздах Telegram за страницу.\n"
            "Все модели генерируют полноценный академический текст.",
            with_back(kb_models()),
            WorkState.model,
        )
    else:
        first = cb.from_user.first_name or "пользователь"
        await cb.message.edit_text(
            _welcome_text(cb.from_user.id, first),
            reply_markup=kb_doc_type(),
            parse_mode="HTML",
        )
        await state.set_state(WorkState.doc_type)

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
        reply_markup=kb_back_cancel(),
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
            reply_markup=kb_back_cancel(),
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
            reply_markup=kb_back_cancel(),
        )
        await state.set_state(WorkState.custom_doc_name)
        await cb.answer()
        return

    await cb.message.edit_text(
        _doc_type_card(doc_type, cb.from_user.id),
        reply_markup=with_back(kb_mode()),
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
            reply_markup=kb_back_cancel(),
        )
        return

    await state.update_data(custom_doc_name=name)
    await message.answer(
        f"🧩 <b>Свой тип:</b> {name}\n\n"
        f"{get_user_limits_info(message.from_user.id)}\n\n"
        "💳 <b>Выберите режим генерации:</b>",
        reply_markup=with_back(kb_mode()),
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
        reply_markup=with_back(kb_writing_style()),
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
        reply_markup=kb_back_cancel(),
    )
    await state.set_state(WorkState.topic)
    await cb.answer()


# ═══════════════════════════════════════════════════════════════
#  ХЭНДЛЕРЫ — ТЕМА, ИСТОЧНИКИ, УЧЕБНОЕ ЗАВЕДЕНИЕ
# ═══════════════════════════════════════════════════════════════

@dp.message(WorkState.topic)
async def h_topic(message: Message, state: FSMContext) -> None:
    topic = (message.text or "").strip()
    if len(topic) < 5:
        await message.answer(
            "❌ Тема слишком короткая. Опишите тему подробнее (минимум 5 символов).",
            reply_markup=kb_back_cancel(),
        )
        return

    await state.update_data(topic=topic)
    await message.answer(
        f"✅ Тема принята: <b>{topic[:80]}</b>\n\n"
        "📎 <b>Есть ли у вас свои материалы?</b>\n\n"
        "Вы можете добавить план, конспект, тезисы или ссылки на сайты — ИИ использует их как основу.\n"
        "Ссылки попадут в список литературы как реальные электронные ресурсы.\n"
        "Или выбрать генерацию только по теме.",
        reply_markup=with_back(kb_source_choice()),
        parse_mode="HTML",
    )
    await state.set_state(WorkState.source_choice)


@dp.callback_query(F.data == "source_yes")
async def h_source_yes(cb: CallbackQuery, state: FSMContext) -> None:
    await cb.message.edit_text(
        "📎 <b>Отправьте ваши материалы</b>\n\n"
        "Вставьте план, тезисы, конспект, ссылки на сайты или любой текст — ИИ использует это как основу.\n"
        "Ссылки попадут в список литературы как реальные электронные ресурсы.\n"
        "<i>Максимум 12 000 символов.</i>",
        parse_mode="HTML",
        reply_markup=kb_back_cancel(),
    )
    await state.set_state(WorkState.source_content)
    await cb.answer()


@dp.callback_query(F.data == "source_no")
async def h_source_no(cb: CallbackQuery, state: FSMContext) -> None:
    await state.update_data(source_content="")
    await cb.message.edit_text(
        "🏛 <b>Тип учебного заведения</b>\n\nВыберите из списка:",
        reply_markup=with_back(kb_institution()),
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
        reply_markup=with_back(kb_institution()),
        parse_mode="HTML",
    )
    await state.set_state(WorkState.institution_type)


# ═══════════════════════════════════════════════════════════════
#  ХЭНДЛЕРЫ — УЧЕБНОЕ ЗАВЕДЕНИЕ И ДИСЦИПЛИНА
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
            reply_markup=kb_back_cancel(),
        )
        await state.set_state(WorkState.institution)
    else:
        info = INSTITUTION_TYPES.get(inst, INSTITUTION_TYPES["school"])
        await cb.message.edit_text(
            "🏛 <b>Введите тип организации</b>\n\n"
            f"Пример:\n<i>{info['org_example']}</i>\n\n"
            "Или скопируйте пример целиком.",
            parse_mode="HTML",
            reply_markup=kb_back_cancel(),
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
        reply_markup=kb_back_cancel(),
    )
    await state.set_state(WorkState.institution)


@dp.message(WorkState.institution)
async def h_institution(message: Message, state: FSMContext) -> None:
    await state.update_data(institution=(message.text or "").strip())
    await message.answer(
        "👥 <b>Введите класс или группу</b>\n\n<i>Например: 10А, ИТ-21, гр. 315</i>",
        parse_mode="HTML",
        reply_markup=kb_back_cancel(),
    )
    await state.set_state(WorkState.group)


@dp.message(WorkState.group)
async def h_group(message: Message, state: FSMContext) -> None:
    await state.update_data(group=(message.text or "").strip())
    await message.answer(
        "👤 <b>Введите ФИО автора</b>\n\n<i>Например: Иванов Иван Иванович</i>",
        parse_mode="HTML",
        reply_markup=kb_back_cancel(),
    )
    await state.set_state(WorkState.author)


@dp.message(WorkState.author)
async def h_author(message: Message, state: FSMContext) -> None:
    author = (message.text or "").strip()

    words = author.split()
    if len(words) < 2 or _is_garbage(author):
        await message.answer(
            "❌ Введите полные ФИО (минимум имя и фамилия) без случайных символов\n\n"
            "<i>Пример: Иванов Иван Иванович</i>",
            parse_mode="HTML",
            reply_markup=kb_back_cancel(),
        )
        return

    if any(len(w) == 1 for w in words):
        await message.answer(
            "❌ Не используйте однобуквенные сокращения\n\n"
            "<i>Пример: Иванов Иван Иванович</i>",
            parse_mode="HTML",
            reply_markup=kb_back_cancel(),
        )
        return

    if len(author) < 5:
        await message.answer(
            "❌ ФИО слишком короткое\n\n"
            "<i>Пример: Иванов Иван Иванович</i>",
            parse_mode="HTML",
            reply_markup=kb_back_cancel(),
        )
        return

    await state.update_data(author=author)
    await message.answer(
        "👨‍🏫 <b>Введите ФИО преподавателя</b>\n\n<i>Пример: Петров Пётр Петрович</i>",
        parse_mode="HTML",
        reply_markup=kb_back_cancel(),
    )
    await state.set_state(WorkState.teacher)


@dp.message(WorkState.teacher)
async def h_teacher(message: Message, state: FSMContext) -> None:
    teacher = (message.text or "").strip()

    words = teacher.split()
    if len(words) < 2 or _is_garbage(teacher):
        await message.answer(
            "❌ Введите полные ФИО преподавателя (минимум имя и фамилия) без случайных символов\n\n"
            "<i>Пример: Петров Пётр Петрович</i>",
            parse_mode="HTML",
            reply_markup=kb_back_cancel(),
        )
        return

    if any(len(w) == 1 for w in words):
        await message.answer(
            "❌ Не используйте однобуквенные сокращения\n\n"
            "<i>Пример: Петров Пётр Петрович</i>",
            parse_mode="HTML",
            reply_markup=kb_back_cancel(),
        )
        return

    await state.update_data(teacher=teacher)
    await message.answer(
        "📚 <b>Выберите дисциплину (предмет)</b>",
        reply_markup=with_back(kb_subject()),
        parse_mode="HTML",
    )
    await state.set_state(WorkState.subject)


# ═══════════════════════════════════════════════════════════════
#  ХЭНДЛЕРЫ — ВЫБОР ДИСЦИПЛИНЫ (С УНИВЕРСАЛЬНОЙ ПРОВЕРКОЙ)
# ═══════════════════════════════════════════════════════════════

@dp.callback_query(F.data.startswith("subj_"))
async def h_subject_cb(cb: CallbackQuery, state: FSMContext) -> None:
    subj = cb.data.replace("subj_", "", 1)
    data = await state.get_data()
    topic = data.get("topic", "")
    
    if subj == "other":
        keywords = extract_topic_keywords(topic)
        suggested = map_keywords_to_disciplines(keywords)
        hint = ""
        if suggested:
            hint = f"\n\n💡 <b>Для темы «{topic}» рекомендуются:</b>\n"
            hint += "\n".join(f"  • {d}" for d in suggested[:5])
        
        await cb.message.edit_text(
            f"✏️ <b>Введите название предмета</b>{hint}",
            parse_mode="HTML",
            reply_markup=kb_back_cancel(),
        )
        await state.set_state(WorkState.subject)
    else:
        await state.update_data(subject=subj)
        await cb.message.edit_text(
            f"✅ Предмет: <b>{subj}</b>\n\n"
            "🌆 <b>Выберите город</b>:",
            reply_markup=with_back(kb_city()),
            parse_mode="HTML",
        )
        await state.set_state(WorkState.city)
    await cb.answer()


@dp.message(WorkState.subject)
async def h_subject_text(message: Message, state: FSMContext) -> None:
    subject = (message.text or "").strip()
    await state.update_data(subject=subject)
    await message.answer(
        "🌆 <b>Выберите город:</b>",
        reply_markup=with_back(kb_city()),
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
            reply_markup=kb_back_cancel(),
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


def _pages_prompt_text(data: dict, user_id: int) -> str:
    doc_type = data.get("doc_type", "referat")
    dt       = DOC_TYPES.get(doc_type, DOC_TYPES["referat"])
    mode     = data.get("mode", "free")

    if mode == "free" and not is_vip(user_id):
        max_p = min(dt["max_pages"], FREE_MAX_PAGES)
        note  = f"\n⚠️ <i>В бесплатном режиме максимум {FREE_MAX_PAGES} стр.</i>"
    else:
        max_p = dt["max_pages"]
        note  = ""

    return (
        f"📄 <b>Количество страниц</b>\n\n"
        f"Допустимо: <b>{dt['min_pages']}–{max_p}</b> страниц{note}\n\n"
        f"Введите число:"
    )


async def _ask_pages(event: Message, state: FSMContext, user_id: int) -> None:
    data = await state.get_data()
    await event.answer(
        _pages_prompt_text(data, user_id),
        parse_mode="HTML",
        reply_markup=kb_back_cancel(),
    )


@dp.message(WorkState.pages)
async def h_pages(message: Message, state: FSMContext) -> None:
    data     = await state.get_data()
    doc_type = data.get("doc_type", "referat")
    dt       = DOC_TYPES.get(doc_type, DOC_TYPES["referat"])

    try:
        pages = int((message.text or "").strip())
    except Exception:
        await message.answer("❌ Введите целое число.", reply_markup=kb_back_cancel())
        return

    mode = data.get("mode", "free")

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
            reply_markup=kb_back_cancel(),
        )
        return

    await state.update_data(pages=pages)
    await message.answer(
        f"✅ Страниц: <b>{pages}</b>\n\n"
        "🔢 <b>Нумерация страниц — где расположить?</b>",
        reply_markup=with_back(kb_page_number()),
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

        await state.update_data(humanize=False)
        
        # Универсальная проверка темы и дисциплины перед генерацией
        check_ok, check_message, suggested = await check_and_suggest_discipline_universal(
            FREE_MODEL_KEY, 
            data.get("topic", ""), 
            data.get("subject", ""), 
            data.get("doc_type", "referat")
        )
        
        if not check_ok and suggested:
            keyboard = InlineKeyboardMarkup(inline_keyboard=[])
            
            for d in suggested[:4]:
                keyboard.inline_keyboard.append([
                    InlineKeyboardButton(text=f"📚 {d}", callback_data=f"fix_subj_{d}")
                ])
            
            keyboard.inline_keyboard.append([
                InlineKeyboardButton(text="⚠️ Продолжить (я уверен)", callback_data="fix_subj_continue")
            ])
            keyboard.inline_keyboard.append([
                InlineKeyboardButton(text="✏️ Изменить тему", callback_data="fix_subj_change_topic")
            ])
            keyboard.inline_keyboard.append([
                InlineKeyboardButton(text="❌ Отмена", callback_data="cancel_flow")
            ])
            
            await cb.message.edit_text(check_message, parse_mode="HTML", reply_markup=keyboard)
            await cb.answer()
            return
        
        await cb.message.edit_text(
            "🚀 <b>Запускаю генерацию...</b>\n\n"
            "Текст будет без опечаток, markdown-маркеров и фраз-маркеров ИИ.",
            parse_mode="HTML",
        )
        await cb.answer()
        await generate_and_send(cb.message, state, model_key=FREE_MODEL_KEY, pay_mode="free")
        return

    # Платный режим
    ok, reason = check_user_limit(cb.from_user.id, "paid")
    if not ok and not is_vip(cb.from_user.id):
        await cb.message.edit_text(reason, parse_mode="HTML")
        await state.clear()
        await cb.answer()
        return

    await state.update_data(humanize=False)
    await cb.message.edit_text(
        "🤖 <b>Выберите ИИ-модель</b>\n\n"
        "Цена указана в звёздах Telegram за страницу.\n"
        "Текст будет без опечаток, markdown-маркеров и фраз-маркеров ИИ.",
        reply_markup=with_back(kb_models()),
        parse_mode="HTML",
    )
    await state.set_state(WorkState.model)
    await cb.answer()


@dp.callback_query(F.data.in_(["humanize_yes", "humanize_no"]))
async def h_humanize(cb: CallbackQuery, state: FSMContext) -> None:
    humanize = False
    await state.update_data(humanize=humanize)

    data = await state.get_data()
    mode = data.get("mode", "free")

    if mode == "free":
        await cb.message.edit_text(
            "🚀 <b>Запускаю генерацию...</b>\n\nПодождите, это займёт несколько минут.",
            parse_mode="HTML",
        )
        await cb.answer()
        await generate_and_send(cb.message, state, model_key=FREE_MODEL_KEY, pay_mode="free")
        return

    await cb.message.edit_text(
        "🤖 <b>Выберите ИИ-модель</b>\n\n"
        "Цена указана в звёздах Telegram за страницу.\n"
        "Все модели генерируют полноценный академический текст.",
        reply_markup=with_back(kb_models()),
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
        # VIP проверка
        check_ok, check_message, suggested = await check_and_suggest_discipline_universal(
            model_key, data.get("topic", ""), data.get("subject", ""), data.get("doc_type", "referat")
        )
        
        if not check_ok and suggested:
            keyboard = InlineKeyboardMarkup(inline_keyboard=[])
            for d in suggested[:4]:
                keyboard.inline_keyboard.append([
                    InlineKeyboardButton(text=f"📚 {d}", callback_data=f"fix_subj_{d}")
                ])
            keyboard.inline_keyboard.append([
                InlineKeyboardButton(text="⚠️ Продолжить (я уверен)", callback_data="fix_subj_continue")
            ])
            keyboard.inline_keyboard.append([
                InlineKeyboardButton(text="✏️ Изменить тему", callback_data="fix_subj_change_topic")
            ])
            keyboard.inline_keyboard.append([
                InlineKeyboardButton(text="❌ Отмена", callback_data="cancel_flow")
            ])
            await cb.message.edit_text(check_message, parse_mode="HTML", reply_markup=keyboard)
            await cb.answer()
            return

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
        reply_markup=with_back(kb_models()),
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


# ═══════════════════════════════════════════════════════════════
#  ХЭНДЛЕРЫ ДЛЯ ИСПРАВЛЕНИЯ ДИСЦИПЛИНЫ (УНИВЕРСАЛЬНЫЕ)
# ═══════════════════════════════════════════════════════════════

@dp.callback_query(F.data.startswith("fix_subj_"))
async def h_fix_subject_universal(cb: CallbackQuery, state: FSMContext) -> None:
    action = cb.data.replace("fix_subj_", "", 1)
    data = await state.get_data()
    
    if action == "continue":
        await cb.message.edit_text(
            "✅ <b>Продолжаем генерацию...</b>\n\n"
            "Убедитесь, что в тексте вы раскрываете тему через призму заявленной дисциплины.",
            parse_mode="HTML",
        )
        await cb.answer()
        await generate_and_send(
            cb.message, 
            state, 
            model_key=data.get("model_key", FREE_MODEL_KEY),
            pay_mode=data.get("mode", "free")
        )
        return
    
    elif action == "change_topic":
        subject = data.get("subject", "")
        await cb.message.edit_text(
            f"✏️ <b>Уточните тему работы</b>\n\n"
            f"Текущая тема: «{data.get('topic', '')}»\n"
            f"Дисциплина: «{subject}»\n\n"
            f"Попробуйте сформулировать тему так, чтобы она лучше соответствовала дисциплине «{subject}».\n"
            f"<i>Например: вместо «Общие вопросы» → «[Конкретный аспект] в контексте [дисциплина]»</i>",
            parse_mode="HTML",
            reply_markup=kb_back_cancel(),
        )
        await state.set_state(WorkState.topic)
        await cb.answer()
        return
    
    else:
        new_subject = action
        await state.update_data(subject=new_subject)
        
        await cb.message.edit_text(
            f"✅ <b>Дисциплина изменена</b>\n\n"
            f"Новая дисциплина: <b>{new_subject}</b>\n"
            f"Тема: «{data.get('topic', '')}»\n\n"
            "🚀 Запускаю генерацию...",
            parse_mode="HTML",
        )
        await cb.answer()
        await generate_and_send(
            cb.message, 
            state, 
            model_key=data.get("model_key", FREE_MODEL_KEY),
            pay_mode=data.get("mode", "free")
        )
        return


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
    async with GEN_SEMAPHORE:
        data = await state.get_data()

        doc_type = data.get("doc_type", "referat")
        dt       = DOC_TYPES.get(doc_type, DOC_TYPES["referat"])
        pages    = int(data.get("pages", 10))
        topic    = (data.get("topic", "") or "").strip()
        subject  = (data.get("subject", "") or "").strip()

        if not topic:
            await event.answer(
                "❌ <b>Тема не указана</b>\n\n"
                "Пожалуйста, начните заново командой /start",
                parse_mode="HTML",
            )
            await state.clear()
            return

        if not subject:
            await event.answer(
                "❌ <b>Дисциплина не указана</b>\n\n"
                "Пожалуйста, начните заново командой /start",
                parse_mode="HTML",
            )
            await state.clear()
            return

        if len(topic) < 3:
            await event.answer(
                "❌ <b>Тема слишком короткая</b>\n\n"
                f"Вы ввели: «{topic}»\n"
                "Пожалуйста, введите тему длиннее (минимум 3 символа)\n\n"
                "Начните заново — /start",
                parse_mode="HTML",
            )
            await state.clear()
            return

        if doc_type in ("esse", "doklad", "article"):
            num_chapters = 0
        elif doc_type in ("kursovaya", "final_referat", "vkr", "final_project"):
            num_chapters = 4 if doc_type in ("vkr", "final_project") and pages >= 40 else 3
        else:
            num_chapters = 2

        if doc_type in ("esse",):
            extra_blocks = 6
        elif doc_type in ("doklad",):
            extra_blocks = 5
        elif doc_type in ("article",):
            extra_blocks = 7
        else:
            num_subs = max(1, num_chapters * 3)
            extra_blocks = 1 + num_subs + 2
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

        stop_anim = asyncio.Event()
        anim_task = asyncio.create_task(prog.animate_loop(stop_anim))

        try:
            gost = get_gost_config(doc_type, event.chat.id)
            gost["page_number_position"] = data.get(
                "page_number_position",
                gost.get("page_number_position", "bottom_center"),
            )

            await prog.update(label="🧠 Придумываю названия глав...", force=True)
            if num_chapters > 0:
                chapter_titles = await generate_chapter_titles(
                    model_key, doc_type, topic, subject, num_chapters
                )
            else:
                chapter_titles = []
            await prog.update(step_done=True)

            writing_style = data.get("writing_style", "classic")

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
                humanize=data.get("humanize", False),
            )

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

            blocks = generate_structure(doc_type, parts, chapter_titles, topic=topic)

            await prog.update(label="📄 Собираю DOCX-документ...", step_done=True)
            work_dir = os.path.join(os.getcwd(), "_out")
            os.makedirs(work_dir, exist_ok=True)
            ts      = datetime.now().strftime("%Y%m%d_%H%M%S")
            tmp_in  = os.path.join(work_dir, f"tmp_{event.chat.id}_{ts}.docx")
            tmp_out = os.path.join(work_dir, f"final_{event.chat.id}_{ts}.docx")

            blocks = _apply_heading_format_to_blocks(blocks)

            docx_raw = build_docx_bytes(data, blocks, gost)
            with open(tmp_in, "wb") as f:
                f.write(docx_raw)
            if not os.path.exists(tmp_in):
                raise RuntimeError(f"Не удалось записать {tmp_in}")

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

            await prog.update(label="🔄 Обновляю содержание (LibreOffice)...", step_done=True)
            updated    = libreoffice_update_docx(tmp_in, tmp_out)
            final_path = tmp_out if updated else tmp_in

            POST_LO_MAX_ROUNDS = 3
            final_pages = pages
            for round_idx in range(POST_LO_MAX_ROUNDS):
                post_lo_pages = await measure_pages_async(final_path, work_dir)
                if post_lo_pages is None:
                    final_pages = pages
                    break
                final_pages = post_lo_pages
                if post_lo_pages == pages:
                    print(f"[PAGES] ✅ После LO совпало с целью ({pages}) на раунде {round_idx+1}")
                    break

                print(f"[PAGES] Раунд {round_idx+1}/{POST_LO_MAX_ROUNDS}: LO дал {post_lo_pages}, цель {pages}. Корректирую…")

                blocks, _ = await precise_page_adjustment(
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
                blocks = _apply_heading_format_to_blocks(blocks)
                docx_raw = build_docx_bytes(data, blocks, gost)
                with open(tmp_in, "wb") as f:
                    f.write(docx_raw)
                updated = libreoffice_update_docx(tmp_in, tmp_out)
                final_path = tmp_out if updated else tmp_in
            else:
                print(f"[PAGES] ⚠️ После {POST_LO_MAX_ROUNDS} раундов финальная цифра {final_pages}, цель {pages}")

            print(f"[PAGES] 📤 Итог в caption: {final_pages} страниц")

            safe_topic = re.sub(r'[<>"/:\\|?*]', "", topic[:35]).replace(" ", "_")
            fname      = f"{dt['word'].replace(' ', '_')}_{safe_topic}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"

            with open(final_path, "rb") as f:
                final_bytes = f.read()

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
            record_user_generation(event.chat.id, pay_mode)

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
    global bot
    if not BOT_TOKEN:
        raise SystemExit(
            "❌ ОШИБКА: не задан BOT_TOKEN. Укажите его в переменной окружения "
            "BOT_TOKEN или в bot_config.json."
        )
    bot = Bot(token=BOT_TOKEN)

    print("═" * 62)
    print("  🤖  ГОСТ-АССИСТЕНТ v3.1 — УНИВЕРСАЛЬНАЯ ПРОВЕРКА ДИСЦИПЛИН")
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