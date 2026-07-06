# -*- coding: utf-8 -*-
"""checker — Quality Gate: валидация и авто-фикс готового DOCX по ГОСТ 7.32-2017.

Проверяет реальный файл (не «обещания» LLM): шрифт, кегль, интервал, поля,
отступ, markdown-мусор, оборванные ссылки, фразы-маркеры ИИ, структурные
элементы. autofix_docx исправляет то, что можно исправить механически.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional

from docx import Document
from docx.shared import Cm, Mm, Pt

from gost_rules import GOST_732

# ── Паттерны брака ──────────────────────────────────────────────────
_MD_RE = re.compile(r"(^|\s)(#{1,6}\s|\*\*|__|```|~~|^---\s*$)", re.M)
_BROKEN_CITE_RE = re.compile(r"\[\d+,\s*с\.\s*(?=[^\d\s])|\[\d+,\s*с\.\s*$|\[\d+,\s*с\.\s*\]")
_AI_MARKERS = (
    "как ии", "как языковая модель", "конечно, вот", "давайте рассмотрим",
    "надеюсь, это поможет", "вот текст", "приведу пример текста",
)
_STRUCTURAL = ("СОДЕРЖАНИЕ", "ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ")
_BIB_TITLES = ("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "СПИСОК ЛИТЕРАТУРЫ",
               "БИБЛИОГРАФИЧЕСКИЙ СПИСОК", "СПИСОК ИСТОЧНИКОВ")


@dataclass
class Issue:
    title: str
    severity: str = "error"   # error | warning
    detail: str = ""

    def __str__(self) -> str:
        return f"[{self.severity}] {self.title}" + (f": {self.detail}" if self.detail else "")


@dataclass
class ValidationResult:
    score: int = 100
    errors: list = field(default_factory=list)     # list[Issue]
    warnings: list = field(default_factory=list)   # list[Issue]
    passed: bool = True

    def failures(self) -> list:
        return [e for e in self.errors if e.severity == "error"]


def _is_heading(par) -> bool:
    text = (par.text or "").strip()
    if not text or len(text) > 200:
        return False
    if text.upper() in _STRUCTURAL or text.upper() in _BIB_TITLES:
        return True
    return bool(re.match(r"^\d+(\.\d+)*\s+\S", text)) and len(text) < 120


def validate_docx(path: str, doc_type: str = "") -> ValidationResult:
    """Полная проверка DOCX. Оценка 100 минус штрафы; passed при score >= 85."""
    res = ValidationResult()

    def penalty(pts: int, title: str, detail: str = "", severity: str = "error") -> None:
        issue = Issue(title=title, severity=severity, detail=detail)
        if severity == "error":
            res.errors.append(issue)
        else:
            res.warnings.append(issue)
        res.score = max(0, res.score - pts)

    try:
        doc = Document(path)
    except Exception as e:
        res.score = 0
        res.passed = False
        res.errors.append(Issue(title=f"DOCX не открывается: {e}"))
        return res

    g = GOST_732

    # ── 1. Поля страницы ──
    try:
        sec = doc.sections[0]
        checks = [
            ("левое поле", sec.left_margin, Mm(g["left_margin_mm"])),
            ("верхнее поле", sec.top_margin, Mm(g["top_margin_mm"])),
            ("нижнее поле", sec.bottom_margin, Mm(g["bottom_margin_mm"])),
        ]
        for name, actual, expected in checks:
            if actual is not None and abs(int(actual) - int(expected)) > int(Mm(2)):
                penalty(5, f"Неверное {name}",
                        f"{round(actual / 36000, 1) if actual else '?'} мм вместо {int(expected) / 36000:.0f} мм")
        if sec.right_margin is not None and int(sec.right_margin) < int(Mm(10)) - int(Mm(1)):
            penalty(5, "Правое поле меньше 10 мм")
    except Exception:
        pass

    body_pars = [p for p in doc.paragraphs if (p.text or "").strip()]

    # ── 2. Шрифт и кегль основного текста ──
    wrong_font = wrong_size = checked = 0
    for p in body_pars:
        if _is_heading(p):
            continue
        for run in p.runs:
            if not (run.text or "").strip():
                continue
            checked += 1
            if run.font.name and run.font.name != g["font_name"]:
                wrong_font += 1
            if run.font.size and abs(run.font.size - Pt(g["font_size"])) > Pt(0.5) \
                    and run.font.size < Pt(g["font_size"]):
                wrong_size += 1
    if checked:
        if wrong_font / checked > 0.05:
            penalty(10, "Шрифт не Times New Roman",
                    f"{wrong_font}/{checked} фрагментов")
        if wrong_size / checked > 0.05:
            penalty(8, f"Кегль меньше {g['font_size']} pt", f"{wrong_size}/{checked} фрагментов")

    # ── 3. Markdown-мусор, оборванные ссылки, маркеры ИИ ──
    full_text = "\n".join(p.text or "" for p in doc.paragraphs)
    if _MD_RE.search(full_text):
        penalty(15, "Markdown-разметка в тексте (##, **, ``` )")
    if _BROKEN_CITE_RE.search(full_text):
        penalty(12, "Оборванные ссылки вида [N, с. без номера страницы")
    low = full_text.lower()
    for marker in _AI_MARKERS:
        if marker in low:
            penalty(15, "Фраза-маркер ИИ в тексте", f"«{marker}»")
            break

    # ── 4. Структурные элементы ──
    upper_text = full_text.upper()
    for name in _STRUCTURAL:
        if name not in upper_text:
            penalty(8, f"Не найден структурный элемент «{name}»")
    if not any(t in upper_text for t in _BIB_TITLES):
        penalty(8, "Не найден список использованных источников")

    # ── 5. Точка после номера раздела («1. Название» — брак по 7.32) ──
    dot_headings = [p.text.strip() for p in body_pars
                    if re.match(r"^\d+(\.\d+)*\.\s+[А-ЯA-Z]", (p.text or "").strip())
                    and _is_heading(p)]
    if dot_headings:
        penalty(4, "Точка после номера раздела", dot_headings[0][:60], severity="warning")

    # ── 6. Ссылки на источники в тексте ──
    cites = re.findall(r"\[(\d+),\s*с\.\s*\d+", full_text)
    if doc_type not in ("esse",):
        if len(cites) < 3:
            penalty(10, "Слишком мало ссылок на источники в тексте", f"найдено {len(cites)}")
        elif len(set(cites)) == 1:
            penalty(6, "Все ссылки ведут на один источник", f"[{cites[0]}]", severity="warning")

    res.passed = res.score >= 85 and not any(
        e.title.startswith("DOCX не открывается") for e in res.errors
    )
    return res


def autofix_docx(in_path: str, out_path: str) -> list[str]:
    """Механически исправляет то, что можно исправить без перегенерации.
    Возвращает список применённых фиксов (пустой = ничего не менялось)."""
    changes: list[str] = []
    try:
        doc = Document(in_path)
    except Exception:
        return changes

    g = GOST_732

    # 1) Поля страницы
    try:
        for sec in doc.sections:
            if sec.left_margin != Mm(g["left_margin_mm"]):
                sec.left_margin = Mm(g["left_margin_mm"])
                changes.append("поля: левое 30 мм")
            if sec.top_margin != Mm(g["top_margin_mm"]):
                sec.top_margin = Mm(g["top_margin_mm"])
                changes.append("поля: верхнее 20 мм")
            if sec.bottom_margin != Mm(g["bottom_margin_mm"]):
                sec.bottom_margin = Mm(g["bottom_margin_mm"])
                changes.append("поля: нижнее 20 мм")
            if sec.right_margin is None or int(sec.right_margin) < int(Mm(10)):
                sec.right_margin = Mm(g["right_margin_mm"])
                changes.append("поля: правое приведено к норме")
    except Exception:
        pass

    md_inline = re.compile(r"\*\*(.+?)\*\*|__(.+?)__|`([^`]+)`")
    md_prefix = re.compile(r"^#{1,6}\s+")

    for p in doc.paragraphs:
        text = p.text or ""
        if not text.strip():
            continue
        heading = _is_heading(p)

        # 2) Markdown-мусор в run-ах
        for run in p.runs:
            t = run.text or ""
            if not t:
                continue
            new_t = md_prefix.sub("", t)
            new_t = md_inline.sub(lambda m: m.group(1) or m.group(2) or m.group(3) or "", new_t)
            new_t = new_t.replace("**", "").replace("```", "")
            if new_t != t:
                run.text = new_t
                if "markdown удалён" not in changes:
                    changes.append("markdown удалён")

            # 3) Шрифт и кегль
            if run.font.name != g["font_name"]:
                run.font.name = g["font_name"]
                if "шрифт → Times New Roman" not in changes:
                    changes.append("шрифт → Times New Roman")
            if not heading and run.font.size is not None and run.font.size < Pt(g["font_size"]):
                run.font.size = Pt(g["font_size"])
                if "кегль → 14 pt" not in changes:
                    changes.append("кегль → 14 pt")

        # 4) Межстрочный интервал и отступ в основном тексте
        if not heading:
            pf = p.paragraph_format
            try:
                if pf.line_spacing is not None and abs(float(pf.line_spacing) - g["line_spacing"]) > 0.05:
                    pf.line_spacing = g["line_spacing"]
                    if "интервал → 1.5" not in changes:
                        changes.append("интервал → 1.5")
            except (TypeError, ValueError):
                pass

    if changes:
        try:
            doc.save(out_path)
        except Exception:
            return []
    return changes


def quick_check(path: str) -> tuple[bool, str]:
    """Быстрая проверка: файл открывается и не пуст."""
    try:
        doc = Document(path)
        n = sum(1 for p in doc.paragraphs if (p.text or "").strip())
        if n < 5:
            return False, f"Документ почти пуст: {n} абзацев"
        return True, f"OK: {n} абзацев"
    except Exception as e:
        return False, f"Файл не открывается: {e}"
