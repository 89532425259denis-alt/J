# -*- coding: utf-8 -*-
"""
gost_validator.py — автоматическая проверка .docx на соответствие ГОСТ 7.32-2017.

Назначение:
  • объективный СКОРИНГ готовой работы (0..100) + детальный отчёт;
  • eval-метрика для подбора промптов и для дообучения LLM;
  • защита DOCX-слоя от регрессий (CI-проверка перед отправкой пользователю).

Зависимости: python-docx (уже есть в проекте). PDF-проверки опциональны.

Использование:
    from gost_validator import validate_docx, GostProfile
    report = validate_docx("work.docx")
    print(report.score, report.passed)
    print(report.as_text())

Интеграция в бота (generate_and_send), перед выдачей файла:
    rep = validate_docx(docx_path)
    if rep.score < 90:
        # запустить авто-починку слабых пунктов / перегенерацию
        ...
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional

try:
    from docx import Document
    from docx.shared import Pt, Mm, Twips
except Exception:  # pragma: no cover
    Document = None  # позволяет импортировать модуль без python-docx для статических проверок


# ─────────────────────────────────────────────────────────────────────────────
#  ПРОФИЛЬ ТРЕБОВАНИЙ ГОСТ 7.32-2017 (значения по умолчанию)
# ─────────────────────────────────────────────────────────────────────────────
@dataclass
class GostProfile:
    """Нормативные требования. Допуски учитывают округления конвертеров."""
    font_name: str = "Times New Roman"
    font_size_pt: float = 14.0            # 12..14 допускается; норма 14
    font_size_min_pt: float = 12.0
    line_spacing: float = 1.5             # полуторный
    line_spacing_tol: float = 0.06
    margin_left_mm: float = 30.0
    margin_right_mm: float = 15.0
    margin_top_mm: float = 20.0
    margin_bottom_mm: float = 20.0
    margin_tol_mm: float = 1.5            # допуск на конвертацию pt↔mm
    first_line_indent_mm: float = 12.5    # абзацный отступ 1.25 см (допуск 1.0..1.5)
    indent_min_mm: float = 10.0
    indent_max_mm: float = 17.0
    min_bibliography_sources: int = 10
    # структурные элементы, ожидаемые в отчёте (нормализованные ключи)
    required_sections: tuple = (
        "содержание", "введение", "заключение",
    )
    bibliography_titles: tuple = (
        "список использованных источников",
        "список литературы",
        "список использованной литературы",
    )


# ─────────────────────────────────────────────────────────────────────────────
#  РЕЗУЛЬТАТ
# ─────────────────────────────────────────────────────────────────────────────
@dataclass
class Check:
    code: str
    title: str
    weight: float
    passed: bool
    detail: str = ""

    @property
    def earned(self) -> float:
        return self.weight if self.passed else 0.0


@dataclass
class GostReport:
    checks: list = field(default_factory=list)

    def add(self, code, title, weight, passed, detail=""):
        self.checks.append(Check(code, title, weight, passed, detail))

    @property
    def total_weight(self) -> float:
        return sum(c.weight for c in self.checks) or 1.0

    @property
    def score(self) -> int:
        return round(100 * sum(c.earned for c in self.checks) / self.total_weight)

    @property
    def passed(self) -> bool:
        # критичные пункты (вес >= 8) не должны падать
        return self.score >= 90 and all(c.passed for c in self.checks if c.weight >= 8)

    def failures(self) -> list:
        return [c for c in self.checks if not c.passed]

    def as_text(self) -> str:
        lines = [f"╔═ ГОСТ 7.32-2017 — оценка: {self.score}/100  "
                 f"({'ПРОЙДЕНО' if self.passed else 'ТРЕБУЕТ ДОРАБОТКИ'})"]
        for c in sorted(self.checks, key=lambda x: (x.passed, -x.weight)):
            mark = "✅" if c.passed else "❌"
            lines.append(f"  {mark} [{c.weight:>4.1f}] {c.title}"
                         + (f" — {c.detail}" if c.detail else ""))
        return "\n".join(lines)


# ─────────────────────────────────────────────────────────────────────────────
#  ВСПОМОГАТЕЛЬНОЕ
# ─────────────────────────────────────────────────────────────────────────────
_CIT_FULL = re.compile(r"\[\s*\d+\s*,\s*с\.\s*\d+(?:\s*[–—-]\s*\d+)?\s*\]")
_CIT_BARE = re.compile(r"\[\s*\d+\s*\](?!\s*\.)")            # [N] без страницы
_CIT_BROKEN = re.compile(r"\[\s*\d+\s*,\s*с\.\s*\]")          # [N, с. ] без цифры


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def _emu_to_mm(emu: Optional[int]) -> Optional[float]:
    if emu is None:
        return None
    return emu / 36000.0  # 1 mm = 36000 EMU


# ─────────────────────────────────────────────────────────────────────────────
#  ОСНОВНАЯ ПРОВЕРКА
# ─────────────────────────────────────────────────────────────────────────────
def validate_docx(path: str, profile: Optional[GostProfile] = None) -> GostReport:
    if Document is None:
        raise RuntimeError("python-docx не установлен — pip install python-docx")
    p = profile or GostProfile()
    doc = Document(path)
    rep = GostReport()

    # ── 1. Поля страницы (критично, вес 10) ───────────────────────────────
    sec = doc.sections[0]
    margins = {
        "лев.": (_emu_to_mm(sec.left_margin), p.margin_left_mm),
        "прав.": (_emu_to_mm(sec.right_margin), p.margin_right_mm),
        "верх.": (_emu_to_mm(sec.top_margin), p.margin_top_mm),
        "нижн.": (_emu_to_mm(sec.bottom_margin), p.margin_bottom_mm),
    }
    bad = [f"{k} {got:.0f}≠{exp:.0f}мм" for k, (got, exp) in margins.items()
           if got is None or abs(got - exp) > p.margin_tol_mm]
    rep.add("margins", "Поля 30/15/20/20 мм", 10, not bad,
            "ОК" if not bad else "; ".join(bad))

    # ── собрать прогон шрифтов/интервалов/отступов по основному тексту ─────
    sizes, fonts, spacings, indents = [], [], [], []
    body_paras = []
    for para in doc.paragraphs:
        txt = para.text.strip()
        if not txt:
            continue
        body_paras.append(para)
        pf = para.paragraph_format
        if pf.line_spacing is not None and isinstance(pf.line_spacing, float):
            spacings.append(pf.line_spacing)
        if pf.first_line_indent is not None:
            indents.append(_emu_to_mm(pf.first_line_indent))
        for run in para.runs:
            if run.text.strip():
                if run.font.size:
                    sizes.append(run.font.size.pt)
                if run.font.name:
                    fonts.append(run.font.name)

    full_text = "\n".join(par.text for par in doc.paragraphs)

    # ── 2. Шрифт Times New Roman (вес 8) ──────────────────────────────────
    if fonts:
        dominant = max(set(fonts), key=fonts.count)
        share = fonts.count(dominant) / len(fonts)
        ok = _norm(dominant) == _norm(p.font_name) and share >= 0.85
        rep.add("font", f"Шрифт {p.font_name}", 8, ok,
                f"{dominant} ({share*100:.0f}%)")
    else:
        rep.add("font", f"Шрифт {p.font_name}", 8, False, "шрифт не задан в runs")

    # ── 3. Кегль 12–14 пт (вес 7) ─────────────────────────────────────────
    if sizes:
        dominant_sz = max(set(sizes), key=sizes.count)
        ok = p.font_size_min_pt - 0.1 <= dominant_sz <= p.font_size_pt + 0.1
        rep.add("size", "Кегль 12–14 пт", 7, ok, f"{dominant_sz:g} пт")
    else:
        rep.add("size", "Кегль 12–14 пт", 7, False, "размер не задан")

    # ── 4. Полуторный интервал (вес 7) ────────────────────────────────────
    if spacings:
        dominant_ls = max(set(spacings), key=spacings.count)
        ok = abs(dominant_ls - p.line_spacing) <= p.line_spacing_tol
        rep.add("spacing", "Межстрочный интервал 1.5", 7, ok, f"{dominant_ls:g}")
    else:
        rep.add("spacing", "Межстрочный интервал 1.5", 7, False, "не задан явно")

    # ── 5. Абзацный отступ ~1.25 см (вес 4) ───────────────────────────────
    valid_ind = [i for i in indents if i and i > 0]
    if valid_ind:
        med = sorted(valid_ind)[len(valid_ind) // 2]
        ok = p.indent_min_mm <= med <= p.indent_max_mm
        rep.add("indent", "Абзацный отступ ~12.5 мм", 4, ok, f"{med:.1f} мм")
    else:
        rep.add("indent", "Абзацный отступ ~12.5 мм", 4, False, "нет первой строки с отступом")

    # ── 6. Обязательные структурные элементы (вес 9) ──────────────────────
    heads = {_norm(par.text) for par in doc.paragraphs if par.text.strip()}
    heads_blob = " | ".join(heads)
    missing = [s for s in p.required_sections
               if not any(s in h for h in heads)]
    rep.add("structure", "Содержание / Введение / Заключение", 9, not missing,
            "все на месте" if not missing else "нет: " + ", ".join(missing))

    # ── 7. Список источников присутствует (вес 8) ─────────────────────────
    has_bib = any(any(bt in h for bt in p.bibliography_titles) for h in heads)
    rep.add("bib_present", "Список использованных источников", 8, has_bib,
            "найден" if has_bib else "раздел не найден")

    # ── 8. Кол-во источников >= минимума (вес 5) ──────────────────────────
    # эвристика: считаем нумерованные строки в хвосте документа
    bib_count = len(re.findall(r"(?m)^\s*\d{1,3}[.\)]\s+\S", full_text))
    ok = bib_count >= p.min_bibliography_sources
    rep.add("bib_count", f"Источников ≥ {p.min_bibliography_sources}", 5, ok,
            f"найдено ~{bib_count}")

    # ── 9. Формат сносок [N, с. X] (вес 8) ────────────────────────────────
    full_cit = len(_CIT_FULL.findall(full_text))
    bare_cit = len(_CIT_BARE.findall(full_text))
    broken_cit = len(_CIT_BROKEN.findall(full_text))
    ok = full_cit > 0 and broken_cit == 0 and bare_cit == 0
    rep.add("citations", "Сноски в формате [N, с. X]", 8, ok,
            f"корректных {full_cit}, без страниц {bare_cit}, битых {broken_cit}")

    # ── 10. Нет markdown-артефактов (вес 4) ───────────────────────────────
    md = len(re.findall(r"(?m)(^#{1,6}\s|\*\*|\s##\s|^\s*\*\s)", full_text))
    rep.add("no_markdown", "Нет markdown-разметки (#, **, ##)", 4, md == 0,
            "чисто" if md == 0 else f"{md} артефактов")

    # ── 11. Нумерация страниц включена (вес 5) ────────────────────────────
    has_pgnum = _docx_has_page_field(doc)
    rep.add("page_numbers", "Нумерация страниц (поле PAGE)", 5, has_pgnum,
            "есть" if has_pgnum else "поле PAGE не найдено")

    # ── 12. Объём документа разумен (вес 3) ───────────────────────────────
    chars = len(full_text)
    rep.add("volume", "Документ не пустой (> 1500 знаков)", 3, chars > 1500,
            f"{chars} знаков")

    return rep


def _docx_has_page_field(doc) -> bool:
    """Ищет инструкцию поля PAGE в колонтитулах (нумерация страниц)."""
    ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    for section in doc.sections:
        for hf in (section.footer, section.header,
                   section.first_page_footer, section.first_page_header):
            try:
                xml = hf._element.xml
            except Exception:
                continue
            if "PAGE" in xml and ("instrText" in xml or "fldSimple" in xml):
                return True
    return False


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("usage: python gost_validator.py work.docx")
        raise SystemExit(1)
    r = validate_docx(sys.argv[1])
    print(r.as_text())
    raise SystemExit(0 if r.passed else 2)
