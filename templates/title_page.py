"""Generate the title page (титульный лист) per СТО 31025229-001-2024 Appendix В."""

from __future__ import annotations

from datetime import datetime

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

from config import COMPANY_SHORT, DIRECTOR_NAME, DIRECTOR_TITLE
from templates.formatting_rules import (
    FONT_NAME,
    FONT_SIZE_TITLE,
    STO_NUMBER_INDENT,
    TITLE_PAGE_INDENT,
)


def add_title_page(
    doc: Document,
    *,
    sto_number: str,
    document_title: str,
    intro_status: str = "Введен впервые",
    year: int | None = None,
) -> None:
    """Add a complete title page to the document."""
    if year is None:
        year = datetime.now().year

    _add_approve_block(doc, year)
    _add_empty(doc, count=2)
    _add_sto_number(doc, sto_number)
    _add_empty(doc, count=2)
    _add_centered_bold(doc, "СТАНДАРТ ОРГАНИЗАЦИИ")
    _add_empty(doc)
    _add_line(doc)
    _add_empty(doc)
    _add_left_bold(doc, "Система менеджмента качества")
    _add_empty(doc)
    _add_document_title(doc, document_title, intro_status)
    _add_line(doc)
    _add_empty(doc, count=2)
    _add_intro_block(doc)
    _add_empty(doc, count=4)
    _add_centered(doc, COMPANY_SHORT)
    _add_empty(doc)
    _add_centered(doc, str(year))


def _run(para, text: str, *, bold: bool = False, size=FONT_SIZE_TITLE):
    r = para.add_run(text)
    r.font.name = FONT_NAME
    r.font.size = size
    r.bold = bold
    return r


def _add_approve_block(doc: Document, year: int) -> None:
    lines = [
        "УТВЕРЖДАЮ",
        DIRECTOR_TITLE,
        COMPANY_SHORT,
        f"______________ {DIRECTOR_NAME}",
        f"«          »                        {year} г.",
    ]
    for line in lines:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        _run(p, line, bold=True)


def _add_sto_number(doc: Document, sto_number: str) -> None:
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = STO_NUMBER_INDENT
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    _run(p, sto_number, bold=True)


def _add_centered_bold(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    _run(p, text, bold=True)


def _add_centered(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    _run(p, text, bold=False)


def _add_left_bold(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    _run(p, text, bold=True)


def _add_line(doc: Document) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    _run(p, "─" * 72, bold=True)


def _add_document_title(doc: Document, title: str, intro_status: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)

    title_upper = title.upper()
    padding = " " * max(1, 60 - len(title_upper))
    _run(p, title_upper)
    _run(p, padding + intro_status)


def _add_intro_block(doc: Document) -> None:
    lines = [
        "Введен в действие Приказом",
        "от _____________№ ______",
    ]
    for line in lines:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf = p.paragraph_format
        pf.first_line_indent = Cm(1.0)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        _run(p, line)

    _add_empty(doc)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    _run(p, "Дата введения ____________")


def _add_empty(doc: Document, count: int = 1) -> None:
    for _ in range(count):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        r = p.add_run("")
        r.font.name = FONT_NAME
        r.font.size = FONT_SIZE_TITLE
