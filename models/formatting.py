"""Formatting rules data models."""

from __future__ import annotations

from pydantic import BaseModel


class PageSetup(BaseModel):
    width_cm: float = 21.0
    height_cm: float = 29.7
    margin_left_cm: float = 2.0
    margin_right_cm: float = 1.0
    margin_top_cm: float = 2.0
    margin_bottom_cm: float = 2.0


class FontSpec(BaseModel):
    name: str = "Times New Roman"
    size_pt: float = 13.0
    bold: bool = False
    italic: bool = False


class FormattingRules(BaseModel):
    page: PageSetup = PageSetup()
    body_font: FontSpec = FontSpec(name="Times New Roman", size_pt=13.0)
    heading_font: FontSpec = FontSpec(
        name="Times New Roman", size_pt=14.0, bold=True
    )
    title_font: FontSpec = FontSpec(
        name="Times New Roman", size_pt=14.0, bold=True
    )
    first_line_indent_cm: float = 1.0
    heading_space_after_emu: int = 152400
