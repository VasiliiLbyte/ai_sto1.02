"""Data models for parsed document content."""

from __future__ import annotations

import base64
from enum import Enum
from typing import Optional

from pydantic import BaseModel, Field


class Alignment(str, Enum):
    LEFT = "LEFT"
    CENTER = "CENTER"
    RIGHT = "RIGHT"
    JUSTIFY = "JUSTIFY"


class RunData(BaseModel):
    """A single run of text with uniform formatting."""

    text: str = ""
    font_name: Optional[str] = None
    font_size_pt: Optional[float] = None
    bold: bool = False
    italic: bool = False
    underline: bool = False


class ParagraphData(BaseModel):
    """A single paragraph with its formatting metadata."""

    index: int = 0
    text: str = ""
    runs: list[RunData] = Field(default_factory=list)
    style_name: str = "Normal"
    alignment: Optional[Alignment] = None
    first_line_indent_cm: Optional[float] = None
    left_indent_cm: Optional[float] = None
    space_before_pt: Optional[float] = None
    space_after_pt: Optional[float] = None
    line_spacing: Optional[float] = None
    is_heading: bool = False
    heading_level: Optional[int] = None


class TableCellData(BaseModel):
    """A single table cell."""

    text: str = ""
    paragraphs: list[ParagraphData] = Field(default_factory=list)
    row_span: int = 1
    col_span: int = 1


class TableData(BaseModel):
    """A parsed table."""

    index: int = 0
    position_after_paragraph: int = 0
    rows: list[list[TableCellData]] = Field(default_factory=list)
    num_rows: int = 0
    num_cols: int = 0


class ImageData(BaseModel):
    """An extracted image."""

    index: int = 0
    position_after_paragraph: int = 0
    image_bytes_b64: str = ""
    content_type: str = "image/png"
    width_cm: Optional[float] = None
    height_cm: Optional[float] = None
    filename: str = ""

    def get_bytes(self) -> bytes:
        return base64.b64decode(self.image_bytes_b64)

    def set_bytes(self, data: bytes) -> None:
        self.image_bytes_b64 = base64.b64encode(data).decode("ascii")


class SectionData(BaseModel):
    """A logical section of the document (e.g. '1 Область применения')."""

    number: str = ""
    title: str = ""
    level: int = 1
    paragraphs: list[ParagraphData] = Field(default_factory=list)
    tables: list[TableData] = Field(default_factory=list)
    images: list[ImageData] = Field(default_factory=list)
    subsections: list[SectionData] = Field(default_factory=list)


class DocumentContent(BaseModel):
    """Full parsed document content."""

    filename: str = ""
    total_paragraphs: int = 0
    total_tables: int = 0
    total_images: int = 0

    paragraphs: list[ParagraphData] = Field(default_factory=list)
    tables: list[TableData] = Field(default_factory=list)
    images: list[ImageData] = Field(default_factory=list)
    sections: list[SectionData] = Field(default_factory=list)

    title: str = ""
    sto_number: str = ""
    introduction_status: str = "Введен впервые"

    appendix_start_index: int = -1
