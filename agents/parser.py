"""Agent 1: DocumentParser — extract structured content from a draft DOCX."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from agents.base import BaseAgent
from models.document import DocumentContent, SectionData
from services.docx_reader import read_docx

_SECTION_NUM_RE = re.compile(
    r"^(\d+(?:\.\d+)*)\.*\s+(.*)", re.UNICODE
)


class DocumentParserAgent(BaseAgent):
    name = "DocumentParser"

    def run(self, *, draft_path: str, **kwargs: Any) -> DocumentContent:
        self.logger.info("Parsing draft document: %s", draft_path)
        content = read_docx(draft_path)
        content.sections = self._build_sections(content)
        self._detect_meta(content)
        self.logger.info(
            "Parsed %d paragraphs, %d tables, %d images, %d sections",
            content.total_paragraphs,
            content.total_tables,
            content.total_images,
            len(content.sections),
        )
        return content

    def _build_sections(self, content: DocumentContent) -> list[SectionData]:
        """Build logical section hierarchy from numbered paragraphs."""
        sections: list[SectionData] = []
        current_section: SectionData | None = None

        for para in content.paragraphs:
            text = para.text.strip()
            if not text:
                if current_section:
                    current_section.paragraphs.append(para)
                continue

            m = _SECTION_NUM_RE.match(text)
            has_bold = any(r.bold for r in para.runs) if para.runs else False

            if m and has_bold:
                num = m.group(1).rstrip(".")
                title = m.group(2).strip()
                level = num.count(".") + 1

                new_sec = SectionData(
                    number=num, title=title, level=level, paragraphs=[]
                )

                if level == 1:
                    if current_section:
                        sections.append(current_section)
                    current_section = new_sec
                elif current_section:
                    current_section.subsections.append(new_sec)
                else:
                    current_section = new_sec
            else:
                if current_section:
                    current_section.paragraphs.append(para)

        if current_section:
            sections.append(current_section)

        # Assign tables and images to sections by position
        for tbl in content.tables:
            owner = self._find_owner_section(sections, tbl.position_after_paragraph)
            if owner:
                owner.tables.append(tbl)

        for img in content.images:
            owner = self._find_owner_section(sections, img.position_after_paragraph)
            if owner:
                owner.images.append(img)

        return sections

    @staticmethod
    def _find_owner_section(
        sections: list[SectionData], para_pos: int
    ) -> SectionData | None:
        owner: SectionData | None = None
        for sec in sections:
            if sec.paragraphs and sec.paragraphs[0].index <= para_pos:
                owner = sec
        return owner

    def _detect_meta(self, content: DocumentContent) -> None:
        """Try to detect document title and STO number from early paragraphs."""
        for p in content.paragraphs[:30]:
            text = p.text.strip()
            if re.search(r"СТО\s*\d+", text):
                m = re.search(r"(СТО\s*[\d\-–]+)", text)
                if m:
                    content.sto_number = m.group(1)
            if "регламент" in text.lower() or "положение" in text.lower():
                content.title = text
