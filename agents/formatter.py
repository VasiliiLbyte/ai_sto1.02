"""Agent 4: DocumentFormatter — assemble the final DOCX document."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from agents.base import BaseAgent
from models.document import DocumentContent, SectionData
from models.mapping import StructureMapping
from services.docx_writer import DocxWriter
from templates.appendix_sheets import (
    add_approval_sheet,
    add_change_registration_sheet,
    add_familiarization_sheet,
)
from templates.title_page import add_title_page
from templates.toc import add_toc_page


class DocumentFormatterAgent(BaseAgent):
    name = "DocumentFormatter"

    def run(
        self,
        *,
        sections: list[SectionData],
        mapping: StructureMapping,
        original: DocumentContent,
        output_path: str,
        **kwargs: Any,
    ) -> Path:
        self.logger.info("Assembling formatted document …")
        writer = DocxWriter()

        sto_number = mapping.sto_number or original.sto_number or "СТО 31025229–XXX–2025"
        doc_title = mapping.document_title or original.title or "Стандарт организации"
        intro_status = mapping.introduction_status or "Введен впервые"

        # 1. Title page
        add_title_page(
            writer.doc,
            sto_number=sto_number,
            document_title=doc_title,
            intro_status=intro_status,
        )

        # 2. Table of contents
        appendix_titles = self._collect_appendix_titles(original)
        add_toc_page(writer.doc, sections, appendix_titles=appendix_titles)

        # 3. Header & footer (on a new section so title page has no header)
        self._setup_body_section(writer, sto_number)

        # 4. Body sections
        for sec in sections:
            writer.write_section(sec)

        # 5. Appendices from the original document
        self._write_appendices(writer, original)

        # 6. Mandatory sheets
        add_change_registration_sheet(writer.doc)
        add_familiarization_sheet(writer.doc)
        add_approval_sheet(writer.doc)

        saved = writer.save(output_path)
        self.logger.info("Document saved: %s", saved)
        return saved

    @staticmethod
    def _setup_body_section(writer: DocxWriter, sto_number: str) -> None:
        """Add a new section for the body so the title page stays header-free."""
        from docx.enum.section import WD_ORIENT
        from templates.formatting_rules import (
            MARGIN_BOTTOM,
            MARGIN_LEFT,
            MARGIN_RIGHT,
            MARGIN_TOP,
            PAGE_HEIGHT,
            PAGE_WIDTH,
        )

        writer.doc.add_section()
        new_sec = writer.doc.sections[-1]
        new_sec.page_width = PAGE_WIDTH
        new_sec.page_height = PAGE_HEIGHT
        new_sec.left_margin = MARGIN_LEFT
        new_sec.right_margin = MARGIN_RIGHT
        new_sec.top_margin = MARGIN_TOP
        new_sec.bottom_margin = MARGIN_BOTTOM
        new_sec.orientation = WD_ORIENT.PORTRAIT

        # Unlink header/footer from previous (title page) section
        new_sec.header.is_linked_to_previous = False
        new_sec.footer.is_linked_to_previous = False

        writer.setup_header_footer(sto_number)

    @staticmethod
    def _collect_appendix_titles(original: DocumentContent) -> list[str]:
        titles: list[str] = []
        for p in original.paragraphs:
            t = p.text.strip()
            if t.lower().startswith("приложение"):
                titles.append(t)
        return titles

    def _write_appendices(self, writer: DocxWriter, original: DocumentContent) -> None:
        """Copy appendix content (tables/images after the last numbered section)."""
        capturing = False
        for p in original.paragraphs:
            t = p.text.strip()
            if t.lower().startswith("приложение"):
                capturing = True
                writer.add_page_break()
                writer.add_section_heading("", t)
                continue
            if capturing and t:
                writer.add_paragraph_data(p)

        # Insert remaining tables not already placed
        placed_tables = set()
        for sec in original.sections:
            for tbl in sec.tables:
                placed_tables.add(tbl.index)

        for tbl in original.tables:
            if tbl.index not in placed_tables:
                writer.add_table(tbl)

        # Insert remaining images not already placed
        placed_images = set()
        for sec in original.sections:
            for img in sec.images:
                placed_images.add(img.index)

        for img in original.images:
            if img.index not in placed_images:
                writer.add_image(img)
