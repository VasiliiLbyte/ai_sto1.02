"""Agent 1: DocumentParser — extract structured content from a draft DOCX."""

from __future__ import annotations

import re
from typing import Any

from agents.base import BaseAgent
from models.document import DocumentContent, SectionData
from services.docx_reader import read_docx

_SECTION_NUM_RE = re.compile(
    r"^(\d+(?:\.\d+)*)\.*\s+(.*)", re.UNICODE
)

_APPENDIX_HEADING_RE = re.compile(
    r"^Приложение\s*[№#]\s*\d", re.UNICODE | re.IGNORECASE
)

_TRAILING_MATTER_STARTS = {
    "разработал", "согласовано", "утверждаю", "лист ознакомления",
    "*примечание", "*справочно", "*м-19", "*м-29",
}


def _first_para_index(sec: SectionData) -> int | None:
    """Return the index of the first non-empty paragraph in a section."""
    for p in sec.paragraphs:
        return p.index
    return None


def _last_para_index(sec: SectionData) -> int | None:
    """Return the index of the last paragraph in a section (including subsections)."""
    last: int | None = None
    for p in sec.paragraphs:
        last = p.index
    for sub in sec.subsections:
        sub_last = _last_para_index(sub)
        if sub_last is not None and (last is None or sub_last > last):
            last = sub_last
    return last


def _compute_body_ranges(sections: list[SectionData]) -> list[tuple[int, int]]:
    """Build (start, end) paragraph ranges for leaf-level sections.

    Uses subsection ranges (not parent) to avoid spanning appendix gaps.
    """
    ranges: list[tuple[int, int]] = []

    for sec in sections:
        if sec.subsections:
            # Use each subsection's own range (plus parent's direct paragraphs)
            if sec.paragraphs:
                first = _first_para_index(sec)
                last = sec.paragraphs[-1].index
                if first is not None:
                    ranges.append((first, last))
            for sub in sec.subsections:
                first = _first_para_index(sub)
                last = sub.paragraphs[-1].index if sub.paragraphs else None
                if first is not None and last is not None:
                    ranges.append((first, last))
        else:
            first = _first_para_index(sec)
            last = sec.paragraphs[-1].index if sec.paragraphs else None
            if first is not None and last is not None:
                ranges.append((first, last))
    return ranges


def _in_body_range(pos: int, ranges: list[tuple[int, int]]) -> bool:
    """Check if a position falls within any body section's paragraph range."""
    for start, end in ranges:
        if start <= pos <= end:
            return True
    return False


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
        """Build logical section hierarchy from numbered paragraphs.

        Correctly assigns paragraphs to subsections (not just parent),
        and detects the appendix boundary to avoid polluting body sections
        with appendix content.
        """
        sections: list[SectionData] = []
        current_section: SectionData | None = None
        current_subsection: SectionData | None = None
        appendix_start: int = -1
        in_appendix_zone = False

        for para in content.paragraphs:
            text = para.text.strip()
            has_bold = any(r.bold for r in para.runs) if para.runs else False

            # Detect appendix zones: bold "Приложение №..." well past the TOC.
            # Can trigger multiple times (appendix material may appear in
            # multiple places in the original draft).
            if (
                has_bold
                and para.index > 80
                and _APPENDIX_HEADING_RE.match(text)
            ):
                if appendix_start == -1:
                    appendix_start = para.index
                    content.appendix_start_index = appendix_start
                    self.logger.info(
                        "Appendix boundary detected at paragraph %d: %s",
                        para.index,
                        text[:60],
                    )
                in_appendix_zone = True

            # Detect trailing matter (signatures, справочно, etc.)
            if (
                appendix_start > 0
                and not in_appendix_zone
                and text
                and any(text.lower().startswith(s) for s in _TRAILING_MATTER_STARTS)
            ):
                in_appendix_zone = True

            # A numbered heading exits the appendix zone (e.g. 7.3 after appendices)
            m = _SECTION_NUM_RE.match(text) if text else None
            if in_appendix_zone and m and has_bold:
                in_appendix_zone = False
                self.logger.info(
                    "Re-entered body at paragraph %d: %s", para.index, text[:60]
                )

            # Skip non-heading paragraphs inside the appendix zone
            if in_appendix_zone:
                continue

            if not text:
                target = current_subsection or current_section
                if target:
                    target.paragraphs.append(para)
                continue

            if m and has_bold:
                num = m.group(1).rstrip(".")
                title = m.group(2).strip()
                level = num.count(".") + 1

                new_sec = SectionData(
                    number=num, title=title, level=level, paragraphs=[]
                )

                if level == 1:
                    current_subsection = None
                    if current_section:
                        sections.append(current_section)
                    current_section = new_sec
                elif current_section:
                    current_subsection = new_sec
                    current_section.subsections.append(new_sec)
                else:
                    current_section = new_sec
            else:
                target = current_subsection or current_section
                if target:
                    target.paragraphs.append(para)

        if current_section:
            sections.append(current_section)

        # Assign tables and images to sections/subsections by position.
        # Only assign if the element falls within a section's actual paragraph range.
        body_ranges = _compute_body_ranges(sections)
        for tbl in content.tables:
            if not _in_body_range(tbl.position_after_paragraph, body_ranges):
                continue
            owner = self._find_owner_section(sections, tbl.position_after_paragraph)
            if owner:
                owner.tables.append(tbl)

        for img in content.images:
            if not _in_body_range(img.position_after_paragraph, body_ranges):
                continue
            owner = self._find_owner_section(sections, img.position_after_paragraph)
            if owner:
                owner.images.append(img)

        return sections

    @staticmethod
    def _find_owner_section(
        sections: list[SectionData], para_pos: int
    ) -> SectionData | None:
        """Find the most specific section/subsection owning a given paragraph position."""
        owner: SectionData | None = None
        for sec in sections:
            first_idx = _first_para_index(sec)
            if first_idx is not None and first_idx <= para_pos:
                owner = sec
            for sub in sec.subsections:
                sub_first = _first_para_index(sub)
                if sub_first is not None and sub_first <= para_pos:
                    owner = sub
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
