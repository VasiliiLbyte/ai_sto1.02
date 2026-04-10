"""Agent 3: ContentEditor — per-section content editing via LLM."""

from __future__ import annotations

from typing import Any

from agents.base import BaseAgent
from config import MODEL_EDITOR
from models.document import DocumentContent, ParagraphData, RunData, SectionData
from models.mapping import ActionType, StructureMapping
from services.llm_client import LLMClient

SYSTEM_PROMPT = """\
Ты — технический редактор стандартов организации (СТО) для ООО «ПКФ «СНАРК».
Твоя задача: отредактировать текст раздела документа для соответствия формату СТО 31025229–001–2024.

Правила:
1. Сохрани смысловую нагрузку полностью — не удаляй важную информацию.
2. Нумерация подразделов: без точки после последнего числа (правильно: «4.1 Заголовок», \
неправильно: «4.1. Заголовок»).
3. Термины оформляются так: **термин**: определение (термин полужирным, затем двоеточие).
4. Перечисления оформляются через дефис «–» или строчные буквы с закрывающей скобкой «а)».
5. Пиши на официальном деловом русском языке.
6. Не добавляй информацию, которой не было в исходном тексте.
7. Убери лишние переносы строк и объедини разбитые предложения.
8. НЕ УДАЛЯЙ содержание — возвращай ВСЕ параграфы, даже если они кажутся избыточными.

Верни результат в формате JSON:
{
  "title": "<заголовок раздела>",
  "paragraphs": [
    {"text": "<текст параграфа>", "bold": false},
    ...
  ]
}
Параграф с bold=true означает подзаголовок или термин.
"""


class ContentEditorAgent(BaseAgent):
    name = "ContentEditor"

    def __init__(self, llm: LLMClient | None = None) -> None:
        super().__init__()
        self.llm = llm or LLMClient()

    def run(
        self,
        *,
        content: DocumentContent,
        mapping: StructureMapping,
        **kwargs: Any,
    ) -> list[SectionData]:
        """Edit each mapped section through LLM and return the new section list."""
        self.logger.info("Editing content for %d mapped sections", len(mapping.mappings))
        edited_sections: list[SectionData] = []

        for sm in mapping.mappings:
            if sm.action == ActionType.ADD:
                edited_sections.append(
                    SectionData(
                        number=sm.sto_number,
                        title=sm.sto_title,
                        level=sm.sto_number.count(".") + 1,
                        paragraphs=[],
                    )
                )
                continue

            source = self._find_source_section(content, sm.draft_number)
            if source is None:
                source_text = self._fallback_text(content, sm.draft_title)
                if source_text.strip():
                    edited = self._edit_flat(
                        sm.sto_number, sm.sto_title, source_text, sm.content_instructions
                    )
                else:
                    edited = SectionData(
                        number=sm.sto_number,
                        title=sm.sto_title,
                        level=sm.sto_number.count(".") + 1,
                    )
                edited_sections.append(edited)
                continue

            if source.subsections:
                edited = self._edit_with_subsections(source, sm.sto_number, sm.sto_title, sm.content_instructions)
            else:
                text = self._paragraphs_to_text(source.paragraphs)
                edited = self._edit_flat(sm.sto_number, sm.sto_title, text, sm.content_instructions)
                edited.tables = list(source.tables)
                edited.images = list(source.images)

            edited_sections.append(edited)

        return edited_sections

    def _edit_with_subsections(
        self,
        source: SectionData,
        sto_number: str,
        sto_title: str,
        instructions: str,
    ) -> SectionData:
        """Edit a section that has subsections, preserving the hierarchy."""
        # Edit the parent's own paragraphs (usually introductory text)
        parent_text = self._paragraphs_to_text(source.paragraphs)
        if parent_text.strip():
            parent_edited = self._edit_flat(
                sto_number, sto_title, parent_text, instructions
            )
        else:
            parent_edited = SectionData(
                number=sto_number, title=sto_title,
                level=sto_number.count(".") + 1,
            )

        parent_edited.tables = list(source.tables)
        parent_edited.images = list(source.images)

        # Edit each subsection separately
        for sub in source.subsections:
            sub_text = self._paragraphs_to_text(sub.paragraphs)
            sub_num = sub.number.replace(".", ".").rstrip(".")
            if sub_text.strip():
                edited_sub = self._edit_flat(
                    sub_num, sub.title, sub_text, instructions
                )
            else:
                edited_sub = SectionData(
                    number=sub_num, title=sub.title,
                    level=sub_num.count(".") + 1,
                )
            edited_sub.tables = list(sub.tables)
            edited_sub.images = list(sub.images)
            parent_edited.subsections.append(edited_sub)

        return parent_edited

    def _edit_flat(
        self, number: str, title: str, source_text: str, instructions: str
    ) -> SectionData:
        """Edit a flat section (no subsections) via LLM."""
        user_msg = (
            f"Раздел: {number} {title}\n"
            f"Инструкции: {instructions}\n\n"
            f"Исходный текст:\n{source_text}"
        )
        try:
            data = self.llm.chat_json(
                model=MODEL_EDITOR,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": user_msg},
                ],
            )
        except Exception as e:
            self.logger.warning("LLM edit failed for section %s: %s", number, e)
            return SectionData(
                number=number,
                title=title,
                level=number.count(".") + 1,
                paragraphs=[ParagraphData(text=source_text)],
            )

        paragraphs: list[ParagraphData] = []
        for p in data.get("paragraphs", []):
            text = p.get("text", "")
            bold = p.get("bold", False)
            paragraphs.append(
                ParagraphData(
                    text=text,
                    runs=[RunData(text=text, bold=bold)],
                )
            )

        raw_title = data.get("title", title)
        raw_title = self._strip_number_prefix(raw_title, number)

        return SectionData(
            number=number,
            title=raw_title,
            level=number.count(".") + 1,
            paragraphs=paragraphs,
        )

    @staticmethod
    def _strip_number_prefix(title: str, number: str) -> str:
        """Remove a leading section number from the title if the LLM included it."""
        stripped = title.strip()
        if number and stripped.startswith(number):
            stripped = stripped[len(number):].lstrip(". ")
        return stripped or title

    @staticmethod
    def _find_source_section(
        content: DocumentContent, draft_number: str | None
    ) -> SectionData | None:
        if not draft_number:
            return None
        clean = draft_number.rstrip(".")
        for sec in content.sections:
            if sec.number.rstrip(".") == clean:
                return sec
            for sub in sec.subsections:
                if sub.number.rstrip(".") == clean:
                    return sub
        return None

    @staticmethod
    def _paragraphs_to_text(paragraphs: list[ParagraphData]) -> str:
        lines: list[str] = []
        for p in paragraphs:
            t = p.text.strip()
            if t:
                lines.append(t)
        return "\n".join(lines)

    @staticmethod
    def _fallback_text(content: DocumentContent, title: str | None) -> str:
        """Try to find content by matching title in paragraphs."""
        if not title:
            return ""
        title_lower = title.lower()
        capturing = False
        lines: list[str] = []
        for p in content.paragraphs:
            t = p.text.strip()
            if not capturing:
                if title_lower in t.lower():
                    capturing = True
                    continue
            else:
                if any(r.bold for r in p.runs) and len(t) < 100 and t:
                    break
                if t:
                    lines.append(t)
        return "\n".join(lines)
