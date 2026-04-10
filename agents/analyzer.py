"""Agent 2: StructureAnalyzer — map draft structure to STO requirements via LLM."""

from __future__ import annotations

import json
from typing import Any

from agents.base import BaseAgent
from config import MODEL_ANALYZER
from models.document import DocumentContent
from models.mapping import ActionType, SectionMapping, StructureMapping
from services.llm_client import LLMClient

STO_REQUIRED_SECTIONS = [
    ("1", "Область применения"),
    ("2", "Нормативные ссылки"),
    ("3", "Термины, определения и сокращения"),
    # Sections 4+ are the main body — variable titles
    # Final mandatory sections:
    # "Отчетные документы"
    # "Ответственность"
]

SYSTEM_PROMPT = """\
Ты — эксперт по стандартам организации (СТО) для ООО «ПКФ «СНАРК».
Твоя задача: проанализировать структуру черновика документа и создать план реструктуризации \
в соответствии с требованиями СТО 31025229–001–2024.

Обязательная структура СТО:
1. Область применения
2. Нормативные ссылки
3. Термины, определения и сокращения
4-N. Основная часть (нумерованные разделы по теме документа)
Предпоследний раздел: Отчетные документы
Последний раздел: Ответственность
После разделов: Приложения

Правила нумерации: без точки после номера раздела (правильно: «1 Область применения», \
неправильно: «1. Область применения»).

Ответь строго в формате JSON (без markdown).
"""

USER_PROMPT_TEMPLATE = """\
Вот структура черновика документа:

Название документа: {title}

Разделы черновика:
{sections_list}

Создай план маппинга в формате JSON:
{{
  "sto_number": "<рекомендуемый номер СТО или оставь пустым>",
  "document_title": "<корректное название для СТО>",
  "introduction_status": "Введен впервые",
  "mappings": [
    {{
      "draft_number": "<номер раздела из черновика или null>",
      "draft_title": "<название в черновике или null>",
      "sto_number": "<номер в СТО>",
      "sto_title": "<название раздела по СТО>",
      "action": "keep|rename|add|merge|reorder",
      "content_instructions": "<инструкции по редактированию содержания>"
    }}
  ],
  "notes": "<общие замечания>"
}}

Важно:
- Обязательно включи разделы «Область применения», «Нормативные ссылки», \
«Термины, определения и сокращения» в начале.
- Последние два раздела основной части должны быть «Отчетные документы» и «Ответственность».
- Если в черновике есть «Общие положения», его содержание нужно перенести в «Область применения».
- Сохрани все содержательные разделы черновика, лишь переименуй/перенумеруй при необходимости.
"""


class StructureAnalyzerAgent(BaseAgent):
    name = "StructureAnalyzer"

    def __init__(self, llm: LLMClient | None = None) -> None:
        super().__init__()
        self.llm = llm or LLMClient()

    def run(self, *, content: DocumentContent, **kwargs: Any) -> StructureMapping:
        self.logger.info("Analyzing document structure …")

        sections_list = self._format_sections(content)
        title = content.title or content.filename

        user_msg = USER_PROMPT_TEMPLATE.format(
            title=title, sections_list=sections_list
        )

        data = self.llm.chat_json(
            model=MODEL_ANALYZER,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_msg},
            ],
        )

        mapping = self._parse_response(data)
        self.logger.info(
            "Structure mapping: %d section mappings", len(mapping.mappings)
        )
        return mapping

    @staticmethod
    def _format_sections(content: DocumentContent) -> str:
        lines: list[str] = []
        for sec in content.sections:
            lines.append(f"  {sec.number}. {sec.title}  (level={sec.level}, "
                         f"paragraphs={len(sec.paragraphs)}, "
                         f"tables={len(sec.tables)}, images={len(sec.images)})")
            for sub in sec.subsections:
                lines.append(f"    {sub.number}. {sub.title}")
        if not lines:
            for p in content.paragraphs[:60]:
                t = p.text.strip()
                if t and any(r.bold for r in p.runs):
                    lines.append(f"  [{p.index}] {t[:120]}")
        return "\n".join(lines) if lines else "(no sections detected)"

    @staticmethod
    def _parse_response(data: dict) -> StructureMapping:
        mappings: list[SectionMapping] = []
        for m in data.get("mappings", []):
            action_str = m.get("action", "keep").lower()
            try:
                action = ActionType(action_str)
            except ValueError:
                action = ActionType.KEEP
            mappings.append(
                SectionMapping(
                    draft_number=m.get("draft_number"),
                    draft_title=m.get("draft_title"),
                    sto_number=m.get("sto_number", ""),
                    sto_title=m.get("sto_title", ""),
                    action=action,
                    content_instructions=m.get("content_instructions", ""),
                )
            )

        return StructureMapping(
            mappings=mappings,
            sto_number=data.get("sto_number", ""),
            document_title=data.get("document_title", ""),
            introduction_status=data.get("introduction_status", "Введен впервые"),
            notes=data.get("notes", ""),
        )
