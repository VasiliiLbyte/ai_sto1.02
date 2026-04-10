"""Pipeline orchestrator — coordinates agents with error handling and retries."""

from __future__ import annotations

import logging
import time
from pathlib import Path

from agents.analyzer import StructureAnalyzerAgent
from agents.editor import ContentEditorAgent
from agents.formatter import DocumentFormatterAgent
from agents.parser import DocumentParserAgent
from agents.quality import QualityControlAgent, QualityReport
from models.document import DocumentContent, SectionData
from models.mapping import StructureMapping
from services.llm_client import LLMClient

logger = logging.getLogger(__name__)

MAX_QC_RETRIES = 2


class Orchestrator:
    """Run the full conversion pipeline: Parse -> Analyze -> Edit -> Format -> QC."""

    def __init__(self, *, api_key: str | None = None) -> None:
        self.llm: LLMClient | None = None
        self._api_key = api_key

    def _get_llm(self) -> LLMClient:
        if self.llm is None:
            self.llm = LLMClient(api_key=self._api_key)
        return self.llm

    def run(self, *, draft_path: str, output_path: str) -> QualityReport:
        t0 = time.time()
        logger.info("=" * 60)
        logger.info("STO Formatter Pipeline START")
        logger.info("Draft: %s", draft_path)
        logger.info("Output: %s", output_path)
        logger.info("=" * 60)

        # Step 1: Parse
        logger.info("\n>>> STEP 1: Parsing draft document")
        parser = DocumentParserAgent()
        content: DocumentContent = parser.run(draft_path=draft_path)
        original_text = "\n".join(p.text for p in content.paragraphs if p.text.strip())

        # Step 2: Analyze structure
        logger.info("\n>>> STEP 2: Analyzing structure")
        llm = self._get_llm()
        analyzer = StructureAnalyzerAgent(llm=llm)
        mapping: StructureMapping = analyzer.run(content=content)

        # Step 3: Edit content
        logger.info("\n>>> STEP 3: Editing content")
        editor = ContentEditorAgent(llm=llm)
        sections: list[SectionData] = editor.run(content=content, mapping=mapping)

        # Step 4 + 5: Format and QC (with retries)
        report = self._format_and_validate(
            sections=sections,
            mapping=mapping,
            original=content,
            original_text=original_text,
            output_path=output_path,
        )

        elapsed = time.time() - t0
        logger.info("=" * 60)
        logger.info("Pipeline COMPLETE in %.1f seconds", elapsed)
        logger.info(report.summary())
        logger.info("=" * 60)
        return report

    def _format_and_validate(
        self,
        *,
        sections: list[SectionData],
        mapping: StructureMapping,
        original: DocumentContent,
        original_text: str,
        output_path: str,
    ) -> QualityReport:
        formatter = DocumentFormatterAgent()
        llm = self._get_llm() if self.llm else None
        qc = QualityControlAgent(llm=llm)

        for attempt in range(1, MAX_QC_RETRIES + 2):
            logger.info("\n>>> STEP 4 (attempt %d): Formatting document", attempt)
            out_path = formatter.run(
                sections=sections,
                mapping=mapping,
                original=original,
                output_path=output_path,
            )

            logger.info("\n>>> STEP 5 (attempt %d): Quality control", attempt)
            report = qc.run(output_path=str(out_path), original_text=original_text)

            if report.passed:
                logger.info("QC PASSED on attempt %d", attempt)
                return report

            logger.warning(
                "QC FAILED (attempt %d): %d errors, %d warnings",
                attempt,
                sum(1 for i in report.issues if i.severity == "error"),
                sum(1 for i in report.issues if i.severity == "warning"),
            )

            if attempt > MAX_QC_RETRIES:
                logger.error("Max retries reached. Returning last result.")
                break

        return report
