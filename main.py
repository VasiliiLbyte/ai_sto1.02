"""CLI entry point for the STO document formatter."""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

# Ensure the package root is on sys.path
sys.path.insert(0, str(Path(__file__).parent))

from config import DRAFT_PATH, OPENROUTER_API_KEY, OUTPUT_DIR
from orchestrator import Orchestrator


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Форматирование черновика документа по СТО 31025229-001-2024"
    )
    parser.add_argument(
        "--draft",
        default=DRAFT_PATH,
        help="Путь к черновику DOCX",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Путь к итоговому DOCX файлу",
    )
    parser.add_argument(
        "--api-key",
        default=None,
        help="OpenRouter API key (или задайте в .env)",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Подробный лог",
    )
    args = parser.parse_args()

    level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
        datefmt="%H:%M:%S",
    )

    draft = args.draft
    if not draft:
        print("Ошибка: укажите путь к черновику через --draft или DRAFT_PATH в .env")
        sys.exit(1)

    if not Path(draft).exists():
        print(f"Ошибка: файл не найден: {draft}")
        sys.exit(1)

    if args.output:
        output = args.output
    else:
        draft_name = Path(draft).stem
        output = str(Path(OUTPUT_DIR) / f"{draft_name}_СТО.docx")

    api_key = args.api_key or OPENROUTER_API_KEY

    orchestrator = Orchestrator(api_key=api_key)
    report = orchestrator.run(draft_path=draft, output_path=output)

    print("\n" + report.summary())
    if report.passed:
        print(f"\nГотовый документ: {output}")
    else:
        print(f"\nДокумент сохранен с замечаниями: {output}")
        sys.exit(1)


if __name__ == "__main__":
    main()
