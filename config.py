import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / ".env")

OPENROUTER_API_KEY: str = os.getenv("OPENROUTER_API_KEY", "")
DRAFT_PATH: str = os.getenv("DRAFT_PATH", "")
OUTPUT_DIR: str = os.getenv("OUTPUT_DIR", str(Path(__file__).parent / "output"))

# LLM models (OpenRouter, no OpenAI / Anthropic)
MODEL_ANALYZER = "google/gemini-2.5-pro-preview-05-06"
MODEL_EDITOR = "deepseek/deepseek-r1"
MODEL_QC = "google/gemini-2.5-flash-preview-05-20"

LLM_TEMPERATURE = 0.3
LLM_MAX_TOKENS = 16_000

# STO formatting constants (extracted from СТО 31025229-001-2024)
FONT_MAIN = "Times New Roman"
FONT_SIZE_BODY_PT = 13
FONT_SIZE_HEADING_PT = 14
FIRST_LINE_INDENT_CM = 1.0
MARGIN_LEFT_CM = 2.0
MARGIN_RIGHT_CM = 1.0
MARGIN_TOP_CM = 2.0
MARGIN_BOTTOM_CM = 2.0
PAGE_WIDTH_CM = 21.0
PAGE_HEIGHT_CM = 29.7

COMPANY_NAME = 'ООО «Производственно-коммерческая фирма «СНАРК»'
COMPANY_SHORT = 'ООО «ПКФ «СНАРК»'
DIRECTOR_NAME = "В.В. Ласточкин"
DIRECTOR_TITLE = "Генеральный директор"
