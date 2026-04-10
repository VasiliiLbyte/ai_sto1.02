"""Central formatting constants extracted from СТО 31025229-001-2024."""

from docx.shared import Cm, Pt, Emu

# Page setup
PAGE_WIDTH = Cm(21.0)
PAGE_HEIGHT = Cm(29.7)
MARGIN_LEFT = Cm(2.0)
MARGIN_RIGHT = Cm(1.0)
MARGIN_TOP = Cm(2.0)
MARGIN_BOTTOM = Cm(2.0)

# Fonts
FONT_NAME = "Times New Roman"
FONT_SIZE_BODY = Pt(13)
FONT_SIZE_HEADING = Pt(14)
FONT_SIZE_TITLE = Pt(14)
FONT_SIZE_TOC_TITLE = Pt(13)

# Paragraph formatting
FIRST_LINE_INDENT = Cm(1.0)
HEADING_SPACE_AFTER = Emu(152400)
TITLE_PAGE_INDENT = Cm(9.0)
STO_NUMBER_INDENT = Cm(9.75)

# Line spacing
LINE_SPACING_SINGLE = 1.0
