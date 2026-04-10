"""Word COM post-processing utilities.

Updates DOCX fields (TOC, page numbers, references) using Microsoft Word.
"""

from __future__ import annotations

import logging
from pathlib import Path

logger = logging.getLogger(__name__)


def update_word_fields(docx_path: str) -> bool:
    """Update all fields in a DOCX via Word COM automation.

    Returns True on success. Returns False if Word COM is unavailable
    or any update error occurs.
    """
    path = Path(docx_path)
    if not path.exists():
        logger.warning("Word postprocess skipped: file not found: %s", path)
        return False

    try:
        import pythoncom  # type: ignore
        import win32com.client as win32  # type: ignore
    except Exception as exc:
        logger.warning("Word postprocess unavailable (pywin32): %s", exc)
        return False

    app = None
    doc = None
    try:
        pythoncom.CoInitialize()
        app = win32.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0

        doc = app.Documents.Open(str(path.resolve()))

        # Update all document fields first.
        doc.Fields.Update()

        # Update TOC fields explicitly to ensure page numbers are refreshed.
        if doc.TablesOfContents.Count > 0:
            for i in range(1, doc.TablesOfContents.Count + 1):
                doc.TablesOfContents(i).Update()

        # Repaginate and save.
        doc.Repaginate()
        doc.Save()
        logger.info("Word postprocess complete: fields updated in %s", path)
        return True
    except Exception as exc:
        logger.warning("Word postprocess failed: %s", exc)
        return False
    finally:
        try:
            if doc is not None:
                doc.Close(SaveChanges=True)
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

