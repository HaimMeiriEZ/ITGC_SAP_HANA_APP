"""Embed files into Excel worksheets as OLE objects (Windows + Excel)."""
from __future__ import annotations

import logging
import sys
from pathlib import Path

_logger = logging.getLogger(__name__)


def embed_file_in_worksheet(
    workbook_path: Path,
    sheet_name: str,
    file_path: Path,
    *,
    left: float = 24,
    top: float = 110,
    width: float = 520,
    height: float = 360,
) -> bool:
    """Embed *file_path* into *sheet_name* of a saved Excel workbook.

    Returns True when embedding succeeded. Requires Windows with Excel installed.
    """
    if sys.platform != "win32":
        _logger.info("OLE embedding skipped — supported on Windows only")
        return False

    workbook_path = Path(workbook_path)
    file_path = Path(file_path)
    if not workbook_path.exists():
        _logger.warning("Workbook not found for OLE embed: %s", workbook_path)
        return False
    if not file_path.exists():
        _logger.warning("File not found for OLE embed: %s", file_path)
        return False

    try:
        import win32com.client  # type: ignore[import-untyped]
    except ImportError:
        _logger.warning("pywin32 not available — cannot embed OLE object")
        return False

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(str(workbook_path.resolve()))
        worksheet = workbook.Worksheets(sheet_name)
        # Positional args required — named kwargs fail with win32com (HRESULT 0x800A17AC).
        # DisplayAsIcon=True embeds ZIP/PDF/DOC as clickable package icons.
        worksheet.OLEObjects().Add(
            None,
            str(file_path.resolve()),
            False,
            True,
            None,
            0,
            None,
            left,
            top,
            width,
            height,
        )
        workbook.Save()
        return True
    except Exception as exc:
        _logger.exception("Failed to embed OLE object in %s: %s", workbook_path, exc)
        return False
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
