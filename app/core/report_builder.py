"""Save the filled template as DOCX and convert it to PDF.

PDF conversion uses LibreOffice (headless) when installed, otherwise
Microsoft Word.
"""
from __future__ import annotations

import os
import subprocess
import tempfile
import threading
from pathlib import Path

from .docx_renderer import render_report
from .models import ReportData
from .os_utils import find_soffice


class ReportBuildError(Exception):
    pass


# One LibreOffice profile reused for the whole session: only the first
# conversion pays LibreOffice's start-up cost. A profile can be used by one
# soffice at a time, so conversions are serialised.
_PROFILE = tempfile.mkdtemp(prefix="cmc_lo_profile_")
_LOCK = threading.Lock()


def render_docx(data: ReportData, out_path: str) -> str:
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    render_report(data).save(out_path)
    return out_path


def _via_libreoffice(soffice: str, docx_path: str, out_dir: str) -> None:
    cmd = [soffice, "--headless", "--norestore", "--nologo", "--nolockcheck",
           f"-env:UserInstallation={Path(_PROFILE).as_uri()}",
           "--convert-to", "pdf", "--outdir", out_dir, docx_path]
    with _LOCK:
        try:
            subprocess.run(cmd, capture_output=True, timeout=120,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        except subprocess.TimeoutExpired as e:
            raise ReportBuildError("PDF conversion timed out") from e


def _via_word(docx_path: str, pdf_path: str) -> None:
    try:
        import pythoncom
        from win32com.client import DispatchEx
    except ImportError as e:
        raise ReportBuildError(
            "No PDF converter found: install LibreOffice or Microsoft Word.") from e
    with _LOCK:
        pythoncom.CoInitialize()
        word = None
        try:
            word = DispatchEx("Word.Application")
            word.DisplayAlerts = 0
            doc = word.Documents.Open(os.path.abspath(docx_path), False, True)
            doc.ExportAsFixedFormat(os.path.abspath(pdf_path), 17)   # 17 = PDF
            doc.Close(False)
        except Exception as e:  # noqa: BLE001
            raise ReportBuildError(f"Word PDF conversion failed: {e}") from e
        finally:
            if word is not None:
                word.Quit()
            pythoncom.CoUninitialize()


def docx_to_pdf(docx_path: str, out_dir: str) -> str:
    """Convert docx_path to <out_dir>/<same name>.pdf and return that path."""
    pdf_path = os.path.join(out_dir, Path(docx_path).stem + ".pdf")
    if os.path.exists(pdf_path):
        os.remove(pdf_path)     # never mistake a stale PDF for a fresh one
    soffice = find_soffice()
    if soffice:
        _via_libreoffice(soffice, docx_path, out_dir)
    else:
        _via_word(docx_path, pdf_path)
    if not os.path.exists(pdf_path):
        raise ReportBuildError("PDF was not produced.")
    return pdf_path
