"""Save the filled template as DOCX and convert it to PDF.

PDF conversion uses LibreOffice (headless) when installed, otherwise
Microsoft Word.

Layout check (build_report): a LAYOUT_CHECKED section that starts near the
bottom of a page and leaves only its heading or fewer than MIN_LINES lines
there is moved whole to the next page. If enough lines fit, it flows on.
Word can't express "at least 3 lines", so the converted PDF is measured and
the report re-rendered with that section starting a new page.
"""
from __future__ import annotations

import atexit
import os
import queue
import re
import shutil
import subprocess
import tempfile
import threading
from pathlib import Path

from .docx_renderer import render_report, section_opening
from .models import ReportData
from .os_utils import find_soffice


class ReportBuildError(Exception):
    pass


# One LibreOffice profile kept between sessions: building a fresh one costs
# ~10 s, so only the very first conversion ever pays it. A profile can be used
# by one soffice at a time, so conversions are serialised.
_PROFILE = os.path.join(tempfile.gettempdir(), "cmc_lo_profile")
_LOCK = threading.Lock()


LAYOUT_CHECKED = ("Methodology",)
MIN_LINES = 3


def render_docx(data: ReportData, out_path: str, new_pages: tuple[str, ...] = ()) -> str:
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    render_report(data, new_pages).save(out_path)
    return out_path


def _squash(text: str) -> str:
    """Letters and digits only: PDF text differs from the DOCX in spacing and
    symbols (a hyphen can come out as U+FFFE)."""
    return re.sub(r"[^0-9A-Za-z]+", "", text)


def _stranded(pdf_path: str, openings: dict[str, str]) -> tuple[str, ...]:
    """The headings in `openings` ({heading: its first paragraph}) that the
    PDF shows at the foot of a page with fewer than MIN_LINES lines of that
    paragraph under them, the rest carried over to the next page."""
    import pypdfium2 as pdfium
    out = []
    doc = pdfium.PdfDocument(pdf_path)
    try:
        pages = [doc[i].get_textpage().get_text_range().splitlines() for i in range(len(doc))]
    finally:
        doc.close()
    for heading, para in openings.items():
        para = _squash(para)
        for lines in pages:
            idx = next((i for i, ln in enumerate(lines) if ln.strip() == heading), None)
            if idx is None:
                continue
            shown = ""
            count = 0
            for ln in lines[idx + 1:]:
                part = _squash(ln)
                if not part or (shown + part) not in para:
                    break
                shown += part
                count += 1
            if len(shown) < len(para) and count < MIN_LINES:
                out.append(heading)
            break
    return tuple(out)


def build_report(data: ReportData, docx_path: str, pdf_dir: str | None = None):
    """Save the report as docx_path and, with pdf_dir, its PDF there.
    Returns (docx path, pdf path or None). Applies the layout check; without
    pdf_dir the PDF is only a probe in a temp folder, and if no converter is
    available the DOCX is saved unchecked."""
    render_docx(data, docx_path)
    probe = pdf_dir or tempfile.mkdtemp(prefix="cmc_probe_")
    try:
        try:
            pdf = docx_to_pdf(docx_path, probe)
        except ReportBuildError:
            if pdf_dir:
                raise
            return docx_path, None
        doc = render_report(data)
        openings = {h: section_opening(doc, h) for h in LAYOUT_CHECKED}
        moved = _stranded(pdf, {h: t for h, t in openings.items() if t})
        if moved:
            render_docx(data, docx_path, moved)
            pdf = docx_to_pdf(docx_path, probe)
        return docx_path, (pdf if pdf_dir else None)
    finally:
        if not pdf_dir:
            shutil.rmtree(probe, ignore_errors=True)


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


class _WordThread:
    """Microsoft Word kept open between conversions, so only the first pays its
    start-up (~2 s). A COM object can only be used from the thread that made
    it, and previews / exports / web requests each run on a thread of their
    own, so one thread owns Word and every conversion is handed to it. Word is
    a private instance (never the user's open Word) and quits at exit."""

    def __init__(self):
        self._jobs: queue.Queue = queue.Queue()
        self._thread: threading.Thread | None = None
        self._start = threading.Lock()

    def call(self, fn):
        """Run fn(word) on the Word thread and return its result."""
        with self._start:
            if self._thread is None or not self._thread.is_alive():
                self._thread = threading.Thread(target=self._loop, name="word-pdf", daemon=True)
                self._thread.start()
        done, box = threading.Event(), {}
        self._jobs.put((fn, done, box))
        done.wait()
        if "error" in box:
            raise box["error"]
        return box.get("result")

    def _loop(self):
        import pythoncom
        from win32com.client import DispatchEx
        pythoncom.CoInitialize()
        word, stopped = None, None
        try:
            while True:
                fn, done, box = self._jobs.get()
                if fn is None:                      # quit(): answered once Word has closed
                    stopped = done
                    return
                try:
                    if word is None:
                        word = DispatchEx("Word.Application")
                        word.Visible = False
                        word.DisplayAlerts = 0
                    box["result"] = fn(word)
                except Exception as e:  # noqa: BLE001 - re-raised in call()
                    box["error"] = e
                    try:                            # Word may have died: start afresh next time
                        word.Quit()
                    except Exception:  # noqa: BLE001
                        pass
                    word = None
                finally:
                    done.set()
        finally:
            if word is not None:
                try:
                    word.Quit()
                except Exception:  # noqa: BLE001
                    pass
            pythoncom.CoUninitialize()
            if stopped is not None:
                stopped.set()

    def quit(self, timeout: float = 15):
        if self._thread is not None and self._thread.is_alive():
            done = threading.Event()
            self._jobs.put((None, done, {}))
            done.wait(timeout)


_WORD = _WordThread()
atexit.register(_WORD.quit)


def _via_word(docx_path: str, pdf_path: str) -> None:
    try:
        import pythoncom  # noqa: F401 - only checks pywin32 is there
        import win32com.client  # noqa: F401
    except ImportError as e:
        raise ReportBuildError(
            "No PDF converter found: install LibreOffice or Microsoft Word.") from e

    def convert(word):
        doc = word.Documents.Open(os.path.abspath(docx_path), False, True)
        try:
            doc.ExportAsFixedFormat(os.path.abspath(pdf_path), 17)   # 17 = PDF
        finally:
            doc.Close(False)

    with _LOCK:
        try:
            _WORD.call(convert)
        except Exception as e:  # noqa: BLE001
            raise ReportBuildError(f"Word PDF conversion failed: {e}") from e


def docx_to_pdf(docx_path: str, out_dir: str) -> str:
    """Convert docx_path to <out_dir>/<same name>.pdf and return that path.

    On Windows Word (COM) is tried first — it reuses the running Word process
    and is typically 5-10× faster than spawning a new LibreOffice process.
    LibreOffice is the fallback (and the only option on Linux/macOS).
    """
    pdf_path = os.path.join(out_dir, Path(docx_path).stem + ".pdf")
    if os.path.exists(pdf_path):
        os.remove(pdf_path)
    if os.name == "nt":
        try:
            _via_word(docx_path, pdf_path)
            if os.path.exists(pdf_path):
                return pdf_path
        except Exception:
            pass
    soffice = find_soffice()
    if soffice:
        _via_libreoffice(soffice, docx_path, out_dir)
    else:
        _via_word(docx_path, pdf_path)
    if not os.path.exists(pdf_path):
        raise ReportBuildError("PDF was not produced.")
    return pdf_path
