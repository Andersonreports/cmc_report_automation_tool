"""Save the filled template as DOCX and convert it to PDF.

PDF conversion uses LibreOffice (headless) when installed, otherwise
Microsoft Word.

Layout check (build_report): a LAYOUT_CHECKED section that starts near the
bottom of a page and leaves only its heading or fewer than MIN_LINES lines
there is moved whole to the next page. If enough lines fit, it flows on.
Word can't express "at least 3 lines", so the converted PDF is measured and
the report re-rendered with that section starting a new page.

Gaps: a page that ends well short of its foot although the next page doesn't
start with a planned page break (Disclaimer, Appendix, ...) means a block kept
in one piece (the classification table with its intro) didn't fit there and
moved on whole. The report is then rendered tight (less padding, no blank line
in that block) and the tight version kept if it closes the gap.
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

from .docx_renderer import page_break_texts, render_report, section_opening
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
# A page whose text ends above this share of the writing area, followed by a
# page that isn't a planned page break, has a gap.
GAP_FILL = 0.80
_FOOTER = re.compile(r"^\s*Page\s+\d+\s+of\s+\d+\s*$", re.I)


def render_docx(data: ReportData, out_path: str, new_pages: tuple[str, ...] = (),
                tight: bool = False) -> str:
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    render_report(data, new_pages, tight).save(out_path)
    return out_path


def _page_layout(pdf_path: str) -> list[tuple[float, float, str]]:
    """Per page: (top, bottom) of its text as a share of the page height from the
    top, and its first line of text. The 'Page x of y' footer is left out."""
    import pypdfium2 as pdfium
    out = []
    doc = pdfium.PdfDocument(pdf_path)
    try:
        for i in range(len(doc)):
            page = doc[i]
            height, width = page.get_height(), page.get_width()
            tp = page.get_textpage()
            top, bottom, first = 1.0, 0.0, ""
            for k in range(tp.count_rects()):
                l, b, r, t = tp.get_rect(k)
                text = tp.get_text_bounded(l, b, r, t).strip()
                # the sideways 'Page x of y' sits in the right margin, in pieces
                if not text or _FOOTER.match(text) or l > 0.93 * width:
                    continue
                if (height - t) / height < top:
                    top, first = (height - t) / height, text
                bottom = max(bottom, (height - b) / height)
            out.append((top, bottom, first))
    finally:
        doc.close()
    return out


def _gaps(pdf_path: str, planned: list[str]) -> int:
    """How many pages end with a gap that isn't followed by a planned page break."""
    layout = _page_layout(pdf_path)
    planned = [_squash(t)[:25] for t in planned if _squash(t)]
    starts = [t for t, _, _ in layout if t < 1.0]
    floor = min(starts) if starts else 0.125        # the writing area's top
    foot = 1.0 - floor                              # same margin at the foot
    count = 0
    for (top, bottom, _), (_, _, first) in zip(layout, layout[1:]):
        nxt = _squash(first)
        if any(nxt.startswith(p) or p.startswith(nxt[:25]) for p in planned if nxt):
            continue
        if bottom < floor + GAP_FILL * (foot - floor):
            count += 1
    return count


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
        tight = False
        gaps = _gaps(pdf, page_break_texts(doc))
        if gaps:
            # A kept-together block left a page half empty: try the tight layout.
            render_docx(data, docx_path, tight=True)
            tight_pdf = docx_to_pdf(docx_path, probe)
            if _gaps(tight_pdf, page_break_texts(render_report(data, tight=True))) < gaps:
                tight, pdf, doc = True, tight_pdf, render_report(data, tight=True)
            else:
                render_docx(data, docx_path)
                pdf = docx_to_pdf(docx_path, probe)
        openings = {h: section_opening(doc, h) for h in LAYOUT_CHECKED}
        moved = _stranded(pdf, {h: t for h, t in openings.items() if t})
        if moved:
            render_docx(data, docx_path, moved, tight)
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
