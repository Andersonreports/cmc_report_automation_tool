"""Read the maternal cell contamination (MCC) table from the lab's MCC report.

The MCC report is a Word file (often the old .doc format) holding one table:

    No | Marker | Maternal genotype | Fetal genotype | No. of shared alleles | % MCC | Remarks
    1  | D3S1358 | 16, 17 | 15, 16 | 1 | - | Informative
    ...
    The average percentage of maternal cell contamination in the fetal sample is 0%   (merged row)

read_mcc() returns its marker rows and that closing line, which the report shows
as Appendix I on POC / Amniotic Fluid reports.
"""
from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from docx import Document

MCC_COLUMNS = ("No", "Marker", "Maternal genotype", "Fetal genotype",
               "No. of shared alleles", "% MCC", "Remarks")


class MCCError(Exception):
    pass


def _doc_to_docx(path: str, out_dir: str) -> str:
    """Convert an old .doc file to .docx with Word (Windows) or LibreOffice."""
    out = os.path.join(out_dir, Path(path).stem + ".docx")
    from .report_builder import _LOCK, _WORD, find_soffice
    if os.name == "nt":
        try:
            import win32com.client  # noqa: F401 - only checks pywin32 is there

            def convert(word):
                doc = word.Documents.Open(os.path.abspath(path), False, True)
                try:
                    doc.SaveAs2(os.path.abspath(out), 16)       # 16 = .docx
                finally:
                    doc.Close(False)

            with _LOCK:
                _WORD.call(convert)
            if os.path.exists(out):
                return out
        except Exception:  # noqa: BLE001 - try LibreOffice next
            pass
    soffice = find_soffice()
    if soffice:
        profile = tempfile.mkdtemp(prefix="cmc_lo_")
        try:
            subprocess.run([soffice, "--headless", "--norestore", "--nologo",
                            f"-env:UserInstallation={Path(profile).as_uri()}",
                            "--convert-to", "docx", "--outdir", out_dir, path],
                           capture_output=True, timeout=120,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        finally:
            shutil.rmtree(profile, ignore_errors=True)
        if os.path.exists(out):
            return out
    raise MCCError("Could not open the .doc file: Microsoft Word or LibreOffice is needed. "
                   "Save it as .docx and upload that instead.")


def _cell_text(cell) -> str:
    return " ".join(p.text.strip() for p in cell.paragraphs if p.text.strip())


def _is_header(cells: list[str]) -> bool:
    low = [c.lower() for c in cells]
    return any("marker" in c for c in low) and any("genotype" in c for c in low)


def read_mcc_docx(path: str) -> dict:
    """{'rows': [[7 cells], ...], 'summary': '...'} from a .docx MCC report."""
    doc = Document(path)
    for table in doc.tables:
        rows = []
        for tr in table.rows:
            cells, seen = [], set()
            for c in tr.cells:              # a merged cell repeats: keep it once
                if id(c._tc) in seen:
                    continue
                seen.add(id(c._tc))
                cells.append(_cell_text(c))
            rows.append(cells)
        start = next((i for i, r in enumerate(rows) if _is_header(r)), None)
        if start is None:
            continue
        markers, summary = [], ""
        for cells in rows[start + 1:]:
            if not any(cells):
                continue
            if len(cells) < len(MCC_COLUMNS):
                # The merged closing line ("The average percentage ...").
                summary = summary or " ".join(c for c in cells if c)
                continue
            markers.append(cells[:len(MCC_COLUMNS)])
        if markers:
            return {"rows": markers, "summary": summary}
    raise MCCError("No maternal cell contamination table was found in the file "
                   "(a table with Marker, Maternal genotype and Fetal genotype columns).")


def read_mcc(path: str) -> dict:
    """The MCC table of a .doc or .docx MCC report."""
    ext = Path(path).suffix.lower()
    if ext not in (".doc", ".docx"):
        raise MCCError("Upload the MCC report as a Word file (.doc or .docx).")
    if ext == ".docx":
        return read_mcc_docx(path)
    tmp = tempfile.mkdtemp(prefix="cmc_mcc_")
    try:
        return read_mcc_docx(_doc_to_docx(path, tmp))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
