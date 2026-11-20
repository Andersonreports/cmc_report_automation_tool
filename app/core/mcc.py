"""Read the maternal cell contamination (MCC) table from the lab's MCC report.

The MCC report is a Word file (often the old .doc format) holding one table:

    No | Marker | Maternal genotype | Fetal genotype | No. of shared alleles | % MCC | Remarks
    1  | D3S1358 | 16, 17 | 15, 16 | 1 | - | Informative
    ...
    The average percentage of maternal cell contamination in the fetal sample is 0%   (merged row)

read_mcc() returns its marker rows, always in that 7-column order, and the closing
line; the report rebuilds the table from them (Appendix I of POC / Amniotic Fluid
reports). The files are typed by hand, so the table is not always built properly;
these are all read the same way:
  - the table nested inside another (a letterhead / layout table);
  - columns in another order, extra blank columns, no "No" column (numbered here);
  - a heading split over two rows ("Maternal" / "genotype");
  - all the markers in one row, one per line inside each cell;
  - the closing line as a paragraph under the table instead of a merged row;
  - no Word table at all: one line per marker, cells separated by tabs.
"""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

MCC_COLUMNS = ("No", "Marker", "Maternal genotype", "Fetal genotype",
               "No. of shared alleles", "% MCC", "Remarks")

# STR markers of the lab's kits; used to recognise marker rows in untidy files.
KNOWN_MARKERS = ("D3S1358", "TH01", "D21S11", "D18S51", "Penta E", "D5S818", "D13S317",
                 "D7S820", "D16S539", "CSF1PO", "Penta D", "Amelogenin", "vWA", "D8S1179",
                 "TPOX", "FGA", "D2S1338", "D19S433", "D1S1656", "D12S391", "D2S441",
                 "D10S1248", "D22S1045", "SE33", "DYS391", "D6S1043", "D2S1772")
_MARKER_RE = re.compile(r"^(D\d{1,2}S\d+|DYS\d+|TH0?1|TPOX|FGA|CSF1PO|VWA|SE33|AMEL\w*|PENTA\s*[A-Z])$",
                        re.I)
_SUMMARY_RE = re.compile(r"average\s+percentage|maternal\s+cell\s+contamination\s+in\s+the", re.I)


class MCCError(Exception):
    pass


# ---------------------------------------------------------------------------
# .doc -> .docx
# ---------------------------------------------------------------------------
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


# ---------------------------------------------------------------------------
# Cleaning helpers
# ---------------------------------------------------------------------------
def _clean(text: str) -> str:
    text = (text or "").replace(" ", " ").replace("–", "-").replace("—", "-")
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"\s*,\s*", ", ", text)               # "16,17" / "16 ,17" -> "16, 17"
    text = re.sub(r"(?i)^non[\s-]*informative$", "Non informative", text)
    text = re.sub(r"(?i)^informative$", "Informative", text)
    return text.rstrip(",").strip()


def _cell_lines(tc) -> list[str]:
    """A cell's text, one entry per paragraph / line break (empty lines dropped)."""
    lines = []
    for p in tc.iter(qn("w:p")):
        cur = []
        for el in p.iter(qn("w:t"), qn("w:br"), qn("w:cr"), qn("w:tab")):
            if el.tag == qn("w:t"):
                cur.append(el.text or "")
            elif el.tag == qn("w:tab"):
                cur.append(" ")
            else:
                cur.append("\n")
        lines.extend("".join(cur).split("\n"))
    return [line.strip() for line in lines if line.strip()]


def _column_of(header: str) -> int | None:
    """Which MCC column a heading names (index into MCC_COLUMNS)."""
    h = re.sub(r"[^a-z%]", "", header.lower())
    if not h:
        return None
    if "matern" in h:
        return 2
    if "fetal" in h or "foetal" in h or "fetus" in h:
        return 3
    if "shared" in h or "allele" in h:
        return 4
    if "%" in h or "mcc" in h or "percent" in h or "contamination" in h:
        return 5
    if "remark" in h or "interpret" in h or "informative" in h or "comment" in h:
        return 6
    if "marker" in h or "locus" in h or "loci" in h or h == "str":
        return 1
    if h in ("no", "sno", "slno", "sno.", "serialno", "sl", "s", "number", "#"):
        return 0
    return None


def _is_marker(text: str) -> bool:
    t = re.sub(r"\s+", " ", text.strip())
    return bool(t) and (bool(_MARKER_RE.match(t)) or t.lower() in (m.lower() for m in KNOWN_MARKERS))


def _canonical_marker(text: str) -> str:
    t = re.sub(r"\s+", " ", text.strip())
    for m in KNOWN_MARKERS:
        if t.lower() == m.lower() or t.replace(" ", "").lower() == m.replace(" ", "").lower():
            return m
    return t


# ---------------------------------------------------------------------------
# Tables
# ---------------------------------------------------------------------------
def _all_tables(doc):
    """Every table in the document, nested ones included (a table inside a
    letterhead / layout table)."""
    return list(doc.element.body.iter(qn("w:tbl")))


def _grid_rows(tbl) -> list[list[list[str]]]:
    """The table as rows of cells (each cell a list of lines). A cell spanning
    several grid columns is repeated so every row lines up with the grid; nested
    tables' cells are not mixed in."""
    rows = []
    for tr in tbl.findall(qn("w:tr")):
        cells = []
        for tc in tr.findall(qn("w:tc")):
            span = tc.find(qn("w:tcPr") + "/" + qn("w:gridSpan"))
            n = int(span.get(qn("w:val"))) if span is not None else 1
            # only this cell's own paragraphs, not those of a table nested in it
            lines = []
            for p in tc.findall(qn("w:p")):
                lines.extend(_cell_lines(p))
            cells.extend([lines] * n)
        rows.append(cells)
    return rows


def _row_text(row) -> str:
    seen, parts = set(), []
    for cell in row:
        key = id(cell)
        if key in seen:
            continue
        seen.add(key)
        parts.append(" ".join(cell))
    return " ".join(p for p in parts if p).strip()


def _parse_table(tbl) -> dict | None:
    rows = _grid_rows(tbl)
    if not rows:
        return None
    width = max(len(r) for r in rows)
    rows = [r + [[]] * (width - len(r)) for r in rows]
    # Header: the rows (one or two) whose cells name the columns.
    header_idx = next((i for i, r in enumerate(rows)
                       if sum(_column_of(" ".join(c)) is not None for c in r) >= 3), None)
    if header_idx is None:
        return None
    heads = [" ".join(c) for c in rows[header_idx]]
    start = header_idx + 1
    # A heading split over two rows: the next row adds to the headings when it has no
    # marker in it ("genotype" / "alleles" under "Maternal" / "No. of shared").
    if start < len(rows) and not any(_is_marker(line) for c in rows[start] for line in c) \
            and not _SUMMARY_RE.search(_row_text(rows[start])) \
            and any(" ".join(c) for c in rows[start]):
        heads = [(a + " " + " ".join(b)).strip() for a, b in zip(heads, rows[start])]
        start += 1
    colmap: dict[int, int] = {}                  # MCC column -> grid column
    for g, h in enumerate(heads):
        col = _column_of(h)
        if col is not None and col not in colmap:
            colmap[col] = g
    if 1 not in colmap:                          # no Marker heading: the column holding markers
        counts = {}
        for r in rows[start:]:
            for g, c in enumerate(r):
                counts[g] = counts.get(g, 0) + sum(_is_marker(x) for x in c)
        if counts and max(counts.values()):
            colmap[1] = max(counts, key=counts.get)
    if 1 not in colmap:
        return None

    markers, summary = [], ""
    for r in rows[start:]:
        text = _row_text(r)
        if not text:
            continue
        if _SUMMARY_RE.search(text) and not any(_is_marker(x) for x in r[colmap[1]]):
            summary = summary or _clean(text)
            continue
        cols = [r[colmap[c]] if c in colmap else [] for c in range(len(MCC_COLUMNS))]
        n = max(len(c) for c in cols)
        if n > 1 and len(cols[1]) > 1:
            # All markers in one row, one per line: split it into rows, line by line.
            for k in range(len(cols[1])):
                markers.append([_clean(c[k]) if k < len(c) else "" for c in cols])
        else:
            markers.append([_clean(" ".join(c)) for c in cols])
    markers = [m for m in markers if m[1] and not _column_of(m[1]) == 1]
    if not markers:
        return None
    return {"rows": markers, "summary": summary}


# ---------------------------------------------------------------------------
# No table: one line per marker
# ---------------------------------------------------------------------------
def _parse_text(lines: list[str]) -> dict | None:
    markers, summary = [], ""
    for line in lines:
        if _SUMMARY_RE.search(line):
            summary = summary or _clean(line)
            continue
        parts = [p.strip() for p in re.split(r"\t+|\s{3,}", line) if p.strip()]
        idx = next((i for i, p in enumerate(parts) if _is_marker(p)), None)
        if idx is None:
            continue
        no = parts[idx - 1] if idx > 0 and parts[idx - 1].isdigit() else ""
        rest = parts[idx + 1:]
        row = [no, parts[idx]] + rest[:5]
        markers.append([_clean(x) for x in row + [""] * (7 - len(row))])
    return {"rows": markers, "summary": summary} if markers else None


def _summary_after(doc) -> str:
    for p in doc.paragraphs:
        if _SUMMARY_RE.search(p.text):
            return _clean(p.text)
    return ""


def _finish(found: dict) -> dict:
    rows = []
    for i, r in enumerate(found["rows"], 1):
        r = list(r)
        r[1] = _canonical_marker(r[1])
        if not r[0] or not r[0].rstrip(".").isdigit():
            r[0] = str(i)                       # no / odd "No" column: numbered here
        r[0] = r[0].rstrip(".")
        rows.append(r)
    return {"rows": rows, "summary": found.get("summary", "")}


def read_mcc_docx(path: str) -> dict:
    """{'rows': [[7 cells], ...], 'summary': '...'} from a .docx MCC report."""
    doc = Document(path)
    best = None
    for tbl in _all_tables(doc):
        found = _parse_table(tbl)
        if found and (best is None or len(found["rows"]) > len(best["rows"])):
            best = found
    if best is None:
        # Tabs kept: they separate the cells of a marker line.
        lines = ["".join((t.text or "") if t.tag == qn("w:t") else "\t"
                         for t in p.iter(qn("w:t"), qn("w:tab")))
                 for p in doc.element.body.iter(qn("w:p"))]
        best = _parse_text(lines)
    if best is None:
        raise MCCError("No maternal cell contamination table was found in the file "
                       "(a table with Marker, Maternal genotype and Fetal genotype columns).")
    if not best["summary"]:
        best["summary"] = _summary_after(doc)
    return _finish(best)


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
