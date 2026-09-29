"""Parse a gene coverage table pasted from Word / Excel / a web page.

Copied tables arrive as tab-separated text, one line per row. Both layouts
are accepted and read row by row, left to right, so gene order is kept:
  * the report's own layout: Gene | % | Gene | % | Gene | % | Gene | %
  * a two-column list:       Gene | %
Header cells ("Gene Name", "Percentage of coding region covered") and empty
or dash cells are skipped.
"""
from __future__ import annotations

import csv
import io
import re

_GENE = re.compile(r"[A-Za-z0-9][A-Za-z0-9._-]*")
_COVERAGE = re.compile(r"\d+(\.\d+)?\s*%?")
_BLANK = {"", "-", "–", "—"}


_HEADER_WORDS = {"gene", "name", "genes", "percentage", "of", "coding", "region",
                 "covered", "coverage", "%"}


def _is_header(cell: str) -> bool:
    """Header text, or a fragment of it: Word splits the multi-line header
    "Percentage / of coding / region / covered" across several lines."""
    words = cell.lower().replace("(", " ").replace(")", " ").split()
    return bool(words) and all(w in _HEADER_WORDS for w in words)


def parse_pasted(text: str) -> tuple[list[tuple[str, str]], list[str]]:
    """Return ([(gene, coverage), ...], [problem cells]). Coverage is kept as
    written (minus any '%'), e.g. '100.0'."""
    genes: list[tuple[str, str]] = []
    problems: list[str] = []
    # csv handles cells that Excel quotes because they contain line breaks.
    for row in csv.reader(io.StringIO(text.replace("\r\n", "\n")), delimiter="\t"):
        cells = [" ".join(c.split()) for c in row]
        if all(c in _BLANK or _is_header(c) for c in cells):
            continue
        for i in range(0, len(cells), 2):
            gene = cells[i]
            cov = cells[i + 1] if i + 1 < len(cells) else ""
            if gene in _BLANK and cov in _BLANK:
                continue
            if _is_header(gene):
                continue
            if not _GENE.fullmatch(gene) or not _COVERAGE.fullmatch(cov):
                problems.append(f"{gene or '(blank)'} | {cov or '(blank)'}")
                continue
            genes.append((gene, cov.rstrip("%").strip()))
    return genes, problems
