"""Filter a coverage CSV by a gene panel to produce a gene coverage table."""
from __future__ import annotations

import csv
import io
import re

import openpyxl


def parse_gene_panel(text: str) -> list[str]:
    """Parse a comma/newline separated gene list, deduplicate, sort A-Z."""
    genes: list[str] = []
    seen: set[str] = set()
    for token in re.split(r"[,\n\r]+", text):
        gene = token.strip().rstrip(".").strip()
        if gene and gene.upper() not in seen:
            genes.append(gene)
            seen.add(gene.upper())
    genes.sort(key=str.upper)
    return genes


def parse_coverage_csv(csv_bytes: bytes | str) -> dict[str, str]:
    """Parse a coverage CSV and return {GENE_NAME_UPPER: coverage_pct}.

    Accepts the CSV used by the Gene Panel Streamlit app whose header is:
    Region, Ref Name, Aliases, Gene_Name, Name, Gene IDs,
    Counted Bases, Mean Depth, Min Depth, Max Depth, % 1x
    """
    if isinstance(csv_bytes, bytes):
        text = csv_bytes.decode("utf-8-sig")
    else:
        text = csv_bytes
    reader = csv.reader(io.StringIO(text))
    header = next(reader, None)
    if not header:
        return {}
    header_clean = [h.strip().lower() for h in header]
    gene_col: int | None = None
    cov_col: int | None = None
    for i, h in enumerate(header_clean):
        if h in ("gene_name", "gene name"):
            gene_col = i
        elif "% 1x" in h or h == "% 1x":
            cov_col = i
    if gene_col is None:
        for i, h in enumerate(header_clean):
            if "gene" in h and "id" not in h:
                gene_col = i
                break
    if cov_col is None:
        for i, h in enumerate(header_clean):
            if h.startswith("%") or "coverage" in h:
                cov_col = i
                break
    if gene_col is None or cov_col is None:
        raise ValueError(
            "Could not find Gene_Name and coverage columns in the CSV. "
            f"Header: {', '.join(header)}"
        )
    result: dict[str, str] = {}
    limit = max(gene_col, cov_col)
    for row in reader:
        if len(row) <= limit:
            continue
        gene = row[gene_col].strip()
        coverage = row[cov_col].strip()
        if gene:
            result[gene.upper()] = coverage
    return result


def parse_mito_excel(data: bytes) -> list[tuple[str, str]]:
    """Parse a mitochondrial coverage Excel (header in row 2).

    The "Name" column holds entries like ``MT-TF/MT-TF/ex1``; the gene name
    is the first segment before ``/``.  Coverage comes from the ``% 1x``
    column.  Returns ``[(gene, coverage), ...]`` sorted alphabetically.
    """
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if len(rows) < 3:
        raise ValueError("Mitochondrial Excel has fewer than 3 rows")
    header = [str(h or "").strip().lower() for h in rows[1]]
    name_col: int | None = None
    cov_col: int | None = None
    for i, h in enumerate(header):
        if h == "name":
            name_col = i
        elif h in ("gene_name", "gene name", "gene"):
            name_col = name_col if name_col is not None else i
        elif "% 1x" in h:
            cov_col = i
    if name_col is None:
        for i, h in enumerate(header):
            if "name" in h or ("gene" in h and "id" not in h):
                name_col = i
                break
    if cov_col is None:
        for i, h in enumerate(header):
            if h.startswith("%") or "coverage" in h:
                cov_col = i
                break
    if name_col is None or cov_col is None:
        raise ValueError(
            "Could not find Name and coverage columns in the mitochondrial Excel. "
            f"Header (row 2): {', '.join(str(h) for h in rows[1])}"
        )
    result: list[tuple[str, str]] = []
    seen: set[str] = set()
    for row in rows[2:]:
        if len(row) <= max(name_col, cov_col):
            continue
        raw = str(row[name_col] or "").strip()
        gene = raw.split("/")[0].strip()
        cov_val = row[cov_col]
        cov = str(round(float(cov_val), 1)) if isinstance(cov_val, (int, float)) else str(cov_val or "").strip()
        if gene and gene.upper() not in seen:
            result.append((gene, cov))
            seen.add(gene.upper())
    result.sort(key=lambda g: g[0].upper())
    return result


def filter_coverage(
    csv_bytes: bytes | str, panel_text: str
) -> tuple[list[tuple[str, str]], list[str]]:
    """Filter coverage CSV by gene panel.

    Returns ([(gene, coverage), ...], [genes_not_found]).
    Gene names in the result use the panel's original casing.
    """
    coverage_map = parse_coverage_csv(csv_bytes)
    panel = parse_gene_panel(panel_text)
    matched: list[tuple[str, str]] = []
    not_found: list[str] = []
    for gene in panel:
        cov = coverage_map.get(gene.upper())
        if cov is not None:
            matched.append((gene, cov))
        else:
            not_found.append(gene)
    return matched, not_found
