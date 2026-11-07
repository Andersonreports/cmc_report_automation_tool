"""MANE Select NM ID lookup from the HGNC REST API."""
from __future__ import annotations

import json
import re
import urllib.request

_HGNC_BY_SYMBOL = "https://rest.genenames.org/fetch/symbol/{symbol}"
_HGNC_BY_ACCESSION = "https://rest.genenames.org/fetch/refseq_accession/{nm}"


class ManeError(Exception):
    pass


def _strip_version(nm: str) -> str:
    return re.sub(r"\.\d+$", "", nm.strip())


def _hgnc_get(url: str) -> list[dict]:
    req = urllib.request.Request(
        url,
        headers={"Accept": "application/json", "User-Agent": "cmc-report-tool"},
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as r:
            data = json.load(r)
    except Exception as e:  # noqa: BLE001
        raise ManeError(f"Could not reach HGNC: {e}") from e
    return (data.get("response") or {}).get("docs") or []


def _mane_from_docs(docs: list[dict]) -> str:
    if not docs:
        return ""
    for entry in (docs[0].get("mane_select") or []):
        for part in entry.split("|"):
            if part.startswith("NM_"):
                return _strip_version(part)
    return ""


def fetch_mane_for_gene(gene: str, original_nm: str = "") -> str:
    """Return the MANE Select NM ID (without version) for the given gene.

    Looks up by gene symbol first. If the symbol is not found (e.g. Excel uses
    'MUT' but HGNC uses 'MMUT'), falls back to looking up by the original Excel
    NM ID to find the right gene entry.

    Returns '' when no MANE Select transcript is found.
    Raises ManeError on network errors.
    """
    # Primary: gene symbol lookup
    if gene:
        docs = _hgnc_get(_HGNC_BY_SYMBOL.format(symbol=gene.strip()))
        result = _mane_from_docs(docs)
        if result:
            return result

    # Fallback: use the original Excel NM ID to find the gene when symbol differs
    if original_nm:
        docs = _hgnc_get(_HGNC_BY_ACCESSION.format(nm=_strip_version(original_nm)))
        result = _mane_from_docs(docs)
        if result:
            return result

    return ""


def check_nm(gene: str, nm_to_check: str, original_nm: str = "") -> tuple[bool | None, str]:
    """Compare nm_to_check against the HGNC MANE Select for the gene.

    original_nm: the NM ID from the Excel input — used only as a fallback
                 to find the gene in HGNC when the symbol lookup fails.

    Returns:
        (True,  mane_nm) — nm_to_check matches MANE Select
        (False, mane_nm) — nm_to_check does not match MANE Select
        (None,  "")      — gene not in HGNC / no MANE Select entry
    Raises ManeError on network errors.
    """
    mane_nm = fetch_mane_for_gene(gene, original_nm or nm_to_check)
    if not mane_nm:
        return None, ""
    return _strip_version(nm_to_check) == mane_nm, mane_nm
