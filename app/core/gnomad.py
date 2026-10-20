"""Variant Frequency from the gnomAD v4 website (its public GraphQL API).

The listed-variants Excel only has allele counts (AC); the percentage needs the
allele number (AN) too, which only gnomAD has: AC / AN * 100. gnomAD's "joint" figures
(exomes + genomes) are used, the website's Total view. The South Asian figures are
used when the variant is seen in that group, else the all-populations figures,
else "not observed".
"""
from __future__ import annotations

import json
import math
import urllib.request

from . import variants as variants_core

API = "https://gnomad.broadinstitute.org/api"
QUERY = ("query($id: String!) { variant(variantId: $id, dataset: gnomad_r4) "
         "{ joint { ac an populations { id ac an } } } }")


class GnomadError(Exception):
    pass


def variant_id(v: dict) -> str:
    """gnomAD id (chrom-pos-ref-alt, GRCh38) from the sheet's Chromosomal Position
    and Ref/Alt; '' for insertions / deletions, whose id needs the base before them."""
    chrom, _, pos = v.get("chrom_pos", "").partition(":")
    ref, _, alt = v.get("ref_alt", "").partition("/")
    ref, alt = ref.strip().upper(), alt.strip().upper()
    if not (chrom and pos.strip().isdigit() and ref.isalpha() and alt.isalpha()):
        return ""
    return f"{chrom.strip().removeprefix('chr')}-{pos.strip()}-{ref}-{alt}"


def percent_text(ac: int, an: int) -> str:
    """AC / AN as a percentage to 3 decimals like the reports (0.00119 -> 0.001);
    a value that would round to 0.000 keeps 3 significant figures instead."""
    pct = ac / an * 100
    if round(pct, 3) > 0:
        return f"{pct:.3f}"
    return f"{pct:.{2 - math.floor(math.log10(pct))}f}"


def _query(vid: str) -> dict | None:
    body = json.dumps({"query": QUERY, "variables": {"id": vid}}).encode()
    req = urllib.request.Request(API, body, {"Content-Type": "application/json",
                                             "User-Agent": "cmc-report-tool"})
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            out = json.load(r)
    except Exception as e:  # noqa: BLE001 - shown to the user
        raise GnomadError(f"Could not reach gnomAD: {e}") from e
    errors = [e.get("message", "") for e in out.get("errors") or []]
    if (out.get("data") or {}).get("variant") is None:
        if all("not found" in m.lower() for m in errors):     # none, or only 'Variant not found'
            return None
        raise GnomadError("gnomAD error: " + "; ".join(errors))
    return out["data"]["variant"]


def _total(joint: dict, group: str | None) -> tuple[int, int]:
    """AC and AN of exomes + genomes, for one ancestry group (None = everyone)."""
    row = joint if group is None else next(
        (p for p in joint["populations"] if p["id"] == group), None)
    return ((row["ac"] or 0), (row["an"] or 0)) if row else (0, 0)


def frequency_text(v: dict) -> str:
    """The Variant Frequency sentence for this variant, from gnomAD v4 (exomes + genomes)."""
    vid = variant_id(v)
    if not vid:
        raise GnomadError("Insertions / deletions can't be looked up automatically "
                          "(type the frequency, or check the variant on the gnomAD site).")
    joint = (_query(vid) or {}).get("joint")
    if not joint:
        return variants_core.FREQUENCY_ABSENT
    ac, an = _total(joint, "sas")
    if ac and an:
        return percent_text(ac, an) + variants_core.FREQUENCY_TEXT
    ac, an = _total(joint, None)
    if ac and an:
        return percent_text(ac, an) + variants_core.FREQUENCY_ALL
    return variants_core.FREQUENCY_ABSENT
