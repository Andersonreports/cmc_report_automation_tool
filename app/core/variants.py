"""Listed-variants Excel (VarSeq export) for the pathogenic templates.

The sheet has one or more blocks, each a title row (e.g. "Secondary Findings",
"LISTED VARIANT"), a header row with GENE and VARIANT columns (starting with
"S.NO", "REMARKS", ...) and one variant per row. Every variant is
kept as a flat dict of strings (see FIELDS) plus a "selected" flag the user
ticks in the app; only the selected ones are written into the report.
"""
from __future__ import annotations

import re
from functools import lru_cache

# Excel header (lower case) -> key in the variant dict.
HEADERS = {
    "s.no": "sno", "gene": "gene", "variant": "variant", "protein pos": "protein",
    "acmg": "acmg", "franklin": "franklin", "clinvar": "clinvar",
    "zygosity": "zygosity", "inheritance": "inheritance", "vaf(%)": "vaf",
    "allelic depth": "depth", "omim id": "omim_id", "omim title": "omim_title",
    "conditions": "conditions", "gnom ad exome sa v4": "gnomad_sa",
    "gnom ad exome all v4": "gnomad_all",
    "alpha missense": "alpha_missense", "revel": "revel",
    "chromosomal position": "chrom_pos", "ref/alt": "ref_alt", "nm id": "nm_id",
    "sequence ontology": "so", "exon": "exon", "disease label": "disease_label",
    # Optional columns: used when the sheet has them, "-" otherwise.
    "lrt": "lrt", "lrt pred": "lrt",
    "cadd": "cadd",
    "mt2": "mt2", "mutationtaster": "mt2", "mutationtaster pred": "mt2",
    "sift prediction (vertebrates)": "sift",
    "polyphen2 prediction (vertebrates)": "polyphen",
    "phylop": "phylop", "phylop prediction (vertebrates)": "phylop",
    "gerp": "gerp", "gerp++": "gerp", "gerp++ prediction (vertebrates)": "gerp",
    "strand": "strand", "transcript strand": "strand",
    "clinvar variant id": "clinvar_link",
    "acmg classification criteria": "criteria",
}
# In silico prediction tools / conservation scores: the labels are fixed; values
# are auto-populated from the Excel when the sheet has them.
SILICO = (("SIFT", "sift"), ("LRT", "lrt"), ("PolyPhen2", "polyphen"),
          ("GeneSplicer", "genesplicer"), ("MaxEntScan", "maxentscan"),
          ("NNSplice", "nnsplice"), ("PWM", "pwm"), ("CADD", "cadd"), ("MT2", "mt2"))     # every tool; see insilico_tools()
CONSERVATION = (("PhyloP", "phylop"), ("GERP++", "gerp"))


def silico_case(text: str) -> str:
    """In silico values have only the first letter capital: 'DA' -> 'Da', 'supp' -> 'Supp'."""
    return text[:1].upper() + text[1:].lower()
FIELDS = tuple(HEADERS.values()) + ("section", "selected")


class VariantError(Exception):
    pass


def _cell(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def _is_header(row: list[str]) -> bool:
    """A block's header row: it has GENE and VARIANT columns (it may start with
    S.NO, REMARKS or anything else)."""
    cells = {c.lower() for c in row}
    return "gene" in cells and "variant" in cells


def read_variants(path: str) -> list[dict]:
    """Every variant in the workbook, in sheet order, none selected."""
    from openpyxl import load_workbook
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as e:  # noqa: BLE001 - shown to the user
        raise VariantError(f"Could not open the Excel file: {e}") from e
    try:
        # VarSeq exports put "Complete Data" before "Listed Variants": read the latter.
        sheets = sorted(wb.worksheets, key=lambda w: w.title.strip().lower() != "listed variants")
        rows = []
        for ws in sheets:
            rows = [[_cell(c) for c in r] for r in ws.iter_rows(values_only=True)]
            if any(r and _is_header(r) for r in rows):
                break
    finally:
        wb.close()
    out, cols, section = [], None, ""
    for i, row in enumerate(rows):
        if row and _is_header(row):               # a header row starts a block
            cols = {HEADERS[h.lower()]: j for j, h in enumerate(row) if h.lower() in HEADERS}
            # The block's title row above it: "Secondary Findings", "LISTED VARIANT", ...
            section = next((r[0] for r in reversed(rows[:i])
                            if r and r[0] and not any(r[1:])), "")
            continue
        if cols is None or cols["gene"] >= len(row) or not row[cols["gene"]]:
            continue
        v = {k: (row[j] if j < len(row) else "") for k, j in cols.items()}
        # Strand is often embedded in the gene cell: "CACNA1S (+)" or "XDH (-)"
        if not v.get("strand"):
            gene_raw = str(v.get("gene", ""))
            m = re.search(r"\(\s*([+\-])\s*\)", gene_raw)
            if m:
                v["strand"] = m.group(1)
                v["gene"] = re.sub(r"\s*\(\s*[+\-]\s*\)", "", gene_raw).strip()
        v["section"], v["selected"] = section, False
        for typed in ("disease", "criteria", "frequency", "evidence", "omim_phenotype", "region",
                      "clinvar_link", "pubmed_link"):    # typed by the user in the app
            v.setdefault(typed, "")
        v["classification"] = (classification_options(v) or [""])[0]
        for _, key in SILICO + CONSERVATION:
            v.setdefault(key, "")   # keep value from Excel if already populated
        v["inheritance"] = ", ".join(  # merged codes; the sheet repeats them per phenotype
            dict.fromkeys(p.strip() for p in re.split(r"[,/]", v.get("inheritance", ""))
                          if p.strip() and "somatic" not in p.lower()))
        # The sheet lists every OMIM phenotype of the gene; the user keeps the right one.
        ids, titles, inh = _omim_phenotypes(v)
        v["omim_options"] = ", ".join(ids)
        v["omim_titles"] = "|".join(titles)       # aligned with the ids, "" where unknown
        v["omim_inh"] = "|".join(inh)             # AD / AR / "" per id, from OMIM genemap2
        v["omim_id"] = ", ".join(ids)           # all of the gene's; the user deletes the rest
        v["disease"] = suggested_disease(v)
        v["gene_mim"] = gene_mim_number(v.get("gene", ""))
        v["omim_phenotype_prefix"] = v["omim_phenotype"] = omim_phenotype_prefix(v)
        v["inheritance"] = suggested_inheritance(v) or v["inheritance"]
        out.append(v)
    if not out:
        raise VariantError("No variants found. Upload the VarSeq listed-variants Excel "
                           "(a header row with GENE and VARIANT columns, then one variant per row).")
    return out


# ---------------------------------------------------------------------------
# OMIM phenotypes (disease names)
# ---------------------------------------------------------------------------

_INH_CODE = {"autosomal dominant": "AD", "autosomal recessive": "AR"}


@lru_cache(maxsize=1)
def _genemap() -> dict[str, list[tuple[str, str]]]:
    """Gene symbol -> its OMIM phenotypes as (number, inheritance code), ascending,
    from the bundled omim_genemap2.txt. Empty if the file is missing."""
    from .os_utils import resource_path
    out: dict[str, list[tuple[str, str]]] = {}
    try:
        with open(resource_path("app", "data", "omim_genemap2.txt"), encoding="utf-8") as f:
            for line in f:
                if line.startswith("#"):
                    continue
                cols = line.rstrip("\n").split("\t")
                if len(cols) > 12 and cols[6].strip():
                    pheno = {m[0]: _INH_CODE.get(m[1].strip().lower(), "") for m in
                             re.findall(r"association, (\d+) \(\d\), ([^;]*)", cols[12])}
                    if pheno:
                        out[cols[6].strip().upper()] = sorted(pheno.items())
    except OSError:
        pass
    return out


@lru_cache(maxsize=1)
def _gene_mim_map() -> dict[str, str]:
    """Gene symbol (upper) -> gene MIM number, from the bundled omim_mim2gene.txt.
    Only 'gene' type entries are included (not phenotype-only MIM numbers)."""
    from .os_utils import resource_path
    out: dict[str, str] = {}
    try:
        with open(resource_path("app", "data", "omim_mim2gene.txt"), encoding="utf-8") as f:
            for line in f:
                if line.startswith("#"):
                    continue
                cols = line.rstrip("\n").split("\t")
                if len(cols) >= 4 and cols[1].strip() == "gene" and cols[3].strip():
                    out[cols[3].strip().upper()] = cols[0].strip()
    except OSError:
        pass
    return out


def gene_mim_number(gene: str) -> str:
    """Return the OMIM gene MIM number for a gene symbol, or '' if not found."""
    return _gene_mim_map().get(gene.strip().upper(), "") if gene else ""


def _omim_phenotypes(v: dict) -> tuple[list[str], list[str], list[str]]:
    """The gene's OMIM numbers and, aligned with them, the disease names the sheet
    lists in OMIM TITLE (same order as its OMIM ID column, joined by a bare comma;
    commas inside a name are followed by a space) and the inheritance codes."""
    sheet_ids = [i for i in re.split(r"[,\s]+", v.get("omim_id", "")) if i.isdigit()]
    titles = [t.strip() for t in re.split(r",(?! )", v.get("omim_title", "")) if t.strip()]
    known = _genemap().get(v.get("gene", "").strip().upper())
    ids = [i for i, _ in known] if known else sheet_ids
    inh = [c for _, c in known] if known else [""] * len(ids)
    if len(titles) == len(ids):
        return ids, titles, inh
    if len(titles) == len(sheet_ids):       # same order as the sheet's own ids
        by_id = dict(zip(sheet_ids, titles))
        return ids, [by_id.get(i, "") for i in ids], inh
    return ids, [""] * len(ids), inh


def _disease_name(title: str) -> str:
    """'BREAST-OVARIAN CANCER, FAMILIAL, SUSCEPTIBILITY TO, 2; BROVCA2' ->
    '{Breast-ovarian cancer, familial, 2}' (OMIM braces mark susceptibility)."""
    name = title.split(";")[0].strip()
    braces = bool(re.search(r",\s*SUSCEPTIBILITY TO(?=,|$)", name, re.I))
    name = re.sub(r",\s*SUSCEPTIBILITY TO(?=,|$)", "", name, flags=re.I)
    words = []
    for w in name.split():              # keep 1A / D1 / III as they are, and "type I"
        roman = re.fullmatch(r"[IVX]{2,}", w) or (
            re.fullmatch(r"[IVX]", w) and words and words[-1] == "type")
        words.append(w if re.search(r"\d", w) or roman else w.lower())
    name = " ".join(words)
    name = name[:1].upper() + name[1:]
    return f"{{{name}}}" if braces else name


def _chosen_omim(v: dict, key: str) -> list[str]:
    """Values of `key` ("omim_titles" / "omim_inh") for the numbers in the OMIM ID box."""
    by_id = dict(zip([i.strip() for i in v.get("omim_options", "").split(",")],
                     v.get(key, "").split("|")))
    return [by_id.get(i, "") for i in re.split(r"[,\s]+", v.get("omim_id", "")) if i]


def suggested_disease(v: dict) -> str:
    """Disease or Disorder from the OMIM number(s) in the OMIM ID box."""
    return "; ".join(_disease_name(t) for t in _chosen_omim(v, "omim_titles") if t)


def omim_ids_in(v: dict, disease: str) -> list[str]:
    """The OMIM numbers whose disease name is still in the Disease box text."""
    ids = [i.strip() for i in v.get("omim_options", "").split(",")]
    return [i for i, t in zip(ids, v.get("omim_titles", "").split("|"))
            if t and _disease_name(t) in disease]


def omim_phenotype_prefix(v: dict) -> str:
    """Start of the OMIM phenotype paragraph: '{Disease} (OMIM#id) ... gene (OMIM*gene_mim)';
    the rest of the sentence is typed by the user."""
    disease, ids = v.get("disease", "").strip(), v.get("omim_id", "").strip()
    if not disease:
        return ""
    text = f"{disease} (OMIM#{ids})" if ids else disease
    gene = v.get("gene", "").strip()
    gmim = v.get("gene_mim", "") or gene_mim_number(gene)
    if gmim and gene:
        text += f" is caused by mutation in the {gene} gene (OMIM*{gmim})."
    return text


def suggested_inheritance(v: dict) -> str:
    """Inheritance codes of the OMIM number(s) in the OMIM ID box, e.g. 'AD, AR'."""
    return ", ".join(dict.fromkeys(c for c in _chosen_omim(v, "omim_inh") if c))


# ---------------------------------------------------------------------------
# Wording for the report
# ---------------------------------------------------------------------------

_ALLELE = {"HET": "Heterozygous", "HOM": "Homozygous", "HEMI": "Hemizygous"}
_INHERITANCE = {"AD": "Autosomal Dominant", "AR": "Autosomal Recessive",
                "XLR": "X-linked Recessive", "XLD": "X-linked Dominant",
                "XL": "X-linked", "MT": "Mitochondrial", "YL": "Y-linked"}


_CLASS = {"p": "Pathogenic", "lp": "Likely Pathogenic", "vus": "Variant of Uncertain Significance",
          "lb": "Likely Benign", "b": "Benign"}


def _class_label(text: str) -> str:
    t = text.replace("↗", "").strip()
    low = t.lower()
    if low in _CLASS:
        return _CLASS[low]
    if low.startswith("vus") or "uncertain" in low:
        return _CLASS["vus"]
    return t


def classification_options(v: dict) -> list[str]:
    """The Franklin classification first (the report's default; ACMG, then ClinVar,
    only when Franklin is empty), then Pathogenic / Likely Pathogenic / VUS if
    missing - the user can switch to any of them."""
    first = next((lab for lab in (_class_label(v.get(k, ""))
                                  for k in ("franklin", "acmg", "clinvar")) if lab), "")
    out = [first] if first else []
    return out + [c for c in (_CLASS["p"], _CLASS["lp"], _CLASS["vus"]) if c not in out]


def classification(v: dict) -> str:
    """The chosen classification (default: the first option)."""
    return v.get("classification") or (classification_options(v) or ["-"])[0]


def class_display(v: dict) -> str:
    """Classification as the report prints it (VUS -> 'Uncertain Significance')."""
    label = classification(v)
    return "Uncertain Significance" if label == _CLASS["vus"] else label


# The result tag the user can pick (Variants tab): key -> sentence in the box;
# {gene} is the ticked variants' gene, in italics.
RESULT_TAGS = (
    ("pathogenic", "Pathogenic variant related to the given phenotype was identified"),
    ("likely", "Likely pathogenic variant related to the given phenotype was identified"),
    ("vus_related", "Variant of uncertain significance related to the given phenotype was identified"),
    # As the lab's newer WES / WES & Mito reports word it (LAB_RESULT_TAGS).
    ("likely_identified", "Likely Pathogenic variant was identified related to the given phenotype"),
    ("het_pathogenic_gene", "Heterozygous Pathogenic variant was identified in {gene} gene"),
    ("vus_gene", "Variant of Uncertain Significance was identified in {gene} gene"),
    ("compound_het", "Likely Compound Heterozygous variants were identified related to the "
                     "given phenotype"),
    ("compound_het_gene", "Likely Compound Heterozygous variants were identified in {gene} gene"),
    # Extended result lines (manually chosen by the user).
    ("no_pathogenic_detected",
     "No pathogenic or likely pathogenic variant(s) causative of the reported phenotype was detected"),
    ("patho_causative_related",
     "Pathogenic variant(s) causative to the related phenotype was identified"),
    ("patho_in_gene",
     "Pathogenic variant(s) was identified in {gene} gene"),
    ("patho_causative_reported",
     "Pathogenic variant(s) causative of the reported phenotype was identified"),
    ("patho_identified",
     "Pathogenic variant(s) was identified"),
    ("patho_of_phenotype",
     "Pathogenic variant(s) of the reported phenotype was identified"),
    ("likely_causative_related",
     "Likely Pathogenic variant(s) causative to the related phenotype was identified"),
    ("likely_causative_reported",
     "Likely pathogenic variant(s) causative of the reported phenotype was identified"),
    ("likely_of_phenotype",
     "Likely pathogenic variant(s) of the reported phenotype was identified"),
    ("likely_in_gene",
     "Likely pathogenic variant(s) was identified in {gene} gene"),
    ("likely_identified_only",
     "Likely Pathogenic variant(s) was identified"),
    ("vus_causative_reported",
     "Variant of uncertain significance causative of the reported phenotype was identified"),
    ("vus_causative_related",
     "Variant of uncertain significance causative to the related phenotype was identified"),
    ("vus_in_gene",
     "Variant of uncertain significance was identified in {gene} gene"),
    ("carrier_likely_gene",
     "Carrier of likely pathogenic variant was identified in {gene} gene"),
    ("carrier_patho_gene",
     "Carrier of pathogenic variant was identified in {gene} gene"),
    ("compound_het_phenotype",
     "Likely compound heterozygous variants related to the phenotype were identified"),
    ("compound_het_in_gene",
     "Likely compound heterozygous variants were identified in {gene} gene"),
    ("compound_het_related",
     "Likely compound heterozygous variants were identified related to the phenotype"),
    ("homo_vus_gene",
     "Homozygous variant of uncertain significance was identified in {gene} gene"),
    ("homo_related",
     "Homozygous variant identified related to the phenotype"),
    ("homo_likely_gene",
     "Homozygous likely pathogenic variant was identified in {gene} gene"),
)
# Offered only for templates with lab_wording (not the CMC ones).
LAB_RESULT_TAGS = frozenset({"likely_identified", "het_pathogenic_gene", "vus_gene",
                             "compound_het", "compound_het_gene"})
# Each tag's colour in the report, as the lab's reports print it.
RESULT_COLORS = {
    "pathogenic": "FF0000", "likely": "FF0000", "vus_related": "0070C0",
    "likely_identified": "FF0000", "het_pathogenic_gene": "002060",
    "vus_gene": "0070C0", "compound_het": "002060", "compound_het_gene": "002060",
    # Extended tags.
    "no_pathogenic_detected": "00B050",
    "patho_causative_related": "FF0000", "patho_in_gene": "FF0000",
    "patho_causative_reported": "FF0000", "patho_identified": "FF0000",
    "patho_of_phenotype": "FF0000",
    "likely_causative_related": "FF0000", "likely_causative_reported": "FF0000",
    "likely_of_phenotype": "FF0000", "likely_in_gene": "FF0000",
    "likely_identified_only": "FF0000",
    "vus_causative_reported": "0070C0", "vus_causative_related": "0070C0",
    "vus_in_gene": "0070C0",
    "carrier_likely_gene": "002060", "carrier_patho_gene": "002060",
    "compound_het_phenotype": "002060", "compound_het_in_gene": "002060",
    "compound_het_related": "002060",
    "homo_vus_gene": "0070C0", "homo_related": "FF0000", "homo_likely_gene": "FF0000",
}
# The kind of result each tag reports: the classification it gives the ticked variants
# (none for compound heterozygous, whose variants keep their own) and the auto heading.
RESULT_KIND = {
    "pathogenic": "pathogenic", "likely": "likely", "vus_related": "vus_related",
    "likely_identified": "likely", "het_pathogenic_gene": "pathogenic",
    "vus_gene": "vus_related", "compound_het": "", "compound_het_gene": "",
    # Extended tags.
    "no_pathogenic_detected": "",
    "patho_causative_related": "pathogenic", "patho_in_gene": "pathogenic",
    "patho_causative_reported": "pathogenic", "patho_identified": "pathogenic",
    "patho_of_phenotype": "pathogenic",
    "likely_causative_related": "likely", "likely_causative_reported": "likely",
    "likely_of_phenotype": "likely", "likely_in_gene": "likely",
    "likely_identified_only": "likely",
    "vus_causative_reported": "vus_related", "vus_causative_related": "vus_related",
    "vus_in_gene": "vus_related",
    "carrier_likely_gene": "pathogenic", "carrier_patho_gene": "pathogenic",
    "compound_het_phenotype": "", "compound_het_in_gene": "", "compound_het_related": "",
    "homo_vus_gene": "vus_related", "homo_related": "pathogenic",
    "homo_likely_gene": "likely",
}

# Titles the user can pick for the Secondary findings table (Variants tab); the
# first is the template's own. Any other wording can be typed in.
SECONDARY_TITLES = (
    "List of Secondary findings identified:",
    "List of pathogenic Secondary findings identified:",
    "List of likely pathogenic Secondary findings identified:",
    "List of Secondary findings of uncertain significance identified:",
)

# Titles the user can pick for the PRIMARY variant list table heading.
PRIMARY_TITLES = (
    "List of significant variant(s) identified related to the phenotype:",
    "List of pathogenic variant(s) identified related to the phenotype:",
    "List of likely pathogenic variant(s) identified related to the phenotype:",
    "List of uncertain significant variant(s) identified related to the phenotype:",
    "List of additional significant variant(s) identified:",
    "List of Incidental significant variant(s) identified:",
    "List of significant variant(s) identified:",
    "List of significant carrier variant(s) identified:",
    # As the lab's newer WES / WES & Mito reports word it (LAB_PRIMARY_TITLES).
    "List of significant variant(s) identified related to the given phenotype:",
    "List of significant variant identified:",
    "List of uncertain significant variant(s) identified:",
    "List of additional uncertain significant variant(s) identified:",
)
# Offered only for templates with lab_wording (not the CMC ones).
LAB_PRIMARY_TITLES = frozenset(PRIMARY_TITLES[-4:])

ADDITIONAL_VARIANTS_DEFAULT_TEXT = (
    "The additional variants identified which are significant but may not be related to "
    "patient’s phenotype. Phenotype – genotype correlation is recommended.")

# The note under the variant list for two variants in one gene (Report options).
COMPOUND_HET_NOTE = ("The variants are reported to be in likely compound heterozygous state. "
                     "These likely compound heterozygous variants are strongly recommended to "
                     "confirm the Cis or Trans status by parental segregation analysis.")


def is_secondary(v: dict) -> bool:
    """A variant for the Secondary findings table: the user's pick in the app
    ("in_secondary"), else whether it came from the Excel's "Secondary Findings" block."""
    if "in_secondary" in v:
        return bool(v["in_secondary"])
    return "secondary" in v.get("section", "").lower()


def report_mode(chosen: list[dict], pick: str = "") -> str:
    """The result tag: the user's pick (a RESULT_TAGS key), else from the ticked
    variants (see _auto_mode)."""
    return pick if pick in dict(RESULT_TAGS) else _auto_mode(chosen)


def _auto_mode(chosen: list[dict]) -> str:
    """What the ticked variants make the report say: 'pathogenic' (any Pathogenic),
    'likely' (Likely Pathogenic, no Pathogenic) or 'vus_related' (neither)."""
    labels = {classification(v) for v in chosen}
    if _CLASS["p"] in labels:
        return "pathogenic"
    return "likely" if _CLASS["lp"] in labels else "vus_related"


def allele_status(v: dict) -> str:
    z = v.get("zygosity", "").strip()
    return _ALLELE.get(z.upper(), z or "-")


def inheritance(v: dict) -> list[str]:
    """One line per word of each inheritance pattern so multi-word labels
    ('Autosomal Dominant') always wrap consistently regardless of column width."""
    parts = []
    for p in re.split(r"[,/]", v.get("inheritance", "")):
        p = p.strip()
        if p and "somatic" not in p.lower() and p.upper() not in [q.upper() for q in parts]:
            parts.append(p)
    lines = []
    for n, p in enumerate(parts):
        words = _INHERITANCE.get(p.upper(), p).split()  # "Autosomal Dominant" → 2 lines
        if n < len(parts) - 1:
            words[-1] += "/"            # "Autosomal Dominant/ Autosomal Recessive", as the lab writes it
        lines.extend(words)
    return lines or ["-"]


def disease_lines(v: dict) -> list[str]:
    """Disease or Disorder cell, as in the lab's reports: each disease the user typed
    (separated by ';', e.g. '{Breast-ovarian cancer, familial, 2}') on its own line
    with its own OMIM number(s) under it as '(OMIM#612555)'. Numbers that match no
    typed disease go on a last line; text that matches none stays as one line."""
    disease = v.get("disease", "").strip()
    chosen = [i for i in re.split(r"[,\s]+", v.get("omim_id", "")) if i]
    if not disease:
        return [f"(OMIM#{', '.join(chosen)})"] if chosen else ["-"]
    names = [(i.strip(), _disease_name(t)) for i, t in
             zip(v.get("omim_options", "").split(","), v.get("omim_titles", "").split("|")) if t]
    parts = [p.strip() for p in disease.split(";") if p.strip()]
    lines, used = [], []
    for part in parts:
        ids = [i for i in chosen if i not in used and any(
            n == i and name and name in part for n, name in names)]
        used += ids
        lines.append(part)
        if ids:
            lines.append(f"(OMIM#{', '.join(ids)})")
    if not used:                    # typed freely: one line, every number under it
        return [disease] + ([f"(OMIM#{', '.join(chosen)})"] if chosen else [])
    rest = [i for i in chosen if i not in used]
    if rest:
        lines.append(f"(OMIM#{', '.join(rest)})")
    return lines


def region(v: dict) -> str:
    """The Region cell: what the user typed, else 'Exon <the Excel's exon(s)>'."""
    if v.get("region", "").strip():
        return v["region"].strip()
    exon = v.get("exon", "")
    if exon:
        return f"Exon {exon}"
    return "Intron" if re.search(r"c\.[-*]?\d+[+-]\d+", v.get("variant", "")) else "-"


_TOOLS_BY_EFFECT = {
    "splice": ("genesplicer", "maxentscan", "nnsplice", "pwm", "cadd", "mt2"),
    "lof": ("lrt", "cadd", "mt2"),                      # stop gained / frameshift
    "missense": ("sift", "lrt", "polyphen", "cadd", "mt2"),
}


def insilico_tools(v: dict) -> tuple[tuple[str, str], ...]:
    """The in silico tools the report lists for this variant, by its Sequence
    Ontology: splice site / region -> the splice predictors + CADD, MT2; stop gained
    or frameshift -> LRT, CADD, MT2; missense -> SIFT, LRT, PolyPhen2, CADD, MT2.
    Anything else (synonymous, in-frame ...) -> CADD, MT2."""
    so = v.get("so", "").lower()
    if "splice" in so:
        keys = _TOOLS_BY_EFFECT["splice"]
    elif "stop_gained" in so or "stop gained" in so or "frameshift" in so:
        keys = _TOOLS_BY_EFFECT["lof"]
    elif "missense" in so:
        keys = _TOOLS_BY_EFFECT["missense"]
    else:
        keys = ("cadd", "mt2")
    return tuple((label, key) for label, key in SILICO if key in keys)


def sequence_ontology(v: dict) -> str:
    return v.get("so", "").replace("_", " ").title() or "-"


def allelic_depth(v: dict) -> str:
    ref, _, alt = (v.get("ref_alt", "") or "/").partition("/")
    depth = [d.strip() for d in v.get("depth", "").split(",")]
    try:
        vaf = f"{float(v['vaf']) * 100:.2f}"
    except (KeyError, ValueError):
        vaf = ""
    if len(depth) != 2:
        return f"VAF: {vaf}%" if vaf else "-"
    return (f"Ref({ref or '-'}): {depth[0]}, Alt({alt or '-'}): {depth[1]}"
            + (f", VAF: {vaf}%" if vaf else ""))


FREQUENCY_TEXT = "% frequency observed in annotated gnomAD v4 South Asian"
FREQUENCY_ALL = "% frequency observed in annotated gnomAD v4 All"
FREQUENCY_ABSENT = "Not observed in annotated gnomAD v4 All"


def default_frequency(v: dict) -> str:
    """The sentence the Variant Frequency box starts with, by the sheet's
    GNOM AD EXOME SA V4 cell: a count -> South Asian sentence (the user types the
    number before it); empty -> 'Not observed in ...'; 0 -> the 'all' sentence."""
    sa = str(v.get("gnomad_sa", "")).strip()
    if not sa:
        return FREQUENCY_ABSENT
    return FREQUENCY_ALL if sa in ("0", "0.0") else FREQUENCY_TEXT


def frequency(v: dict) -> str:
    """Variant Frequency cell: what the box holds, or its default sentence when empty.
    A bare number gets the default sentence after it."""
    text = v.get("frequency", "").strip()
    default = default_frequency(v)
    if not text:
        return default.strip()      # the sentence shows even before the number is typed
    if not re.search(r"[A-Za-z]", text):
        return text.rstrip("%").strip() + default
    return text


_NOISE = re.compile(r"\s*\*?\s*Genetic test results are based on the recommendation.*$|"
                    r"\s*No other variant that warrants to be reported.*$|"
                    r"\s*Page \d+ of \d+", re.I | re.S)


def clean_narrative(text: str) -> str:
    """A paragraph pasted from a PDF / report: line breaks become spaces, spaces are
    collapsed, and the notes that follow it in the report ('*Genetic test results are
    based on...', 'No other variant that warrants...', page numbers) are cut off."""
    return " ".join(_NOISE.sub("", " ".join(text.split())).split())


def suggested_evidence(v: dict) -> str:
    """Starting text of "Clinical and Literature Evidences" from the links given:
    a ClinVar link -> the ClinVar sentence, a PubMed link -> the literature
    sentence, both -> both. The user can edit it afterwards."""
    parts = []
    if v.get("clinvar_link", "").strip():
        label = _class_label(v.get("clinvar", "")) or classification(v)
        if label == _CLASS["vus"]:
            label = "uncertain significance"
        parts.append(f"This variant has been classified as {label} in ClinVar database.")
    if v.get("pubmed_link", "").strip():
        kind = sequence_ontology(v).lower()
        kind = f"{kind} variant" if kind != "-" else "variant"
        article = "An" if kind[0] in "aeiou" else "A"
        pos = ", ".join(x for x in (v.get("variant", ""), v.get("protein", "")) if x)
        names = [re.sub(r",\s*\d+$", "", n.strip().strip("{}").strip())
                 for n in v.get("disease", "").split(";")]
        disease = "; ".join(n for n in names if n) or "the disease"
        state = allele_status(v).lower()
        parts.append(
            f"This variant has been previously reported in patients affected with {disease} "
            f"in {state} state. "
            f"{article} {kind} in the same amino acid position ({pos}) has been "
            f"previously reported in patients affected with {disease} "
            f"in {state} state."
        )
    return " ".join(parts)


# RefSeq chromosome accessions of GRCh38 (the sheet's coordinates are GRCh38).
NC_GRCH38 = {
    "1": "NC_000001.11", "2": "NC_000002.12", "3": "NC_000003.12", "4": "NC_000004.12",
    "5": "NC_000005.10", "6": "NC_000006.12", "7": "NC_000007.14", "8": "NC_000008.11",
    "9": "NC_000009.12", "10": "NC_000010.11", "11": "NC_000011.10", "12": "NC_000012.12",
    "13": "NC_000013.11", "14": "NC_000014.9", "15": "NC_000015.10", "16": "NC_000016.10",
    "17": "NC_000017.11", "18": "NC_000018.10", "19": "NC_000019.10", "20": "NC_000020.11",
    "21": "NC_000021.9", "22": "NC_000022.11", "X": "NC_000023.11", "Y": "NC_000024.10",
    "M": "NC_012920.1", "MT": "NC_012920.1",
}


def _hgvs_g(pos: int, ref: str, alt: str) -> str:
    """g. change from the sheet's VCF-style position / ref / alt. Substitutions and
    plain deletions / insertions only; an insertion that is really a duplication
    needs the genome sequence, so the user can correct it in the Genomic Position box."""
    if ref == "-":                                                  # insertion, no anchor base
        return f"g.{pos}_{pos + 1}ins{alt}"
    if alt == "-":                                                  # deletion, no anchor base
        end = pos + len(ref) - 1
        return f"g.{pos}del" if end == pos else f"g.{pos}_{end}del"
    if len(ref) == 1 and len(alt) == 1:
        return f"g.{pos}{ref}>{alt}"
    if len(ref) > len(alt) and ref.startswith(alt):                 # deletion
        start, end = pos + len(alt), pos + len(ref) - 1
        return f"g.{start}del" if start == end else f"g.{start}_{end}del"
    if len(alt) > len(ref) and alt.startswith(ref):                 # insertion
        return f"g.{pos + len(ref) - 1}_{pos + len(ref)}ins{alt[len(ref):]}"
    return f"g.{pos}{ref}>{alt}"


def expected_nc_number(v: dict) -> int:
    """The number inside the NC accession for the variant's chromosome (13 for
    NC_000013.x, 23 for X); 0 if unknown. Used to check an accession read from an image."""
    chrom = re.sub(r"^chr", "", v.get("chrom_pos", "").partition(":")[0].strip(), flags=re.I).upper()
    nc = NC_GRCH38.get(chrom, "")
    return int(nc[3:9]) if nc.startswith("NC_0000") else 0


def auto_genomic_position(v: dict, nc: str = "") -> str:
    """'Chr13:NC_000013.11:g.32338140C>A' from the Excel's position and ref/alt;
    `nc` replaces the built-in accession (e.g. one read from an image)."""
    chrom, _, pos = v.get("chrom_pos", "").strip().partition(":")
    chrom = re.sub(r"^chr", "", chrom.strip(), flags=re.I).upper()
    ref, _, alt = v.get("ref_alt", "").strip().partition("/")
    nc = nc or NC_GRCH38.get(chrom)
    if not (chrom and pos.strip().isdigit() and nc and ref and alt):
        return f"Chr{chrom}:{pos} {v.get('ref_alt', '').replace('/', '>')}".strip()
    return f"Chr{chrom}:{nc}:{_hgvs_g(int(pos), ref, alt)}"


def genomic_text(v: dict, nc: str = "") -> str:
    """What the Genomic Position box starts with: the built position and, after it,
    the Excel's transcript in brackets: 'Chr13:NC_000013.11:g.32338140C>A (NM_000059.4)'."""
    nm = v.get("nm_id", "").strip()
    return auto_genomic_position(v, nc) + (f" ({nm})" if nm else "")


def genomic_position(v: dict) -> tuple[str, str]:
    """Genomic Position cell as two lines: the position and, when the text ends with
    a bracketed transcript, that '(NM_...)' on its own line. The user's own text is
    used when typed, otherwise the built one."""
    text = v.get("genomic_position", "").strip() or genomic_text(v)
    m = re.search(r"\s*(\([A-Z]{2}_[^)]*\))\s*$", text)
    return (text[:m.start()].strip(), m.group(1)) if m else (text, "")


PSEUDOGENE_GENES: frozenset[str] = frozenset({
    "BRCA1", "PMS2", "CHEK2", "PKD1", "STRC", "CYP21A2",
    "GBA", "GBA1", "SMN1", "SMN2", "VWF", "IKBKG", "HYDIN",
})


def pseudogene_genes(chosen: list[dict]) -> list[str]:
    """Selected variant genes that have known pseudogene(s), in order."""
    seen: set[str] = set()
    out: list[str] = []
    for v in chosen:
        g = v.get("gene", "").strip()
        if g.upper() in {s.upper() for s in PSEUDOGENE_GENES} and g not in seen:
            seen.add(g)
            out.append(g)
    return out


CLINVAR_URL = "https://www.ncbi.nlm.nih.gov/clinvar/variation/{}/"
PUBMED_URL = "https://pubmed.ncbi.nlm.nih.gov/{}/"


def evidence_link(text: str, url_template: str):
    """What the user typed for a ClinVar / PubMed link -> (evidence number, url),
    shown as a hyperlink. A bare number gets the standard address; the number is
    the last run of digits of a pasted link. Nothing typed -> '-'."""
    text = text.strip()
    nums = re.findall(r"\d+", text)
    if not text or not nums:
        return "-"
    url = text if re.match(r"https?://", text, re.I) else url_template.format(nums[-1])
    return (nums[-1], url)
