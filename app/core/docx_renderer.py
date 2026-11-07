"""Fill a bundled WES Word template with a ReportData.

Only the user-entered parts are touched: the patient demography table, the
clinical history paragraph, the Sequence data attributes values, the
Methodology gene list link ("Click here"), the Methodology opening, which
follows the chosen specimen, the Appendix 1 gene coverage table and the
signature block, which follows the chosen clinical reviewer. All other
content (results, CNV findings, recommendations, methodology, disclaimer,
references, signatures, appendix) is left exactly as in the template.

Layout-only fixes keep the fixed content well laid out whatever the length
of the clinical history:
  * section headings are kept with the text that follows them;
  * the footer's typed "of 4" becomes a live NUMPAGES field;
  * the empty paragraphs padding page 3 are replaced by a page break before
    the appendix, so it always starts on a fresh page.
Elements are located by their text, not their position.
"""
from __future__ import annotations

import copy
import os
import re

from docx import Document
from docx.oxml import OxmlElement
from docx.shared import Cm, Pt
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.text.run import Run
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as _RT
from docx.text.paragraph import Paragraph

from ..config.templates import REVIEWERS, effective, get_template, option_values
from .models import ReportData
from .os_utils import resource_path

PATIENT_CELLS = [  # (row, col, label, attribute) of the usual layout; filled by label
    (0, 0, "Patient ID", "patient_id"), (0, 1, "PIN", "pin"),
    (1, 0, "Age", "age"), (1, 1, "Gender", "gender"),
    (2, 0, "Hospital/Clinic", "hospital"), (2, 1, "Sample Number", "sample_number"),
    (3, 0, "Referring Clinician", "referring_clinician"),
    (3, 1, "Sample Collection Date", "collection_date"),
    (4, 0, "Specimen", "specimen"), (4, 1, "Sample Received Date", "received_date"),
    (5, 1, "Report Date", "report_date"),
]


class TemplateError(Exception):
    pass


# ---------------------------------------------------------------------------
# XML helpers
# ---------------------------------------------------------------------------

_HYPERLINK = _RT.HYPERLINK


def _runs(p_el):
    return [r for r in p_el.iter(qn("w:r")) if r.find(qn("w:t")) is not None]


def _text(p_el) -> str:
    return "".join(t.text or "" for t in p_el.iter(qn("w:t")))


def _set_text(p_el, text: str) -> None:
    """Replace a paragraph's text, keeping the first run's formatting.
    Line breaks in `text` become Word line breaks."""
    runs = _runs(p_el)
    if runs:
        first = runs[0]
        for r in runs[1:]:
            r.getparent().remove(r)
        for child in list(first):
            if child.tag != qn("w:rPr"):
                first.remove(child)
    else:
        first = OxmlElement("w:r")
        rpr = p_el.find(qn("w:pPr") + "/" + qn("w:rPr"))
        if rpr is not None:
            first.append(copy.deepcopy(rpr))
        p_el.append(first)
    for i, line in enumerate(text.split("\n")):
        if i:
            first.append(OxmlElement("w:br"))
        t = OxmlElement("w:t")
        t.text = line
        t.set(qn("xml:space"), "preserve")
        first.append(t)


def _find_para(doc, prefix: str, what: str):
    for p in doc.element.body.iter(qn("w:p")):
        if _text(p).strip().lower().startswith(prefix.lower()):
            return p
    raise TemplateError(f"Template is missing: {what}")


def _find_table(doc, first_cell: str, what: str):
    for table in doc.tables:
        if table.rows and table.rows[0].cells[0].text.strip().lower().startswith(first_cell.lower()):
            return table
    raise TemplateError(f"Template is missing: {what}")


# ---------------------------------------------------------------------------
# User-entered sections
# ---------------------------------------------------------------------------

def _fill_patient(doc, data: ReportData):
    """Each box of the table is filled by its label, so templates may arrange
    the boxes differently or leave some out (e.g. no Hospital/Clinic)."""
    table = _find_table(doc, "Patient", "patient details table")
    id_label = _cfg(data).patient_id_label
    attrs = {label: attr for _, _, label, attr in PATIENT_CELLS}
    attrs["Patient name"] = "patient_id"
    for row in table.rows:
        for cell in row.cells:
            p = cell.paragraphs[0]._p
            attr = attrs.get(_text(p).split(":")[0].strip())
            if not attr:
                continue                # an empty box stays empty
            label = id_label if attr == "patient_id" else _text(p).split(":")[0].strip()
            value = (getattr(data.patient, attr, "") or "").strip()
            _set_text(p, f"{label}: {value}")
    if _cfg(data).own_styles:
        return                      # the template's own padding and indents stay
    # No paragraph indentation or spacing inside the demography table; the gap
    # to the cell borders comes from the table's cell margins instead, so the
    # text doesn't touch the lines.
    tbl_pr = table._tbl.tblPr
    old = tbl_pr.find(qn("w:tblCellMar"))
    if old is not None:
        tbl_pr.remove(old)
    mar = OxmlElement("w:tblCellMar")
    for side in ("left", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), "113")            # 0.2 cm
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tbl_pr.append(mar)
    for row in table.rows:
        for cell in row.cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for p in cell.paragraphs:
                pf = p.paragraph_format
                pf.left_indent = pf.right_indent = pf.first_line_indent = 0
                pf.space_before = pf.space_after = 0
                pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT


def _fill_history(doc, data: ReportData):
    head = _find_para(doc, "Clinical History", "'Clinical History' heading")
    body = head.getnext()
    while body is not None and (body.tag != qn("w:p") or not _text(body).strip()):
        body = body.getnext()
    if body is None:
        raise TemplateError("Template is missing the clinical history paragraph")
    # One flowing, justified paragraph: a single line break (e.g. from pasted,
    # pre-wrapped text) becomes a space, as a forced break would stretch the
    # line before it. A blank line still starts a new paragraph.
    paragraphs = [" ".join(part.split())
                  for part in re.split(r"\n\s*\n", data.clinical_history.strip())]
    paragraphs = [p for p in paragraphs if p] or [""]
    _set_text(body, paragraphs[0])
    anchor = body
    for text in paragraphs[1:]:
        extra = copy.deepcopy(body)
        _set_text(extra, text)
        anchor.addnext(extra)
        anchor = extra


def _gene_list_link(doc):
    """The Methodology paragraph's 'Click here' hyperlink element."""
    for p in doc.element.body.iter(qn("w:p")):
        link = p.find(qn("w:hyperlink"))
        if link is not None and link.get(qn("r:id")):
            return link
    raise TemplateError("Template is missing the gene list 'Click here' link")


def _fill_gene_list_url(doc, data: ReportData):
    url = data.gene_list_url.strip()
    if url:
        doc.part.rels[_gene_list_link(doc).get(qn("r:id"))]._target = url


def _normalize_hyperlink_colors(doc, hex_color: str = "0563C1"):
    """Set every hyperlink run's colour to hex_color throughout the document."""
    for link in doc.element.body.iter(qn("w:hyperlink")):
        for run in link.iter(qn("w:r")):
            rpr = run.find(qn("w:rPr"))
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                run.insert(0, rpr)
            color = rpr.find(qn("w:color"))
            if color is None:
                color = OxmlElement("w:color")
                rpr.append(color)
            color.set(qn("w:val"), hex_color)


def template_gene_list_url(template_key: str) -> str:
    """The gene list link stored in a template ('' if it has none)."""
    doc = Document(template_path(template_key))
    try:
        return doc.part.rels[_gene_list_link(doc).get(qn("r:id"))].target_ref
    except TemplateError:
        return ""


# Specimens named in the Methodology opening: "DNA was used to perform whole exome
# using whole exome capture kit .The" becomes "DNA extracted from the <specimen>
# was used to perform whole exome using whole exome capture kit. The" (or clinical
# exome, as the template says). DNA keeps the template wording as is.
METHODOLOGY_SPECIMENS = ("Peripheral Blood", "Amniotic fluid", "CVS", "Product Of Conception")

# Specimens that require the PCPNDT notice in the report.
_PCPNDT_SPECIMENS = frozenset({"Amniotic fluid", "CVS", "Product Of Conception"})
_PCPNDT_TEXT = ("This test confers to the PCPNDT act and does not determine "
                "the sex of the fetus")


def _replace_in_para(p_el, old: str, new: str) -> bool:
    """Replace `old` with `new` in a paragraph even when `old` spans several
    runs. The new text goes into the first run touched, so it keeps that
    run's formatting; the rest of the paragraph is left as is."""
    ts = [t for t in p_el.iter(qn("w:t"))]
    full = "".join(t.text or "" for t in ts)
    start = full.find(old)
    if start < 0:
        return False
    end, pos, placed = start + len(old), 0, False
    for t in ts:
        text = t.text or ""
        a, b = pos, pos + len(text)
        pos = b
        if b <= start or a >= end:
            continue
        keep_head = text[:max(start - a, 0)]
        keep_tail = text[max(end - a, 0):] if b > end else ""
        t.text = keep_head + ("" if placed else new) + keep_tail
        t.set(qn("xml:space"), "preserve")
        placed = True
    return True


def _insert_pcpndt_note(doc, data: ReportData):
    """Insert the PCPNDT notice (italic, centred) immediately after the
    demography table for specimens that require it.
    Font matches the document body: Arial 9 pt, color 242629."""
    if data.patient.specimen.strip() not in _PCPNDT_SPECIMENS:
        return
    tbl_el = _find_table(doc, "Patient", "patient details table")._tbl

    # paragraph properties: BodyText style, centred, body-text spacing
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    style_el = OxmlElement("w:pStyle")
    style_el.set(qn("w:val"), "BodyText")
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "110")
    spacing.set(qn("w:line"), "312")
    spacing.set(qn("w:lineRule"), "auto")
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "center")
    # paragraph-level rPr (bold+italic mark for the whole paragraph)
    p_rpr = OxmlElement("w:rPr")
    p_rpr.append(OxmlElement("w:b"))
    p_rpr.append(OxmlElement("w:bCs"))
    p_rpr.append(OxmlElement("w:i"))
    p_rpr.append(OxmlElement("w:iCs"))
    for child in (style_el, spacing, jc, p_rpr):
        ppr.append(child)
    p.append(ppr)

    # run with matching font / size / colour
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        fonts.set(qn(attr), "Arial")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "242629")
    w_el = OxmlElement("w:w")
    w_el.set(qn("w:val"), "105")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "18")
    szcs = OxmlElement("w:szCs")
    szcs.set(qn("w:val"), "18")
    i_el = OxmlElement("w:i")
    ics_el = OxmlElement("w:iCs")
    b_el = OxmlElement("w:b")
    bcs_el = OxmlElement("w:bCs")
    for child in (fonts, color, w_el, sz, szcs, b_el, bcs_el, i_el, ics_el):
        rpr.append(child)
    r.append(rpr)
    t = OxmlElement("w:t")
    t.text = f"({_PCPNDT_TEXT})"
    t.set(qn("xml:space"), "preserve")
    r.append(t)
    p.append(r)
    tbl_el.addnext(p)


def _insert_pcpndt_disclaimer(doc, data: ReportData):
    """Append the PCPNDT statement as a new bullet at the end of the
    Disclaimer section for specimens that require it."""
    if data.patient.specimen.strip() not in _PCPNDT_SPECIMENS:
        return
    others = {h.lower() for h in SECTION_HEADINGS} - {"disclaimer"}
    in_disclaimer = False
    last_bullet = None
    for p in doc.element.body.iterchildren(qn("w:p")):
        text = _text(p).strip().rstrip(":").lower()
        if in_disclaimer and text in others:
            break
        if text == "disclaimer":
            in_disclaimer = True
            continue
        if in_disclaimer and _text(p).strip():
            last_bullet = p
    if last_bullet is None:
        return
    new_p = copy.deepcopy(last_bullet)
    _set_text(new_p, _PCPNDT_TEXT)
    last_bullet.addnext(new_p)


_SECONDARY_FINDING_SENTENCE = (
    "Incidental or secondary findings that meet the ACMG guidelines "
    "can be given upon request [3]."
)


def _para_containing(doc, fragment: str):
    """Return the first paragraph element whose full text contains `fragment`
    (case-insensitive), or None."""
    fragment_lower = fragment.lower()
    for p in doc.element.body.iter(qn("w:p")):
        if fragment_lower in _text(p).lower():
            return p
    return None


def _adjust_secondary_sentence(doc, data: ReportData):
    """For pathogenic templates (variants Excel upload):
      - Secondary variants selected → sentence stays in Methodology.
      - No secondary variants → sentence moves to Disclaimer.
    For all other templates the sentence is left exactly as in the template.
    Also removes the duplicate 'Gene list used in the assay.' that precedes it."""
    cfg = _cfg(data)

    # Find the paragraph that actually contains the secondary-findings sentence.
    p = _para_containing(doc, "Incidental or secondary findings")
    if p is None:
        return

    # Always remove the duplicate "Gene list used in the assay." before the sentence.
    # Try common variants of the separator (space, no-space after period).
    for old, new in (
        ("Gene list used in the assay. Incidental", "Incidental"),
        ("Gene list used in the assay.Incidental",  "Incidental"),
        ("Gene list used in the assay. \nIncidental", "Incidental"),
    ):
        if _replace_in_para(p, old, new):
            break

    if not cfg.variants:
        # Non-pathogenic template: leave the sentence in methodology.
        return

    has_secondary = any(
        _is_secondary(v) and v.get("selected")
        for v in (data.variants or [])
    )

    if not has_secondary:
        removed = _replace_in_para(p, _SECONDARY_FINDING_SENTENCE + " ", "")
        if not removed:
            _replace_in_para(p, _SECONDARY_FINDING_SENTENCE, "")
        _insert_sentence_to_disclaimer(doc, _SECONDARY_FINDING_SENTENCE)


def _insert_sentence_to_disclaimer(doc, sentence: str):
    """Append sentence as the last item in the Disclaimer section,
    unless it is already present there."""
    others = {h.lower() for h in SECTION_HEADINGS} - {"disclaimer"}
    in_disclaimer = False
    proto = None        # last non-empty bullet to clone formatting from
    last_p = None       # absolute last paragraph in the disclaimer (may be empty)
    for p in doc.element.body.iterchildren(qn("w:p")):
        text = _text(p).strip().rstrip(":").lower()
        if in_disclaimer and text in others:
            break
        if text == "disclaimer":
            in_disclaimer = True
            continue
        if in_disclaimer:
            last_p = p
            if _text(p).strip():
                if sentence.lower() in _text(p).strip().lower():
                    return      # already in the disclaimer — don't duplicate
                proto = p
    if proto is None:
        return
    new_p = copy.deepcopy(proto)
    _set_text(new_p, sentence)
    # Always insert after the last NON-EMPTY bullet (proto), so any trailing
    # empty spacer paragraphs stay at the very end and don't create a visible gap.
    proto.addnext(new_p)


def _fill_methodology(doc, data: ReportData):
    specimen = data.patient.specimen.strip()
    if specimen in METHODOLOGY_SPECIMENS:
        para = _gene_list_link(doc).getparent()
        for kind in ("whole", "clinical"):
            if _replace_in_para(
                    para, f"DNA was used to perform {kind} exome using {kind} exome capture kit .The",
                    f"DNA extracted from the {specimen} was used to perform {kind} exome "
                    f"using {kind} exome capture kit. The"):
                break
    _adjust_secondary_sentence(doc, data)


def _gene_table(doc):
    return _find_table(doc, "Gene Name", "appendix gene coverage table")


def template_gene_table(template_key: str) -> list[tuple[str, str]]:
    """The (gene, coverage) pairs in a template's appendix, row by row."""
    try:
        table = _gene_table(Document(template_path(template_key)))
    except TemplateError:
        return []
    out = []
    for row in table.rows[1:]:
        cells = [c.text.strip() for c in row.cells]
        for i in range(0, len(cells) - 1, 2):
            if cells[i] and cells[i] not in ("-", "–"):
                out.append((cells[i], cells[i + 1]))
    return out


def _fill_gene_table(doc, data: ReportData, blank_unused: bool = True):
    """Rebuild the appendix rows from data.genes, filled row by row like the
    template (gene names keep the template's italic style)."""
    table = _gene_table(doc)
    if not data.genes:
        if not blank_unused:
            return
        # No genes supplied — remove all data rows, keep only the header.
        tbl = table._tbl
        for tr in tbl.findall(qn("w:tr"))[1:]:
            tbl.remove(tr)
        return
    tbl = table._tbl
    rows = tbl.findall(qn("w:tr"))
    proto = copy.deepcopy(rows[1])
    for tr in rows[1:]:
        tbl.remove(tr)
    per_row = len(proto.findall(qn("w:tc"))) // 2
    for start in range(0, len(data.genes), per_row):
        tr = copy.deepcopy(proto)
        tcs = tr.findall(qn("w:tc"))
        chunk = data.genes[start:start + per_row]
        for i in range(per_row):
            # Unused cells in the last row are left empty.
            gene, cov = chunk[i] if i < len(chunk) else ("", "")
            _set_text(tcs[2 * i].find(qn("w:p")), gene)
            _set_text(tcs[2 * i + 1].find(qn("w:p")), cov)
        tbl.append(tr)


def _set_cell_lines(tc, lines: list[str]) -> None:
    """Write one paragraph per line into a table cell. Each keeps the formatting
    of the template paragraph at the same position (the last one is cloned when
    there are more lines); unused template paragraphs are removed."""
    paras = tc.findall(qn("w:p"))
    for i, line in enumerate(lines):
        if i >= len(paras):
            extra = copy.deepcopy(paras[-1])
            paras[-1].addnext(extra)
            paras.append(extra)
        _set_text(paras[i], line)
    for p in paras[max(len(lines), 1):]:
        tc.remove(p)


def _set_pairs(p_el, pairs: list[tuple[str, "str | tuple[str, str]"]], part=None,
               bold: bool = False) -> None:
    """'Label: value' pairs in one paragraph, the first run's style for the
    labels and the second's (bold) for the values, as in the template. A value
    given as (text, url) becomes a hyperlink to `url` (needs the document `part`)."""
    runs = _runs(p_el)
    real = [r for r in runs if _text(r).strip()] or runs     # not the spacing runs
    label_proto = real[0]
    value_proto = real[1] if len(real) > 1 else real[0]
    protos = [copy.deepcopy(label_proto), copy.deepcopy(value_proto)]
    for child in list(p_el):        # the runs, and any links holding runs
        if child.tag in (qn("w:r"), qn("w:hyperlink")):
            p_el.remove(child)
    for i, (label, value) in enumerate(pairs):
        for n, (proto, text) in enumerate(((protos[0], ("      " if i else "") + f"{label}: "),
                                           (protos[1], value))):
            url = ""
            if isinstance(text, tuple):
                text, url = text
            r = copy.deepcopy(proto)
            for child in list(r):
                if child.tag != qn("w:rPr"):
                    r.remove(child)
            t = OxmlElement("w:t")
            t.text = text
            t.set(qn("xml:space"), "preserve")
            r.append(t)
            if bold and n == 1:         # the values bold, the labels as in the template
                Run(r, None).bold = True
            if url and part is not None:
                rpr = r.get_or_add_rPr()
                color = OxmlElement("w:color")          # schema order: ... color, sz, szCs, u
                color.set(qn("w:val"), "0563C1")
                sz = rpr.find(qn("w:sz"))
                (sz.addprevious(color) if sz is not None else rpr.append(color))
                under = OxmlElement("w:u")
                under.set(qn("w:val"), "single")
                rpr.append(under)
                link = OxmlElement("w:hyperlink")
                link.set(qn("r:id"), part.relate_to(url, _HYPERLINK, is_external=True))
                link.append(r)
                p_el.append(link)
            else:
                p_el.append(r)


def _center_line(p_el) -> None:
    """Center a paragraph on one line: no tab stops, no indent."""
    pf = Paragraph(p_el, None).paragraph_format
    for tabs in p_el.findall(qn("w:pPr") + "/" + qn("w:tabs")):
        tabs.getparent().remove(tabs)
    pf.left_indent = pf.right_indent = pf.first_line_indent = 0
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _spacer_row(tbl):
    """An empty, borderless row across the whole table (a gap between blocks)."""
    cols = len(tbl.findall(qn("w:tblGrid") + "/" + qn("w:gridCol"))) or 1
    tr = OxmlElement("w:tr")
    trpr = OxmlElement("w:trPr")
    cant = OxmlElement("w:cantSplit")
    height = OxmlElement("w:trHeight")
    height.set(qn("w:val"), "160")
    height.set(qn("w:hRule"), "exact")
    trpr.extend([cant, height])
    tr.append(trpr)
    tc = OxmlElement("w:tc")
    tcpr = OxmlElement("w:tcPr")
    span = OxmlElement("w:gridSpan")
    span.set(qn("w:val"), str(cols))
    borders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tcpr.extend([span, borders])
    tc.append(tcpr)
    para = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), "20")
    spacing.set(qn("w:lineRule"), "exact")
    ppr.append(spacing)
    para.append(ppr)
    tc.append(para)
    tr.append(tc)
    return tr


def _no_split(tr) -> None:
    """Keep a table row on one page."""
    trpr = tr.get_or_add_trPr()
    if trpr.find(qn("w:cantSplit")) is None:
        trpr.insert(0, OxmlElement("w:cantSplit"))


def _cell_no_wrap(tc) -> None:
    """Prevent text in a table cell from wrapping (w:noWrap)."""
    tcpr = tc.find(qn("w:tcPr"))
    if tcpr is None:
        tcpr = OxmlElement("w:tcPr")
        tc.insert(0, tcpr)
    if tcpr.find(qn("w:noWrap")) is None:
        tcpr.append(OxmlElement("w:noWrap"))


_BANNER = "Pathogenic variant related to the given phenotype was identified"


OMIM_URL = "https://www.omim.org/entry/{}"


def _set_disease_cell(tc, lines: list[str]) -> None:
    """A Disease or Disorder cell: one paragraph per line, a disease in the style of
    the template's first paragraph and an '(OMIM#...)' line in that of its last
    (bold), whatever order and count the lines come in."""
    paras = tc.findall(qn("w:p"))
    text_proto, omim_proto = copy.deepcopy(paras[0]), copy.deepcopy(paras[-1])
    anchor = paras[0]
    for line in lines:
        p = copy.deepcopy(omim_proto if line.startswith("(OMIM#") else text_proto)
        _set_text(p, line)
        anchor.addnext(p)
        anchor = p
    for p in paras:
        tc.remove(p)


def _link_omim(tc, part) -> None:
    """Every '(OMIM#id, id)' line of a Disease cell: each number a hyperlink to its
    OMIM entry (blue, underlined), the rest as in the template."""
    for p in tc.findall(qn("w:p")):
        _link_omim_line(p, part)


def _link_omim_line(p, part) -> None:
    m = re.fullmatch(r"\(OMIM#([\d,\s]+)\)", _text(p).strip())
    runs = _runs(p)
    if not m or not runs:
        return
    ids = [i for i in re.split(r"[,\s]+", m.group(1)) if i]
    proto = copy.deepcopy(runs[0])
    for r in runs:
        r.getparent().remove(r)
    pieces = [("(OMIM#", "")]
    for n, i in enumerate(ids):
        pieces += ([(", ", "")] if n else []) + [(i, OMIM_URL.format(i))]
    pieces.append((")", ""))
    for text, url in pieces:
        r = copy.deepcopy(proto)
        for child in list(r):
            if child.tag != qn("w:rPr"):
                r.remove(child)
        t = OxmlElement("w:t")
        t.text = text
        t.set(qn("xml:space"), "preserve")
        r.append(t)
        if not url:
            p.append(r)
            continue
        rpr = r.get_or_add_rPr()
        for color in rpr.findall(qn("w:color")):
            rpr.remove(color)
        color = OxmlElement("w:color")          # schema order: ... color, sz, szCs, u
        color.set(qn("w:val"), "0563C1")
        sz = rpr.find(qn("w:sz"))
        (sz.addprevious(color) if sz is not None else rpr.append(color))
        under = OxmlElement("w:u")
        under.set(qn("w:val"), "single")
        rpr.append(under)
        link = OxmlElement("w:hyperlink")
        link.set(qn("r:id"), part.relate_to(url, _HYPERLINK, is_external=True))
        link.append(r)
        p.append(link)


def _color_class(tc, v) -> None:
    """The classification text in the cell's first paragraph: blue for an uncertain
    significance, red for Pathogenic / Likely Pathogenic. Always set: a template's
    sample row may be either colour."""
    from . import variants as vr
    vus = vr.classification(v) == "Variant of Uncertain Significance"
    for r in tc.find(qn("w:p")).iter(qn("w:r")):
        rpr = r.get_or_add_rPr()
        color = rpr.find(qn("w:color"))
        if color is None:
            color = OxmlElement("w:color")
            sz = rpr.find(qn("w:sz"))
            (sz.addprevious(color) if sz is not None else rpr.append(color))
        color.set(qn("w:val"), "0070C0" if vus else "FF0000")


def _split_run(r, parts) -> None:
    """Replace run r by one copy of it per (text, italic) part."""
    anchor = r
    for text, italic in parts:
        nr = copy.deepcopy(r)
        ts = nr.findall(qn("w:t"))
        for t in ts[1:]:
            nr.remove(t)
        ts[0].text = text
        ts[0].set(qn("xml:space"), "preserve")
        if italic:
            Run(nr, None).italic = True
        anchor.addnext(nr)
        anchor = nr
    r.getparent().remove(r)


def _set_banner(doc, mode: str, genes: str = "") -> None:
    """The boxed result line under 'Results' (it exists twice in the XML: the
    DrawingML shape and its VML fallback). genes: for a line that names the gene."""
    from . import variants as vr
    if mode == "pathogenic":
        return
    if mode == "none":              # as the template's own no-variant result line
        mito = next((p for p in doc.element.body.iter(qn("w:p"))
                     if p.getparent().tag == qn("w:txbxContent")
                     and "".join(t.text or "" for t in p.iter(qn("w:t"))).startswith(
                         "No pathogenic or likely pathogenic")), None)
        if mito is None:
            return
        text = "".join(t.text or "" for t in mito.iter(qn("w:t")))
        found = mito.find(qn("w:r")).get_or_add_rPr().find(qn("w:color"))
        color = found.get(qn("w:val")) if found is not None else "0070C0"
    else:
        text, color = dict(vr.RESULT_TAGS)[mode], vr.RESULT_COLORS[mode]
    for t in [t for t in doc.element.body.iter(qn("w:t")) if (t.text or "").strip() == _BANNER]:
        t.text = text
        r = t.getparent()
        r.get_or_add_rPr().find(qn("w:color")).set(qn("w:val"), color)
        if "{gene}" in text:            # the gene name in italics
            head, tail = text.split("{gene}")
            _split_run(r, [(head, False), (genes, True), (tail, False)])


def _variant_sets(doc) -> list[tuple]:
    """The template's (results table, variant interpretation table) pairs, in order.
    A results table has 2 rows (header + sample); an interpretation table has 8 per variant."""
    res, interp = [], []
    for t in doc.tables:
        if t.rows and t.rows[0].cells[0].text.strip() == "Gene":
            (interp if len(t.rows) >= 8 else res).append(t._tbl)
    return list(zip(res, interp))


def _is_secondary(v: dict) -> bool:
    from . import variants as vr
    return vr.is_secondary(v)


def _fill_variants(doc, data: ReportData):
    """Replace the template's sample variants with the variants the user ticked.
    A template with two sets of variant tables (primary / Secondary findings) gets
    the Excel's Secondary Findings variants in the second set; the second set is
    removed when none of them is ticked."""
    chosen = [v for v in data.variants if v.get("selected")]
    if not data.variants or not _cfg(data).variants:
        return                      # no Excel uploaded: the template's sample stays
    if _cfg(data).variant_captions:
        _fill_captioned_variants(doc, data, chosen)
        return
    sets = _variant_sets(doc)
    if not sets:
        raise TemplateError("Template is missing: variant results table")
    if len(sets) == 2:
        groups = [[v for v in chosen if not _is_secondary(v)],
                  [v for v in chosen if _is_secondary(v)]]
    else:
        groups = [chosen]
    for n, ((res, tbl), group) in enumerate(zip(sets, groups)):
        if n and not group:
            _drop_secondary_findings(doc, res, tbl)
        elif len(sets) == 2 and not group:      # only Secondary findings: no primary variant
            notes = []                          # the notes under the interpretation table
            el = tbl.getnext()
            while el is not None and el.tag != qn("w:tbl"):
                if el.tag == qn("w:p") and _text(el).strip().lower().startswith(
                        ("*genetic test results", "no other variant")):
                    notes.append(el)
                el = el.getnext()
            for p in notes:
                p.getparent().remove(p)
            _drop_interpretation(doc, tbl, whole_doc=False)
            res.getparent().remove(res)
            heading = _find_para(doc, "List of significant variant(s)", "variant list heading")
            heading.getparent().remove(heading)
            _set_banner(doc, "none")
            # Remove empty paragraphs left between the banner box and the
            # secondary-findings heading — they were spacers around the primary table.
            body = doc.element.body
            _VML_SHAPE = "{urn:schemas-microsoft-com:vml}shape"
            banner_p = next(
                (p for p in body.iterchildren(qn("w:p"))
                 if p.find(".//" + qn("w:drawing")) is not None or
                    p.find(".//" + _VML_SHAPE) is not None),
                None)
            if banner_p is not None:
                nxt = banner_p.getnext()
                while nxt is not None and nxt.tag == qn("w:p") and not _text(nxt).strip():
                    gone, nxt = nxt, nxt.getnext()
                    gone.getparent().remove(gone)
        else:
            _fill_variant_set(doc, data, res, tbl, group, banner=not n,
                              whole_doc=len(sets) == 1)
            if n and data.secondary_title.strip():
                _set_text(_find_para(doc, "List of Secondary", "Secondary findings heading"),
                          data.secondary_title.strip())
            _keep_list_together(res)


def _keep_list_together(res) -> None:
    """Page breaks around a variant list: its 'List of ...' heading (and the blank
    lines under it - one without keep-with-next breaks the chain) stays with the
    header row and the first variant, so a heading is never left alone at the foot
    of a page. A long table may go on over the page, its header row repeated there;
    the last variant stays with the notes under the table, which never split."""
    heading = res.getprevious()
    while heading is not None and not (heading.tag == qn("w:p") and _text(heading).strip()):
        heading = heading.getprevious()
    keep = []
    el = heading if heading is not None and _text(heading).startswith("List of") else res
    while el is not None and el is not res:
        if el.tag == qn("w:p"):
            keep.append(el)
        el = el.getnext()
    rows = res.findall(qn("w:tr"))
    keep += [p for tr in (rows[:1] + rows[-1:]) for p in tr.iter(qn("w:p"))]
    for p in keep:
        Paragraph(p, None).paragraph_format.keep_with_next = True
    tr_pr = rows[0].get_or_add_trPr()       # the header row again on the next page
    if tr_pr.find(qn("w:tblHeader")) is None:
        tr_pr.append(OxmlElement("w:tblHeader"))
    note = res.getnext()
    while note is not None and note.tag == qn("w:p") and not _text(note).strip():
        note = note.getnext()
    if note is not None and note.tag == qn("w:p"):
        Paragraph(note, None).paragraph_format.keep_together = True


def _remove_between(first, stop_prefixes: tuple[str, ...]) -> None:
    """Remove `first` and the elements after it up to the next paragraph starting
    with one of `stop_prefixes` (that one stays)."""
    el = first
    while el is not None and (el is first or not (
            el.tag == qn("w:p") and _text(el).strip().lower().startswith(stop_prefixes))):
        nxt = el.getnext()
        el.getparent().remove(el)
        el = nxt


def _fill_captioned_variants(doc, data: ReportData, chosen: list[dict]) -> None:
    """Templates with variant_captions: the variant list (and the Secondary findings
    list) on top, then under 'Variant Interpretation' one block per variant -
    'Variant N', its table, the abbreviation line, its evidence and OMIM paragraphs.
    The template holds one sample row per list and one sample block."""
    from . import variants as vr
    cfg = _cfg(data)
    gene_tables = [t._tbl for t in doc.tables if t.rows and t.rows[0].cells[0].text.strip() == "Gene"]
    lists = [t for t in gene_tables if len(t.findall(qn("w:tr"))) < 8]
    interp = next(t for t in gene_tables if len(t.findall(qn("w:tr"))) >= 8)
    groups = ([[v for v in chosen if not vr.is_secondary(v)], [v for v in chosen if vr.is_secondary(v)]]
              if cfg.secondary_findings and len(lists) > 1 else [chosen])
    head = _find_para(doc, "Variant Interpretation", "'Variant Interpretation' heading")

    # -- the lists: one row per variant; a list with none is removed with its notes
    titles = [data.primary_title, data.secondary_title]
    for res, group, title in zip(lists, groups, titles):
        heading = res.getprevious()             # its 'List of ...' line (past blank lines / marks)
        while heading is not None and not (heading.tag == qn("w:p") and _text(heading).strip()):
            heading = heading.getprevious()
        if not group:               # its heading, table and notes, up to the next list / section
            _remove_between(heading, ("list of", "variant interpretation"))
            continue
        rows = res.findall(qn("w:tr"))
        proto = rows[1]
        for tr in rows[1:]:
            res.remove(tr)
        _add_result_rows(doc, res, proto, group)
        if title.strip() and heading is not None and _text(heading).startswith("List of"):
            _set_text(heading, title.strip())
        _keep_list_together(res)

    # -- the sample block: caption, table, abbreviation, evidence / OMIM paragraphs
    caption = interp.getprevious()
    while caption is not None and not (caption.tag == qn("w:p") and
                                       _text(caption).strip().lower().startswith("variant ")):
        caption = caption.getprevious()
    after = []                      # the paragraphs under the table that belong to the block
    el = interp.getnext()
    while el is not None and el.tag == qn("w:p") and not _text(el).strip().lower().startswith(
            ("*genetic test results", "no other variant", "cnv findings")):
        after.append(el)
        el = el.getnext()
    evidence = next((p for p in after if _text(p).strip().lstrip("#").lower().startswith("clinical")), None)
    omim = next((p for p in after if _text(p).strip().lower().startswith("omim phenotype")), None)
    ordered = [v for group in groups for v in group]
    if not ordered:                 # nothing ticked: no Additional variants part at all;
        start = _find_para(doc, "Additional Variant", "'Additional Variant(s)' heading")
        cnv = _find_para(doc, "CNV Findings", "'CNV Findings' heading")
        el = start.getprevious()        # notes already under the result line?
        while el is not None and not (el.tag == qn("w:p") and _text(el).strip().lower().startswith(
                ("*genetic test results", "results"))):
            el = el.getprevious()
        noted = el is not None and _text(el).strip().lower().startswith("*genetic test results")
        last_note = None if noted else cnv.getprevious()    # else keep the last notes
        while last_note is not None and not _text(last_note).strip().lower().startswith(
                "*genetic test results"):
            last_note = last_note.getprevious()
        el = start
        while el is not None and el is not (last_note if last_note is not None else cnv):
            nxt = el.getnext()
            el.getparent().remove(el)
            el = nxt
        return
    block_rows = [copy.deepcopy(r) for r in interp.findall(qn("w:tr"))[:8]]
    anchor = head
    for n, v in enumerate(ordered, 1):
        cap = copy.deepcopy(caption)
        _set_text(cap, f"Variant {n}")
        Paragraph(cap, None).paragraph_format.keep_with_next = True
        tbl = copy.deepcopy(interp)
        for tr in tbl.findall(qn("w:tr")):
            tbl.remove(tr)
        for tr in _interpretation_rows(doc, block_rows, v):
            tbl.append(tr)
        parts = [cap, tbl]
        for p in after:             # the template's lines under the table, blank ones included
            key = "evidence" if p is evidence else "omim_phenotype" if p is omim else ""
            if not key:
                parts.append(copy.deepcopy(p))
                continue
            text = vr.clean_narrative(v.get(key, ""))
            if text:                # an empty evidence / OMIM paragraph is left out
                para = copy.deepcopy(p)
                _set_para_tail(para, text)
                if key == "omim_phenotype":
                    _link_omim_in_para(para, doc.part)
                parts.append(para)
        for part in parts:
            anchor.addnext(part)
            anchor = part
    for el in [caption, interp] + after:
        if el is not None and el.getparent() is not None:
            el.getparent().remove(el)


def _fill_variants_interleaved(doc, tbl, block, chosen: list[dict], whole_doc: bool) -> None:
    """Multiple variants: each gets its own interpretation table with narrative
    immediately below it, rather than all narratives grouped at the end."""
    from . import variants as vr
    omim_para = _para_near(doc, tbl, "OMIM phenotype", "", required=False, whole_doc=whole_doc)
    evidence_el = _para_near(doc, tbl, "#Clinical and Literature Evidences", "", required=False,
                             whole_doc=whole_doc)
    if omim_para is None and evidence_el is None:
        raise TemplateError("Template is missing: OMIM phenotype / evidence paragraph")
    if evidence_el is None:
        evidence_el = copy.deepcopy(omim_para)
        _set_lead(evidence_el, "Clinical and Literature Evidences: ")
    empty_tbl = copy.deepcopy(tbl)   # snapshot of empty table structure for reuse
    anchor = tbl                     # tracks insertion point: each element goes after anchor
    for n, v in enumerate(chosen):
        if n == 0:
            current_tbl = tbl
        else:
            current_tbl = copy.deepcopy(empty_tbl)
            anchor.addnext(current_tbl)
        for tr in _interpretation_rows(doc, block, v):
            current_tbl.append(tr)
        anchor = current_tbl
        last_inserted = None
        protos = [(evidence_el, "evidence")]
        if omim_para is not None:
            protos.append((omim_para, "omim_phenotype"))
        for proto, key in protos:
            text = vr.clean_narrative(v.get(key, ""))
            if text:
                para = copy.deepcopy(proto)
                _set_para_tail(para, text)
                if key == "omim_phenotype":
                    _link_omim_in_para(para, doc.part)
                anchor.addnext(para)
                anchor = para
                last_inserted = para
        if last_inserted is not None:
            Paragraph(last_inserted, None).paragraph_format.space_after = Pt(8)
    for p in (evidence_el, omim_para):
        if p is not None and p.getparent() is not None:
            p.getparent().remove(p)


def _fill_variant_set(doc, data: ReportData, res, tbl, chosen: list[dict], banner: bool,
                      whole_doc: bool = True):
    """One set: a row in the Results table and one block (8 rows) in the Variant
    Interpretation table for each variant, plus their evidence / OMIM paragraphs."""
    from . import variants as vr
    # -- Results table: Gene | Variant | Allele | Region | Disease | Inheritance | Class
    rows = res.findall(qn("w:tr"))
    proto = rows[1]
    for tr in rows[1:]:
        res.remove(tr)
    if not chosen:                  # nothing ticked: no variant section at all
        _drop_interpretation(doc, tbl, whole_doc)
        return
    mode = vr.report_mode(chosen, data.result_tag)
    fixed = _cfg(data).fixed_result
    if banner:                      # the result line and its wording are the primary set's
        if not fixed:               # else the template's own result line stays
            _set_banner(doc, mode, ", ".join(dict.fromkeys(v["gene"] for v in chosen)))
        # Heading: user pick overrides auto; auto only changes for VUS mode.
        heading_text = data.primary_title or (
            "List of uncertain significant variant(s) identified related to the phenotype:"
            if vr.RESULT_KIND.get(mode) == "vus_related" and not fixed else None)
        if heading_text:
            heading = res.getprevious()             # the 'List of ...' line over the table
            if heading is None or heading.tag != qn("w:p") or not _text(heading).startswith("List of"):
                heading = _find_para(doc, "List of significant variant(s)", "variant list heading")
            _set_text(heading, heading_text)
    _add_result_rows(doc, res, proto, chosen)
    # -- Variant Interpretation: the template's 8-row block, once per variant
    rows = tbl.findall(qn("w:tr"))
    block = [copy.deepcopy(r) for r in rows[:8]]
    for tr in rows:
        tbl.remove(tr)
    if len(chosen) <= 1:
        for tr in _interpretation_rows(doc, block, chosen[0] if chosen else {}):
            tbl.append(tr)
        _fill_narrative(doc, tbl, chosen, whole_doc)
    else:
        _fill_variants_interleaved(doc, tbl, block, chosen, whole_doc)


def _add_result_rows(doc, res, proto, chosen: list[dict]) -> None:
    """One row of the Results table per variant, made from the template row `proto`:
    Gene | Variant | Allele | Region | Disease | Inheritance | Classification."""
    from . import variants as vr
    for v in chosen:
        tr = copy.deepcopy(proto)
        cells = [[v["gene"]] + ([f"({v['strand']})"] if v.get("strand") else []),
                 [v["variant"]] + ([f"({v['protein']})"] if v.get("protein") else []),
                 [vr.allele_status(v)], [vr.region(v)],
                 vr.disease_lines(v),
                 vr.inheritance(v), [vr.class_display(v)]]
        crit = " ".join(v.get("criteria", "").split())
        if crit:                    # the criteria go under the classification, bold black
            cells[-1].append(crit if crit.startswith("(") else f"({crit})")
        tcs_row = tr.findall(qn("w:tc"))
        for n, (tc, lines) in enumerate(zip(tcs_row, cells)):
            (_set_disease_cell if n == 4 else _set_cell_lines)(tc, lines)
        _cell_no_wrap(tcs_row[0])   # Gene column — never wrap the gene name
        _color_class(tcs_row[-1], v)
        _link_omim(tcs_row[4], doc.part)
        if crit:
            for r in tr.findall(qn("w:tc"))[-1].findall(qn("w:p"))[-1].iter(qn("w:r")):
                rpr = r.get_or_add_rPr()
                for color in rpr.findall(qn("w:color")):
                    rpr.remove(color)
                Run(r, None).bold = True
        _no_split(tr)
        res.append(tr)


def _conservation_display(value: str) -> str:
    """'Conserved' -> 'C', 'Not Conserved' / blank -> '-', anything else uppercased."""
    low = value.lower().strip()
    if not low or "not" in low or low == "-":
        return "-"
    if low == "conserved":
        return "C"
    return value.upper()


def _interpretation_rows(doc, block, v) -> list:
    """The template's 8-row Variant Interpretation block (a list of w:tr) filled
    with one variant, kept together on one page."""
    from . import variants as vr
    pos1, pos2 = vr.genomic_position(v)
    trs = [copy.deepcopy(r) for r in block]
    tcs = [r.findall(qn("w:tc")) for r in trs]
    for n, (tc, lines) in enumerate(zip(tcs[1], [
            [v["gene"]] + ([f"({v['strand']})"] if v.get("strand") else []),
            [v["variant"]] + ([f"({v['protein']})"] if v.get("protein") else []),
            vr.inheritance(v), vr.disease_lines(v),
            [v.get("criteria", "").strip() or "-"], [vr.class_display(v)]])):
        (_set_disease_cell if n == 3 else _set_cell_lines)(tc, lines)
    _color_class(tcs[1][-1], v)
    _link_omim(tcs[1][3], doc.part)
    for tc, lines in zip(tcs[3], [
            [vr.region(v)], [vr.allele_status(v)], [vr.sequence_ontology(v)],
            [vr.frequency(v)], [vr.allelic_depth(v)]]):
        _set_cell_lines(tc, lines)
    _set_pairs(tcs[5][0].find(qn("w:p")),
               [(k, vr.silico_case(v.get(key, "").strip()) or "-") for k, key in vr.insilico_tools(v)],
               bold=True)
    _set_pairs(tcs[5][1].find(qn("w:p")),
               [(k, _conservation_display(v.get(key, "").strip())) for k, key in vr.CONSERVATION],
               bold=True)
    for tc in tcs[5]:               # one centered line, whatever tabs / spaces the template has
        _center_line(tc.find(qn("w:p")))
    _set_cell_lines(tcs[7][0], [pos1] + ([pos2] if pos2 else []))
    _set_pairs(tcs[7][1].find(qn("w:p")),
               [("ClinVar", vr.evidence_link(v.get("clinvar_link", ""), vr.CLINVAR_URL)),
                ("PubMed", vr.evidence_link(v.get("pubmed_link", ""), vr.PUBMED_URL))],
               part=doc.part)
    for n, tr in enumerate(trs):
        _no_split(tr)
        if n < len(trs) - 1:        # keep the whole block on one page
            for p in tr.iter(qn("w:p")):
                Paragraph(p, None).paragraph_format.keep_with_next = True
    return trs


def _fill_narrative(doc, tbl, chosen: list[dict], whole_doc: bool = True) -> None:
    """The evidence / OMIM phenotype paragraphs under an interpretation table, one
    pair per variant, from the template's own paragraphs."""
    from . import variants as vr
    # -- narrative: "#Clinical and Literature Evidences" / "OMIM phenotype"
    # A template may lack one of the two (e.g. no OMIM phenotype paragraph), never both.
    omim_para = _para_near(doc, tbl, "OMIM phenotype", "", required=False, whole_doc=whole_doc)
    evidence = _para_near(doc, tbl, "#Clinical and Literature Evidences", "", required=False,
                          whole_doc=whole_doc)
    if omim_para is None and evidence is None:
        raise TemplateError("Template is missing: OMIM phenotype / evidence paragraph")
    if evidence is None:            # a template without it: the OMIM paragraph's layout
        evidence = copy.deepcopy(omim_para)
        _set_lead(evidence, "Clinical and Literature Evidences: ")
    protos = [(evidence, "evidence")] + ([(omim_para, "omim_phenotype")]
                                         if omim_para is not None else [])
    start = anchor = omim_para if omim_para is not None else evidence
    for v in chosen:                # both paragraphs are typed by the user
        for proto, key in protos:
            text = vr.clean_narrative(v.get(key, ""))
            if text:
                para = copy.deepcopy(proto)
                _set_para_tail(para, text)
                if key == "omim_phenotype":
                    _link_omim_in_para(para, doc.part)
                anchor.addnext(para)
                anchor = para
    if anchor is not start:         # as the template's last paragraph, the last one leaves a gap
        Paragraph(anchor, None).paragraph_format.space_after = Pt(8)   # before the notes under it
    for p in (evidence, omim_para):
        if p is not None and p.getparent() is not None:
            p.getparent().remove(p)


def _superscript_hash(doc) -> None:
    """Strip the leading '#' from any run whose text starts with '#' + letter
    (e.g. '#Clinical and Literature Evidences' in the table header and narrative)."""
    for r in list(doc.element.body.iter(qn("w:r"))):
        t = r.find(qn("w:t"))
        if t is None or not t.text or len(t.text) < 2 or t.text[0] != "#" \
                or not t.text[1].isalpha():
            continue
        t.text = t.text[1:]
        t.set(qn("xml:space"), "preserve")


def _align_result_notes(doc) -> None:
    """The notes under the variant tables and the headings above them start a few twips
    in from the margin while the abbreviation line and the text start at it: line them
    up. All variant tables are centered, the same width, so the sets line up too."""
    for p in doc.paragraphs:
        if p.text.strip().startswith(("*Genetic test results", "No other variant",
                                      "List of ", "Variant Interpretation")):
            p.paragraph_format.left_indent = 0
            first = next((t for t in p._p.iter(qn("w:t")) if t.text), None)
            if first is not None:           # a stray leading space in the template's text
                first.text = first.text.lstrip()
    for res, interp in _variant_sets(doc):
        for tbl in (res, interp):
            pr = tbl.tblPr
            for tag in ("w:tblInd", "w:jc"):
                for el in pr.findall(qn(tag)):
                    pr.remove(el)
            jc = OxmlElement("w:jc")
            jc.set(qn("w:val"), "center")
            tblw = pr.find(qn("w:tblW"))
            (tblw.addnext(jc) if tblw is not None else pr.append(jc))


def _fit_variant_layout(doc) -> None:
    """Variant tables as wide as the text (not inset from it), and everything around
    them - headings, notes, evidence / OMIM paragraphs - starting at the same left edge."""
    sec = doc.sections[0]
    width = round((sec.page_width - sec.left_margin - sec.right_margin) / 635)     # twips
    for tbl in [t._tbl for t in doc.tables if t.rows and t.rows[0].cells[0].text.strip() == "Gene"]:
        grid = tbl.find(qn("w:tblGrid"))
        cols = grid.findall(qn("w:gridCol"))
        total = sum(int(c.get(qn("w:w"))) for c in cols)
        k = width / total
        for c in cols:
            c.set(qn("w:w"), str(round(int(c.get(qn("w:w"))) * k)))
        for tc in tbl.iter(qn("w:tc")):
            w = tc.find(qn("w:tcPr") + "/" + qn("w:tcW"))
            if w is not None and w.get(qn("w:type")) == "dxa":
                w.set(qn("w:w"), str(round(int(w.get(qn("w:w"))) * k)))
        pr = tbl.tblPr
        for tag in ("w:tblInd", "w:jc", "w:tblW"):
            for el in pr.findall(qn(tag)):
                pr.remove(el)
        tblw = OxmlElement("w:tblW")
        tblw.set(qn("w:w"), str(width))
        tblw.set(qn("w:type"), "dxa")
        jc = OxmlElement("w:jc")
        jc.set(qn("w:val"), "center")
        style = pr.find(qn("w:tblStyle"))
        (style.addnext(tblw) if style is not None else pr.insert(0, tblw))
        tblw.addnext(jc)
    for p in doc.paragraphs:
        text = p.text.strip()
        if text.startswith(("*Genetic test results", "No other variant", "List of ",
                            "Variant Interpretation", "OMIM phenotype", "#Clinical and Literature")):
            p.paragraph_format.left_indent = 0
            first = next((t for t in p._p.iter(qn("w:t")) if t.text), None)
            if first is not None:           # a stray leading space in the template's text
                first.text = first.text.lstrip()
        elif text.startswith("Abbreviation:"):      # centered, it starts inside the margin
            p.paragraph_format.left_indent = 0
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT


def _para_near(doc, tbl, prefix: str, what: str, required: bool = True, whole_doc: bool = True):
    """The paragraph starting with `prefix` that belongs to a variant interpretation
    table: the first one after it (before the next table), else the first in the
    document (a template with one set; `whole_doc` False: only after the table). None when
    not required and not found."""
    el = tbl.getnext()
    while el is not None and el.tag != qn("w:tbl"):
        if el.tag == qn("w:p") and _text(el).strip().lower().startswith(prefix.lower()):
            return el
        el = el.getnext()
    try:
        if not whole_doc:
            raise TemplateError(f"Template is missing: {what}")
        return _find_para(doc, prefix, what)
    except TemplateError:
        if required:
            raise
        return None


def _drop_interpretation(doc, tbl, whole_doc: bool = True) -> None:
    """Remove a Variant Interpretation heading, table and paragraphs."""
    paras = [_para_near(doc, tbl, prefix, prefix, required=False, whole_doc=whole_doc) for prefix in
             ("Abbreviation:", "#Clinical and Literature Evidences", "OMIM phenotype")]
    el = tbl.getprevious()                      # its heading, just before it
    while el is not None and el.tag != qn("w:tbl"):
        if el.tag == qn("w:p") and _text(el).strip().lower().startswith("variant interpretation"):
            paras.append(el)
            break
        el = el.getprevious()
    tbl.getparent().remove(tbl)
    for p in paras:
        if p is not None and p.getparent() is not None:
            p.getparent().remove(p)


def _drop_secondary_findings(doc, res, tbl) -> None:
    """No Secondary findings ticked: remove that whole part, from its 'List of
    Secondary findings' heading to the heading that follows it (CNV Findings)."""
    start = res.getprevious()
    while start is not None and not (
            start.tag == qn("w:p") and _text(start).strip().lower().startswith("list of secondary")):
        start = start.getprevious()
    end = tbl.getnext()
    while end is not None and not (
            end.tag == qn("w:p") and _text(end).strip().lower().startswith("cnv findings")):
        end = end.getnext()
    if start is None or end is None:
        raise TemplateError("Template is missing: Secondary findings section")
    el = start
    while el is not end:
        nxt = el.getnext()
        el.getparent().remove(el)
        el = nxt


_OMIM_REF = re.compile(r"\(OMIM([#*])(\d+(?:,\s*\d+)*)\)")


def _link_omim_in_para(p_el, part) -> None:
    """'(OMIM#610205)' / '(OMIM*600275)' in a paragraph: the numbers become hyperlinks
    to their OMIM entry (blue, underlined), the rest of the text stays as it is."""
    runs = _runs(p_el)
    if len(runs) < 2:
        return
    tail = runs[-1]
    text = _text(tail)
    if not _OMIM_REF.search(text):
        return
    pieces, pos = [], 0
    for m in _OMIM_REF.finditer(text):
        pieces.append((text[pos:m.start()] + f"(OMIM{m.group(1)}", ""))
        ids = re.split(r",\s*", m.group(2))
        for n, i in enumerate(ids):
            pieces += ([(", ", "")] if n else []) + [(i, OMIM_URL.format(i))]
        pieces.append((")", ""))
        pos = m.end()
    pieces.append((text[pos:], ""))
    anchor = tail
    for chunk, url in pieces:
        if not chunk:
            continue
        r = copy.deepcopy(tail)
        for child in list(r):
            if child.tag != qn("w:rPr"):
                r.remove(child)
        t = OxmlElement("w:t")
        t.text = chunk
        t.set(qn("xml:space"), "preserve")
        r.append(t)
        if url:
            rpr = r.get_or_add_rPr()
            color = OxmlElement("w:color")
            color.set(qn("w:val"), "0563C1")
            sz = rpr.find(qn("w:sz"))
            (sz.addprevious(color) if sz is not None else rpr.append(color))
            under = OxmlElement("w:u")
            under.set(qn("w:val"), "single")
            rpr.append(under)
            node = OxmlElement("w:hyperlink")
            node.set(qn("r:id"), part.relate_to(url, _HYPERLINK, is_external=True))
            node.append(r)
        else:
            node = r
        anchor.addnext(node)
        anchor = node
    tail.getparent().remove(tail)


def _is_bold(r) -> bool:
    rpr = r.find(qn("w:rPr"))
    return rpr is not None and rpr.find(qn("w:b")) is not None


def _set_lead(p_el, text: str):
    """Make a paragraph's bold lead-in ('OMIM phenotype: ') one run holding `text`, whatever
    number of runs the template splits it into. Returns the first non-bold run (the
    formatting of the rest of the paragraph), or None if the paragraph has no such run."""
    runs = _runs(p_el)
    k = next((i for i, r in enumerate(runs) if not _is_bold(r)), None)
    lead = runs[:k] if k else runs[:1]
    for r in lead[1:]:
        r.getparent().remove(r)
    ts = list(lead[0].iter(qn("w:t")))
    ts[0].text = text
    ts[0].set(qn("xml:space"), "preserve")
    for t in ts[1:]:
        t.getparent().remove(t)
    return runs[k] if k else None


def _set_para_tail(p_el, text: str) -> None:
    """Keep the paragraph's bold lead-in ('OMIM phenotype: ') and replace the
    rest with `text`, in the formatting of the first non-bold run."""
    runs = _runs(p_el)
    k = next((i for i, r in enumerate(runs) if not _is_bold(r)), None)
    lead_text = "".join(_text(r) for r in (runs[:k] if k else runs[:1]))
    tail = _set_lead(p_el, lead_text)
    if tail is None:
        return
    for child in list(tail):
        if child.tag != qn("w:rPr"):
            tail.remove(child)
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    tail.append(t)
    node = tail
    while node.getparent() is not p_el:          # a run inside a hyperlink: start after it
        node = node.getparent()
    for later in list(node.itersiblings()):      # the template's sample text, links included
        if later.tag in (qn("w:r"), qn("w:hyperlink")):
            p_el.remove(later)


def _fill_qc(doc, data: ReportData, restyle: bool = True):
    table = _find_table(doc, "Total Read", "'Sequence data attributes' table")
    for row, value in ((0, data.total_reads), (1, data.q30)):
        _set_text(table.rows[row].cells[1].paragraphs[0]._p, value.strip())
    if not restyle:                 # the template's own cell layout stays
        return
    # Indent 0.2 cm left and right, 3 pt before and after, in every cell.
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                pf = p.paragraph_format
                pf.left_indent = pf.right_indent = Cm(0.2)
                pf.first_line_indent = 0
                pf.space_before = pf.space_after = Pt(3)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


# ---------------------------------------------------------------------------
# Layout-only fixes
# ---------------------------------------------------------------------------

def _appendix_page_break(doc):
    body = doc.element.body
    try:
        app_head = _find_para(doc, "Appendix 1", "appendix heading")
    except TemplateError:
        return      # template without an appendix
    sig = next((p for p in body if p.tag == qn("w:p")
                and p.find(".//" + qn("a:blip")) is not None), None)
    if sig is None or app_head.getparent() is not body:
        return
    n = sig.getnext()
    while n is not None and n is not app_head:
        nxt = n.getnext()
        if n.tag == qn("w:p") and not _text(n).strip() \
                and n.find(".//" + qn("w:drawing")) is None:
            body.remove(n)
        n = nxt
    ppr = app_head.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        app_head.insert(0, ppr)
    if ppr.find(qn("w:pageBreakBefore")) is not None:
        return
    pb = OxmlElement("w:pageBreakBefore")
    # CT_PPr order: pStyle, keepNext, keepLines, pageBreakBefore, ...
    before = [ppr.find(qn(t)) for t in ("w:keepLines", "w:keepNext", "w:pStyle")]
    before = next((e for e in before if e is not None), None)
    if before is not None:
        before.addnext(pb)
    else:
        ppr.insert(0, pb)


SECTION_HEADINGS = ("Clinical History", "Results", "Variant Interpretation", "CNV Findings",
                    "Recommendations", "Methodology", "Sequence data attributes", "Disclaimer", "References",
                    "Whole Exome Sequencing Analysis", "Whole Mitochondrial Genome Sequencing",
                    "This report has been reviewed and approved by")


def _keep_headings_with_next(doc):
    """A section heading never sits alone at the bottom of a page: it moves to
    the next page together with the start of its section."""
    wanted = {h.lower() for h in SECTION_HEADINGS}
    for p in doc.element.body.iterchildren(qn("w:p")):
        if _text(p).strip().rstrip(":").lower() not in wanted:
            continue
        ppr = p.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            p.insert(0, ppr)
        if ppr.find(qn("w:keepNext")) is not None:
            continue
        kn = OxmlElement("w:keepNext")
        # CT_PPr order: pStyle, keepNext, ...
        style = ppr.find(qn("w:pStyle"))
        if style is not None:
            style.addnext(kn)
        else:
            ppr.insert(0, kn)


def _tidy_classification_table(doc, pad: int = 180):
    """Classification column: one paragraph per label, centred vertically.
    The template has an empty paragraph above 'Variant of Uncertain
    Significance' and splits 'Likely Pathogenic' over two paragraphs, which
    leaves uneven space in the cells."""
    try:
        table = _find_table(doc, "Classification", "classification table")
    except TemplateError:
        return
    for row in table.rows[1:]:
        cell = row.cells[0]
        texts = [p.text.strip() for p in cell.paragraphs if p.text.strip()]
        if not texts:
            continue
        paras = cell.paragraphs
        _set_text(paras[next(i for i, p in enumerate(paras) if p.text.strip())]._p,
                  " ".join(texts))
        keep = next(p for p in paras if p.text.strip())
        for p in paras:
            if p is not keep:
                p._p.getparent().remove(p._p)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    # Breathing room between the text and the cell lines, and every cell
    # centred vertically, so a wrapped label never touches the border.
    tbl_pr = table._tbl.tblPr
    old = tbl_pr.find(qn("w:tblCellMar"))
    if old is not None:
        tbl_pr.remove(old)
    mar = OxmlElement("w:tblCellMar")
    for side, w in (("top", pad), ("left", 50), ("bottom", pad), ("right", 50)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(w))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tbl_pr.append(mar)
    for row in table.rows[1:]:
        # The template's fixed row heights pad some rows more than others;
        # let the text plus the cell margins set every row's height.
        trpr = row._tr.trPr
        height = trpr.find(qn("w:trHeight")) if trpr is not None else None
        if height is not None:
            trpr.remove(height)
        for cell in row.cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    # Interpretation column: heading centred, the descriptions justified.
    for i, row in enumerate(table.rows):
        for p in row.cells[1].paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i == 0 else WD_ALIGN_PARAGRAPH.JUSTIFY


def _title_size(doc, points: int = 20):
    """The report title ('Whole Exome Sequencing ...') in one size."""
    half = str(points * 2)
    for p in doc.element.body.iterchildren(qn("w:p")):
        ppr = p.find(qn("w:pPr"))
        style = ppr.find(qn("w:pStyle")) if ppr is not None else None
        if style is None or style.get(qn("w:val")) != "Title":
            continue
        for r in p.iterchildren(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                r.insert(0, rpr)
            for tag in ("w:sz", "w:szCs"):
                el = rpr.find(qn(tag))
                if el is None:
                    el = OxmlElement(tag)
                    rpr.append(el)      # sz/szCs sit late in CT_RPr order
                el.set(qn("w:val"), half)
        break


def _result_tag_size(doc, points: int = 10):
    """The boxed result line ('No pathogenic or likely pathogenic variant(s)
    causative of ...' / 'Pathogenic variant related to ...') in one size; the templates mix 9, 9.5 and 10 pt."""
    half = str(points * 2)
    for p in doc.element.body.iter(qn("w:p")):
        text = _text(p).strip().lower()
        if not ((text.startswith("no pathogenic") and "causative" in text)
                or text.startswith(("pathogenic variant related to",
                                    "likely pathogenic variant related to",
                                    "variant of uncertain significance related to"))):
            continue
        for r in p.iterchildren(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                continue
            for tag in ("w:sz", "w:szCs"):
                el = rpr.find(qn(tag))
                if el is not None:
                    el.set(qn("w:val"), half)
        _fit_result_box(doc, p)


def _fit_result_box(doc, p):
    """At 10 pt the line no longer fits the template's text box beside the
    paragraph's left indent (it wraps and the second line is cut off). Keep the
    template's box narrower than the text area, so it stays inset from the section
    headings, and drop the paragraph's left indent and space before, centred."""
    box = p
    while box is not None and box.tag != qn("w:txbxContent"):
        box = box.getparent()
    if box is None:
        return                      # plain paragraph, not in a text box
    anchor = box
    while anchor is not None and anchor.tag != qn("wp:anchor"):
        anchor = anchor.getparent()
    if anchor is not None:
        # Wide enough for the 10 pt line, but 0.5 cm in from the text margins
        # on each side, so the box is inset from the section headings.
        sec = doc.sections[0]
        width = str(int(sec.page_width - sec.left_margin - sec.right_margin - Cm(1)))
        anchor.find(qn("wp:extent")).set("cx", width)
        for ext in anchor.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}ext"):
            ext.set("cx", width)
        body_pr = anchor.find(".//{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}bodyPr")
        if body_pr is not None:
            body_pr.set("anchor", "ctr")
    for para in box.iterchildren(qn("w:p")):
        pf = Paragraph(para, None).paragraph_format
        pf.left_indent = pf.first_line_indent = 0
        pf.space_before = 0
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _justify_section(doc, heading: str):
    """Justify every paragraph of a section (up to the next section heading)."""
    wanted = heading.lower()
    others = {h.lower() for h in SECTION_HEADINGS} - {wanted}
    inside = False
    for p in doc.element.body.iterchildren(qn("w:p")):
        text = _text(p).strip().rstrip(":").lower()
        if text in others:
            inside = False
        if inside and text:
            Paragraph(p, None).alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if text == wanted:
            inside = True


# Section layout taken from the reference report (MEL_17139): headings flush
# left, bullets and the disclaimer with the same hanging indents, whichever
# template is used. Values are EMU (python-docx lengths).
_LEFT_HEADINGS = ("clinical history", "results", "variant interpretation", "cnv findings", "recommendations",
                  "methodology", "sequence data attributes", "disclaimer", "references",
                  "this report has been reviewed and approved by")


def _match_reference_layout(doc):
    section = ""
    for el in doc.element.body.iterchildren(qn("w:p")):
        text = _text(el).strip().rstrip(":").lower()
        para = Paragraph(el, None)
        if text in _LEFT_HEADINGS:
            section = text
            para.paragraph_format.left_indent = 0
            para.paragraph_format.first_line_indent = None
            continue
        if text.startswith(("whole exome sequencing", "whole mitochondrial", "appendix")):
            section = ""
            continue
        if not text:
            continue
        pf = para.paragraph_format
        if section == "disclaimer":
            pf.left_indent, pf.first_line_indent, pf.line_spacing = 255270, -226695, 1.3
        elif section == "recommendations" and el.pPr is not None and el.pPr.numPr is not None:
            pf.left_indent, pf.first_line_indent, pf.line_spacing = 228600, -228600, 1.3
        elif section == "results" and text.startswith(("*genetic test", "no other variant")):
            pf.left_indent = 27940
        elif section == "references" and el.pPr is not None and el.pPr.numPr is not None:
            pf.left_indent = 27940


def section_opening(doc, heading: str) -> str:
    """Text of the first non-empty paragraph after a section heading ('' if
    the heading isn't there)."""
    found = False
    for p in doc.element.body.iterchildren(qn("w:p")):
        text = _text(p).strip()
        if found and text:
            return text
        found = found or text.rstrip(":").lower() == heading.lower()
    return ""


def _new_pages(doc, headings):
    """Start each of these section headings on a new page."""
    wanted = {h.lower() for h in headings}
    for p in doc.element.body.iterchildren(qn("w:p")):
        if _text(p).strip().rstrip(":").lower() not in wanted:
            continue
        ppr = p.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            p.insert(0, ppr)
        if ppr.find(qn("w:pageBreakBefore")) is not None:
            continue
        pb = OxmlElement("w:pageBreakBefore")
        # CT_PPr order: pStyle, keepNext, keepLines, pageBreakBefore, ...
        before = [ppr.find(qn(t)) for t in ("w:keepLines", "w:keepNext", "w:pStyle")]
        before = next((e for e in before if e is not None), None)
        if before is not None:
            before.addnext(pb)
        else:
            ppr.insert(0, pb)


def _field_runs(rpr, instr: str):
    out = []
    for kind in ("begin", "instr", "separate", "text", "end"):
        r = OxmlElement("w:r")
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
        if kind == "instr":
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = instr
        elif kind == "text":
            el = OxmlElement("w:t")
            el.text = "1"
        else:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
        r.append(el)
        out.append(r)
    return out


def _fix_page_count(doc):
    """Replace the footer's typed 'of 4' with a live NUMPAGES field."""
    for section in doc.sections:
        for txbx in section.footer._element.iter(qn("w:txbxContent")):
            for p in txbx.iter(qn("w:p")):
                runs = _runs(p)
                for prev, r in zip(runs, runs[1:]):
                    if (prev.find(qn("w:t")).text or "").strip() == "of" and \
                            (r.find(qn("w:t")).text or "").strip().isdigit():
                        for nr in reversed(_field_runs(r.find(qn("w:rPr")),
                                                       " NUMPAGES \\* MERGEFORMAT ")):
                            r.addnext(nr)
                        p.remove(r)
                        break


def _fill_signatures(doc, data: ReportData):
    """Put the chosen clinical reviewer's signature block in place of the
    template's: the first picture after "This report has been reviewed and
    approved by". It keeps the template's width; the height follows the image."""
    reviewer = REVIEWERS.get(data.reviewer or _cfg(data).reviewer)
    if reviewer is None:
        return
    body = list(doc.element.body.iterchildren(qn("w:p")))
    start = next((i for i, p in enumerate(body)
                  if _text(p).strip().startswith("This report has been reviewed and approved by")),
                 None)
    if start is None:
        raise TemplateError("Signature block not found in the template.")
    for p in body[start:]:
        if p is not body[start] and _text(p).strip().startswith("Appendix"):
            break
        blip = next(p.iter(qn("a:blip")), None)
        if blip is None:
            continue
        path = resource_path("app", "templates", "signatures", reviewer.signature_file)
        old, (rid, image) = blip.get(qn("r:embed")), doc.part.get_or_add_image(path)
        blip.set(qn("r:embed"), rid)
        if old != rid:
            doc.part.drop_rel(old)          # the template's picture, now unused
        drawing = blip
        while drawing.tag not in (qn("wp:inline"), qn("wp:anchor")):
            drawing = drawing.getparent()
        extent = drawing.find(qn("wp:extent"))
        cx = int(extent.get("cx"))
        cy = str(round(cx * image.px_height / image.px_width))
        extent.set("cy", cy)
        for ext in drawing.iter(qn("a:ext")):
            if ext.get("cy") is not None:
                ext.set("cy", cy)
        return
    raise TemplateError("Signature picture not found in the template.")


def _space_after_tables(doc):
    """A little room under the Sequence data attributes table and the variant
    tables: the note under them and the 'Abbreviation:' line otherwise touch
    the table's bottom border (often through a nearly empty spacer paragraph)."""
    notes = ("Abbreviation:", "Genetic test results are reported")
    for tbl in doc.element.body.iterchildren(qn("w:tbl")):
        nxt = tbl.getnext()
        gap = None
        while nxt is not None and nxt.tag == qn("w:p") and not _text(nxt).strip():
            gap = gap if gap is not None else nxt
            nxt = nxt.getnext()
        if nxt is None or nxt.tag != qn("w:p") or not _text(nxt).strip().startswith(notes):
            continue
        para = gap if gap is not None else nxt
        pf = Paragraph(para, None).paragraph_format
        if pf.space_before is None or pf.space_before < Pt(8):
            pf.space_before = Pt(8)


def _tighten_signature_gap(doc):
    """Drop the empty paragraphs between 'This report has been reviewed and
    approved by' and the signature picture."""
    heading = next((p for p in doc.element.body.iterchildren(qn("w:p"))
                    if _text(p).strip().startswith("This report has been reviewed and approved by")),
                   None)
    if heading is None:
        return
    nxt = heading.getnext()
    while nxt is not None and nxt.tag == qn("w:p") and not _text(nxt).strip()             and next(nxt.iter(qn("w:drawing"), qn("w:pict"), qn("w:sectPr")), None) is None:
        gone, nxt = nxt, nxt.getnext()
        gone.getparent().remove(gone)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def template_path(template_key: str) -> str:
    return resource_path("app", "templates", get_template(template_key).docx_file)


def _drop_analysis_word(doc):
    """'Whole Exome Sequencing Analysis' -> 'Whole Exome Sequencing' in the
    title and the section heading (run last: the layout steps find the heading
    by its template text)."""
    for t in doc.element.body.iter(qn("w:t")):
        if t.text and "Sequencing Analysis" in t.text:
            t.text = t.text.replace("Sequencing Analysis", "Sequencing").rstrip() or t.text


def _cfg(data: ReportData):
    """The report's template, with an options template's layout resolved."""
    return effective(data.template_key, data.options)


# The patient table's boxes (row by row, left / right) with and without the
# Hospital/Clinic box, as the lab's reports lay them out.
_PATIENT_LAYOUTS = {
    True: [("Patient name", "PIN"), ("Age", "Gender"), ("Hospital/Clinic", "Sample Number"),
           ("Referring Clinician", "Sample Collection Date"), ("Specimen", "Sample Received Date"),
           ("", "Report Date")],
    False: [("Patient name", "PIN"), ("Age", "Gender"), ("Referring Clinician", "Sample Number"),
            ("Specimen", "Sample Collection Date"), ("Report Date", "Sample Received Date")],
}

PARENTS_RECOMMENDATION = ("Sequencing the variant(s) in the parents and the other affected and "
                          "unaffected members of the family is recommended to confirm the significance.")


def _patient_layout(doc, hospital: bool) -> None:
    """Relabel the patient table's boxes for the chosen layout (with or without the
    Hospital/Clinic box), adding or removing its last row; _fill_patient then fills
    them by label."""
    tbl = _find_table(doc, "Patient", "patient details table")._tbl
    layout = _PATIENT_LAYOUTS[hospital]
    rows = tbl.findall(qn("w:tr"))
    while len(rows) < len(layout):
        rows[-1].addnext(copy.deepcopy(rows[-1]))
        rows = tbl.findall(qn("w:tr"))
    for tr in rows[len(layout):]:
        tbl.remove(tr)
    for tr, labels in zip(tbl.findall(qn("w:tr")), layout):
        for tc, label in zip(tr.findall(qn("w:tc")), labels):
            paras = tc.findall(qn("w:p"))
            for p in paras[1:]:
                tc.remove(p)
            _set_text(paras[0], f"{label}: XX" if label else "")


_SECTION_STARTS = ("list of", "additional variant", "cnv findings", "whole mitochondrial")


def _drop_interpretation_section(doc) -> None:
    """No Variant Interpretation page: remove each Variant Interpretation part (its
    heading, tables, evidence and the notes repeated under it) up to the blank lines
    before what follows it: the Secondary findings list or CNV Findings."""
    for head in [p for p in doc.element.body.iterchildren(qn("w:p"))
                 if _text(p).strip().rstrip(":").lower() == "variant interpretation"]:
        stop = head.getnext()
        while stop is not None and not (stop.tag == qn("w:p") and
                                        _text(stop).strip().lower().startswith(_SECTION_STARTS)):
            stop = stop.getnext()
        end = stop.getprevious() if stop is not None else None
        while end is not None and end is not head and end.tag == qn("w:p") and not _text(end).strip():
            end = end.getprevious()
        el = head
        while el is not None and el is not stop:
            nxt = el.getnext()
            el.getparent().remove(el)
            if el is end:
                break
            el = nxt


def _compound_het_note(doc) -> None:
    """The likely compound heterozygous note, bold, under the first variant list (before
    its '*Genetic test results' note), with Cis and Trans in italics."""
    from . import variants as vr
    res = next((t._tbl for t in doc.tables if t.rows and t.rows[0].cells[0].text.strip() == "Gene"),
               None)
    if res is None:
        return
    after = res.getnext()
    while after is not None and after.tag == qn("w:p") and not _text(after).strip():
        after = after.getnext()
    model = after if after is not None and after.tag == qn("w:p") else None
    note = copy.deepcopy(model) if model is not None else OxmlElement("w:p")
    _set_text(note, vr.COMPOUND_HET_NOTE)
    pf = Paragraph(note, None).paragraph_format     # the page's full width, single spaced
    pf.right_indent, pf.line_spacing, pf.space_after = 0, 1.1, Pt(4)
    r = _runs(note)[0]
    Run(r, None).bold = True
    head, rest = vr.COMPOUND_HET_NOTE.split("Cis", 1)
    mid, tail = rest.split("Trans", 1)
    _split_run(r, [(head, False), ("Cis", True), (mid, False), ("Trans", True), (tail, False)])
    (model.addprevious if model is not None else res.addnext)(note)


def _parents_recommendation(doc, wanted: bool) -> None:
    """Add or remove the parents-testing line, the first of the Recommendations."""
    head = _find_para(doc, "Recommendations", "'Recommendations' heading")
    first = head.getnext()
    while first is not None and first.tag == qn("w:p") and not _text(first).strip():
        first = first.getnext()
    if first is None or first.tag != qn("w:p"):
        return
    has = _text(first).strip().startswith("Sequencing the variant(s) in the parents")
    if has and not wanted:
        first.getparent().remove(first)
    elif wanted and not has:
        line = copy.deepcopy(first)
        _set_text(line, PARENTS_RECOMMENDATION)
        first.addprevious(line)


def _pseudogene_recommendation(doc, data: ReportData) -> None:
    """Add a pseudogene note as the first item in Recommendations."""
    from . import variants as vr
    chosen = [v for v in data.variants if v.get("selected")]
    genes = vr.pseudogene_genes(chosen)
    if not genes:
        return
    head = _find_para(doc, "Recommendations", "'Recommendations' heading")
    first = head.getnext()
    while first is not None and first.tag == qn("w:p") and not _text(first).strip():
        first = first.getnext()
    if first is None or first.tag != qn("w:p"):
        return
    for gene in reversed(genes):
        line = copy.deepcopy(first)
        _set_text(line, f"The {gene} gene has pseudogene(s) in the human genome.")
        first.addprevious(line)


def _title_with_analysis(doc) -> None:
    """'Whole Exome Sequencing' -> 'Whole Exome Sequencing Analysis' in the title and
    in a WES & Mito report's 'Whole Exome Sequencing' section heading."""
    paras = [p for p in doc.element.body.iterchildren(qn("w:p")) if _text(p).strip()]
    for p in [paras[0]] + [p for p in paras[1:] if _text(p).strip() == "Whole Exome Sequencing"]:
        for t in p.iter(qn("w:t")):
            if t.text and "Whole Exome Sequencing" in t.text and "Sequencing Analysis" not in t.text:
                t.text = t.text.replace("Whole Exome Sequencing", "Whole Exome Sequencing Analysis", 1)


def render_report(data: ReportData, new_pages: tuple[str, ...] = ()):
    """The filled template. new_pages: section headings to start on a new
    page on top of the template's own (see report_builder's layout check)."""
    cfg = _cfg(data)
    path = resource_path("app", "templates", cfg.docx_file)
    if not os.path.exists(path):
        raise TemplateError(f"Template file not found: {path}")
    doc = Document(path)
    options = option_values(data.template_key, data.options)
    if options:
        _patient_layout(doc, options.get("hospital", "yes") == "yes")
    _fill_patient(doc, data)
    _insert_pcpndt_note(doc, data)
    _fill_history(doc, data)
    keep = _cfg(data).keep_layout
    _fill_qc(doc, data, restyle=not keep)
    _fill_gene_list_url(doc, data)
    _fill_methodology(doc, data)
    _insert_pcpndt_disclaimer(doc, data)
    _fill_gene_table(doc, data, blank_unused=not keep)
    _fill_variants(doc, data)
    if options:
        if options["interpretation"] == "no":
            _drop_interpretation_section(doc)
        _parents_recommendation(doc, options.get("parents", "no") == "yes")
        if options["compound_note"] == "yes" and cfg.variants:
            _compound_het_note(doc)
    if cfg.variants:
        _pseudogene_recommendation(doc, data)
        if cfg.keep_layout:
            _fit_variant_layout(doc)
        else:
            _align_result_notes(doc)
        _superscript_hash(doc)
    _fill_signatures(doc, data)
    _tighten_signature_gap(doc)
    _appendix_page_break(doc)
    if not cfg.keep_layout:
        _tidy_classification_table(doc)
        _result_tag_size(doc)
        _title_size(doc)
        _justify_section(doc, "Disclaimer")
        _match_reference_layout(doc)
    else:
        # Laid out as the finished report, but the same table padding, sizes
        # and alignment as every other template. The pages are packed tight, so
        # the classification rows get less padding than elsewhere.
        _tidy_classification_table(doc, pad=70)
        _result_tag_size(doc)
        if not cfg.own_styles:
            _title_size(doc)
        _justify_section(doc, "Disclaimer")
    _space_after_tables(doc)
    _keep_headings_with_next(doc)
    _new_pages(doc, cfg.new_page_before + tuple(new_pages))
    _fix_page_count(doc)
    _drop_analysis_word(doc)
    if options.get("title") == "analysis":
        _title_with_analysis(doc)
    _normalize_hyperlink_colors(doc, hex_color="0000ED")
    return doc
