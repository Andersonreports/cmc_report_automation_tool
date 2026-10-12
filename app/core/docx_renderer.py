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
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from ..config.templates import REVIEWERS, get_template
from .models import ReportData
from .os_utils import resource_path

PATIENT_CELLS = [  # (row, col, label, attribute)
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
    table = _find_table(doc, "Patient", "patient details table")
    id_label = get_template(data.template_key).patient_id_label
    for row, col, label, attr in PATIENT_CELLS:
        if attr == "patient_id":
            label = id_label          # "Patient ID" or "Patient name"
        value = (getattr(data.patient, attr, "") or "").strip()
        _set_text(table.rows[row].cells[col].paragraphs[0]._p, f"{label}: {value}")
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


def template_gene_list_url(template_key: str) -> str:
    """The gene list link stored in a template ('' if it has none)."""
    doc = Document(template_path(template_key))
    try:
        return doc.part.rels[_gene_list_link(doc).get(qn("r:id"))].target_ref
    except TemplateError:
        return ""


# Methodology opening per specimen: (template text, replacement). DNA keeps
# the template wording as is.
METHODOLOGY_OPENING = {
    "Peripheral Blood": (
        "DNA was used to perform whole exome using whole exome capture kit .The",
        "DNA extracted from the Peripheral Blood was used to perform whole exome "
        "using whole exome capture kit. The"),
}


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


def _fill_methodology(doc, data: ReportData):
    swap = METHODOLOGY_OPENING.get(data.patient.specimen.strip())
    if swap:
        _replace_in_para(_gene_list_link(doc).getparent(), *swap)


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


def _fill_gene_table(doc, data: ReportData):
    """Rebuild the appendix rows from data.genes, filled row by row like the
    template (gene names keep the template's italic style)."""
    table = _gene_table(doc)
    if not data.genes:
        # Keep the template's table, but leave its unused cells empty.
        for row in table.rows[1:]:
            for cell in row.cells:
                if cell.text.strip() in ("-", "–", "—"):
                    _set_text(cell.paragraphs[0]._p, "")
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


def _fill_qc(doc, data: ReportData):
    table = _find_table(doc, "Total Read", "'Sequence data attributes' table")
    for row, value in ((0, data.total_reads), (1, data.q30)):
        _set_text(table.rows[row].cells[1].paragraphs[0]._p, value.strip())
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


SECTION_HEADINGS = ("Clinical History", "Results", "CNV Findings", "Recommendations",
                    "Methodology", "Sequence data attributes", "Disclaimer", "References",
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


def _tidy_classification_table(doc):
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
    for side, w in (("top", 180), ("left", 50), ("bottom", 180), ("right", 50)):
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
    """The report title ('Whole Exome Sequencing Analysis ...') in one size."""
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
    causative of ...') in one size; the templates mix 9, 9.5 and 10 pt."""
    half = str(points * 2)
    for p in doc.element.body.iter(qn("w:p")):
        text = _text(p).strip().lower()
        if not (text.startswith("no pathogenic") and "causative" in text):
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
_LEFT_HEADINGS = ("clinical history", "results", "cnv findings", "recommendations",
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
    reviewer = REVIEWERS.get(data.reviewer or get_template(data.template_key).reviewer)
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


def render_report(data: ReportData, new_pages: tuple[str, ...] = ()):
    """The filled template. new_pages: section headings to start on a new
    page on top of the template's own (see report_builder's layout check)."""
    path = template_path(data.template_key)
    if not os.path.exists(path):
        raise TemplateError(f"Template file not found: {path}")
    doc = Document(path)
    _fill_patient(doc, data)
    _fill_history(doc, data)
    _fill_qc(doc, data)
    _fill_gene_list_url(doc, data)
    _fill_methodology(doc, data)
    _fill_gene_table(doc, data)
    _fill_signatures(doc, data)
    _tighten_signature_gap(doc)
    _appendix_page_break(doc)
    _tidy_classification_table(doc)
    _result_tag_size(doc)
    _title_size(doc)
    _justify_section(doc, "Disclaimer")
    _match_reference_layout(doc)
    _keep_headings_with_next(doc)
    _new_pages(doc, get_template(data.template_key).new_page_before + tuple(new_pages))
    _fix_page_count(doc)
    return doc
