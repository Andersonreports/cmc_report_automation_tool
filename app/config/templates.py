"""Report template registry.

Each entry points at a bundled Word template (app/templates/<file>). The user
fills in only the demography table, the clinical history and the Sequence
data attributes values; all other template content is left exactly as is.
A new template must therefore keep: the patient table (first cell starting
"Patient ID"), a "Clinical History" heading followed by its paragraph, and the
"Total Read Generated" / "Data ≥ Q30" table.

To add a template:
  1. Drop the .docx into app/templates/.
  2. Add a TemplateConfig to TEMPLATES below (pathogenic=True for a report with
     a pathogenic result).
The UI's template picker lists it automatically, under "No pathogenic" or
"Pathogenic".
"""
from __future__ import annotations

from dataclasses import dataclass, replace


@dataclass(frozen=True)
class TemplateConfig:
    key: str
    name: str                       # shown in the template picker
    docx_file: str                  # file name inside app/templates/
    # Label of the first patient-table cell: "Patient ID" or "Patient name".
    patient_id_label: str = "Patient ID"
    # Pre-filled when the template is chosen (still editable by the user).
    hospital: str = ""
    referring_clinician: str = ""
    # Specimen for every report of this template (the sheet's Sample Type is
    # then ignored); "" = take it from the sheet.
    specimen: str = ""
    # Which sheet Fetch reads the patient details from: "" = the patient
    # sheet, "WES" = the Whole Exome Sequencing sheet (see sheet_client).
    patient_sheet: str = ""
    # Patient sheet lookup: when the row's "Client name" contains
    # patient_id_client, Patient ID becomes "<prefix> - <5 digits from Name>".
    patient_id_client: str = ""
    patient_id_prefix: str = ""
    file_suffix: str = "whole_exome_report"   # <PatientID>_<suffix>_<date>
    file_date_sep: str = "-"                  # 29-09-2026 / 29_09_2026
    # Section headings that always start on a new page, e.g. ("Disclaimer",).
    new_page_before: tuple[str, ...] = ()
    # Clinical reviewer whose signature block is used unless the user picks
    # another one: a key of REVIEWERS.
    reviewer: str = "ROBERT"
    # The template has the variant tables, which can be filled from an uploaded
    # listed-variants Excel (the UI's Variants tab).
    variants: bool = False
    # A report with a pathogenic result: the picker lists it under "Pathogenic",
    # every other template under "No pathogenic".
    pathogenic: bool = False
    # The .docx is already laid out exactly as the finished report: its classification
    # table, headings, notes and spacing are left as they are (no layout fixes).
    keep_layout: bool = False
    # The template has a second set of tables for the Excel's Secondary Findings; the
    # result tag and its classification then apply to the other variants only.
    secondary_findings: bool = False
    # The template's result line and variant list heading stay as written, whatever
    # the classifications (e.g. "No pathogenic ... causative" over incidental findings);
    # the UI then hides the Result picker.
    fixed_result: bool = False
    # The patient table has a Hospital/Clinic box; without it the UI hides the field.
    hospital_row: bool = True
    # Variant Interpretation as one captioned block per variant ("Variant 1": its own
    # table, abbreviation line, evidence and OMIM paragraphs), after the variant
    # list and the Secondary findings list, instead of one shared table per list.
    variant_captions: bool = False
    # Keep the template's own title size and patient-table padding instead of the
    # app's usual 20 pt title and 0.2 cm cell margins (for a lab's exact layout).
    own_styles: bool = False
    # One template for every result: the user picks the result and the report's
    # options (OPTIONS below) and the layout follows (see effective()).
    options: bool = False


@dataclass(frozen=True)
class Reviewer:
    name: str                       # shown in the Clinical reviewer picker
    signature_file: str             # whole signature block, app/templates/signatures/


# The signature block under "This report has been reviewed and approved by"
# is one picture of all five signatures; only the clinical reviewer differs.
REVIEWERS: dict[str, Reviewer] = {
    "ROBERT": Reviewer("Dr. Robert Patrick Selvam", "robert.png"),
    "SARATH": Reviewer("Dr. Sarath R.S", "sarath.jpeg"),
}


TEMPLATES: dict[str, TemplateConfig] = {
    "ENDOCRINOLOGY": TemplateConfig(
        key="ENDOCRINOLOGY",
        name="Endocrinology",
        docx_file="endocrinology.docx",
        hospital="Christian Medical College - Molecular Endocrinology",
        referring_clinician="Dr. Aaron Chapla",
        patient_id_client="CHRISTIAN MEDICAL COLLEGE - MOLECULAR ENDOCRINOLOGY",
        patient_id_prefix="MEL",
    ),
    # Renders with the Endocrinology layout until its own .docx is supplied.
    "NEPHROLOGY": TemplateConfig(
        key="NEPHROLOGY",
        name="Nephrology",
        docx_file="endocrinology.docx",
        hospital="Christian Medical College - Nephrology",
        referring_clinician="Dr. Suceena Alexander, MD., DM., FRCP(Lon), FASN, PhD.",
    ),
    # Endocrinology layout, but no default Hospital/Clinic or Referring
    # Clinician and no MEL Patient ID rule; patient details from the WES sheet.
    "WES": TemplateConfig(
        key="WES",
        name="Whole Exome Sequencing",
        docx_file="endocrinology.docx",
        patient_id_label="Patient name",
        specimen="Peripheral Blood",
        patient_sheet="WES",
    ),
    # Clinical Exome Sequencing with no pathogenic variant: the Endocrinology layout,
    # "clinical exome" in the title and Methodology, as the lab's G Shivananda report.
    "CLINICAL_EXOME": TemplateConfig(
        key="CLINICAL_EXOME",
        name="Clinical Exome Sequencing",
        docx_file="clinical_exome.docx",
        patient_id_label="Patient name",
        file_suffix="Clinical_Exome_Report",
    ),
    # Endocrinology layout with a pathogenic result: red result line, variant
    # list, a Variant Interpretation page and the parents-testing recommendation.
    # The variant tables hold the sample variant and are edited in the .docx.
    "CMC_PATHO": TemplateConfig(
        key="CMC_PATHO",
        name="CMC - Pathogenic",
        docx_file="whole_exome_patho.docx",
        hospital="Christian Medical College - Molecular Endocrinology",
        referring_clinician="Dr. Aaron Chapla",
        patient_id_client="CHRISTIAN MEDICAL COLLEGE - MOLECULAR ENDOCRINOLOGY",
        patient_id_prefix="MEL",
        specimen="Peripheral Blood",
        new_page_before=("Variant Interpretation",),
        variants=True,
        pathogenic=True,
    ),
    # Whole Exome + Mitochondrial report with a result: the exome part has two sets of
    # variant tables, the primary and the Secondary findings (the Excel's
    # "Secondary Findings" block); the mitochondrial result stays as in the template.
    "WES_MITO_PATHO": TemplateConfig(
        key="WES_MITO_PATHO",
        name="Whole Exome & Mitochondrial - Pathogenic",
        docx_file="whole_exome_mito_patho.docx",
        patient_id_label="Patient name",
        specimen="Peripheral Blood",
        file_suffix="WES_&_Mito_report",
        new_page_before=("Variant Interpretation",),
        keep_layout=True,
        secondary_findings=True,
        reviewer="SARATH",
        patient_sheet="WES",
        variants=True,
        pathogenic=True,
    ),
    # Whole Exome report with no causative variant but incidental finding(s): green
    # "No pathogenic ... causative" result line over "List of Incidental significant
    # variant(s)", a Variant Interpretation page, and no Hospital/Clinic in the table.
    "WES_INCIDENTAL": TemplateConfig(
        key="WES_INCIDENTAL",
        name="Whole Exome Sequencing - Incidental finding",
        docx_file="whole_exome_incidental.docx",
        patient_id_label="Patient name",
        specimen="Peripheral Blood",
        patient_sheet="WES",
        file_suffix="wes_report",
        new_page_before=("Variant Interpretation",),
        reviewer="SARATH",
        variants=True,
        pathogenic=True,
        fixed_result=True,
        hospital_row=False,
    ),
    # Whole Exome + Mitochondrial report with no causative variant but additional
    # variant(s), as the lab's Baby Infant Snoba report: green "No pathogenic ...
    # causative" result line, an "Additional Variant(s)" list, an optional Secondary
    # findings list, and one captioned interpretation block per variant ("Variant 1",
    # "Variant 2", ...) following on the same page.
    "WES_MITO_ADDITIONAL": TemplateConfig(
        key="WES_MITO_ADDITIONAL",
        name="Whole Exome & Mitochondrial - Additional variants",
        docx_file="whole_exome_mito_additional.docx",
        patient_id_label="Patient name",
        specimen="Peripheral Blood",
        file_suffix="WES_&_Mito_report",
        keep_layout=True,
        secondary_findings=True,
        reviewer="SARATH",
        patient_sheet="WES",
        variants=True,
        pathogenic=True,
        fixed_result=True,
        variant_captions=True,
        own_styles=True,
    ),
    "WES_MITO": TemplateConfig(
        key="WES_MITO",
        name="Whole Exome & Mitochondrial",
        docx_file="whole_exome_mito.docx",
        patient_id_label="Patient name",
        file_suffix="Whole_Exome_Whole_mitochondrial_Report",
        file_date_sep="_",
        new_page_before=("Disclaimer",),
        reviewer="SARATH",
        patient_sheet="WES",
    ),
    # Trial: the WES reports (no pathogenic / pathogenic / incidental) as one template;
    # the Result option picks the layout and the other options its details. It has
    # its own .docx files: nothing is shared with the CMC templates.
    "WES_ALL": TemplateConfig(
        key="WES_ALL",
        name="Whole Exome Sequencing - All results (trial)",
        docx_file="wes_no_pathogenic.docx",
        patient_id_label="Patient name",
        specimen="Peripheral Blood",
        patient_sheet="WES",
        options=True,
    ),
}

DEFAULT_TEMPLATE = "ENDOCRINOLOGY"

# The report options of an options template: key -> (label, [(value, text)]);
# the first choice is the default unless the Result's RESULT_DEFAULTS say otherwise.
OPTIONS: dict[str, tuple[str, list[tuple[str, str]]]] = {
    "result": ("Result", [
        ("none", "No pathogenic variant"),
        ("pathogenic", "Pathogenic / Likely pathogenic / VUS"),
        ("incidental", "Incidental finding(s)"),
    ]),
    "interpretation": ("Variant Interpretation page", [("yes", "Yes"), ("no", "No")]),
    "parents": ("Parents-testing recommendation", [("yes", "Yes"), ("no", "No")]),
    "title": ("Title", [("short", "Whole Exome Sequencing"),
                        ("analysis", "Whole Exome Sequencing Analysis")]),
    "hospital": ("Hospital/Clinic row", [("yes", "Yes"), ("no", "No")]),
    "file_suffix": ("File name", [("whole_exome_report", "<name>_whole_exome_report_<date>"),
                                  ("wes_report", "<name>_wes_report_<date>")]),
}

# What each Result brings: its layout (TemplateConfig fields) and the options'
# starting values, as the lab's existing WES / CMC - Pathogenic / Incidental reports.
RESULT_DEFAULTS: dict[str, tuple[dict, dict]] = {
    "none": (dict(docx_file="wes_no_pathogenic.docx"),
             dict(interpretation="no", parents="no", hospital="yes",
                  file_suffix="whole_exome_report", reviewer="ROBERT")),
    "pathogenic": (dict(docx_file="wes_pathogenic.docx", variants=True, pathogenic=True),
                   dict(interpretation="yes", parents="yes", hospital="yes",
                        file_suffix="whole_exome_report", reviewer="ROBERT")),
    "incidental": (dict(docx_file="whole_exome_incidental.docx", variants=True, pathogenic=True,
                        fixed_result=True),
                   dict(interpretation="yes", parents="no", hospital="no",
                        file_suffix="wes_report", reviewer="SARATH")),
}


def get_template(key: str) -> TemplateConfig:
    return TEMPLATES.get(key) or TEMPLATES[DEFAULT_TEMPLATE]


def option_values(key: str, options: dict | None) -> dict:
    """An options template's choices: the user's, else the Result's defaults, else
    each option's first choice. Empty for any other template."""
    if not get_template(key).options:
        return {}
    options = options or {}
    result = options.get("result") if options.get("result") in RESULT_DEFAULTS else "none"
    out = {k: choices[0][0] for k, (_, choices) in OPTIONS.items()}
    out.update(RESULT_DEFAULTS[result][1], result=result)
    out.update({k: v for k, v in options.items()
                if k in OPTIONS and v in dict(OPTIONS[k][1])})
    return out


def effective(key: str, options: dict | None = None) -> TemplateConfig:
    """The template as used for a report: an options template takes its layout from
    the chosen Result and options; every other template is returned as it is."""
    cfg = get_template(key)
    if not cfg.options:
        return cfg
    values = option_values(key, options)
    layout = RESULT_DEFAULTS[values["result"]][0]
    interp = layout.get("variants") and values["interpretation"] == "yes"
    return replace(cfg, **layout, hospital_row=values["hospital"] == "yes",
                   file_suffix=values["file_suffix"],
                   reviewer=RESULT_DEFAULTS[values["result"]][1]["reviewer"],
                   new_page_before=("Variant Interpretation",) if interp else ())
