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
    # options (a key of OPTIONS below) and the layout follows (see effective()).
    options: str = ""
    # Also offer the result lines and list titles of the lab's newer WES / WES & Mito
    # reports (variants.LAB_RESULT_TAGS / LAB_PRIMARY_TITLES) in the Variants tab.
    lab_wording: bool = False


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
    # Clinical Exome Sequencing with no pathogenic variant: the Endocrinology layout,
    # "clinical exome" in the title and Methodology, as the lab's G Shivananda report.
    "CLINICAL_EXOME": TemplateConfig(
        key="CLINICAL_EXOME",
        name="Clinical Exome Sequencing",
        docx_file="clinical_exome.docx",
        patient_id_label="Patient name",
        file_suffix="Clinical_Exome_Report",
    ),
    # The WES reports (no pathogenic / pathogenic / VUS) as one template;
    # the Result option picks the layout and the other options its details.
    "WES_ALL": TemplateConfig(
        key="WES_ALL",
        name="Whole Exome Sequencing - All results",
        docx_file="wes_no_pathogenic.docx",
        patient_id_label="Patient name",
        specimen="Peripheral Blood",
        patient_sheet="WES",
        options="wes",
        lab_wording=True,
    ),
    # POC (Product Of Conception) and Amniotic Fluid reports share one WES
    # template; the user picks the specimen from the dropdown.
    "WES_POC_AF": TemplateConfig(
        key="WES_POC_AF",
        name="Whole Exome Sequencing - POC / Amniotic Fluid",
        docx_file="wes_no_pathogenic.docx",
        patient_id_label="Patient name",
        patient_sheet="WES",
        options="wes_poc_af",
        lab_wording=True,
    ),
    # The Whole Exome & Mitochondrial reports (no pathogenic / pathogenic /
    # VUS / additional variants) as one template, the same way.
    "WES_MITO_ALL": TemplateConfig(
        key="WES_MITO_ALL",
        name="Whole Exome & Mitochondrial - All results",
        docx_file="whole_exome_mito.docx",
        patient_id_label="Patient name",
        specimen="Peripheral Blood",
        reviewer="SARATH",
        patient_sheet="WES",
        options="wes_mito",
        lab_wording=True,
    ),
}

DEFAULT_TEMPLATE = "ENDOCRINOLOGY"

_YES_NO = [("yes", "Yes"), ("no", "No")]


def _options(results, title, files) -> dict[str, tuple[str, list[tuple[str, str]]]]:
    opts: dict[str, tuple[str, list[tuple[str, str]]]] = {
        "result": ("Result", results),
        "interpretation": ("Variant Interpretation page", _YES_NO),
        "compound_note": ("Compound heterozygous note", [("no", "No"), ("yes", "Yes")]),
    }
    if title is not None:
        opts["title"] = ("Title", [("short", title), ("analysis", title.replace(
            "Whole Exome Sequencing", "Whole Exome Sequencing Analysis", 1))])
    if files is not None:
        opts["file_suffix"] = ("File name", files)
    return opts


# The report options of each options template (TemplateConfig.options): key ->
# (label, [(value, text)]); the first choice is the default unless the Result's
# RESULT_DEFAULTS say otherwise. A file_suffix value is "<suffix>|<date separator>".
OPTIONS: dict[str, dict[str, tuple[str, list[tuple[str, str]]]]] = {
    "wes": _options(
        [("none", "No pathogenic variant"),
         ("pathogenic", "Pathogenic / Likely pathogenic"),
         ("vus", "VUS")],
        None,
        None),
    "wes_poc_af": _options(
        [("none", "No pathogenic variant"),
         ("pathogenic", "Pathogenic / Likely pathogenic"),
         ("vus", "VUS")],
        None,
        None),
    "wes_mito": _options(
        [("none", "No pathogenic variant"),
         ("pathogenic", "Pathogenic / Likely pathogenic"),
         ("vus", "VUS"),
         ("additional", "Additional variant(s)")],
        None,
        None),
}

# What each Result brings: its layout (TemplateConfig fields) and the options'
# starting values, as the lab's existing reports of that kind.
RESULT_DEFAULTS: dict[str, dict[str, tuple[dict, dict]]] = {
    "wes": {
        "none": (dict(docx_file="wes_pathogenic.docx", variants=True, pathogenic=True),
                 dict(interpretation="yes",
                      file_suffix="whole_exome_report|-", reviewer="ROBERT")),
        "pathogenic": (dict(docx_file="wes_pathogenic.docx", variants=True, pathogenic=True),
                       dict(interpretation="yes",
                            file_suffix="whole_exome_report|-", reviewer="ROBERT")),
        "vus": (dict(docx_file="wes_pathogenic.docx", variants=True, pathogenic=True),
                dict(interpretation="yes",
                     file_suffix="whole_exome_report|-", reviewer="ROBERT")),
    },
    "wes_poc_af": {
        "none": (dict(docx_file="wes_pathogenic.docx", variants=True, pathogenic=True),
                 dict(interpretation="yes",
                      file_suffix="whole_exome_report|_", reviewer="ROBERT")),
        "pathogenic": (dict(docx_file="wes_pathogenic.docx", variants=True, pathogenic=True),
                       dict(interpretation="yes",
                            file_suffix="whole_exome_report|_", reviewer="ROBERT")),
        "vus": (dict(docx_file="wes_pathogenic.docx", variants=True, pathogenic=True),
                dict(interpretation="yes",
                     file_suffix="whole_exome_report|_", reviewer="ROBERT")),
    },
    "wes_mito": {
        "none": (dict(docx_file="whole_exome_mito.docx", new_page_before=("Disclaimer",)),
                 dict(interpretation="no",
                      file_suffix="Whole_Exome_Whole_mitochondrial_Report|_", reviewer="SARATH")),
        "pathogenic": (dict(docx_file="whole_exome_mito_patho.docx", variants=True,
                            pathogenic=True, keep_layout=True, secondary_findings=True),
                       dict(interpretation="yes",
                            file_suffix="WES_Mito_report|_", reviewer="SARATH")),
        "vus": (dict(docx_file="whole_exome_mito_patho.docx", variants=True,
                     pathogenic=True, keep_layout=True, secondary_findings=True),
                dict(interpretation="yes",
                     file_suffix="WES_Mito_report|_", reviewer="SARATH")),
        "additional": (dict(docx_file="whole_exome_mito_additional.docx", variants=True,
                            pathogenic=True, keep_layout=True, secondary_findings=True,
                            fixed_result=True, variant_captions=True, own_styles=True),
                       dict(interpretation="yes",
                            file_suffix="WES_Mito_report|_", reviewer="SARATH")),
    },
}


def get_template(key: str) -> TemplateConfig:
    return TEMPLATES.get(key) or TEMPLATES[DEFAULT_TEMPLATE]


def template_options(key: str) -> dict[str, tuple[str, list[tuple[str, str]]]]:
    """The report options of an options template; empty for any other template."""
    return OPTIONS.get(get_template(key).options, {})


def option_values(key: str, options: dict | None) -> dict:
    """An options template's choices: the user's, else the Result's defaults, else
    each option's first choice. Empty for any other template."""
    opts = template_options(key)
    if not opts:
        return {}
    defaults = RESULT_DEFAULTS[get_template(key).options]
    options = options or {}
    result = options.get("result") if options.get("result") in defaults else "none"
    out = {k: choices[0][0] for k, (_, choices) in opts.items()}
    out.update({k: v for k, v in defaults[result][1].items() if k in opts}, result=result)
    out.update({k: v for k, v in options.items() if k in opts and v in dict(opts[k][1])})
    return out


def effective(key: str, options: dict | None = None) -> TemplateConfig:
    """The template as used for a report: an options template takes its layout from
    the chosen Result and options; every other template is returned as it is."""
    cfg = get_template(key)
    if not cfg.options:
        return cfg
    values = option_values(key, options)
    layout, defaults = RESULT_DEFAULTS[cfg.options][values["result"]]
    layout = dict(layout)
    pages = layout.pop("new_page_before", ())
    if values["interpretation"] == "no":
        pages = tuple(p for p in pages if p != "Variant Interpretation")
    raw_suffix = values.get("file_suffix") or defaults.get("file_suffix", "whole_exome_report|-")
    suffix, sep = raw_suffix.split("|")
    hospital_row = (values["hospital"] == "yes") if "hospital" in values else True
    return replace(cfg, **layout, hospital_row=hospital_row,
                   file_suffix=suffix, file_date_sep=sep, reviewer=defaults["reviewer"],
                   new_page_before=pages)
