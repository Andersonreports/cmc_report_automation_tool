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

from dataclasses import dataclass


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
}

DEFAULT_TEMPLATE = "ENDOCRINOLOGY"


def get_template(key: str) -> TemplateConfig:
    return TEMPLATES.get(key) or TEMPLATES[DEFAULT_TEMPLATE]
