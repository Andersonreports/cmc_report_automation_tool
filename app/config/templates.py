"""Report template registry.

Each entry points at a bundled Word template (app/templates/<file>). The user
fills in only the demography table, the clinical history and the Sequence
data attributes values; all other template content is left exactly as is.
A new template must therefore keep: the patient table (first cell starting
"Patient ID"), a "Clinical History" heading followed by its paragraph, and the
"Total Read Generated" / "Data ≥ Q30" table.

To add a template:
  1. Drop the .docx into app/templates/.
  2. Add a TemplateConfig to TEMPLATES below.
The UI's template picker lists it automatically.
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
