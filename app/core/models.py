"""Data model for the CMC Whole Exome Sequencing report.

The user enters the patient demography table, the clinical history, the
Sequence data attributes values and the gene list link. Everything else is
fixed template content.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass, field, fields
from datetime import date


def today_str() -> str:
    return date.today().strftime("%d/%m/%Y")


def _from_dict(cls, d: dict):
    names = {f.name for f in fields(cls)}
    return cls(**{k: v for k, v in (d or {}).items() if k in names})


@dataclass
class PatientInfo:
    """Demography table - every field is typed in by the user (no defaults)."""
    patient_id: str = ""
    pin: str = ""
    age: str = ""
    gender: str = ""
    hospital: str = ""
    sample_number: str = ""
    referring_clinician: str = ""
    collection_date: str = ""
    specimen: str = ""
    received_date: str = ""
    report_date: str = ""


@dataclass
class ReportData:
    template_key: str = "ENDOCRINOLOGY"   # key into config.templates.TEMPLATES
    patient: PatientInfo = field(default_factory=PatientInfo)
    clinical_history: str = ""
    total_reads: str = ""                 # Total Read Generated, e.g. "12.6 GB"
    q30: str = ""                         # Data >= Q30, e.g. "96.34 %"
    gene_list_url: str = ""               # Methodology "Click here"; blank = template's link
    genes: list = field(default_factory=list)   # [(gene, coverage)]; empty = template's table
    reviewer: str = ""                    # key into config.templates.REVIEWERS; blank = template's

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "ReportData":
        d = dict(d or {})
        data = _from_dict(cls, {k: v for k, v in d.items() if k != "patient"})
        data.patient = _from_dict(PatientInfo, d.get("patient", {}))
        data.genes = [(str(g), str(c)) for g, c in d.get("genes", [])]
        return data
