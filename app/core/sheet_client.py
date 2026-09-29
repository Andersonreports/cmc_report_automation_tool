"""Look up a patient's demography in the live Google Sheet by PIN.

The sheet is read fresh on every lookup (CSV export of one tab), so new or
corrected rows are picked up immediately. Its link lives in
app/config/sheet_config.py, which is git-ignored: the sheet holds patient
details and is readable by anyone with the link, so the link must never be
committed. See app/config/sheet_config.example.py.

Column mapping (matched by header name):
    Anderson ID        -> PIN (lookup key)
    Sample Number      -> Sample Number
    Name               -> Patient ID, Age, Gender   e.g. "MS.12345 (40Y/F)-10"
                          (the template's Patient ID rule may apply, see
                          TemplateConfig.patient_id_prefix)
    Received Date      -> Sample Received Date and Sample Collection Date
                          (the same date), dd-mm-yyyy -> dd/mm/yyyy
    Sample Type        -> Specimen                  DNA / PB (Peripheral Blood)
    Client name        -> Hospital/Clinic           matched to the dropdown options
    Client Doctor Name -> Referring Clinician       only when filled in
"""
from __future__ import annotations

import csv
import io
import re
import urllib.request

try:
    from ..config.sheet_config import SHEET_CSV_URL
except ImportError:          # no local config: lookup is disabled
    SHEET_CSV_URL = ""

HOSPITALS = {
    "ENDOCRINOLOGY": "Christian Medical College - Molecular Endocrinology",
    "NEPHROLOGY": "Christian Medical College - Nephrology",
}
SPECIMENS = {"DNA": "DNA", "PB": "Peripheral Blood", "PERIPHERAL BLOOD": "Peripheral Blood"}
AGE_UNITS = {"Y": "Years", "M": "Months", "D": "Days"}
GENDERS = {"F": "Female", "M": "Male"}


class SheetError(Exception):
    pass


def is_configured() -> bool:
    return bool(SHEET_CSV_URL)


def _norm(h: str) -> str:
    return re.sub(r"\s+", " ", (h or "").strip().lower())


def _parse_name(name: str) -> dict[str, str]:
    """'MS.12345 (40Y/F)-10' -> patient_id '12345', age '40 Years',
    gender 'Female'. Titles (MS./MR./MASTER./BABY.) and the '-NN' after the
    brackets are dropped; the ID is otherwise kept as written."""
    out: dict[str, str] = {}
    m = re.search(r"\((\d+)\s*([YMD])\s*/\s*([MF])\)", name, re.I)
    if m:
        out["age"] = f"{int(m.group(1))} {AGE_UNITS[m.group(2).upper()]}"
        out["gender"] = GENDERS[m.group(3).upper()]
    core = re.sub(r"\)\s*-\s*\d+\s*$", ")", name.strip())   # "(40Y/F)-10" suffix
    core = re.sub(r"\(.*?\)", " ", core)
    core = re.sub(r"^\s*(MS|MR|MRS|MASTER|BABY|MISS)\.\s*", "", core, flags=re.I).strip()
    if core:
        out["patient_id"] = core
    return out


def _date(value: str) -> str:
    m = re.fullmatch(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})", value.strip())
    return f"{int(m.group(1)):02d}/{int(m.group(2)):02d}/{m.group(3)}" if m else value.strip()


def _template_patient_id(cfg, name: str, client: str) -> str:
    """Template rule, e.g. Endocrinology: client 'CHRISTIAN MEDICAL COLLEGE -
    MOLECULAR ENDOCRINOLOGY ...' + name 'MS.12345 (40Y/F)-10' -> 'MEL - 12345'.
    '' when the rule doesn't apply."""
    if cfg is None or not cfg.patient_id_prefix or not cfg.patient_id_client:
        return ""
    if " ".join(cfg.patient_id_client.upper().split()) not in " ".join(client.upper().split()):
        return ""
    digits = re.search(r"(?<!\d)\d{5}(?!\d)", re.sub(r"\(.*?\)", " ", name))
    return f"{cfg.patient_id_prefix} - {digits.group(0)}" if digits else ""


def fetch_patient(pin: str, cfg=None, timeout: int = 20) -> dict[str, str] | None:
    """Return the demography fields found for `pin`, or None if not in the sheet.
    Only fields the sheet actually has a value for are returned. `cfg` is the
    selected TemplateConfig, whose Patient ID rule is applied."""
    if not SHEET_CSV_URL:
        raise SheetError("The patient sheet link is not set up "
                         "(app/config/sheet_config.py is missing).")
    try:
        with urllib.request.urlopen(SHEET_CSV_URL, timeout=timeout) as resp:
            text = resp.read().decode("utf-8-sig", errors="replace")
    except Exception as e:  # noqa: BLE001
        raise SheetError(f"Could not reach the patient sheet: {e}") from e
    if text.lstrip().startswith("<"):
        raise SheetError("The patient sheet is not shared for reading "
                         "(Google returned a sign-in page).")

    rows = list(csv.reader(io.StringIO(text)))
    if not rows:
        return None
    col = {_norm(h): i for i, h in enumerate(rows[0])}

    def get(row, header):
        i = col.get(_norm(header))
        return row[i].strip() if i is not None and i < len(row) else ""

    key = pin.strip().upper()
    for row in rows[1:]:
        if get(row, "Anderson ID").upper() != key:
            continue
        out = {"pin": key}
        out.update(_parse_name(get(row, "Name")))
        pid = _template_patient_id(cfg, get(row, "Name"), get(row, "Client name"))
        if pid:
            out["patient_id"] = pid
        if get(row, "Sample Number"):
            out["sample_number"] = get(row, "Sample Number")
        if get(row, "Received Date"):
            # The lab's collection date is the same as the received date.
            out["received_date"] = out["collection_date"] = _date(get(row, "Received Date"))
        specimen = SPECIMENS.get(get(row, "Sample Type").upper())
        if specimen:
            out["specimen"] = specimen
        client = get(row, "Client name").upper()
        hospital = next((v for k, v in HOSPITALS.items() if k in client), "")
        if hospital:
            out["hospital"] = hospital
        doctor = get(row, "Client Doctor Name")
        if doctor:
            doctor = re.sub(r"^\s*DR\.?\s*", "", doctor, flags=re.I).title()
            out["referring_clinician"] = f"Dr. {doctor}"
        return out
    return None
