"""Look up a patient's demography in the live Google Sheet by PIN.

The sheet is read fresh on every lookup, so new or corrected rows are picked
up immediately. Two ways, set in app/config/sheet_config.py (git-ignored -
never commit these links; see sheet_config.example.py):
  * APPS_SCRIPT_URL (preferred): the web app in tools/apps_script/Code.gs.
    The sheet can stay private and only one patient's row is returned.
  * SHEET_CSV_URL (fallback): the sheet's CSV export; the sheet must be
    shared as "anyone with the link".

The Sequence data attributes come from the sequencing QC sheet, matched on
Anderson_ID: returned alongside the patient row by the Apps Script, or read
from QC_CSV_URL (its CSV export) when no Apps Script is set:
    After Data -> Total Read Generated   "9.96 GB"
    Q30        -> Data >= Q30            "-95.85%" -> "95.85 %"

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

Templates with patient_sheet "WES" (Whole Exome Sequencing) read the WES
sheet instead: the Apps Script with sheet=wes, or WES_SHEET_CSV_URL (its CSV
export). Same columns, except the header is on row 2 under a group row, and
the client column is headed only by "Client" in that group row. A PIN listed
more than once there gives its latest (lowest) row.
"""
from __future__ import annotations

import csv
import io
import json
import re
import urllib.parse
import urllib.request

try:
    from ..config import sheet_config as _cfg
except ImportError:          # no local config: lookup is disabled
    _cfg = None
# Preferred: the Apps Script web app (sheet can stay private, returns one row).
APPS_SCRIPT_URL = getattr(_cfg, "APPS_SCRIPT_URL", "").strip()
# Fallback: the sheet's CSV export (needs "anyone with the link" sharing).
SHEET_CSV_URL = getattr(_cfg, "SHEET_CSV_URL", "").strip()
# Sequencing QC tab (After Data / Q30), CSV export.
QC_CSV_URL = getattr(_cfg, "QC_CSV_URL", "").strip()
# Whole Exome Sequencing sheet, CSV export (fallback to the Apps Script).
WES_SHEET_CSV_URL = getattr(_cfg, "WES_SHEET_CSV_URL", "").strip()

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
    return bool(APPS_SCRIPT_URL or SHEET_CSV_URL or QC_CSV_URL or WES_SHEET_CSV_URL)


def qc_configured() -> bool:
    return bool(APPS_SCRIPT_URL or QC_CSV_URL)


def patient_configured(cfg=None) -> bool:
    if cfg is not None and cfg.patient_sheet == "WES":
        return bool(APPS_SCRIPT_URL or WES_SHEET_CSV_URL)
    return bool(APPS_SCRIPT_URL or SHEET_CSV_URL)


def _norm(h: str) -> str:
    return re.sub(r"\s+", " ", (h or "").strip().lower())


def _parse_name(name: str) -> dict[str, str]:
    """'MS.12345 (40Y/F)-10' -> patient_id '12345', age '40 Years',
    gender 'Female'. '(40YF)' and '( 40Y|F )' are read the same way; tags in
    square brackets ('MRS.ABINAYA (30YF) [POC]') are dropped. Titles (MS./MR./MASTER./BABY.) and the '-NN' after the
    brackets are dropped; the ID is otherwise kept as written."""
    out: dict[str, str] = {}
    m = re.search(r"\(\s*(\d+)\s*([YMD])\s*[/|]?\s*([MF])\s*\)", name, re.I)
    if m:
        out["age"] = f"{int(m.group(1))} {AGE_UNITS[m.group(2).upper()]}"
        out["gender"] = GENDERS[m.group(3).upper()]
    core = re.sub(r"\)\s*-\s*\d+\s*$", ")", name.strip())   # "(40Y/F)-10" suffix
    core = re.sub(r"\(.*?\)|\[.*?\]", " ", core)
    core = " ".join(core.split())
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


# The Apps Script reply's QC part, kept for fetch_qc() so one lookup is one
# request: {pin: (qc row or None, qc error)}.
_apps_script_qc: dict[str, tuple[dict | None, str]] = {}


def _apps_script(pin: str, timeout: int, sheet: str = "") -> dict:
    """The Apps Script web app's reply for `pin` (tools/apps_script/Code.gs).
    sheet "wes" looks the patient up in the WES sheet."""
    query = {"pin": pin, **({"sheet": sheet} if sheet else {})}
    url = f"{APPS_SCRIPT_URL}?{urllib.parse.urlencode(query)}"
    try:
        with urllib.request.urlopen(url, timeout=timeout) as resp:
            reply = json.loads(resp.read().decode("utf-8"))
    except json.JSONDecodeError as e:
        raise SheetError("The patient lookup service returned an unexpected reply "
                         "(check its deployment: Who has access = Anyone).") from e
    except Exception as e:  # noqa: BLE001
        raise SheetError(f"Could not reach the patient lookup service: {e}") from e
    if not reply.get("ok"):
        raise SheetError(reply.get("error") or "The patient lookup failed.")
    if sheet and reply.get("sheet") != sheet:   # older script: wrong sheet
        raise SheetError("The patient lookup service can't read the WES sheet "
                         "yet (redeploy Code.gs).")
    if "qc" in reply:           # older script versions don't send it
        _apps_script_qc[pin] = (reply["qc"], reply.get("qc_error", ""))
    return reply


def _row_via_apps_script(pin: str, timeout: int, sheet: str = "") -> dict[str, str] | None:
    reply = _apps_script(pin, timeout, sheet)
    return ({_norm(k): str(v) for k, v in (reply.get("row") or {}).items()}
            if reply.get("found") else None)


def _row_via_csv(pin: str, timeout: int, url: str = "", pin_column: str = "Anderson ID",
                 last: bool = False) -> dict[str, str] | None:
    """One patient's row from a sheet's CSV export (sheet shared by link).
    The header is the first row holding `pin_column`; a blank header cell
    takes the label above it (a group row). `last`: the PIN's lowest row."""
    try:
        with urllib.request.urlopen(url or SHEET_CSV_URL, timeout=timeout) as resp:
            text = resp.read().decode("utf-8-sig", errors="replace")
    except Exception as e:  # noqa: BLE001
        raise SheetError(f"Could not reach the patient sheet: {e}") from e
    if text.lstrip().startswith("<"):
        raise SheetError("The patient sheet is not shared for reading "
                         "(Google returned a sign-in page).")
    rows = list(csv.reader(io.StringIO(text)))
    top = next((i for i, r in enumerate(rows[:10])
                if _norm(pin_column) in map(_norm, r)), None)
    if top is None:
        return None
    above = rows[top - 1] if top else []
    header = [_norm(h) or _norm(above[i] if i < len(above) else "")
              for i, h in enumerate(rows[top])]
    found = None
    for row in rows[top + 1:]:
        cells = {}
        for i, h in enumerate(header):
            if h and not cells.get(h):          # first non-blank column of a name
                cells[h] = row[i].strip() if i < len(row) else ""
        if cells.get(_norm(pin_column), "").upper() == pin:
            found = cells
            if not last:
                break
    return found


def fetch_patient(pin: str, cfg=None, timeout: int = 20) -> dict[str, str] | None:
    """Return the demography fields found for `pin`, or None if not in the sheet.
    Only fields the sheet actually has a value for are returned. `cfg` is the
    selected TemplateConfig, whose Patient ID rule is applied.
    Uses the Apps Script service when configured, else the CSV export."""
    if not patient_configured(cfg):
        raise SheetError("The patient sheet link is not set up "
                         "(app/config/sheet_config.py is missing).")
    key = pin.strip().upper()
    if cfg is not None and cfg.patient_sheet == "WES":
        cells = (_row_via_csv(key, timeout, WES_SHEET_CSV_URL, last=True) if WES_SHEET_CSV_URL
                 else _row_via_apps_script(key, timeout, "wes"))
    else:
        cells = (_row_via_apps_script(key, timeout) if APPS_SCRIPT_URL
                 else _row_via_csv(key, timeout))
    if cells is None:
        return None

    def get(header):
        return (cells.get(_norm(header)) or "").strip()

    out = {"pin": key}
    # The WES sheet heads the client column just "Client".
    client = get("Client name") or get("Client")
    out.update(_parse_name(get("Name")))
    pid = _template_patient_id(cfg, get("Name"), client)
    if pid:
        out["patient_id"] = pid
    if get("Sample Number"):
        out["sample_number"] = get("Sample Number")
    if get("Received Date"):
        # The lab's collection date is the same as the received date.
        out["received_date"] = out["collection_date"] = _date(get("Received Date"))
    specimen = (cfg.specimen if cfg is not None and cfg.specimen
                else SPECIMENS.get(get("Sample Type").upper()))
    if specimen:
        out["specimen"] = specimen
    hospital = next((v for k, v in HOSPITALS.items() if k in client.upper()), "")
    if hospital:
        out["hospital"] = hospital
    doctor = get("Client Doctor Name")
    if doctor:
        doctor = re.sub(r"^\s*DR\.?\s*", "", doctor, flags=re.I).title()
        out["referring_clinician"] = f"Dr. {doctor}"
    return out


def fetch_qc(pin: str, timeout: int = 20) -> dict[str, str] | None:
    """Return {"total_reads", "q30"} from the sequencing QC sheet for `pin`,
    or None if the PIN isn't there (or no QC source is set)."""
    pin = pin.strip().upper()
    if APPS_SCRIPT_URL:
        if pin not in _apps_script_qc:
            _apps_script(pin, timeout)
        qc, error = _apps_script_qc.pop(pin, (None, "The patient lookup service doesn't "
                                              "return QC values yet (redeploy Code.gs)."))
        if error:
            raise SheetError(error)
        if not qc:
            return None
        cells = {_norm(k): str(v) for k, v in qc.items()}
    elif QC_CSV_URL:
        cells = _row_via_csv(pin, timeout, QC_CSV_URL, "Anderson_ID")
        if cells is None:
            return None
    else:
        return None
    out = {}
    reads = " ".join((cells.get(_norm("After Data")) or "").split())
    if reads:
        out["total_reads"] = re.sub(r"(?<=\d)(?=[A-Za-z])", " ", reads)   # "9.96GB" -> "9.96 GB"
    # The sheet writes Q30 as "-95.85%"; the report shows "95.85 %".
    q30 = re.search(r"\d+(\.\d+)?", cells.get(_norm("Q30")) or "")
    if q30:
        out["q30"] = f"{q30.group(0)} %"
    return out or None
