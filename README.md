# CMC Report Automation

Desktop app (Python + PySide6) that produces the CMC **Whole Exome Sequencing**
report (`.docx` / `.pdf`) with a live preview of the real output.
Templates:
* **Endocrinology** (CMC - Molecular Endocrinology)
* **Nephrology** (uses the Endocrinology layout until its own template is added)
* **Whole Exome Sequencing** (Endocrinology layout; first patient field is
  *Patient name*; no default Hospital/Clinic or Referring Clinician; Specimen is always Peripheral Blood; patient
  details come from the WES sheet, see below)
* **Whole Exome & Mitochondrial** (WES + whole mitochondrial genome; first
  patient field is *Patient name*)

## What the user enters

Only these parts of the report change per patient. They are typed in by
the user and are never pre-filled or generated:

1. **Patient details**: the demography table (Patient ID, PIN, Age, Gender,
   Hospital/Clinic, Sample Number, Referring Clinician, Sample Collection
   Date, Specimen, Sample Received Date, Report Date).
2. **Clinical history**.
3. **Sequence data attributes** values: Total Read Generated, Data ≥ Q30.
4. **Gene list link**: the Methodology "Click here" link. It is pre-filled
   with the template's current link and can be edited per report.
5. **Gene coverage table** (Appendix 1), on the *Gene Coverage* tab: copy the
   whole table from Word, Excel or the coverage report and click
   *Paste table* (or Ctrl+V) to replace it. Accepts the report's layout
   (4 gene/coverage pairs per row) or 2 columns (Gene, Coverage). Starts with
   the template's table; *Restore template table* brings it back.
6. **Clinical reviewer** (next to *Template*): Dr. Robert Patrick Selvam or
   Dr. Sarath R.S. Picks the signature block under "This report has been
   reviewed and approved by" (images in `app/templates/signatures/`). Each
   template has a default (`reviewer` in `app/config/templates.py`).

**Patient sheet lookup:** typing (or pasting) a complete PIN, the Anderson
ID such as `ADK0000001234`, fills the demography fields from the live
patient sheet, read fresh on every lookup. The Fetch button or Enter
re-runs it. Filled: Patient ID, Age, Gender (from the sheet's *Name* column,
e.g. `MS.12345 (40Y/F)-10`), Sample Number, Sample Received Date, Specimen
(DNA / PB), Hospital/Clinic, and Referring Clinician when present.
The sheet's Received Date fills both Sample Received Date and Sample
Collection Date (they are the same). Report Date is not in the sheet; it
defaults to today and can be changed.
Endocrinology template only: when the row's *Client name* contains
"CHRISTIAN MEDICAL COLLEGE - MOLECULAR ENDOCRINOLOGY", Patient ID becomes
`MEL - ` + the 5 digits in *Name* (`MS.12345 (40Y/F)-10` → `MEL - 12345`).
Otherwise the ID is copied as written. Set per template in
`app/config/templates.py` (`patient_id_client`, `patient_id_prefix`).

Two ways to read the sheet, set in `app/config/sheet_config.py`
(**git-ignored** - never commit these links; on a new machine copy
`app/config/sheet_config.example.py` to `sheet_config.py`):

* **`APPS_SCRIPT_URL` (recommended):** deploy `tools/apps_script/Code.gs` from
  the sheet (*Extensions -> Apps Script*; setup steps are at the top of that
  file). The sheet can then be **private**; the script returns only the
  requested patient's 7 demography columns and allows 60 lookups a minute.
* **`SHEET_CSV_URL` (fallback):** the sheet's CSV export; the sheet must be
  shared as "anyone with the link", and the whole list is downloaded.

The **Whole Exome Sequencing** template looks patients up in the WES sheet
instead (header on row 2; a PIN listed twice gives its latest row): through
the Apps Script with `sheet=wes` (fill in `WES_SPREADSHEET_ID` in Code.gs and
redeploy), or `WES_SHEET_CSV_URL`, its CSV export, when that is set.

After changing any link, rebuild the exe (`build.bat`) - the links are
built into it.

Everything else (results, CNV findings, recommendations, methodology,
disclaimer, references, signatures, appendix) is **fixed** template content.

## How it works

```
Form ──▶ ReportData ──▶ docx_renderer (fills app/templates/<template>.docx) ──▶ .docx ──▶ LibreOffice / Word ──▶ .pdf
                                                                                              │
                                                                     live preview (pypdfium2) ◀┘
```

* The lab's Word file **is** the template. The renderer only replaces the
  user-entered parts, so all other content and formatting stay exactly
  as in the original.
* Layout-only safeguards: the footer's page count is a real field (the
  original had a typed "of 4"), and the appendix always starts on a new page,
  so a long clinical history can't break the layout.
* The preview refreshes automatically after each edit and renders at the
  monitor's physical pixel density, so it is sharp with Windows scaling
  above 100%.
* Export as DOCX or PDF; the file name defaults to
  `<PatientID>_whole_exome_report_<report date>`.

## Run

```bat
pip install -r requirements.txt
run.bat            (or:  python -m app.main)
```

PDF export and the preview need **LibreOffice** (found on PATH or in
Program Files) or **Microsoft Word** as a fallback.

## Project layout

| Path | Purpose |
|---|---|
| `app/main.py` | entry point |
| `app/config/templates.py` | template registry (name, template file) |
| `app/templates/*.docx` | the report templates (patient details replaced by placeholders) |
| `app/core/models.py` | `ReportData`, `PatientInfo` |
| `app/core/docx_renderer.py` | fills the template |
| `app/core/report_builder.py` | DOCX save + PDF conversion (LibreOffice / Word) |
| `app/core/os_utils.py` | resource paths, LibreOffice lookup |
| `app/ui/main_window.py` | PySide6 UI with live preview |

## Adding a template

1. Put the Word template in `app/templates/` (e.g. `nephrology.docx`). It
   must keep the patient table (first cell "Patient ID: …"), the
   "Clinical History" heading with its paragraph, and the
   "Total Read Generated" / "Data ≥ Q30" table.
2. Add or update its `TemplateConfig` in `app/config/templates.py`.

It then appears in the **Template** drop-down automatically.

## Windows and Linux builds (GitHub Actions)

`.github/workflows/build.yml` builds both apps on every push to `main` (or
*Actions -> Build Windows + Linux apps -> Run workflow*). Download them from
the run page, under **Artifacts**:

* `CMCReportAutomation-windows` -> `CMCReportAutomation.exe` (double-click).
* `CMCReportAutomation-linux` -> `CMCReportAutomation-linux.tar.gz`. On the
  Linux PC: extract it, then run `./install.sh` inside the folder once. It
  adds the app to the applications menu and desktop, and installs LibreOffice
  Writer, a Qt library and the report fonts (Arial, Trebuchet MS, Carlito for
  Calibri) so the reports look the same as on Windows.

The patient lookup address is read from the repository secret
`APPS_SCRIPT_URL` (*Settings -> Secrets and variables -> Actions*) and built
into both apps, so keep the repository **private** and share the builds only
with staff. Without the secret the apps still build, with the lookup off.

## Building a Windows exe

`build.bat` (or `pyinstaller build.spec --noconfirm --clean` after
`pip install -r requirements.txt -r requirements-build.txt`). Output:
`dist/CMCReportAutomation.exe`. The user's PC needs LibreOffice or
Microsoft Word for the preview and PDF export.
