# CMC Report Automation

Desktop app (Python + PySide6) that produces the CMC **Whole Exome Sequencing**
report (`.docx` / `.pdf`) with a live preview of the real output.
Templates: **Endocrinology**, **Nephrology** (name only for now; uses the
Endocrinology layout until its own template is added).

## What the user enters

Only three parts of the report change per patient. They are typed in by
the user and are never pre-filled or generated:

1. **Patient details**: the demography table (Patient ID, PIN, Age, Gender,
   Hospital/Clinic, Sample Number, Referring Clinician, Sample Collection
   Date, Specimen, Sample Received Date, Report Date).
2. **Clinical history**.
3. **Sequence data attributes** values: Total Read Generated, Data ≥ Q30.

Everything else (results, CNV findings, recommendations, methodology,
disclaimer, references, signatures, appendix) is **fixed** template content.

## How it works

```
Form ──▶ ReportData ──▶ docx_renderer (fills app/templates/<template>.docx) ──▶ .docx ──▶ LibreOffice / Word ──▶ .pdf
                                                                                              │
                                                                     live preview (pypdfium2) ◀┘
```

* The lab's Word file **is** the template. The renderer only replaces the
  three user-entered parts, so all other content and formatting stay exactly
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
| `app/templates/endocrinology.docx` | the Endocrinology report template |
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

## Building a Windows exe

`build.bat` (or `pyinstaller build.spec --noconfirm --clean` after
`pip install -r requirements.txt -r requirements-build.txt`). Output:
`dist/CMCReportAutomation.exe`. The user's PC needs LibreOffice or
Microsoft Word for the preview and PDF export.
