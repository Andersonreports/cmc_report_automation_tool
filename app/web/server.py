"""Flask web server for CMC Report Automation."""
from __future__ import annotations

import base64
import io
import json
import os
import shutil
import socket
import tempfile
import threading
import traceback
import zipfile
from concurrent.futures import ThreadPoolExecutor

from flask import Flask, jsonify, render_template, request, send_file

from ..config.templates import DEFAULT_TEMPLATE, REVIEWERS, TEMPLATES, get_template
from ..core import bulk, gnomad, ocr, sheet_client
from ..core import variants as variants_core
from ..core.docx_renderer import template_gene_list_url, template_gene_table
from ..core.gene_table import parse_pasted
from ..core.models import PatientInfo, ReportData, today_str
from ..core.report_builder import build_report, docx_to_pdf, render_docx

app = Flask(__name__, template_folder="templates")

PATIENT_CHOICES = {
    "hospital": [
        "Christian Medical College - Molecular Endocrinology",
        "Christian Medical College - Nephrology",
        "Rajiv Gandhi Government General Hospital, MMC - Chennai",
    ],
    "referring_clinician": [
        "Dr. Aaron Chapla",
        "Dr. Suceena Alexander, MD., DM., FRCP(Lon), FASN, PhD.",
        "Dr. N. Sheik Sulthan Alavudeen MD., DM.,",
    ],
    "specimen": ["DNA", "Peripheral Blood", "Amniotic fluid", "Product Of Conception", "CVS"],
}
DATE_FIELDS = {"collection_date", "received_date", "report_date"}


def _deferred_rmtree(path: str, delay: int = 60) -> None:
    threading.Timer(delay, lambda: shutil.rmtree(path, ignore_errors=True)).start()


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/init")
def api_init():
    templates_list = []
    for key, cfg in TEMPLATES.items():
        templates_list.append({
            "key": key, "name": cfg.name, "pathogenic": cfg.pathogenic,
            "variants": cfg.variants, "secondary_findings": cfg.secondary_findings,
            "hospital": cfg.hospital, "referring_clinician": cfg.referring_clinician,
            "specimen": cfg.specimen, "reviewer": cfg.reviewer,
            "patient_id_label": cfg.patient_id_label,
            "patient_id_prefix": cfg.patient_id_prefix or "",
        })
    reviewers_list = [{"key": k, "name": r.name} for k, r in REVIEWERS.items()]
    result_tags = [{"key": k, "text": t} for k, t in variants_core.RESULT_TAGS]
    return jsonify({
        "templates": templates_list,
        "reviewers": reviewers_list,
        "default_template": DEFAULT_TEMPLATE,
        "patient_choices": PATIENT_CHOICES,
        "date_fields": list(DATE_FIELDS),
        "sheet_configured": sheet_client.is_configured(),
        "today": today_str(),
        "result_tags": result_tags,
        "secondary_titles": list(variants_core.SECONDARY_TITLES),
        "primary_titles": list(variants_core.PRIMARY_TITLES),
        "frequency_text": variants_core.FREQUENCY_TEXT,
        "frequency_all": variants_core.FREQUENCY_ALL,
        "frequency_absent": variants_core.FREQUENCY_ABSENT,
        "silico": [[label, key] for label, key in variants_core.SILICO],
        "conservation": [[label, key] for label, key in variants_core.CONSERVATION],
        # Only a browser on this machine can use the server's folder picker
        # and folders; other computers get the report as a download.
        "local": _is_local_request(),
    })


@app.route("/api/template-info/<key>")
def api_template_info(key):
    cfg = get_template(key)
    genes = template_gene_table(key)
    try:
        url = template_gene_list_url(key)
    except Exception:
        url = ""
    return jsonify({
        "key": cfg.key, "name": cfg.name, "pathogenic": cfg.pathogenic,
        "variants": cfg.variants, "secondary_findings": cfg.secondary_findings,
        "fixed_result": cfg.fixed_result, "hospital_row": cfg.hospital_row,
        "hospital": cfg.hospital, "referring_clinician": cfg.referring_clinician,
        "specimen": cfg.specimen, "reviewer": cfg.reviewer,
        "patient_id_label": cfg.patient_id_label,
        "patient_id_prefix": cfg.patient_id_prefix or "",
        "genes": genes, "gene_list_url": url,
    })


@app.route("/api/preview", methods=["POST"])
def api_preview():
    try:
        data = ReportData.from_dict(request.json or {})
    except Exception as e:
        return jsonify({"error": str(e), "pages": []}), 400
    tmp = tempfile.mkdtemp(prefix="cmc_preview_")
    try:
        docx_path = os.path.join(tmp, "preview.docx")
        pdf_path_result = None
        try:
            _, pdf_path_result = build_report(data, docx_path, tmp)
        except Exception:
            # Try without layout check
            try:
                render_docx(data, docx_path)
                pdf_path_result = docx_to_pdf(docx_path, tmp)
            except Exception as e2:
                return jsonify({"error": str(e2), "pages": []})
        if not pdf_path_result or not os.path.exists(pdf_path_result):
            return jsonify({"error": "PDF was not produced.", "pages": []})
        import pypdfium2 as pdfium
        max_pages = int((request.json or {}).get("max_pages", 0)) or None
        doc = pdfium.PdfDocument(pdf_path_result)
        pages_b64 = []
        try:
            total = len(doc)
            limit = min(total, max_pages) if max_pages else total
            for i in range(limit):
                page = doc[i]
                scale = 780 / page.get_width()
                img = page.render(scale=scale).to_pil().convert("RGB")
                buf = io.BytesIO()
                img.save(buf, format="JPEG", quality=80)
                pages_b64.append(base64.b64encode(buf.getvalue()).decode())
        finally:
            doc.close()
        return jsonify({"pages": pages_b64, "total_pages": total})
    except Exception:
        return jsonify({"error": traceback.format_exc().strip().splitlines()[-1], "pages": []})
    finally:
        _deferred_rmtree(tmp, 30)


def _save_to_folder(folder: str, *paths: str) -> str:
    """Copy files to the output folder pasted in the page — a path as this
    server sees it (a local folder or a network share). Returns the folder
    when every file was copied, else ""."""
    if not folder or not os.path.isdir(folder):
        return ""
    try:
        for p in paths:
            if p and os.path.exists(p):
                shutil.copy2(p, os.path.join(folder, os.path.basename(p)))
    except OSError:
        return ""
    return folder


def _send(path: str, mimetype: str, name: str, saved: str = ""):
    """send_file, telling the page which folder (if any) also got a copy."""
    resp = send_file(path, mimetype=mimetype, as_attachment=True, download_name=name)
    resp.headers["X-Saved-To"] = saved
    return resp


def _is_local_request() -> bool:
    addr = request.remote_addr or ""
    if addr in ("127.0.0.1", "::1"):
        return True
    try:
        return addr in socket.gethostbyname_ex(socket.gethostname())[2]
    except OSError:
        return False


@app.route("/api/read-clipboard", methods=["POST"])
def api_read_clipboard():
    """Read plain text from the OS clipboard on the server side."""
    # The server's clipboard is only the user's clipboard when the browser
    # runs on this same machine; remote browsers paste with Ctrl+V instead.
    if not _is_local_request():
        return jsonify({"text": "", "remote": True})
    # Try win32clipboard first (reliable on Windows, handles Excel clipboard data)
    try:
        import win32clipboard
        win32clipboard.OpenClipboard()
        try:
            if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
            elif win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_TEXT):
                raw = win32clipboard.GetClipboardData(win32clipboard.CF_TEXT)
                text = raw.decode("ansi", errors="replace")
            else:
                text = ""
        finally:
            win32clipboard.CloseClipboard()
        return jsonify({"text": text})
    except ImportError:
        pass
    except Exception as e:
        pass
    # Fallback: tkinter
    try:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        text = root.clipboard_get()
        root.destroy()
        return jsonify({"text": text})
    except Exception as e:
        return jsonify({"error": str(e), "text": ""}), 200


@app.route("/api/pick-folder", methods=["POST"])
def api_pick_folder():
    """Open a native folder-picker dialog on the server and return chosen path."""
    if not _is_local_request():
        return jsonify({"folder": "", "remote": True})
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        folder = filedialog.askdirectory(title="Choose output folder")
        root.destroy()
        return jsonify({"folder": folder or ""})
    except Exception as e:
        return jsonify({"error": str(e), "folder": ""}), 200


@app.route("/api/export", methods=["POST"])
def api_export():
    body = request.json or {}
    try:
        data = ReportData.from_dict(body.get("data", {}))
        kind = body.get("kind", "docx")
        out_folder = body.get("output_folder", "").strip()
    except Exception as e:
        return jsonify({"error": str(e)}), 400
    tmp = tempfile.mkdtemp(prefix="cmc_export_")
    try:
        stem = bulk.file_stem(data)
        docx_path = os.path.join(tmp, stem + ".docx")
        if kind == "docx":
            build_report(data, docx_path)
            saved = _save_to_folder(out_folder, docx_path)
            _deferred_rmtree(tmp)
            return _send(docx_path,
                         "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                         stem + ".docx", saved)
        elif kind == "pdf":
            _, pdf = build_report(data, docx_path, tmp)
            if not pdf:
                raise RuntimeError("PDF conversion not available on this server.")
            saved = _save_to_folder(out_folder, pdf)
            _deferred_rmtree(tmp)
            return _send(pdf, "application/pdf", stem + ".pdf", saved)
        else:  # both
            _, pdf = build_report(data, docx_path, tmp)
            if not pdf:
                raise RuntimeError("PDF conversion not available on this server.")
            saved = _save_to_folder(out_folder, docx_path, pdf)
            zip_path = os.path.join(tmp, stem + ".zip")
            with zipfile.ZipFile(zip_path, "w") as zf:
                zf.write(docx_path, stem + ".docx")
                zf.write(pdf, stem + ".pdf")
            _deferred_rmtree(tmp)
            return _send(zip_path, "application/zip", stem + ".zip", saved)
    except Exception as e:
        shutil.rmtree(tmp, ignore_errors=True)
        return jsonify({"error": str(e)}), 500


@app.route("/api/fetch-patient", methods=["POST"])
def api_fetch_patient():
    body = request.json or {}
    pin = body.get("pin", "").strip().upper()
    template_key = body.get("template_key", DEFAULT_TEMPLATE)
    cfg = get_template(template_key)
    fields, notes = None, []
    try:
        if sheet_client.patient_configured(cfg):
            fields = sheet_client.fetch_patient(pin, cfg)
    except sheet_client.SheetError as e:
        return jsonify({"error": str(e)}), 500
    qc = None
    if sheet_client.qc_configured():
        try:
            qc = sheet_client.fetch_qc(pin)
        except sheet_client.SheetError as e:
            notes.append(f"Sequence data attributes: {e}")
        else:
            if qc is None:
                notes.append(f"{pin} is not in the sequencing QC sheet — "
                             "enter the Sequence data attributes manually.")
    if fields is None and qc is None:
        if notes:
            return jsonify({"missing": True, "note": "\n".join(notes)})
        return jsonify({"missing": True, "note": f"{pin} was not found in the patient sheet."})
    out = dict(fields or {"pin": pin})
    out.update(qc or {})
    if fields is None and sheet_client.patient_configured(cfg):
        notes.insert(0, f"{pin} was not found in the patient sheet — "
                        "enter the patient details manually.")
    return jsonify({"fields": out, "note": "\n".join(notes)})


@app.route("/api/upload-variants", methods=["POST"])
def api_upload_variants():
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "No file uploaded"}), 400
    tmp = tempfile.mkdtemp(prefix="cmc_var_")
    try:
        path = os.path.join(tmp, file.filename or "variants.xlsx")
        file.save(path)
        found = variants_core.read_variants(path)
        return jsonify({"variants": found})
    except variants_core.VariantError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


@app.route("/api/parse-genes", methods=["POST"])
def api_parse_genes():
    text = (request.json or {}).get("text", "")
    genes, problems = parse_pasted(text)
    return jsonify({"genes": [[g, c] for g, c in genes], "problems": problems})


@app.route("/api/gnomad", methods=["POST"])
def api_gnomad():
    variant = request.json or {}
    try:
        text = gnomad.frequency_text(variant)
        return jsonify({"text": text})
    except gnomad.GnomadError as e:
        return jsonify({"error": str(e)}), 400


@app.route("/api/ocr", methods=["POST"])
def api_ocr():
    file = request.files.get("image")
    if not file:
        return jsonify({"error": "No image uploaded"}), 400
    variant_json = request.form.get("variant", "{}")
    try:
        variant = json.loads(variant_json)
    except Exception:
        variant = {}
    tmp = tempfile.mkdtemp(prefix="cmc_ocr_")
    try:
        ext = os.path.splitext(file.filename or ".png")[1] or ".png"
        path = os.path.join(tmp, "image" + ext)
        file.save(path)
        expected = variants_core.expected_nc_number(variant)
        nc = ocr.nc_from_image(path, expected)
        if not nc:
            return jsonify({"nc": "", "error": "No NC accession found in the image."})
        text = variants_core.genomic_text(variant, nc)
        return jsonify({"nc": nc, "text": text})
    except ocr.OcrError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


@app.route("/api/bulk-upload", methods=["POST"])
def api_bulk_upload():
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "No file uploaded"}), 400
    tmp = tempfile.mkdtemp(prefix="cmc_bulk_")
    try:
        path = os.path.join(tmp, file.filename or "bulk.xlsx")
        file.save(path)
        rows = bulk.read_rows(path)
        return jsonify({"rows": [
            {"number": r.number, "pin": r.pin, "history": r.history,
             "reads": r.reads, "q30": r.q30, "url": r.url}
            for r in rows
        ]})
    except bulk.BulkError as e:
        return jsonify({"error": str(e)}), 400
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


@app.route("/api/bulk-prepare", methods=["POST"])
def api_bulk_prepare():
    body = request.json or {}
    rows_data = body.get("rows", [])
    template_key = body.get("template_key", DEFAULT_TEMPLATE)
    reviewer = body.get("reviewer", "")
    report_date = body.get("report_date", today_str())

    rows = [
        bulk.BulkRow(r["number"], r["pin"], r.get("history", ""),
                     r.get("reads", ""), r.get("q30", ""), r.get("url", ""))
        for r in rows_data
    ]

    results = [None] * len(rows)
    with ThreadPoolExecutor(max_workers=8) as pool:
        futures = {
            pool.submit(bulk.prepare, row, template_key, reviewer, report_date): i
            for i, row in enumerate(rows)
        }
        for fut, i in futures.items():
            results[i] = fut.result()

    items = []
    for p in results:
        item = {
            "row": {"number": p.row.number, "pin": p.row.pin},
            "error": p.error,
            "notes": p.notes,
            "data": p.data.to_dict() if p.data else None,
        }
        items.append(item)
    return jsonify({"items": items})


@app.route("/api/bulk-export", methods=["POST"])
def api_bulk_export():
    body = request.json or {}
    items_data = body.get("items", [])
    kind = body.get("kind", "docx")
    out_folder = (body.get("output_folder") or "").strip()

    reports = [ReportData.from_dict(item["data"]) for item in items_data if item.get("data")]
    if not reports:
        return jsonify({"error": "No reports to export"}), 400

    tmp = tempfile.mkdtemp(prefix="cmc_bexp_")
    try:
        prepared = [bulk.Prepared(bulk.BulkRow(0, r.patient.pin), r) for r in reports]
        stems_list = bulk.stems(prepared)
        export_dir = os.path.join(tmp, "exports")
        os.makedirs(export_dir)

        for data, stem in zip(reports, stems_list):
            try:
                bulk.export(data, export_dir, stem, kind)
            except Exception:
                pass

        files = sorted(f for f in os.listdir(export_dir)
                       if os.path.isfile(os.path.join(export_dir, f)))
        if not files:
            return jsonify({"error": "No reports were generated"}), 500
        saved = _save_to_folder(out_folder, *(os.path.join(export_dir, f) for f in files))

        if len(files) == 1:
            fpath = os.path.join(export_dir, files[0])
            ext = os.path.splitext(files[0])[1].lstrip(".").lower()
            mimes = {"docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                     "pdf": "application/pdf"}
            _deferred_rmtree(tmp)
            return _send(fpath, mimes.get(ext, "application/octet-stream"), files[0], saved)

        zip_path = os.path.join(tmp, "reports.zip")
        with zipfile.ZipFile(zip_path, "w") as zf:
            for f in files:
                zf.write(os.path.join(export_dir, f), f)
        _deferred_rmtree(tmp)
        return _send(zip_path, "application/zip", "reports.zip", saved)
    except Exception as e:
        shutil.rmtree(tmp, ignore_errors=True)
        return jsonify({"error": str(e)}), 500


@app.route("/api/variant-defaults", methods=["POST"])
def api_variant_defaults():
    v = request.json or {}
    opts = variants_core.classification_options(v)
    tools = [[label, key] for label, key in variants_core.insilico_tools(v)]
    disease = v.get("disease", "")
    return jsonify({
        "classification_options": opts,
        "frequency_default": variants_core.default_frequency(v),
        "genomic_text": variants_core.genomic_text(v),
        "suggested_disease": variants_core.suggested_disease(v),
        "suggested_inheritance": variants_core.suggested_inheritance(v),
        "suggested_evidence": variants_core.suggested_evidence(v),
        "omim_phenotype_prefix": variants_core.omim_phenotype_prefix(v),
        "insilico_tools": tools,
        "omim_ids_in": variants_core.omim_ids_in(v, disease),
    })


def run(host="0.0.0.0", port=5000, debug=False):
    app.run(host=host, port=port, debug=debug)
