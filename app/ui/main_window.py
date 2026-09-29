"""PySide6 desktop UI for CMC report automation.

The user types in the patient demography table, the clinical history and the
Sequence data attributes values; everything else in the report is fixed
template content. The right-hand pane previews the real output (DOCX -> PDF
-> page images) and refreshes automatically after each edit.
"""
from __future__ import annotations

import json
import os
import re
import shutil
import sys
import tempfile
import traceback
from datetime import datetime

from PySide6.QtCore import QSettings, Qt, QThread, QTimer, Signal
from PySide6.QtGui import QImage, QPixmap
from PySide6.QtWidgets import (
    QApplication, QComboBox, QFileDialog, QFormLayout, QGroupBox, QHBoxLayout,
    QLabel, QLineEdit, QMainWindow, QMessageBox, QPlainTextEdit, QPushButton,
    QScrollArea, QSplitter, QVBoxLayout, QWidget,
)

from ..config.templates import DEFAULT_TEMPLATE, TEMPLATES, get_template
from ..core.models import PatientInfo, ReportData

APP_NAME = "CMC Report Automation"

# (attribute, label, example shown greyed in the empty box - never printed)
PATIENT_FIELDS = [
    ("patient_id", "Patient ID", "e.g. MEL - 00000"),
    ("pin", "PIN", "e.g. ADK0000000000"),
    ("age", "Age", "52 Years"),
    ("gender", "Gender", "Female"),
    ("hospital", "Hospital/Clinic", "Hospital / clinic name"),
    ("sample_number", "Sample Number", "Sample number"),
    ("referring_clinician", "Referring Clinician", "Clinician name"),
    ("collection_date", "Sample Collection Date", "29/08/2026"),
    ("specimen", "Specimen", "DNA"),
    ("received_date", "Sample Received Date", "29/08/2026"),
    ("report_date", "Report Date", "21/09/2026"),
]


class RenderWorker(QThread):
    """Render the report off the UI thread. Emits the saved path(s), one per line.
    kind "docx": out_path is the .docx.  kind "pdf": out_path is the .pdf.
    kind "both": out_path is the .docx; the .pdf is saved next to it."""
    done = Signal(str)
    failed = Signal(str)

    def __init__(self, data: ReportData, out_path: str, kind: str):
        super().__init__()
        self.data, self.out_path, self.kind = data, out_path, kind

    def run(self):
        try:
            from ..core.report_builder import docx_to_pdf, render_docx
            if self.kind == "docx":
                self.done.emit(render_docx(self.data, self.out_path))
                return
            if self.kind == "both":
                docx = render_docx(self.data, self.out_path)
                pdf = docx_to_pdf(docx, os.path.dirname(docx))
                self.done.emit(f"{docx}\n{pdf}")
                return
            tmp = tempfile.mkdtemp(prefix="cmc_render_")
            try:
                stem = os.path.splitext(os.path.basename(self.out_path))[0]
                pdf = docx_to_pdf(render_docx(self.data, os.path.join(tmp, stem + ".docx")), tmp)
                shutil.move(pdf, self.out_path)
            finally:
                shutil.rmtree(tmp, ignore_errors=True)
            self.done.emit(self.out_path)
        except Exception:  # noqa: BLE001 - shown to the user
            self.failed.emit(traceback.format_exc())


class PreviewScroll(QScrollArea):
    resized = Signal()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self.resized.emit()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        # Fit the screen's usable area (excludes the taskbar) so nothing opens
        # hidden behind it; main() then shows the window maximised.
        avail = QApplication.primaryScreen().availableGeometry()
        self.resize(min(1400, int(avail.width() * 0.9)), min(900, int(avail.height() * 0.9)))
        self.move(avail.center() - self.rect().center())
        self.settings = QSettings("AndersonDiagnostics", "CMCReportAutomation")
        self._preview_dir = tempfile.mkdtemp(prefix="cmc_preview_")
        self._preview_pdf = ""
        self._preview_seq = 0
        self._preview_worker: RenderWorker | None = None
        self._export_worker: RenderWorker | None = None
        self._pending = False            # an edit arrived while a preview was rendering

        # Short delay so several keystrokes share one render; not restarted by
        # further typing, so the preview keeps updating while you type.
        self._preview_timer = QTimer(self, singleShot=True, interval=300)
        self._preview_timer.timeout.connect(self._preview)
        self._redraw_timer = QTimer(self, singleShot=True, interval=120)
        self._redraw_timer.timeout.connect(self._redraw)

        self._template_defaults: dict[str, str] = {}
        self._build_ui()
        self._template_changed()

    # -- layout ------------------------------------------------------------
    def _build_ui(self):
        form = QWidget()
        lay = QVBoxLayout(form)

        row = QHBoxLayout()
        row.addWidget(QLabel("<b>Template:</b>"))
        self.template_combo = QComboBox()
        for key, cfg in TEMPLATES.items():
            self.template_combo.addItem(cfg.name, key)
        self.template_combo.currentIndexChanged.connect(self._template_changed)
        row.addWidget(self.template_combo)
        row.addStretch(1)
        lay.addLayout(row)

        pgrp = QGroupBox("Patient details")
        pf = QFormLayout(pgrp)
        self.fields: dict[str, QLineEdit] = {}
        for key, label, example in PATIENT_FIELDS:
            le = QLineEdit()
            le.setPlaceholderText(example)
            le.textChanged.connect(self._edited)
            self.fields[key] = le
            pf.addRow(label, le)
        lay.addWidget(pgrp)

        hgrp = QGroupBox("Clinical history")
        hl = QVBoxLayout(hgrp)
        self.history_edit = QPlainTextEdit()
        self.history_edit.setTabChangesFocus(True)
        self.history_edit.setMinimumHeight(100)
        self.history_edit.setPlaceholderText(
            "Proband, <Patient ID>, is suspected to be affected with <indication> "
            "and has been evaluated for pathogenic variations.")
        self.history_edit.textChanged.connect(self._edited)
        hl.addWidget(self.history_edit)
        lay.addWidget(hgrp, 1)

        qgrp = QGroupBox("Sequence data attributes")
        qf = QFormLayout(qgrp)
        self.reads_edit = QLineEdit(placeholderText="12.6 GB")
        self.q30_edit = QLineEdit(placeholderText="96.34 %")
        for le in (self.reads_edit, self.q30_edit):
            le.textChanged.connect(self._edited)
        qf.addRow("Total Read Generated", self.reads_edit)
        qf.addRow("Data ≥ Q30", self.q30_edit)
        lay.addWidget(qgrp)

        fgrp = QGroupBox("Export folder")
        fl = QHBoxLayout(fgrp)
        self.folder_edit = QLineEdit(self.settings.value("out_dir", "") or
                                     os.path.join(os.path.expanduser("~"), "Downloads"))
        self.folder_edit.setReadOnly(True)
        fl.addWidget(self.folder_edit, 1)
        browse = QPushButton("Browse…")
        browse.clicked.connect(self._choose_folder)
        fl.addWidget(browse)
        lay.addWidget(fgrp)

        row = QHBoxLayout()
        for text, kind in (("Export DOCX", "docx"), ("Export PDF", "pdf"),
                           ("Export DOCX + PDF", "both")):
            b = QPushButton(text)
            b.setMinimumHeight(32)
            b.clicked.connect(lambda _=False, k=kind: self._export(k))
            row.addWidget(b)
        lay.addLayout(row)

        row = QHBoxLayout()
        for text, slot in (("Save Draft", self._save_draft), ("Load Draft", self._load_draft)):
            b = QPushButton(text)
            b.setMinimumHeight(32)
            b.clicked.connect(slot)
            row.addWidget(b)
        lay.addLayout(row)

        preview = QWidget()
        pl = QVBoxLayout(preview)
        bar = QHBoxLayout()
        refresh = QPushButton("Refresh Preview")
        refresh.clicked.connect(self._preview)
        bar.addWidget(refresh)
        self.preview_status = QLabel()
        self.preview_status.setStyleSheet("color:gray;font-style:italic;")
        bar.addWidget(self.preview_status, 1)
        pl.addLayout(bar)
        self.scroll = PreviewScroll()
        self.scroll.setWidgetResizable(True)
        self.scroll.setStyleSheet("background:#606060;")
        self.scroll.resized.connect(self._redraw_timer.start)
        pages = QWidget()
        self.pages = QVBoxLayout(pages)
        self.pages.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        self.scroll.setWidget(pages)
        pl.addWidget(self.scroll, 1)

        split = QSplitter(Qt.Horizontal)
        # Scroll the form on short screens instead of squashing its fields.
        form_scroll = QScrollArea()
        form_scroll.setWidgetResizable(True)
        form_scroll.setFrameShape(QScrollArea.NoFrame)
        form_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        form_scroll.setWidget(form)
        form_scroll.setMinimumWidth(form.minimumSizeHint().width()
                                    + form_scroll.verticalScrollBar().sizeHint().width() + 4)
        split.addWidget(form_scroll)
        split.addWidget(preview)
        split.setSizes([520, 880])
        self.setCentralWidget(split)

    # -- data --------------------------------------------------------------
    def _collect(self) -> ReportData:
        return ReportData(
            template_key=self.template_combo.currentData() or DEFAULT_TEMPLATE,
            patient=PatientInfo(**{k: le.text().strip() for k, le in self.fields.items()}),
            clinical_history=self.history_edit.toPlainText().strip(),
            total_reads=self.reads_edit.text().strip(),
            q30=self.q30_edit.text().strip(),
        )

    def _edited(self, *_):
        if not self._preview_timer.isActive():
            self._preview_timer.start()

    def _template_changed(self, *_):
        """Pre-fill the chosen template's Hospital/Clinic and Referring
        Clinician. A field is only replaced if it is empty or still holds the
        previous template's default, so the user's own edits are kept."""
        cfg = get_template(self.template_combo.currentData())
        new = {"hospital": cfg.hospital, "referring_clinician": cfg.referring_clinician}
        for key, value in new.items():
            le = self.fields[key]
            if not le.text().strip() or le.text() == self._template_defaults.get(key):
                le.setText(value)
        self._template_defaults = new
        self._preview()

    # -- drafts ------------------------------------------------------------
    def _save_draft(self):
        data = self._collect()
        pid = re.sub(r"\s+", "", data.patient.patient_id) or "report"
        folder = self.settings.value("draft_dir", "") or self.folder_edit.text()
        path, _ = QFileDialog.getSaveFileName(
            self, "Save draft", os.path.join(folder, f"{pid}_draft.json"), "Draft (*.json)")
        if not path:
            return
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(data.to_dict(), fh, indent=2, ensure_ascii=False)
        self.settings.setValue("draft_dir", os.path.dirname(path))
        QMessageBox.information(self, "Draft saved", f"Saved:\n{os.path.normpath(path)}")

    def _load_draft(self):
        folder = self.settings.value("draft_dir", "") or self.folder_edit.text()
        path, _ = QFileDialog.getOpenFileName(self, "Load draft", folder, "Draft (*.json)")
        if not path:
            return
        try:
            with open(path, encoding="utf-8") as fh:
                data = ReportData.from_dict(json.load(fh))
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Load draft", f"Could not read the draft:\n{e}")
            return
        self.settings.setValue("draft_dir", os.path.dirname(path))
        # Select the draft's template without its defaults overwriting the draft.
        cfg = get_template(data.template_key)
        self.template_combo.blockSignals(True)
        self.template_combo.setCurrentIndex(max(self.template_combo.findData(cfg.key), 0))
        self.template_combo.blockSignals(False)
        self._template_defaults = {"hospital": cfg.hospital,
                                   "referring_clinician": cfg.referring_clinician}
        for key, le in self.fields.items():
            le.setText(getattr(data.patient, key, ""))
        self.history_edit.setPlainText(data.clinical_history)
        self.reads_edit.setText(data.total_reads)
        self.q30_edit.setText(data.q30)
        self._preview()

    def _choose_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Choose export folder", self.folder_edit.text())
        if folder:
            folder = os.path.normpath(folder)
            self.folder_edit.setText(folder)
            self.settings.setValue("out_dir", folder)

    # -- preview -----------------------------------------------------------
    def _preview(self):
        if self._preview_worker is not None and self._preview_worker.isRunning():
            self._pending = True           # re-render once this one finishes
            return
        self._pending = False
        self._preview_seq += 1
        out = os.path.join(self._preview_dir, f"preview_{self._preview_seq}.pdf")
        self.preview_status.setText("Rendering preview…")
        self._preview_worker = RenderWorker(self._collect(), out, "pdf")
        self._preview_worker.done.connect(self._preview_ready)
        self._preview_worker.failed.connect(self._preview_failed)
        self._preview_worker.finished.connect(
            lambda: self._pending and self._preview())
        self._preview_worker.start()

    def _preview_ready(self, pdf: str):
        old, self._preview_pdf = self._preview_pdf, pdf
        n = self._redraw()
        self.preview_status.setText(
            f"Updated {datetime.now().strftime('%H:%M:%S')} — {n} page(s)")
        if old and os.path.exists(old):
            try:
                os.remove(old)
            except OSError:
                pass        # still open elsewhere; the temp folder is removed on exit

    def _preview_failed(self, tb: str):
        self.preview_status.setText("Preview failed: " + tb.strip().splitlines()[-1])

    def _redraw(self) -> int:
        """Draw each page at the screen's physical pixel density, so text is
        sharp on displays with Windows scaling above 100%."""
        if not self._preview_pdf or not os.path.exists(self._preview_pdf):
            return 0
        import pypdfium2 as pdfium
        vbar = self.scroll.verticalScrollBar()
        pos = vbar.value() / vbar.maximum() if vbar.maximum() else 0.0
        dpr = self.scroll.devicePixelRatioF() or 1.0
        width = max(self.scroll.viewport().width() - 28, 240)
        doc = pdfium.PdfDocument(self._preview_pdf)
        try:
            count = len(doc)
            while self.pages.count() > count:
                self.pages.takeAt(self.pages.count() - 1).widget().deleteLater()
            while self.pages.count() < count:
                lbl = QLabel(alignment=Qt.AlignHCenter)
                lbl.setStyleSheet("background:white;margin:6px 0;")
                self.pages.addWidget(lbl)
            for i in range(count):
                page = doc[i]
                img = page.render(scale=width * dpr / page.get_width()).to_pil().convert("RGB")
                qimg = QImage(img.tobytes(), img.width, img.height, img.width * 3,
                              QImage.Format_RGB888).copy()
                px = QPixmap.fromImage(qimg)
                px.setDevicePixelRatio(dpr)
                self.pages.itemAt(i).widget().setPixmap(px)
        finally:
            doc.close()
        QTimer.singleShot(0, lambda: vbar.setValue(int(pos * vbar.maximum())))
        return count

    # -- export ------------------------------------------------------------
    def _export(self, kind: str):
        if self._export_worker is not None and self._export_worker.isRunning():
            return
        data = self._collect()
        pid = re.sub(r"\s+", "", data.patient.patient_id) or "report"
        date = data.patient.report_date.replace("/", "-") or datetime.now().strftime("%d-%m-%Y")
        stem = re.sub(r'[\\/:*?"<>|]+', "_",
                      f"{pid}_{get_template(data.template_key).file_suffix}_{date}")
        folder = self.folder_edit.text()
        if not os.path.isdir(folder):
            self._choose_folder()
            folder = self.folder_edit.text()
            if not os.path.isdir(folder):
                return
        exts = ["docx", "pdf"] if kind == "both" else [kind]
        existing = [f"{stem}.{e}" for e in exts if os.path.exists(os.path.join(folder, f"{stem}.{e}"))]
        if existing and QMessageBox.question(
                self, "File exists",
                "Already in the folder:\n" + "\n".join(existing) + "\n\nReplace?"
        ) != QMessageBox.Yes:
            return
        path = os.path.join(folder, f"{stem}.{exts[0]}")
        self.setCursor(Qt.WaitCursor)
        self._export_worker = RenderWorker(data, path, kind)
        self._export_worker.done.connect(self._exported)
        self._export_worker.failed.connect(self._export_failed)
        self._export_worker.start()

    def _exported(self, path: str):
        self.unsetCursor()
        QMessageBox.information(self, "Export complete", f"Saved:\n{path}")

    def _export_failed(self, tb: str):
        self.unsetCursor()
        QMessageBox.critical(self, "Export failed", tb)

    def closeEvent(self, e):
        for w in (self._preview_worker, self._export_worker):
            if w is not None:
                w.wait(15000)
        shutil.rmtree(self._preview_dir, ignore_errors=True)
        super().closeEvent(e)


def main():
    # Use the monitor's exact scale (125% / 150%) instead of rounding it.
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setStyle("Fusion")
    win = MainWindow()
    win.showMaximized()
    sys.exit(app.exec())
