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

from PySide6.QtCore import QDate, QPoint, QSettings, QSize, Qt, QThread, QTimer, Signal
from PySide6.QtGui import QColor, QIcon, QImage, QKeySequence, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QComboBox, QFileDialog, QFormLayout,
    QGridLayout, QGroupBox, QHBoxLayout, QHeaderView, QLabel, QLineEdit,
    QMainWindow, QMenu, QMessageBox, QPlainTextEdit, QPushButton, QScrollArea,
    QSplitter, QTableWidget, QTableWidgetItem, QTabWidget, QToolButton,
    QVBoxLayout, QWidget, QWidgetAction,
)

from ..config.templates import DEFAULT_TEMPLATE, REVIEWERS, TEMPLATES, get_template
from ..core.docx_renderer import template_gene_table
from ..core.gene_table import parse_pasted
from ..core import sheet_client
from ..core.models import PatientInfo, ReportData, today_str

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
    ("collection_date", "Sample Collection Date", "dd/mm/yyyy"),
    ("specimen", "Specimen", "DNA"),
    ("received_date", "Sample Received Date", "dd/mm/yyyy"),
    ("report_date", "Report Date", "dd/mm/yyyy"),
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
            from ..core.report_builder import build_report
            if self.kind == "docx":
                self.done.emit(build_report(self.data, self.out_path)[0])
                return
            if self.kind == "both":
                docx, pdf = build_report(self.data, self.out_path,
                                         os.path.dirname(os.path.abspath(self.out_path)))
                self.done.emit(f"{docx}\n{pdf}")
                return
            tmp = tempfile.mkdtemp(prefix="cmc_render_")
            try:
                stem = os.path.splitext(os.path.basename(self.out_path))[0]
                _, pdf = build_report(self.data, os.path.join(tmp, stem + ".docx"), tmp)
                shutil.move(pdf, self.out_path)
            finally:
                shutil.rmtree(tmp, ignore_errors=True)
            self.done.emit(self.out_path)
        except Exception:  # noqa: BLE001 - shown to the user
            self.failed.emit(traceback.format_exc())


class FetchWorker(QThread):
    """Look up one PIN in the live patient and sequencing QC sheets off the
    UI thread."""
    found = Signal(str, dict, str)  # pin, fields, note ('' when all found)
    missing = Signal(str)           # pin
    failed = Signal(str, str)       # pin, message

    def __init__(self, pin: str, template_key: str):
        super().__init__()
        self.pin, self.template_key = pin, template_key

    def run(self):
        fields, notes = None, []
        cfg = get_template(self.template_key)
        try:
            if sheet_client.patient_configured(cfg):
                fields = sheet_client.fetch_patient(self.pin, cfg)
        except sheet_client.SheetError as e:
            self.failed.emit(self.pin, str(e))
            return
        # The QC values are looked up separately: a problem there must not
        # lose the patient details.
        qc = None
        if sheet_client.qc_configured():
            try:
                qc = sheet_client.fetch_qc(self.pin)
            except sheet_client.SheetError as e:
                notes.append(f"Sequence data attributes: {e}")
            else:
                if qc is None:
                    notes.append(f"{self.pin} is not in the sequencing QC sheet - "
                                 "enter the Sequence data attributes manually.")
        if fields is None and qc is None:
            if notes and not notes[0].startswith(self.pin):
                self.failed.emit(self.pin, notes[0])
            else:
                self.missing.emit(self.pin)
            return
        if fields is None and sheet_client.patient_configured(cfg):
            notes.insert(0, f"{self.pin} was not found in the patient sheet - "
                            "enter the patient details manually.")
        out = dict(fields or {"pin": self.pin})
        out.update(qc or {})
        self.found.emit(self.pin, out, "\n".join(notes))


PIN_PATTERN = re.compile(r"ADK\d{10}", re.I)


class ChoiceBox(QComboBox):
    """Dropdown with the same text()/setText()/textChanged API as QLineEdit,
    so it slots into the patient fields like any other box. The mouse wheel
    is ignored unless it has focus, so scrolling the form can't change it."""
    textChanged = Signal(str)

    def __init__(self, options: list[str]):
        super().__init__()
        self.addItems([""] + options)        # blank first: nothing pre-selected
        self.setFocusPolicy(Qt.StrongFocus)
        # Don't force the panel as wide as the longest option (the open list
        # still shows it in full), so the preview keeps its room.
        self.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.setMinimumContentsLength(12)
        self.currentTextChanged.connect(self.textChanged)

    def text(self) -> str:
        return self.currentText()

    def setText(self, value: str):
        if value and self.findText(value) < 0:
            self.addItem(value)              # e.g. a value from an older draft
        self.setCurrentText(value)

    def wheelEvent(self, e):
        if self.hasFocus():
            super().wheelEvent(e)
        else:
            e.ignore()


class OtherChoiceBox(ChoiceBox):
    """Dropdown whose last entry is 'Other': choosing it turns the box into a
    text field for a name that isn't in the list. 'Other' itself is never
    returned as the value."""
    OTHER = "Other"

    def __init__(self, options: list[str], placeholder: str = "Type the value"):
        QComboBox.__init__(self)
        self._options = list(options)
        self._placeholder = placeholder
        self.addItems(self._options + [self.OTHER])
        self.setFocusPolicy(Qt.StrongFocus)
        self.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.setMinimumContentsLength(12)
        self.activated.connect(self._chosen)
        self.currentTextChanged.connect(lambda *_: self.textChanged.emit(self.text()))
        self._set_other("")

    def _set_other(self, value: str):
        self.blockSignals(True)
        self.setCurrentIndex(self.count() - 1)
        self.setEditable(True)
        self.lineEdit().setPlaceholderText(self._placeholder)
        self.lineEdit().setText(value)
        self.blockSignals(False)
        self.textChanged.emit(self.text())

    def _chosen(self, index: int):
        if index == self.count() - 1:          # "Other": type a name
            self._set_other("")
            self.lineEdit().setFocus()
        else:
            self.setEditable(False)
            self.setCurrentIndex(index)
            self.textChanged.emit(self.text())

    def text(self) -> str:
        if self.isEditable():
            value = self.lineEdit().text().strip()
            return "" if value == self.OTHER else value
        return self.currentText()

    def setText(self, value: str):
        value = (value or "").strip()
        if value in self._options:
            self.blockSignals(True)
            self.setEditable(False)
            self.setCurrentIndex(self._options.index(value))
            self.blockSignals(False)
            self.textChanged.emit(value)
        else:
            self._set_other(value)             # blank or a name not in the list


def calendar_icon() -> QIcon:
    """A small calendar glyph, painted at several sizes so it stays sharp at
    any Windows display scaling (Qt has no built-in calendar icon)."""
    icon = QIcon()
    for size in (16, 20, 24, 32, 40, 48, 64):
        px = QPixmap(size, size)
        px.fill(Qt.transparent)
        p = QPainter(px)
        p.setRenderHint(QPainter.Antialiasing)
        u = size / 16.0                      # drawn on a 16-unit grid
        blue, grey = QColor("#1A467D"), QColor("#5A6270")
        # page
        p.setPen(QPen(grey, max(1.0, u)))
        p.setBrush(QColor("white"))
        p.drawRoundedRect(1.5 * u, 3 * u, 13 * u, 11.5 * u, 1.5 * u, 1.5 * u)
        # header band
        p.setPen(Qt.NoPen)
        p.setBrush(blue)
        p.drawRoundedRect(1.5 * u, 3 * u, 13 * u, 3.5 * u, 1.5 * u, 1.5 * u)
        p.drawRect(1.5 * u, 5 * u, 13 * u, 1.5 * u)
        # binder rings
        p.setPen(QPen(grey, max(1.0, 1.2 * u), c=Qt.RoundCap))
        p.drawLine(5 * u, 1.5 * u, 5 * u, 4.5 * u)
        p.drawLine(11 * u, 1.5 * u, 11 * u, 4.5 * u)
        # day grid
        p.setPen(Qt.NoPen)
        p.setBrush(grey)
        for row in range(2):
            for col in range(3):
                p.drawRect((3.5 + col * 3.3) * u, (8.2 + row * 3) * u, 2 * u, 1.8 * u)
        p.end()
        icon.addPixmap(px)
    return icon


def chevron_icon(direction: str) -> QIcon:
    """Thin '<' / '>' arrows for the calendar's month buttons."""
    icon = QIcon()
    for size in (16, 20, 24, 32, 48):
        px = QPixmap(size, size)
        px.fill(Qt.transparent)
        p = QPainter(px)
        p.setRenderHint(QPainter.Antialiasing)
        p.setPen(QPen(QColor("#374151"), size / 9, c=Qt.RoundCap, j=Qt.RoundJoin))
        u = size / 16.0
        xs = (10, 6, 10) if direction == "left" else (6, 10, 6)
        p.drawPolyline([QPoint(int(xs[0] * u), int(4 * u)), QPoint(int(xs[1] * u), int(8 * u)),
                        QPoint(int(xs[2] * u), int(12 * u))])
        p.end()
        icon.addPixmap(px)
    return icon


class DatePicker(QWidget):
    """Clean month-view date picker: arrows to change month, round day
    buttons, the chosen day as a filled circle and today as a ring.
    Same small API as QCalendarWidget: clicked(QDate), selectedDate(),
    setSelectedDate()."""
    clicked = Signal(QDate)

    STYLE = """
    QWidget#picker { background: white; }
    QLabel#title { color: #111827; font-weight: 600; font-size: 10.5pt; }
    QLabel#weekday { color: #6b7280; font-size: 8.5pt; font-weight: 600; }
    QToolButton#nav { border: none; border-radius: 14px; background: transparent; }
    QToolButton#nav:hover { background: #eef2f7; }
    QToolButton#day {
        border: none; border-radius: 15px; background: transparent;
        color: #111827; font-size: 9.5pt; }
    QToolButton#day:hover { background: #e8eef7; }
    QToolButton#day[other="true"] { color: #c0c6cf; }
    QToolButton#day[today="true"] { border: 1.5px solid #1A467D; color: #1A467D; font-weight: 600; }
    QToolButton#day[selected="true"] { background: #1A467D; color: white; font-weight: 600; }
    """

    def __init__(self):
        super().__init__()
        self.setObjectName("picker")
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setStyleSheet(self.STYLE)
        self._selected = QDate.currentDate()
        self._month = QDate(self._selected.year(), self._selected.month(), 1)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(10, 8, 10, 10)
        lay.setSpacing(6)
        head = QHBoxLayout()
        prev_btn, next_btn = QToolButton(objectName="nav"), QToolButton(objectName="nav")
        for btn, direction, step in ((prev_btn, "left", -1), (next_btn, "right", 1)):
            btn.setIcon(chevron_icon(direction))
            btn.setIconSize(QSize(14, 14))
            btn.setFixedSize(28, 28)
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda _=False, st=step: self._shift(st))
        self.title = QLabel(objectName="title", alignment=Qt.AlignCenter)
        head.addWidget(prev_btn)
        head.addWidget(self.title, 1)
        head.addWidget(next_btn)
        lay.addLayout(head)

        grid = QGridLayout()
        grid.setHorizontalSpacing(2)
        grid.setVerticalSpacing(2)
        for c, name in enumerate(("Mo", "Tu", "We", "Th", "Fr", "Sa", "Su")):
            grid.addWidget(QLabel(name, objectName="weekday", alignment=Qt.AlignCenter), 0, c)
        self._days = []
        for i in range(42):
            b = QToolButton(objectName="day")
            b.setFixedSize(30, 30)
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(lambda _=False, idx=i: self._pick(idx))
            grid.addWidget(b, 1 + i // 7, i % 7)
            self._days.append(b)
        lay.addLayout(grid)
        self._refresh()

    def selectedDate(self) -> QDate:
        return self._selected

    def setSelectedDate(self, d: QDate):
        self._selected = d
        self._month = QDate(d.year(), d.month(), 1)
        self._refresh()

    def _first_cell(self) -> QDate:
        return self._month.addDays(-(self._month.dayOfWeek() - 1))   # Monday first

    def _shift(self, months: int):
        self._month = self._month.addMonths(months)
        self._refresh()

    def _pick(self, idx: int):
        d = self._first_cell().addDays(idx)
        self._selected = d
        self._refresh()
        self.clicked.emit(d)

    def _refresh(self):
        self.title.setText(self._month.toString("MMMM yyyy"))
        first, today = self._first_cell(), QDate.currentDate()
        for i, b in enumerate(self._days):
            d = first.addDays(i)
            b.setText(str(d.day()))
            b.setProperty("other", d.month() != self._month.month())
            b.setProperty("today", d == today)
            b.setProperty("selected", d == self._selected)
            b.style().unpolish(b)          # re-apply the stylesheet for new properties
            b.style().polish(b)


class DateField(QWidget):
    """Date box (typed or picked from a calendar), dd/mm/yyyy. Starts empty.
    Same text()/setText()/textChanged API as QLineEdit."""
    textChanged = Signal(str)
    FORMAT = "dd/MM/yyyy"

    def __init__(self, placeholder: str = ""):
        super().__init__()
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(4)
        self.edit = QLineEdit(placeholderText=placeholder or "dd/mm/yyyy")
        self.edit.textChanged.connect(self.textChanged)
        lay.addWidget(self.edit, 1)
        self.calendar = DatePicker()
        self.calendar.clicked.connect(self._picked)
        self._menu = QMenu(self)
        self._menu.setStyleSheet("QMenu { background: white; border: 1px solid #d0d7de; "
                                 "border-radius: 8px; padding: 2px; }")
        act = QWidgetAction(self._menu)
        act.setDefaultWidget(self.calendar)
        self._menu.addAction(act)
        self._menu.aboutToShow.connect(self._sync_calendar)
        # Calendar icon inside the box, at its right edge.
        pick = self.edit.addAction(calendar_icon(), QLineEdit.TrailingPosition)
        pick.setToolTip("Choose the date from a calendar")
        pick.triggered.connect(self._open_calendar)

    def _open_calendar(self):
        # Drop the calendar below the box, right-aligned with it.
        below = self.edit.mapToGlobal(self.edit.rect().bottomRight())
        self._menu.popup(below - QPoint(self._menu.sizeHint().width(), 0))

    def _sync_calendar(self):
        d = QDate.fromString(self.edit.text().strip(), self.FORMAT)
        self.calendar.setSelectedDate(d if d.isValid() else QDate.currentDate())

    def _picked(self, d: QDate):
        self.edit.setText(d.toString(self.FORMAT))
        self._menu.close()

    def text(self) -> str:
        return self.edit.text()

    def setText(self, value: str):
        self.edit.setText(value)


# Patient fields shown as dropdowns / date pickers instead of plain text.
PATIENT_CHOICES = {
    "hospital": ["Christian Medical College - Molecular Endocrinology",
                 "Christian Medical College - Nephrology"],
    "specimen": ["DNA", "Peripheral Blood"],
}
OTHER_PLACEHOLDERS = {
    "hospital": "Type the hospital / clinic name",
    "specimen": "Type the specimen",
}
DATE_FIELDS = {"collection_date", "received_date", "report_date"}


class GeneTableEditor(QWidget):
    """Appendix 1 gene coverage table: paste a whole table to replace it."""
    changed = Signal()

    def __init__(self):
        super().__init__()
        lay = QVBoxLayout(self)
        info = QLabel("Copy the whole gene coverage table (from Word, Excel or the "
                      "coverage report) and click <b>Paste table</b> to replace the "
                      "table in Appendix 1. Both layouts work: 4 gene/coverage pairs "
                      "per row, or 2 columns (Gene, Coverage).")
        info.setWordWrap(True)
        lay.addWidget(info)
        row = QHBoxLayout()
        paste = QPushButton("Paste table")
        paste.setToolTip("Replace the table with the one on the clipboard (Ctrl+V)")
        paste.clicked.connect(self.paste)
        row.addWidget(paste)
        self.restore_btn = QPushButton("Restore template table")
        row.addWidget(self.restore_btn)
        row.addStretch(1)
        self.count = QLabel()
        row.addWidget(self.count)
        lay.addLayout(row)
        self.status = QLabel()
        self.status.setWordWrap(True)
        lay.addWidget(self.status)
        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["Gene Name", "% of coding region covered"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.itemChanged.connect(lambda *_: self._changed())
        lay.addWidget(self.table, 1)

    def genes(self) -> list[tuple[str, str]]:
        out = []
        for r in range(self.table.rowCount()):
            g, c = self.table.item(r, 0), self.table.item(r, 1)
            gene = g.text().strip() if g else ""
            if gene:
                out.append((gene, c.text().strip() if c else ""))
        return out

    def set_genes(self, genes, status: str = "", color: str = "gray"):
        self.table.blockSignals(True)
        self.table.setRowCount(len(genes))
        for r, (gene, cov) in enumerate(genes):
            self.table.setItem(r, 0, QTableWidgetItem(gene))
            self.table.setItem(r, 1, QTableWidgetItem(cov))
        self.table.blockSignals(False)
        self.status.setText(status)
        self.status.setStyleSheet(f"color:{color};")
        self._changed()

    def _changed(self):
        self.count.setText(f"{len(self.genes())} genes")
        self.changed.emit()

    def paste(self):
        genes, problems = parse_pasted(QApplication.clipboard().text())
        if not genes:
            self.status.setText("No gene table found on the clipboard. Copy the table "
                                "(including the gene and coverage cells) and try again.")
            self.status.setStyleSheet("color:#c00000;")
            return
        msg, color = f"Pasted {len(genes)} genes.", "#1a7f37"
        if problems:
            shown = "; ".join(problems[:5]) + (" …" if len(problems) > 5 else "")
            msg += f" Skipped {len(problems)} cell pair(s) that aren't a gene + coverage: {shown}"
            color = "#b35c00"
        self.set_genes(genes, msg, color)

    def keyPressEvent(self, e):
        if e.matches(QKeySequence.Paste):
            self.paste()
        else:
            super().keyPressEvent(e)


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
        self._template_reviewer = ""
        self._template_genes: list[tuple[str, str]] = []
        self._last_received = ""
        self._loading_draft = False
        self._build_ui()
        # Report date defaults to today; the user can change it.
        self.fields["report_date"].setText(today_str())
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
        row.addSpacing(16)
        row.addWidget(QLabel("<b>Clinical reviewer:</b>"))
        self.reviewer_combo = QComboBox()
        self.reviewer_combo.setToolTip("Whose signature goes in the signature block")
        for key, rev in REVIEWERS.items():
            self.reviewer_combo.addItem(rev.name, key)
        self.reviewer_combo.currentIndexChanged.connect(self._edited)
        row.addWidget(self.reviewer_combo)
        row.addStretch(1)
        lay.addLayout(row)

        # Two tabs: patient demography / the rest of the report inputs.
        patient_page, report_page = QWidget(), QWidget()
        play, rlay = QVBoxLayout(patient_page), QVBoxLayout(report_page)

        pgrp = QGroupBox("Patient details")
        pf = QFormLayout(pgrp)
        self.fields: dict[str, QLineEdit | ChoiceBox | DateField] = {}
        for key, label, example in PATIENT_FIELDS:
            if key in PATIENT_CHOICES:     # dropdown + "Other" (typed by the user)
                le = OtherChoiceBox(PATIENT_CHOICES[key], OTHER_PLACEHOLDERS[key])
            elif key in DATE_FIELDS:
                le = DateField(example)
            else:
                le = QLineEdit()
                le.setPlaceholderText(example)
            le.textChanged.connect(self._edited)
            self.fields[key] = le
            if key == "pin":
                pf.addRow(label, self._pin_row(le))
                self.fetch_status = QLabel()
                self.fetch_status.setWordWrap(True)
                pf.addRow("", self.fetch_status)
                self._set_fetch_status(
                    "Enter the PIN (Anderson ID) to fill the details from the patient sheet."
                    if sheet_client.is_configured() else
                    "Patient sheet not set up - enter the details manually.", "gray")
            else:
                pf.addRow(label, le)
            if key == "patient_id":
                self.patient_id_label = pf.labelForField(le)
        play.addWidget(pgrp)
        self.fields["received_date"].textChanged.connect(self._received_changed)
        play.addStretch(1)

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
        rlay.addWidget(hgrp, 1)

        qgrp = QGroupBox("Sequence data attributes")
        qf = QFormLayout(qgrp)
        self.reads_edit = QLineEdit(placeholderText="12.6 GB")
        self.q30_edit = QLineEdit(placeholderText="96.34 %")
        for le in (self.reads_edit, self.q30_edit):
            le.textChanged.connect(self._edited)
        qf.addRow("Total Read Generated", self.reads_edit)
        qf.addRow("Data ≥ Q30", self.q30_edit)
        rlay.addWidget(qgrp)

        lgrp = QGroupBox("Gene list link (Methodology “Click here”)")
        ll = QVBoxLayout(lgrp)
        self.url_edit = QLineEdit(placeholderText="https://…")
        self.url_edit.setToolTip("Paste this run's gene list link.")
        self.url_edit.textChanged.connect(self._edited)
        ll.addWidget(self.url_edit)
        rlay.addWidget(lgrp)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._scroll_page(patient_page), "Patient Details")
        self.tabs.addTab(self._scroll_page(report_page), "Report Details")
        self.gene_editor = GeneTableEditor()
        self.gene_editor.changed.connect(self._edited)
        self.gene_editor.restore_btn.clicked.connect(self._restore_gene_table)
        self.tabs.addTab(self.gene_editor, "Gene Coverage")
        lay.addWidget(self.tabs, 1)

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
        split.addWidget(form)
        split.addWidget(preview)
        split.setStretchFactor(0, 0)          # extra width goes to the preview
        split.setStretchFactor(1, 1)
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
            gene_list_url=self.url_edit.text().strip(),
            genes=self.gene_editor.genes(),
            reviewer=self.reviewer_combo.currentData() or "",
        )

    def _edited(self, *_):
        if not self._preview_timer.isActive():
            self._preview_timer.start()

    def _template_changed(self, *_):
        """Pre-fill the chosen template's Hospital/Clinic, Referring Clinician
        and Specimen. A field is only replaced if it is empty or still holds the
        previous template's default, so the user's own edits are kept."""
        key = self.template_combo.currentData()
        new = self._defaults_for(key)
        for name, value in new.items():
            le = self.fields[name]
            if not le.text().strip() or le.text() == self._template_defaults.get(name):
                le.setText(value)
        self._template_defaults = new
        cfg = get_template(key)
        # Clinical reviewer: the template's, unless the user picked another.
        if self.reviewer_combo.currentData() in (None, self._template_reviewer):
            self._set_reviewer(cfg.reviewer)
        self._template_reviewer = cfg.reviewer
        # "Patient ID" or "Patient name", as the template's table says.
        self.patient_id_label.setText(cfg.patient_id_label)
        self.fields["patient_id"].setPlaceholderText(
            f"e.g. {cfg.patient_id_prefix} - 00000" if cfg.patient_id_prefix
            else "e.g. 12345" if cfg.patient_id_label == "Patient ID" else "e.g. Baby. Name")
        # Same rule for the gene table: follow the template unless it was replaced.
        genes = template_gene_table(key)
        current = self.gene_editor.genes()
        if not current or current == self._template_genes:
            self.gene_editor.set_genes(genes, "Showing the template's gene table.")
        self._template_genes = genes
        self._preview()

    def _set_reviewer(self, key: str):
        self.reviewer_combo.setCurrentIndex(max(self.reviewer_combo.findData(key), 0))

    def _restore_gene_table(self):
        self.gene_editor.set_genes(self._template_genes, "Restored the template's gene table.")

    def _received_changed(self, text: str):
        """Collection date is the same as the received date: keep it in step,
        unless the user has set a different collection date themselves."""
        coll = self.fields["collection_date"]
        if not coll.text().strip() or coll.text() == self._last_received:
            coll.setText(text)
        self._last_received = text

    # -- patient sheet lookup ----------------------------------------------
    @staticmethod
    def _scroll_page(page: QWidget) -> QScrollArea:
        """Tab page that scrolls on short screens instead of squashing its fields."""
        sc = QScrollArea()
        sc.setWidgetResizable(True)
        sc.setFrameShape(QScrollArea.NoFrame)
        sc.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        sc.setWidget(page)
        sc.setMinimumWidth(page.minimumSizeHint().width()
                           + sc.verticalScrollBar().sizeHint().width() + 4)
        return sc

    def _pin_row(self, pin_edit: QLineEdit) -> QWidget:
        row = QWidget()
        lay = QHBoxLayout(row)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(4)
        lay.addWidget(pin_edit, 1)
        self.fetch_btn = QPushButton("Fetch")
        self.fetch_btn.setToolTip("Fill the patient details and Sequence data "
                                  "attributes from the live sheets")
        self.fetch_btn.setEnabled(sheet_client.is_configured())
        self.fetch_btn.clicked.connect(lambda: self._fetch(pin_edit.text()))
        lay.addWidget(self.fetch_btn)
        pin_edit.returnPressed.connect(lambda: self._fetch(pin_edit.text()))
        # Look up automatically once a complete PIN has been typed or pasted.
        pin_edit.textChanged.connect(self._pin_typed)
        self._fetch_worker: FetchWorker | None = None
        self._fetched_pin = ""
        return row

    def _set_fetch_status(self, text: str, color: str):
        self.fetch_status.setText(text)
        self.fetch_status.setStyleSheet(f"color:{color};")
        self.fetch_status.setVisible(bool(text))    # no empty gap when blank

    def _pin_typed(self, text: str):
        pin = text.strip().upper()
        if self._loading_draft:
            return
        if not pin:
            self._clear_patient()
        elif PIN_PATTERN.fullmatch(pin) and pin != self._fetched_pin:
            self._fetch(pin)

    def _clear_patient(self):
        """PIN removed: empty what a Fetch fills (patient details and Sequence
        data attributes). Hospital/Clinic, Referring Clinician and Specimen go back to
        the template's defaults; the report date is kept."""
        for key, le in self.fields.items():
            if key not in ("pin", "report_date"):
                le.setText(self._template_defaults.get(key, ""))
        self.reads_edit.clear()
        self.q30_edit.clear()
        self._fetched_pin = self._last_received = ""
        self._set_fetch_status("", "gray")
        self._edited()

    def _fetch(self, pin: str):
        pin = pin.strip().upper()
        if not sheet_client.is_configured() or not pin:
            return
        if self._fetch_worker is not None and self._fetch_worker.isRunning():
            return
        self._fetched_pin = pin
        self.fetch_btn.setEnabled(False)
        self._set_fetch_status(f"Looking up {pin} in the patient sheet…", "gray")
        self._fetch_worker = FetchWorker(pin, self.template_combo.currentData())
        self._fetch_worker.found.connect(self._fetch_found)
        self._fetch_worker.missing.connect(lambda p: self._set_fetch_status(
            f"{p} was not found in the patient sheet - enter the details manually.", "#b35c00"))
        self._fetch_worker.failed.connect(lambda p, msg: self._fetch_failed(msg))
        self._fetch_worker.finished.connect(lambda: self.fetch_btn.setEnabled(True))
        self._fetch_worker.start()

    def _fetch_failed(self, msg: str):
        self._fetched_pin = ""              # allow retrying the same PIN
        self._set_fetch_status(msg, "#c00000")

    def _fetch_found(self, pin: str, fields: dict, note: str):
        if self.fields["pin"].text().strip().upper() != pin:
            return                          # PIN changed while looking up
        for key, value in fields.items():
            if key != "pin" and key in self.fields and value:
                self.fields[key].setText(value)
        qc_edits = {"total_reads": self.reads_edit, "q30": self.q30_edit}
        for key, edit in qc_edits.items():
            if fields.get(key):
                edit.setText(fields[key])
        # Filled: nothing to report unless part of it wasn't found.
        self._set_fetch_status(note, "#b35c00" if note else "gray")

    # -- drafts ------------------------------------------------------------
    def _save_draft(self):
        data = self._collect()
        pid = self._file_name_part(data.patient.patient_id) or "report"
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
        self._template_defaults = self._defaults_for(cfg.key)
        self._template_reviewer = cfg.reviewer
        self._set_reviewer(data.reviewer or cfg.reviewer)
        # A draft's saved details win: don't let its PIN trigger a sheet lookup.
        self._loading_draft = True
        for key, le in self.fields.items():
            le.setText(getattr(data.patient, key, ""))
        self._loading_draft = False
        if not self.fields["report_date"].text().strip():
            self.fields["report_date"].setText(today_str())
        self._fetched_pin = data.patient.pin.strip().upper()
        self.history_edit.setPlainText(data.clinical_history)
        self.reads_edit.setText(data.total_reads)
        self.q30_edit.setText(data.q30)
        self.url_edit.setText(data.gene_list_url)
        self._template_genes = template_gene_table(cfg.key)
        self.gene_editor.set_genes(data.genes or self._template_genes,
                                   "Gene table from the draft." if data.genes
                                   else "Showing the template's gene table.")
        self._preview()

    @staticmethod
    def _defaults_for(key: str) -> dict[str, str]:
        cfg = get_template(key)
        return {"hospital": cfg.hospital, "referring_clinician": cfg.referring_clinician,
                "specimen": cfg.specimen}

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
    @staticmethod
    def _file_name_part(text: str) -> str:
        """'MEL - 12345' -> 'MEL-12345'; 'Baby. Arun K' -> 'Baby_Arun_K'."""
        text = re.sub(r"\s*-\s*", "-", text.strip())
        return re.sub(r"[^A-Za-z0-9-]+", "_", text).strip("_")

    def _file_stem(self, data: ReportData) -> str:
        cfg = get_template(data.template_key)
        date = data.patient.report_date or datetime.now().strftime("%d/%m/%Y")
        return "_".join(filter(None, (self._file_name_part(data.patient.patient_id) or "report",
                                      cfg.file_suffix,
                                      date.replace("/", cfg.file_date_sep))))

    def _export(self, kind: str):
        if self._export_worker is not None and self._export_worker.isRunning():
            return
        data = self._collect()
        stem = self._file_stem(data)
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
