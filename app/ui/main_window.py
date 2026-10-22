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
from PySide6.QtGui import QColor, QFont, QIcon, QImage, QKeySequence, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QCheckBox, QComboBox, QFileDialog, QFormLayout, QFrame,
    QGridLayout, QGroupBox, QHBoxLayout, QHeaderView, QLabel, QLineEdit,
    QListWidget, QListWidgetItem, QMainWindow, QMenu, QMessageBox, QPlainTextEdit, QProgressDialog, QPushButton, QScrollArea,
    QSplitter, QStackedWidget, QTableWidget, QTableWidgetItem, QTabBar, QTabWidget, QToolButton,
    QVBoxLayout, QWidget, QWidgetAction,
)

from ..config.templates import DEFAULT_TEMPLATE, REVIEWERS, TEMPLATES, get_template
from ..core.docx_renderer import template_gene_table
from ..core.gene_table import parse_pasted
from ..core import bulk, gnomad, ocr, sheet_client, variants as variants_core
from . import style
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


PIN_PATTERN = re.compile(r"ADK\d{10}|PH\d{11}", re.I)    # patient sheet / WES sheet PINs


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


class NarrativeEdit(QPlainTextEdit):
    """Paragraph box for the report: pasted text (e.g. copied from a PDF) is cleaned -
    line breaks joined and the notes that follow it in the report dropped."""

    def insertFromMimeData(self, source):
        if source.hasText():
            self.insertPlainText(variants_core.clean_narrative(source.text()))
        else:
            super().insertFromMimeData(source)


class StableCombo(QComboBox):
    """Dropdown the mouse wheel never changes (even when it has focus), so
    scrolling a form can't alter the value by accident; the scroll passes on to the form."""

    def __init__(self):
        super().__init__()
        self.setFocusPolicy(Qt.StrongFocus)

    def wheelEvent(self, e):
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
        blue, grey = QColor("#0f4c5c"), QColor("#5A6270")
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
    QToolButton#day[today="true"] { border: 1.5px solid #0f4c5c; color: #0f4c5c; font-weight: 600; }
    QToolButton#day[selected="true"] { background: #0f4c5c; color: white; font-weight: 600; }
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
                 "Christian Medical College - Nephrology",
                 "Rajiv Gandhi Government General Hospital, MMC - Chennai"],
    "referring_clinician": ["Dr. Aaron Chapla",
                            "Dr. Suceena Alexander, MD., DM., FRCP(Lon), FASN, PhD.",
                            "Dr. N. Sheik Sulthan Alavudeen MD., DM.,"],
    "specimen": ["DNA", "Peripheral Blood"],
}
OTHER_PLACEHOLDERS = {
    "hospital": "Type the hospital / clinic name",
    "referring_clinician": "Type the clinician name",
    "specimen": "Type the specimen",
}
DATE_FIELDS = {"collection_date", "received_date", "report_date"}


class GeneTableEditor(QWidget):
    """Appendix 1 gene coverage table: paste a whole table to replace it."""
    changed = Signal()

    def __init__(self):
        super().__init__()
        lay = QVBoxLayout(self)
        row = QHBoxLayout()
        paste = QPushButton("Paste table")
        paste.setToolTip("Replace the table with the one on the clipboard (Ctrl+V)")
        paste.clicked.connect(self.paste)
        row.addWidget(paste)
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


class PrepareWorker(QThread):
    """Look up every Excel row's patient off the UI thread. Nothing is saved."""
    progress = Signal(int, int, object)     # done so far, total, bulk.Prepared

    def __init__(self, rows, template_key, reviewer, report_date):
        super().__init__()
        self.rows, self.args = rows, (template_key, reviewer, report_date)

    def run(self):
        # Several patients are looked up at once (the wait is the network);
        # results are still listed in the Excel's order.
        from concurrent.futures import ThreadPoolExecutor
        with ThreadPoolExecutor(max_workers=8) as pool:
            futures = [pool.submit(bulk.prepare, row, *self.args) for row in self.rows]
            for i, fut in enumerate(futures, start=1):
                self.progress.emit(i, len(self.rows), fut.result())


class BulkExportWorker(QThread):
    """Save the chosen reports off the UI thread."""
    progress = Signal(int, str)             # done so far, line for the details
    finished_all = Signal(int, int)         # reports saved, reports failed

    def __init__(self, items, stems, folder, kind):
        super().__init__()
        self.items, self.stems, self.folder, self.kind = items, stems, folder, kind
        self.cancelled = False

    def run(self):
        made = failed = 0
        for i, (item, stem) in enumerate(zip(self.items, self.stems), start=1):
            if self.cancelled:
                break
            label = f"Row {item.row.number} {item.row.pin}"
            try:
                saved = bulk.export(item.data, self.folder, stem, self.kind)
                made += 1
                self.progress.emit(i, f"{label}: saved {saved}")
            except Exception as e:  # noqa: BLE001 - reported per report, the rest go on
                failed += 1
                self.progress.emit(i, f"{label}: FAILED - {type(e).__name__}: {e}")
        self.finished_all.emit(made, failed)


class DropZone(QFrame):
    """Dashed "click to browse or drag & drop" box for the Excel file."""
    chosen = Signal(str)            # dropped file path, or "" to open the file dialog

    EXTS = (".xlsx", ".xlsm")

    def __init__(self, label: str = "Click to browse or drag & drop the Excel file (.xlsx)",
                 tip: str = "Choose the filled Excel file with the patients' details"):
        super().__init__()
        self.setObjectName("dropzone")
        self.setCursor(Qt.PointingHandCursor)
        self.setAcceptDrops(True)
        self.setFixedHeight(40)
        self.setToolTip(tip)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(14, 2, 14, 2)
        lay.setSpacing(12)
        lay.addStretch(1)
        icon = QLabel()
        icon.setPixmap(self._icon(20))
        lay.addWidget(icon)
        text = QLabel(label)
        text.setObjectName("dropText")
        lay.addWidget(text)
        lay.addStretch(1)
        for w in (icon, text):
            w.setAttribute(Qt.WA_TransparentForMouseEvents)

    @staticmethod
    def _icon(size: int) -> QPixmap:
        px = QPixmap(size, size)
        px.fill(Qt.transparent)
        p = QPainter(px)
        p.setRenderHint(QPainter.Antialiasing)
        p.setPen(Qt.NoPen)
        p.setBrush(QColor("#1f8a4c"))
        p.drawRoundedRect(3, 1, size - 6, size - 2, 4, 4)
        p.setPen(QPen(QColor("white"), size / 9, c=Qt.RoundCap))
        a, b = size * 0.34, size * 0.66
        p.drawLine(QPoint(int(a), int(size * .32)), QPoint(int(b), int(size * .68)))
        p.drawLine(QPoint(int(b), int(size * .32)), QPoint(int(a), int(size * .68)))
        p.end()
        return px

    def _hover(self, on: bool):
        if bool(self.property("hover")) == on:
            return                  # already in that state: no restyle
        self.setProperty("hover", on)
        self.style().unpolish(self)
        self.style().polish(self)

    def enterEvent(self, e):
        self._hover(True)

    def leaveEvent(self, e):
        self._hover(False)

    def mouseReleaseEvent(self, e):
        if e.button() == Qt.LeftButton:
            self.chosen.emit("")

    @classmethod
    def _excel_in(cls, mime) -> str:
        """The first Excel file in a drag / drop, '' if there is none."""
        for u in mime.urls():
            path = u.toLocalFile()
            if path.lower().endswith(cls.EXTS):
                return path
        return ""

    def dragEnterEvent(self, e):
        if self._excel_in(e.mimeData()):
            e.acceptProposedAction()
            self._hover(True)
        else:
            e.ignore()

    def dragMoveEvent(self, e):
        # Without this the cursor flickers between "allowed" and "blocked" while moving.
        if self._excel_in(e.mimeData()):
            e.acceptProposedAction()
        else:
            e.ignore()

    def dragLeaveEvent(self, e):
        self._hover(False)

    def dropEvent(self, e):
        self._hover(False)
        path = self._excel_in(e.mimeData())
        if not path:
            e.ignore()
            return
        e.acceptProposedAction()
        # Handle the file after the drop has finished: the handlers can open
        # dialogs or start work, which would freeze Explorer's drag until done.
        QTimer.singleShot(0, lambda: self.chosen.emit(path))


class ExcelDropPanel:
    """Mixin for a panel holding a DropZone: an Excel file can be dropped
    anywhere on the panel, not only on the small dashed box."""

    def _accept_excel_drops(self, zone: DropZone) -> None:
        self._drop_zone = zone
        self.setAcceptDrops(True)

    def dragEnterEvent(self, e):
        self._drop_zone.dragEnterEvent(e)

    def dragMoveEvent(self, e):
        self._drop_zone.dragMoveEvent(e)

    def dragLeaveEvent(self, e):
        self._drop_zone.dragLeaveEvent(e)

    def dropEvent(self, e):
        self._drop_zone.dropEvent(e)


class VariantsPanel(ExcelDropPanel, QWidget):
    """Variants of an uploaded listed-variants Excel: tick the ones the report
    should show (CMC - Pathogenic template)."""
    changed = Signal()
    upload_requested = Signal(str)  # the Excel file dropped on the box, or "" to browse
    CHECK_NC_CHROMOSOME = True         # refuse an image's NC id of another chromosome
    RESULT_COLORS = {"pathogenic": "#d00000", "likely": "#d00000",    # as in the report
                     "vus_related": "#0070C0"}
    RESULT_CLASS = {"pathogenic": "Pathogenic", "likely": "Likely Pathogenic",   # classification each
                    "vus_related": "Variant of Uncertain Significance"}      # result tag sets
    LIST_COLS = ("Gene", "Variant")      # the picker: every variant of the Excel
    DETAIL_COLS = ("Gene", "Classification", "Disease or Disorder", "Criteria",
                   "Variant Frequency")  # the typed-in details of the ticked variants
    TYPED = {2: "disease", 3: "criteria", 4: "frequency"}      # detail columns filled in by the user
    TYPED_TIPS = {2: "Type the disease with its braces, e.g. {Breast-ovarian cancer, familial, 2}. The (OMIM#id) line is added from the Excel",
                  3: "Type the criteria, e.g. PVS1, PM2_Supp",
                  4: "Type the number before the sentence, e.g. 0.012; the wording can be edited too"}
    SEARCH_KEYS = ("gene", "variant", "protein", "zygosity", "exon", "inheritance",
                   "classification", "disease", "criteria", "frequency")

    def __init__(self):
        super().__init__()
        self._variants: list[dict] = []
        self._current = -1                    # variant index shown in the detail form
        self.secondary_separate = False       # the template has its own Secondary findings set
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 12, 0, 0)
        lay.setSpacing(10)
        zone = DropZone("Click to browse or drag & drop the listed-variants Excel (.xlsx)",
                        "Choose the VarSeq listed-variants Excel for this patient")
        zone.chosen.connect(self.upload_requested)
        self._accept_excel_drops(zone)
        top = QHBoxLayout()                 # the upload box and the search box side by side
        top.addWidget(zone, 3)
        self.search = QLineEdit(placeholderText="Search variants…")
        self.search.setClearButtonEnabled(True)
        self.search.textChanged.connect(self._filter)
        top.addWidget(self.search, 2, Qt.AlignVCenter)
        lay.addLayout(top)
        self.status = QLabel("Upload the listed-variants Excel, then tick the variants "
                             "to show in the report. With none ticked the template's "
                             "sample variant is kept.")
        self.status.setObjectName("hint")
        self.status.setWordWrap(True)
        lay.addWidget(self.status)
        self.result_row = QWidget()         # the report's result tag (boxed line under Results)
        rr = QHBoxLayout(self.result_row)
        rr.setContentsMargins(0, 0, 0, 0)
        rr.addWidget(QLabel("Result:"))
        self.f_result = StableCombo()
        self.f_result.addItem("Automatic (from the classifications)", "")
        for key, text in variants_core.RESULT_TAGS:
            self.f_result.addItem(text, key)
            self.f_result.setItemData(self.f_result.count() - 1,
                                      QColor(self.RESULT_COLORS[key]), Qt.ForegroundRole)
        self.f_result.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.f_result.setMinimumContentsLength(12)
        self.f_result.currentIndexChanged.connect(self._result_changed)
        rr.addWidget(self.f_result, 1)
        lay.addWidget(self.result_row)
        self.select_all = QCheckBox("Select all")      # sits on top of the variant list
        self.select_all.setObjectName("selectall")
        self.select_all.setCursor(Qt.PointingHandCursor)
        self.select_all.clicked.connect(self._select_all)

        # Left: the variants to tick. Right: the details of the highlighted one.
        self.body = QSplitter(Qt.Horizontal)
        self.table = QListWidget()
        self.table.setObjectName("variantlist")
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.table.setTextElideMode(Qt.ElideRight)
        self.table.setMinimumWidth(220)
        self.table.itemChanged.connect(self._ticked)
        self.table.currentRowChanged.connect(self._show_detail)
        self.list_box = QWidget()           # the variant list with its collapse symbol
        ll = QVBoxLayout(self.list_box)
        ll.setContentsMargins(0, 0, 0, 0)
        ll.setSpacing(0)
        head_bar = QFrame()                 # header of the list: Select all + collapse symbol
        head_bar.setObjectName("listhead")
        head = QHBoxLayout(head_bar)
        head.setContentsMargins(8, 2, 4, 2)
        head.addWidget(self.select_all)
        head.addStretch(1)
        self.hide_list = self._symbol_button("left", "Hide the variant list (more room for the details)")
        head.addWidget(self.hide_list)
        ll.addWidget(head_bar)
        ll.addWidget(self.table, 1)
        self.body.addWidget(self.list_box)

        self.detail = QWidget()
        dl = QFormLayout(self.detail)
        dl.setVerticalSpacing(10)
        dl.setHorizontalSpacing(12)
        dl.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        dl.setRowWrapPolicy(QFormLayout.WrapAllRows)     # label above, the box gets the full width
        self.detail_title = QLabel()        # stays above the scrolling form: which variant is open
        self.detail_title.setObjectName("hint")
        self.detail_title.setWordWrap(True)
        self.f_gene = QLineEdit()
        self.f_gene.setReadOnly(True)
        self.f_variant = QLineEdit()
        self.f_variant.setReadOnly(True)
        self.f_class = StableCombo()
        self.f_class.setToolTip("Classification shown in the report: Franklin / ACMG / ClinVar")
        self.f_class.currentTextChanged.connect(self._classified)
        self.f_typed: dict[str, QLineEdit] = {}
        def section(title: str):
            head = QLabel(title)
            head.setObjectName("section")
            dl.addRow(head)

        section("Variant Interpretation")
        dl.addRow("Gene:", self.f_gene)
        dl.addRow("Variant:", self.f_variant)
        dl.addRow("Classification:", self.f_class)
        omim_tip = ("OMIM number(s) shown as (OMIM#id); the Excel lists every phenotype of "
                    "the gene; delete the ones that do not apply")
        inh_tip = "e.g. AD, AR (duplicates are merged)"

        def typed_row(key: str, label: str, tip: str):
            edit = QLineEdit()
            edit.setToolTip(tip)
            edit.textEdited.connect(lambda text, k=key: self._typed(k, text))
            self.f_typed[key] = edit
            if key == "genomic_position":       # + a backup: read the NC id from an image
                row = QWidget()
                rl2 = QHBoxLayout(row)
                rl2.setContentsMargins(0, 0, 0, 0)
                rl2.addWidget(edit, 1)
                btn = QPushButton("From image…")
                btn.setCursor(Qt.PointingHandCursor)
                btn.setToolTip("Pick a screenshot / photo that shows the NC accession "
                               "(e.g. NC_000013.11); it replaces the one built from the Excel")
                btn.clicked.connect(self._nc_from_image)
                rl2.addWidget(btn)
                dl.addRow(row)                  # the section heading already says what it is
                return
            if key != "frequency":
                dl.addRow(label, edit)
                return
            combo = StableCombo()           # the sentence to start from; still editable
            combo.setEditable(True)
            combo.setLineEdit(edit)
            combo.setInsertPolicy(QComboBox.NoInsert)
            combo.addItems([variants_core.FREQUENCY_TEXT, variants_core.FREQUENCY_ALL,
                            variants_core.FREQUENCY_ABSENT])
            combo.textActivated.connect(lambda text: (self._typed("frequency", text),
                                                      edit.setCursorPosition(0)))
            row = QWidget()                 # + a button that looks the number up on gnomAD
            rl2 = QHBoxLayout(row)
            rl2.setContentsMargins(0, 0, 0, 0)
            rl2.addWidget(combo, 1)
            btn = QPushButton("From gnomAD")
            btn.setCursor(Qt.PointingHandCursor)
            btn.setToolTip("Look the variant up on the gnomAD website (v4 exomes + genomes) and fill in "
                           "the South Asian frequency = allele count / allele number x 100")
            btn.clicked.connect(self._frequency_from_gnomad)
            rl2.addWidget(btn)
            dl.addRow(label, row)

        typed_row("omim_id", "OMIM ID:", omim_tip)          # above the disease it fills
        for c, key in self.TYPED.items():
            typed_row(key, self.DETAIL_COLS[c] + ":", self.TYPED_TIPS[c])
        typed_row("inheritance", "Inheritance:", inh_tip)
        self.detail_form = dl
        self.f_silico: dict[str, QLineEdit] = {}       # shown by the variant's Sequence Ontology
        for title, group in (("In Silico Predictions", variants_core.SILICO),
                             ("Conservation Scores", variants_core.CONSERVATION)):
            section(title)
            for label, key in group:
                edit = QLineEdit()
                case = (variants_core.silico_case if group is variants_core.SILICO
                        else str.upper)             # Da / Ds / Supp, but conservation in capitals
                edit.setToolTip(f"{label} value for the report, e.g. D / T / a score")
                edit.textEdited.connect(lambda text, e=edit, c=case: self._upper(e, text, c))
                edit.textEdited.connect(lambda text, k=key, c=case: self._typed(k, c(text)))
                self.f_typed[key] = edit
                dl.addRow(f"{label}:", edit)
                if group is variants_core.SILICO:
                    self.f_silico[key] = edit
        section("Genomic Position")
        typed_row("genomic_position", "Genomic Position:",
                  "Built from the Excel (chromosome, position, ref/alt, GRCh38 NC accession) with the "
                  "transcript in brackets at the end, which the report puts on its own line; "
                  "edit it to replace it, e.g. for an insertion that is a duplication")
        section("Evidence Links")
        for key, label in (("clinvar_link", "ClinVar link:"), ("pubmed_link", "PubMed link:")):
            edit = QLineEdit()
            edit.setPlaceholderText("Paste the link (or just the number)")
            edit.setToolTip("The report shows the evidence number as a hyperlink to this link")
            edit.textEdited.connect(lambda text, k=key: self._typed(k, text))
            self.f_typed[key] = edit
            dl.addRow(label, edit)
        self.f_notes: dict[str, QPlainTextEdit] = {}      # the report's free-text paragraphs
        section("Narrative")
        for key, label in (("evidence", "Clinical and Literature Evidences:"),
                           ("omim_phenotype", "OMIM phenotype:")):
            box = NarrativeEdit()
            box.setFixedHeight(90)
            box.setTabChangesFocus(True)
            box.textChanged.connect(lambda k=key, b=box: self._typed(k, b.toPlainText()))
            self.f_notes[key] = box
            dl.addRow(label, box)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setWidget(self.detail)
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.setSpacing(6)
        title_row = QHBoxLayout()
        self.show_list = self._symbol_button("right", "Show the variant list")
        self.show_list.hide()
        title_row.addWidget(self.show_list)
        title_row.addWidget(self.detail_title, 1)
        rl.addLayout(title_row)
        rl.addWidget(scroll, 1)
        self.body.addWidget(right)
        self._detail_box = right
        self.body.setStretchFactor(0, 2)
        self.body.setStretchFactor(1, 3)
        lay.addWidget(self.body, 1)
        self._detail_scroll = scroll
        self._show_tools(False)

    def _show_tools(self, on: bool):
        for w in (self.search, self.result_row, self.body):
            w.setVisible(on)
        if on:
            self._toggle_list(shown=True)

    def _symbol_button(self, direction: str, tip: str) -> QPushButton:
        """A small » / « button; the one that brings the list back is solid, so it
        stands out beside the form."""
        btn = QPushButton()
        shown = direction == "right"
        btn.setIcon(style.chevrons_icon(direction, "white" if shown else style.ACCENT))
        btn.setIconSize(QSize(18, 18))
        btn.setFixedSize(34, 30)
        if shown:
            btn.setProperty("solid", True)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setToolTip(tip)
        btn.clicked.connect(lambda *_: self._toggle_list())
        return btn

    def _toggle_list(self, *_, shown: bool | None = None):
        """Hide / show the variant list (the picker) next to the details form."""
        if shown is None:
            shown = not self.list_box.isVisible()
        for w in (self.list_box, self.search):
            w.setVisible(shown)
        self.show_list.setVisible(not shown)

    def set_variants(self, variants: list[dict]):
        """List these variants (each carries its own "selected" flag)."""
        self._variants = [dict(v) for v in variants]
        self.table.blockSignals(True)
        self.table.clear()
        for v in self._variants:
            it = QListWidgetItem(f"{v.get('gene', '')}  {v.get('variant', '')}")
            it.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
            it.setCheckState(Qt.Checked if v.get("selected") else Qt.Unchecked)
            self.table.addItem(it)
        self.table.blockSignals(False)
        self.search.clear()
        self._show_tools(bool(self._variants))
        self.status.setVisible(not self._variants)   # the hint is only for the empty tab
        self._count()
        self.table.setCurrentRow(0 if self._variants else -1)
        self._show_detail(self.table.currentRow())

    def _show_detail(self, idx: int):
        """Fill the form with the highlighted variant's details."""
        self._current = idx if 0 <= idx < len(self._variants) else -1
        self._detail_box.setVisible(self._current >= 0)
        if self._current < 0:
            return
        v = self._variants[self._current]
        self.detail_title.setText(
            f"<b>{v.get('gene', '')} {v.get('variant', '')}</b> &mdash; "
            + ("ticked: its details go in the report." if v.get("selected")
               else "not ticked: tick it in the list to include it in the report."))
        self.f_gene.setText(v.get("gene", ""))
        self.f_variant.setText(v.get("variant", ""))
        options = variants_core.classification_options(v) or [""]
        if v.get("classification") and v["classification"] not in options:
            options.append(v["classification"])
        self.f_class.blockSignals(True)
        self.f_class.clear()
        self.f_class.addItems(options)
        self.f_class.setCurrentText(v.get("classification", options[0]))
        self.f_class.blockSignals(False)
        for key, edit in self.f_typed.items():
            edit.setText(v.get(key, ""))
        freq = self.f_typed["frequency"]
        if not v.get("frequency"):      # the sentence for this variant; only the number is typed
            freq.setText(variants_core.default_frequency(v))
            freq.setCursorPosition(0)
        freq.setToolTip(f"{self.TYPED_TIPS[4]}. Excel counts: South Asian "
                        f"{v.get('gnomad_sa') or '-'}, all {v.get('gnomad_all') or '-'}")
        for key, box in self.f_notes.items():
            box.blockSignals(True)
            text = v.get(key, "")
            if key == "omim_phenotype" and text and text == v.get("omim_phenotype_prefix"):
                text += " "                 # so the typed sentence continues after it
            box.setPlainText(text)
            box.blockSignals(False)
        shown = {key for _, key in variants_core.insilico_tools(v)}
        for key, edit in self.f_silico.items():
            self.detail_form.setRowVisible(edit, key in shown)
        if not v.get("genomic_position"):       # show the built value so it can be edited
            self.f_typed["genomic_position"].setText(variants_core.genomic_text(v))
        opts = v.get("omim_options", "")
        self.f_typed["omim_id"].setPlaceholderText(f"Sheet has: {opts}" if opts else "")

    def variants(self) -> list[dict]:
        return [dict(v) for v in self._variants]

    def _nc_from_image(self):
        """Backup for the Genomic Position: the NC accession read from an image
        replaces the built-in one. It must belong to the variant's chromosome."""
        if not 0 <= self._current < len(self._variants):
            return
        v = self._variants[self._current]
        path, _ = QFileDialog.getOpenFileName(
            self, "Image showing the NC accession", "",
            "Images (*.png *.jpg *.jpeg *.bmp *.tif *.tiff)")
        if not path:
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            nc = ocr.nc_from_image(path, variants_core.expected_nc_number(v))
        except Exception as e:  # noqa: BLE001 - shown to the user
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(self, "Genomic Position", f"Could not read the image.\n\n{e}")
            return
        QApplication.restoreOverrideCursor()
        if not nc:
            QMessageBox.information(
                self, "Genomic Position",
                "No NC accession (like NC_000013.11) was found in the image.\n\n"
                "Use a clearer or larger picture of the part of the screen that shows it, "
                "or type the Genomic Position yourself.")
            return
        want = variants_core.expected_nc_number(v)
        if self.CHECK_NC_CHROMOSOME and want and int(nc[3:9]) != want:
            QMessageBox.warning(
                self, "Genomic Position",
                f"The image shows {nc}, which is not the chromosome of {v.get('gene', '')} "
                f"{v.get('variant', '')} (chromosome {v.get('chrom_pos', '').partition(':')[0]}). "
                "It was not used - check the image.")
            return
        text = variants_core.genomic_text(v, nc)
        v["genomic_position"] = text
        self.f_typed["genomic_position"].setText(text)
        self.changed.emit()

    def _frequency_from_gnomad(self):
        """Fill the Variant Frequency of the shown variant from the gnomAD website."""
        if not 0 <= self._current < len(self._variants):
            return
        v = self._variants[self._current]
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            text = gnomad.frequency_text(v)
        except gnomad.GnomadError as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(self, "Variant Frequency", str(e))
            return
        QApplication.restoreOverrideCursor()
        v["frequency"] = text
        freq = self.f_typed["frequency"]
        freq.setText(text)
        freq.setCursorPosition(0)
        self.changed.emit()

    @staticmethod
    def _upper(edit: QLineEdit, text: str, case=str.upper):
        """The value's capitals as the user types (D, Da, C ...), keeping the cursor where it is."""
        if text != case(text):
            pos = edit.cursorPosition()
            edit.setText(case(text))
            edit.setCursorPosition(pos)

    RESULT_TINTS = {"pathogenic": "#fdeeee", "likely": "#fdeeee", "vus_related": "#eaf4fc"}

    def _result_changed_style(self):
        """The box looks like the report's result tag: the tag's colour for the text
        and a 2 px frame, on a light tint of it."""
        key = self.result_tag()
        color = self.RESULT_COLORS.get(key)
        self.f_result.setStyleSheet(
            f"QComboBox {{ color: {color}; font-weight: 700; background: {self.RESULT_TINTS[key]};"
            f" border: 2px solid {color}; border-radius: 8px; padding: 6px 30px 6px 12px; }}"
            f"QComboBox:hover, QComboBox:focus {{ border: 2px solid {color}; padding: 6px 30px 6px 12px; }}"
            if color else "")

    def _result_changed(self, *_):
        """Picking a result also sets the classification of the ticked variants
        (the open one when none is ticked) to match it."""
        self._result_changed_style()
        label = self.RESULT_CLASS.get(self.result_tag())
        if label:
            ticked = [v for v in self._variants if v.get("selected")]
            targets = ticked or ([self._variants[self._current]]
                                 if 0 <= self._current < len(self._variants) else [])
            if self.secondary_separate:     # Secondary findings keep their own classification
                targets = [v for v in targets if not variants_core.is_secondary(v)]
            for v in targets:
                v["classification"] = label
            if 0 <= self._current < len(self._variants) and                     self._variants[self._current] in targets:
                self.f_class.blockSignals(True)
                if self.f_class.findText(label) < 0:
                    self.f_class.addItem(label)
                self.f_class.setCurrentText(label)
                self.f_class.blockSignals(False)
        self.changed.emit()

    def result_tag(self) -> str:
        return self.f_result.currentData() or ""

    def set_result_tag(self, key: str):
        self.f_result.blockSignals(True)
        self.f_result.setCurrentIndex(max(0, self.f_result.findData(key or "")))
        self.f_result.blockSignals(False)
        self._result_changed_style()

    def _classified(self, text: str):
        if 0 <= self._current < len(self._variants):
            self._variants[self._current]["classification"] = text
            self.changed.emit()

    def _ticked(self, item):
        r = self.table.row(item)
        if r >= len(self._variants):
            return
        self._variants[r]["selected"] = item.checkState() == Qt.Checked
        self._count()
        if r == self._current:
            self._show_detail(r)
        self.changed.emit()

    def _typed(self, key: str, text: str):
        if 0 <= self._current < len(self._variants):
            v = self._variants[self._current]
            v[key] = text.strip()
            if key == "disease":        # typed by hand: stop auto-filling (until cleared)
                v["disease_manual"] = bool(text.strip())
                self._sync_omim_to_disease()
            elif key == "inheritance":
                v["inheritance_manual"] = bool(text.strip())
            elif key == "omim_id":
                self._refresh_disease()
            if key == "evidence":       # typed by hand: stop auto-filling (until cleared)
                v["evidence_manual"] = bool(text.strip())
            elif key in ("clinvar_link", "pubmed_link"):
                self._refresh_evidence()
            if key in ("disease", "omim_id"):
                self._refresh_omim_phenotype()
            self.changed.emit()

    def _refresh_omim_phenotype(self):
        """The OMIM phenotype paragraph starts with '{Disease} (OMIM#id)'. When the
        disease or OMIM number changes, that start is replaced and the sentence the
        user typed after it is kept."""
        v = self._variants[self._current]
        new, old = variants_core.omim_phenotype_prefix(v), v.get("omim_phenotype_prefix", "")
        cur = v.get("omim_phenotype", "")
        if cur and not (old and cur.startswith(old)):
            return                          # the user wrote the paragraph themselves
        text = new + (cur[len(old):] if cur else "")
        v["omim_phenotype"], v["omim_phenotype_prefix"] = text, new
        box = self.f_notes["omim_phenotype"]
        box.blockSignals(True)
        box.setPlainText(text + (" " if text and text == new else ""))
        box.blockSignals(False)

    def _sync_omim_to_disease(self):
        """Diseases deleted from the Disease box: keep only the OMIM numbers (and their
        inheritance) of the ones that remain."""
        v = self._variants[self._current]
        kept = variants_core.omim_ids_in(v, v["disease"])
        if not kept:
            return
        v["omim_id"] = ", ".join(kept)
        self.f_typed["omim_id"].setText(v["omim_id"])
        inh = variants_core.suggested_inheritance(v)
        if inh and not v.get("inheritance_manual"):
            v["inheritance"] = inh
            self.f_typed["inheritance"].setText(inh)

    def _refresh_disease(self):
        """Fill Disease or Disorder and Inheritance from the chosen OMIM number(s),
        unless the user has typed their own."""
        v = self._variants[self._current]
        inh = variants_core.suggested_inheritance(v)
        if inh and not v.get("inheritance_manual"):
            v["inheritance"] = inh
            self.f_typed["inheritance"].setText(inh)
        if v.get("disease_manual"):
            return
        v["disease"] = variants_core.suggested_disease(v)
        self.f_typed["disease"].setText(v["disease"])
        if not v.get("evidence_manual") and v.get("pubmed_link"):
            self._refresh_evidence()

    def _refresh_evidence(self):
        """Fill the evidence paragraph from the ClinVar / PubMed links, unless the
        user has written their own text."""
        v = self._variants[self._current]
        if v.get("evidence_manual"):
            return
        v["evidence"] = variants_core.suggested_evidence(v)
        box = self.f_notes["evidence"]
        box.blockSignals(True)
        box.setPlainText(v["evidence"])
        box.blockSignals(False)

    def _count(self):
        n = sum(1 for v in self._variants if v.get("selected"))
        total = len(self._variants)                 # the count next to Select all
        listed = sum(1 for r in range(self.table.count()) if not self.table.item(r).isHidden())
        self.select_all.setText(f"Select all ({listed} of {total})" if listed != total
                                else f"Select all ({total})" if total else "Select all")
        self.select_all.blockSignals(True)
        self.select_all.setChecked(bool(self._variants) and n == len(self._variants))
        self.select_all.blockSignals(False)

    def _select_all(self, state: bool):
        """Tick / clear every variant that is currently listed (the search hides the rest)."""
        self.table.blockSignals(True)
        for r, v in enumerate(self._variants):
            if not self.table.item(r).isHidden():
                v["selected"] = state
                self.table.item(r).setCheckState(Qt.Checked if state else Qt.Unchecked)
        self.table.blockSignals(False)
        self._count()
        self._show_detail(self._current)
        self.changed.emit()


    def _filter(self, text: str):
        text = text.strip().lower()
        for r, v in enumerate(self._variants):
            hay = " ".join(v.get(k, "") for k in self.SEARCH_KEYS).lower()
            self.table.item(r).setHidden(bool(text) and text not in hay)
        self._count()


class BulkPanel(ExcelDropPanel, QWidget):
    """The patients of an uploaded Excel file, looked up and listed inside the
    main window. Nothing is saved until the user ticks the ones they want and
    uses the window's Export buttons; clicking a row previews that report."""
    COLS = ("Patients",)
    shown = Signal(object)          # ReportData of the clicked row, or None
    edited = Signal()               # the patient editor changed the shown report
    upload = Signal(str)            # the Excel file dropped on the box, or "" to browse

    def __init__(self):
        super().__init__()
        self.items: list = []
        self.worker: PrepareWorker | None = None
        self.export_worker: BulkExportWorker | None = None
        self._export_log: list[str] = []
        self._total = 0
        self._built = None          # (template key, reviewer key) the list was built with
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 12, 0, 0)
        lay.setSpacing(10)

        zone = DropZone()
        zone.chosen.connect(self.upload)
        self._accept_excel_drops(zone)
        lay.addWidget(zone)
        self.search = QLineEdit(placeholderText="Search patients…")
        self.search.setClearButtonEnabled(True)
        self.search.textChanged.connect(self._filter)
        left = QWidget()
        ll = QVBoxLayout(left)
        ll.setContentsMargins(0, 0, 0, 0)
        ll.setSpacing(10)
        ll.addWidget(self.search)

        bar = QHBoxLayout()
        self.select_all = QCheckBox("Select all")
        self.select_all.setCursor(Qt.PointingHandCursor)
        self.select_all.setEnabled(False)
        self.select_all.clicked.connect(self._select_all)     # ticked = all, cleared = none
        bar.addSpacing(8)
        bar.addWidget(self.select_all)
        bar.addStretch(1)
        self.count_label = QLabel("0 selected")
        self.count_label.setObjectName("chip")
        bar.addWidget(self.count_label)
        ll.addLayout(bar)
        self._show_tools(False)             # only once a file has been uploaded

        self.info = QLabel()                # only shown for the stale-template warning
        self.info.setWordWrap(True)
        self.info.setVisible(False)
        lay.addWidget(self.info)
        self.status = QLabel()
        self.status.setObjectName("hint")
        self.status.setVisible(False)
        lay.addWidget(self.status)

        self.table = QTableWidget(0, len(self.COLS))
        self.table.horizontalHeader().setVisible(False)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(38)
        self.table.setShowGrid(False)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.setTextElideMode(Qt.ElideRight)
        self.table.itemChanged.connect(lambda *_: self._count())
        self.table.itemSelectionChanged.connect(self._clicked)

        empty = QLabel("<div style='text-align:center'><p style='font-size:13pt;font-weight:600'>"
                       "No patients yet</p></div>")
        empty.setAlignment(Qt.AlignCenter)
        empty.setWordWrap(True)
        empty.setObjectName("hint")
        self.body = QStackedWidget()
        self.body.addWidget(empty)
        self.body.addWidget(self.table)
        ll.addWidget(self.body, 1)

        # Patient editor: the clicked patient's details, edited in place.
        self.editor_fields: dict[str, QLineEdit | OtherChoiceBox] = {}
        form = QWidget()
        fl = QFormLayout(form)
        fl.setVerticalSpacing(8)
        fl.setHorizontalSpacing(12)
        title = QLabel("Patient editor")
        title.setObjectName("caption")
        fl.addRow(title)
        for key, label in (("patient_id", "Patient ID"), ("pin", "PIN"), ("age", "Age"),
                           ("gender", "Gender"), ("hospital", "Hospital/Clinic"),
                           ("sample_number", "Sample Number"),
                           ("referring_clinician", "Referring Clinician"),
                           ("specimen", "Specimen"), ("collection_date", "Collection Date"),
                           ("received_date", "Received Date"), ("report_date", "Report Date")):
            if key in ("hospital", "referring_clinician", "specimen"):   # dropdown + "Other", as in Manual entry
                le = OtherChoiceBox(PATIENT_CHOICES[key], OTHER_PLACEHOLDERS[key])
            else:
                le = QLineEdit()
            le.textChanged.connect(self._field_edited)
            self.editor_fields[key] = le
            fl.addRow(label, le)
        self.ed_history = QPlainTextEdit()
        self.ed_history.setMinimumHeight(90)
        self.ed_history.setTabChangesFocus(True)
        self.ed_history.textChanged.connect(self._field_edited)
        fl.addRow("Clinical History", self.ed_history)
        self.ed_reads, self.ed_q30, self.ed_url = QLineEdit(), QLineEdit(), QLineEdit()
        for le in (self.ed_reads, self.ed_q30, self.ed_url):
            le.textChanged.connect(self._field_edited)
        fl.addRow("Total Read Generated", self.ed_reads)
        fl.addRow("Data ≥ Q30", self.ed_q30)
        fl.addRow("Gene list link", self.ed_url)
        self._editing = None            # the ReportData being edited
        self.editor = QScrollArea()
        self.editor.setWidgetResizable(True)
        self.editor.setFrameShape(QScrollArea.NoFrame)
        self.editor.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.editor.setWidget(form)
        self.editor.setVisible(False)

        split = QSplitter(Qt.Horizontal)
        split.addWidget(left)
        split.addWidget(self.editor)
        split.setStretchFactor(0, 1)
        split.setStretchFactor(1, 2)
        split.setChildrenCollapsible(False)
        lay.addWidget(split, 1)


    def _show_tools(self, on: bool):
        for w in (self.search, self.select_all, self.count_label,
                  getattr(self, "editor", None)):
            if w is not None:
                w.setVisible(on)

    def check_current(self, template_key: str, reviewer: str):
        """Warn when the template or reviewer chosen at the top is no longer
        the one the list was built with."""
        if self._built is None:
            return
        stale = self._built != (template_key, reviewer)
        self.info.setText("<span style='color:#b35c00'>The template or reviewer has "
                          "changed - upload the Excel again to use it.</span>")
        self.info.setVisible(stale)

    @property
    def busy(self) -> bool:
        return any(w is not None and w.isRunning() for w in (self.worker, self.export_worker))

    # -- listing -----------------------------------------------------------
    def load(self, rows, template_key, reviewer, reviewer_name, report_date):
        self._built = (template_key, reviewer)
        self.info.setVisible(False)
        self.items = []
        self.table.setRowCount(0)
        self._total = len(rows)
        self.count_label.setText("0 selected")
        self.select_all.setChecked(False)
        self._show_tools(True)
        self.body.setCurrentWidget(self.table)
        self.status.setVisible(True)
        self.status.setText(f"Looking up the patients… 0 of {self._total}")
        self.worker = PrepareWorker(rows, template_key, reviewer, report_date)
        self.worker.progress.connect(self._added)
        self.worker.start()

    def restore(self, datas, template_key, reviewer):
        """List reports loaded from a draft file (already looked up)."""
        self._built = (template_key, reviewer)
        self.info.setVisible(False)
        self.items = []
        self.table.setRowCount(0)
        self._total = len(datas)
        self.select_all.setChecked(False)
        self._show_tools(True)
        self.body.setCurrentWidget(self.table)
        for i, d in enumerate(datas, start=1):
            self._added(i, len(datas), bulk.Prepared(bulk.BulkRow(i, d.patient.pin), d))

    def set_reviewer(self, key: str):
        """Sign every listed report with this reviewer."""
        for item in self.items:
            if item.data:
                item.data.reviewer = key
        if self._built is not None:
            self._built = (self._built[0], key)
        self.info.setVisible(False)

    def draft_reports(self) -> list:
        """The ticked reports, or every report that looked up fine if none is ticked."""
        rows = self._checked() or [r for r, i in enumerate(self.items) if i.data]
        return [self.items[r].data for r in rows]

    def _added(self, done: int, total: int, item):
        self.items.append(item)
        self.table.blockSignals(True)
        r = self.table.rowCount()
        self.table.insertRow(r)
        p = item.data.patient if item.data else None
        name = (p.patient_id if p and p.patient_id else "") or item.row.pin or f"Row {item.row.number}"
        it = QTableWidgetItem(name)
        flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if item.data:
            it.setFlags(flags | Qt.ItemIsUserCheckable)
            it.setCheckState(Qt.Unchecked)
            if item.notes:
                it.setForeground(QColor(style.WARN))
        else:
            it.setFlags(flags)
            it.setForeground(QColor(style.BAD))
            name = f"{name} - {item.error}"
            it.setText(name)
        it.setToolTip(name if item.data else item.error or "")
        if item.data and item.notes:
            it.setToolTip("; ".join(item.notes))
        self.table.setItem(r, 0, it)
        self.table.blockSignals(False)
        if done < total:
            self.status.setText(f"Looking up the patients… {done} of {total}")
            return
        self.status.setVisible(False)   # the Status column says what needs attention
        first = next((r for r, i in enumerate(self.items) if i.data), None)
        if first is not None:
            self.table.selectRow(first)
        self._count()

    def _clicked(self):
        rows = self.table.selectionModel().selectedRows()
        item = self.items[rows[0].row()] if rows and rows[0].row() < len(self.items) else None
        if item is not None and item.data:
            self._load_editor(item.data)
            self.shown.emit(item.data)

    def _load_editor(self, data):
        self._editing = None            # filling the boxes must not count as an edit
        for key, le in self.editor_fields.items():
            le.setText(getattr(data.patient, key, ""))
        self.ed_history.setPlainText(data.clinical_history)
        self.ed_reads.setText(data.total_reads)
        self.ed_q30.setText(data.q30)
        self.ed_url.setText(data.gene_list_url)
        self._editing = data

    def _field_edited(self, *_):
        data = self._editing
        if data is None:
            return
        for key, le in self.editor_fields.items():
            setattr(data.patient, key, le.text().strip())
        data.clinical_history = self.ed_history.toPlainText().strip()
        data.total_reads = self.ed_reads.text().strip()
        data.q30 = self.ed_q30.text().strip()
        data.gene_list_url = self.ed_url.text().strip()
        item = next((i for i in self.items if i.data is data), None)
        if item is not None:
            r = self.items.index(item)
            name = data.patient.patient_id or data.patient.pin or f"Row {item.row.number}"
            self.table.item(r, 0).setText(name)
        self.edited.emit()

    def _checked(self) -> list[int]:
        return [r for r in range(self.table.rowCount())
                if self.table.item(r, 0).checkState() == Qt.Checked]

    def _filter(self, text: str):
        text = text.strip().lower()
        for r in range(self.table.rowCount()):
            self.table.setRowHidden(r, bool(text) and text not in self.table.item(r, 0).text().lower())

    def _select_all(self, state: bool):
        self.table.blockSignals(True)
        for r, item in enumerate(self.items[:self.table.rowCount()]):
            if item.data:
                self.table.item(r, 0).setCheckState(Qt.Checked if state else Qt.Unchecked)
        self.table.blockSignals(False)
        self._count()

    def _count(self):
        checked = len(self._checked())
        self.count_label.setText(f"{checked} selected")
        usable = sum(1 for i in self.items[:self.table.rowCount()] if i.data)
        self.select_all.setEnabled(usable > 0)
        self.select_all.blockSignals(True)
        self.select_all.setChecked(usable > 0 and checked == usable)
        self.select_all.blockSignals(False)

    # -- export ------------------------------------------------------------
    def export(self, kind: str, folder_edit: QLineEdit):
        """Save the ticked reports into the export folder."""
        win = self.window()
        if self.busy:
            return
        if len(self.items) < self._total:
            QMessageBox.information(win, "Bulk upload", "Still looking up the patients.")
            return
        chosen = [self.items[r] for r in self._checked()]
        if not chosen:
            QMessageBox.information(win, "Bulk upload",
                                    "Tick the patients whose reports you want to export.")
            return
        folder = folder_edit.text()
        if not os.path.isdir(folder):
            folder = QFileDialog.getExistingDirectory(win, "Choose export folder", folder)
            if not folder:
                return
            folder = os.path.normpath(folder)
            folder_edit.setText(folder)
        stems = bulk.stems(chosen)
        exts = ["docx", "pdf"] if kind == "both" else [kind]
        existing = [f"{s}.{e}" for s in stems for e in exts
                    if os.path.exists(os.path.join(folder, f"{s}.{e}"))]
        if existing and QMessageBox.question(
                win, "Files exist", "Already in the folder:\n" + "\n".join(existing[:15])
                + ("\n…" if len(existing) > 15 else "") + "\n\nReplace?") != QMessageBox.Yes:
            return
        self._export_log = []
        self._folder = folder
        self._progress = QProgressDialog("Saving reports…", "Stop", 0, len(chosen), win)
        self._progress.setWindowTitle("Bulk upload")
        self._progress.setWindowModality(Qt.WindowModal)
        self._progress.setMinimumDuration(0)
        self._progress.setValue(0)
        self.export_worker = BulkExportWorker(chosen, stems, folder, kind)
        self._progress.canceled.connect(lambda: setattr(self.export_worker, "cancelled", True))
        self.export_worker.progress.connect(self._export_step)
        self.export_worker.finished_all.connect(self._export_done)
        self.export_worker.start()

    def _export_step(self, done: int, line: str):
        self._progress.setValue(done)
        self._progress.setLabelText(f"Saving reports… {done} of {self._progress.maximum()}")
        self._export_log.append(line)

    def _export_done(self, made: int, failed: int):
        self._progress.close()
        box = QMessageBox(QMessageBox.Warning if failed else QMessageBox.Information,
                          "Bulk upload", f"{made} report(s) saved"
                          f"{f', {failed} failed' if failed else ''}.\n\n{self._folder}",
                          QMessageBox.Ok, self.window())
        box.setDetailedText("\n".join(self._export_log))
        box.exec()

    def shutdown(self):
        if self.export_worker is not None:
            self.export_worker.cancelled = True
        for w in (self.worker, self.export_worker):
            if w is not None:
                w.wait(60000)


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
        self._preview_snapshot: dict | None = None
        self._preview_seq = 0
        self._preview_worker: RenderWorker | None = None
        self._export_worker: RenderWorker | None = None
        self._bulk_data: ReportData | None = None   # the bulk row being previewed
        self._pending = False            # an edit arrived while a preview was rendering

        # Short delay so several keystrokes share one render; not restarted by
        # further typing, so the preview keeps updating while you type.
        self._preview_timer = QTimer(self, singleShot=True, interval=150)
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
    @staticmethod
    def _caption(text: str) -> QLabel:
        label = QLabel(text)
        label.setObjectName("caption")
        return label

    @staticmethod
    def _button(text: str, primary: bool = False, link: bool = False,
                solid: bool = False) -> QPushButton:
        b = QPushButton(text)
        b.setCursor(Qt.PointingHandCursor)
        if primary:
            b.setProperty("primary", True)
        if link:
            b.setProperty("link", True)
        if solid:
            b.setProperty("solid", True)
        return b

    @staticmethod
    def _shrinkable(combo: QComboBox, chars: int) -> None:
        """Let a dropdown narrow below its longest entry, so the form column
        can be made slimmer and the preview gets the room."""
        combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        combo.setMaxVisibleItems(12)        # show the whole list, no scroll arrows
        combo.setMinimumContentsLength(chars)

    def _build_ui(self):
        form = QWidget()
        lay = QVBoxLayout(form)
        lay.setContentsMargins(14, 12, 8, 14)
        lay.setSpacing(12)

        row = QHBoxLayout()
        row.setSpacing(10)
        row.addWidget(self._caption("Template"))
        self.template_combo = QComboBox()
        for heading, patho in (("No pathogenic", False), ("Pathogenic", True)):
            group = [(key, cfg) for key, cfg in TEMPLATES.items() if cfg.pathogenic == patho]
            if not group:
                continue
            self.template_combo.addItem(f"— {heading} —")     # a heading, not a choice
            item = self.template_combo.model().item(self.template_combo.count() - 1)
            item.setFlags(item.flags() & ~(Qt.ItemIsSelectable | Qt.ItemIsEnabled))
            head_font = QFont(self.template_combo.font())      # highlighted, so it reads as a title
            head_font.setBold(True)
            item.setData(head_font, Qt.FontRole)
            item.setData(QColor(style.ACCENT), Qt.ForegroundRole)
            item.setData(QColor("#e3eef0"), Qt.BackgroundRole)
            for key, cfg in group:
                self.template_combo.addItem(cfg.name, key)
        self.template_combo.setCurrentIndex(max(self.template_combo.findData(DEFAULT_TEMPLATE), 0))
        self.template_combo.currentIndexChanged.connect(self._template_changed)
        self.template_combo.currentIndexChanged.connect(self._bulk_stale)
        self._shrinkable(self.template_combo, 14)
        row.addWidget(self.template_combo, 1)
        row.addSpacing(14)
        self.mode_bar = QTabBar()          # Manual entry / Bulk upload, drives mode_tabs
        self.mode_bar.setDrawBase(False)
        self.mode_bar.setExpanding(False)
        self.mode_bar.addTab("Manual entry")
        self.mode_bar.addTab("Bulk upload")
        row.addWidget(self.mode_bar)
        lay.addLayout(row)

        # Two tabs: patient demography / the rest of the report inputs.
        patient_page, report_page = QWidget(), QWidget()
        play, rlay = QVBoxLayout(patient_page), QVBoxLayout(report_page)

        pgrp = QGroupBox("Patient details")
        pf = QFormLayout(pgrp)
        pf.setVerticalSpacing(10)
        pf.setHorizontalSpacing(14)
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
                    "" if sheet_client.is_configured() else
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
        self.qc_group = qgrp

        lgrp = QGroupBox("Gene list link (Methodology “Click here”)")
        ll = QVBoxLayout(lgrp)
        self.url_edit = QLineEdit(placeholderText="https://…")
        self.url_edit.setToolTip("Paste this run's gene list link.")
        self.url_edit.textChanged.connect(self._edited)
        ll.addWidget(self.url_edit)
        rlay.addWidget(lgrp)

        self.tabs = QTabWidget()
        self.tabs.setObjectName("innerTabs")
        self.tabs.tabBar().setExpanding(True)     # tabs share the full width evenly
        self.tabs.tabBar().setDrawBase(False)
        self.tabs.addTab(self._scroll_page(patient_page), "Patient Details")
        self.tabs.addTab(self._scroll_page(report_page), "Report Details")
        self.gene_editor = GeneTableEditor()
        self.gene_editor.changed.connect(self._edited)
        self.tabs.addTab(self.gene_editor, "Gene Coverage")
        # Variants: only for templates with the variant tables (CMC - Pathogenic).
        self.variants_panel = VariantsPanel()
        self.variants_panel.changed.connect(self._edited)
        self.variants_panel.upload_requested.connect(self._upload_variants)
        self._variants_tab = self.tabs.addTab(self.variants_panel, "Variants")
        # Manual entry and bulk upload are separate tabs.
        self.bulk_panel = BulkPanel()
        self.bulk_panel.shown.connect(self._bulk_show)
        self.bulk_panel.edited.connect(self._edited)
        self.bulk_panel.upload.connect(self._bulk_upload)
        self.mode_tabs = QTabWidget()
        self.mode_tabs.addTab(self.tabs, "Manual entry")
        self.mode_tabs.addTab(self.bulk_panel, "Bulk upload")
        self.mode_tabs.tabBar().hide()
        self.mode_bar.currentChanged.connect(self.mode_tabs.setCurrentIndex)
        self.mode_tabs.currentChanged.connect(self.mode_bar.setCurrentIndex)
        self.mode_tabs.currentChanged.connect(lambda _: self._preview())
        lay.addWidget(self.mode_tabs, 1)

        foot = QFrame()
        foot.setObjectName("card")
        fl = QVBoxLayout(foot)
        fl.setContentsMargins(14, 12, 14, 14)
        fl.setSpacing(10)
        row = QHBoxLayout()
        row.addWidget(self._caption("Save to"))
        self.folder_edit = QLineEdit(self.settings.value("out_dir", "") or
                                     os.path.join(os.path.expanduser("~"), "Downloads"))
        self.folder_edit.setReadOnly(True)
        row.addWidget(self.folder_edit, 1)
        browse = self._button("Change folder…")
        browse.clicked.connect(self._choose_folder)
        row.addWidget(browse)
        fl.addLayout(row)
        row = QHBoxLayout()
        row.setSpacing(10)
        row.addWidget(self._caption("Clinical reviewer"))
        self.reviewer_combo = QComboBox()
        self.reviewer_combo.setToolTip("Whose signature goes in the signature block")
        for key, rev in REVIEWERS.items():
            self.reviewer_combo.addItem(rev.name, key)
        self.reviewer_combo.currentIndexChanged.connect(self._edited)
        self.reviewer_combo.currentIndexChanged.connect(self._reviewer_changed)
        self._shrinkable(self.reviewer_combo, 16)
        row.addWidget(self.reviewer_combo, 1)
        for text, slot, tip in (("Save draft", self._save_draft,
                                 "Save this report's details to a file to continue later"),
                                ("Load draft", self._load_draft,
                                 "Open a report you saved earlier")):
            b = self._button(text)
            b.setToolTip(tip)
            b.clicked.connect(slot)
            row.addWidget(b)
        fl.addLayout(row)
        row = QHBoxLayout()
        row.setSpacing(10)
        for text, kind, primary in (("Export DOCX", "docx", False), ("Export PDF", "pdf", False),
                                    ("Export DOCX + PDF", "both", True)):
            b = self._button(text, primary=primary, solid=not primary)
            b.setMinimumHeight(38)
            b.clicked.connect(lambda _=False, k=kind: self._export(k))
            row.addWidget(b, 1)
        fl.addLayout(row)
        self.actions_toggle = QToolButton()
        self.actions_toggle.setCursor(Qt.PointingHandCursor)
        self.actions_toggle.setToolTip("Hide or show the export / draft bar")
        self.actions_toggle.setStyleSheet("QToolButton { border: none; color: #5A6270; "
                                          "padding: 2px 6px; font-weight: 600; }")
        self.actions_toggle.clicked.connect(
            lambda: self._set_actions_shown(foot.isHidden()))
        lay.addWidget(self.actions_toggle, 0, Qt.AlignRight)
        lay.addWidget(foot)
        self._actions_bar = foot
        self._set_actions_shown(str(self.settings.value("actions_shown", "true")).lower() != "false")

        preview = QWidget()
        pl = QVBoxLayout(preview)
        bar = QHBoxLayout()
        refresh = self._button("Refresh preview")
        refresh.clicked.connect(lambda: self._preview(force=True))
        bar.addWidget(refresh)
        self.preview_status = QLabel()
        self.preview_status.setObjectName("hint")
        bar.addWidget(self.preview_status, 1)
        pl.addLayout(bar)
        self.scroll = PreviewScroll()
        self.scroll.setWidgetResizable(True)
        self.scroll.setStyleSheet("QScrollArea { background:#e5e7eb; border:1px solid #d1d5db; "
                                  "border-radius:8px; }")
        self.scroll.resized.connect(self._redraw_timer.start)
        pages = QWidget()
        self.pages = QVBoxLayout(pages)
        self.pages.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        self.scroll.setWidget(pages)
        pl.addWidget(self.scroll, 1)

        split = QSplitter(Qt.Horizontal)
        split.addWidget(form)
        split.addWidget(preview)
        split.setStretchFactor(0, 1)          # form and preview share the width equally
        split.setStretchFactor(1, 1)
        split.setSizes([1000, 1000])
        split.setChildrenCollapsible(False)
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
            variants=self.variants_panel.variants(),
            result_tag=self.variants_panel.result_tag(),
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
        self._show_variants_tab(cfg.variants)
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

    def _set_actions_shown(self, shown: bool):
        """Show or hide the Save to / reviewer / export bar (remembered)."""
        self._actions_bar.setVisible(shown)
        self.actions_toggle.setText("Hide actions ▾" if shown else "Show actions ▴")
        self.settings.setValue("actions_shown", "true" if shown else "false")

    def _show_variants_tab(self, on: bool):
        self.variants_panel.secondary_separate = get_template(
            self.template_combo.currentData()).secondary_findings
        self.tabs.setTabVisible(self._variants_tab, on)
        if not on and self.tabs.currentIndex() == self._variants_tab:
            self.tabs.setCurrentIndex(0)

    def _upload_variants(self, path: str = ""):
        """Pick the listed-variants Excel: its variants are listed in the
        Variants tab for the user to tick."""
        if not path:
            path, _ = QFileDialog.getOpenFileName(
                self, "Choose the listed-variants Excel", "", "Excel files (*.xlsx *.xlsm)")
        if not path:
            return
        try:
            found = variants_core.read_variants(path)
        except variants_core.VariantError as e:
            QMessageBox.warning(self, "Variants", str(e))
            return
        self.variants_panel.set_variants(found)
        self._edited()

    def _set_reviewer(self, key: str):
        self.reviewer_combo.setCurrentIndex(max(self.reviewer_combo.findData(key), 0))

    def _received_changed(self, text: str):
        """Collection date is the same as the received date: keep it in step,
        unless the user has set a different collection date themselves."""
        if self._loading_draft:             # a draft carries its own collection date
            return
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
        """Manual entry: this report. Bulk upload: the ticked reports (all of
        them when none is ticked) go into one draft file."""
        if self._in_bulk():
            reports = self.bulk_panel.draft_reports()
            if not reports:
                QMessageBox.information(self, "Save draft", "There are no reports to save yet.")
                return
            payload, pid = {"reports": [r.to_dict() for r in reports]}, "bulk"
        else:
            data = self._collect()
            payload = data.to_dict()
            pid = self._file_name_part(data.patient.patient_id) or "report"
        folder = self.settings.value("draft_dir", "") or self.folder_edit.text()
        path, _ = QFileDialog.getSaveFileName(
            self, "Save draft", os.path.join(folder, f"{pid}_draft.json"), "Draft (*.json)")
        if not path:
            return
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, indent=2, ensure_ascii=False)
        self.settings.setValue("draft_dir", os.path.dirname(path))
        count = len(payload["reports"]) if "reports" in payload else 1
        QMessageBox.information(self, "Draft saved", f"Saved {count} report"
                                f"{'s' if count != 1 else ''}:\n{os.path.normpath(path)}")

    def _load_draft(self):
        folder = self.settings.value("draft_dir", "") or self.folder_edit.text()
        path, _ = QFileDialog.getOpenFileName(self, "Load draft", folder, "Draft (*.json)")
        if not path:
            return
        try:
            with open(path, encoding="utf-8") as fh:
                raw = json.load(fh)
            if isinstance(raw, dict) and "reports" in raw:
                reports = [ReportData.from_dict(r) for r in raw["reports"]]
            else:
                reports = [ReportData.from_dict(raw)]
            if not reports:
                raise ValueError("the draft holds no reports")
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Load draft", f"Could not read the draft:\n{e}")
            return
        self.settings.setValue("draft_dir", os.path.dirname(path))
        if len(reports) > 1:
            self._load_bulk_draft(reports)
            return
        data = reports[0]
        # Select the draft's template without its defaults overwriting the draft.
        cfg = get_template(data.template_key)
        self.template_combo.blockSignals(True)
        self.template_combo.setCurrentIndex(max(self.template_combo.findData(cfg.key), 0))
        self.template_combo.blockSignals(False)
        self._template_defaults = self._defaults_for(cfg.key)
        self._template_reviewer = cfg.reviewer
        self._set_reviewer(data.reviewer or cfg.reviewer)
        self.patient_id_label.setText(cfg.patient_id_label)
        self.fields["patient_id"].setPlaceholderText(
            f"e.g. {cfg.patient_id_prefix} - 00000" if cfg.patient_id_prefix
            else "e.g. 12345" if cfg.patient_id_label == "Patient ID" else "e.g. Baby. Name")
        self.mode_tabs.setCurrentWidget(self.tabs)      # a draft is a manual-entry report
        # A draft's saved details win: don't let its PIN trigger a sheet lookup.
        self._loading_draft = True
        for key, le in self.fields.items():
            le.setText(getattr(data.patient, key, ""))
        self._loading_draft = False
        self._last_received = self.fields["received_date"].text()
        self._set_fetch_status("", "gray")
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
        self._show_variants_tab(cfg.variants)
        self.variants_panel.set_variants(data.variants)
        self.variants_panel.set_result_tag(data.result_tag)
        self._preview()

    def _load_bulk_draft(self, reports: list):
        """A draft of several reports: list them in the Bulk upload tab to
        pick one, edit it in Manual entry, or export the ticked ones."""
        first = reports[0]
        cfg = get_template(first.template_key)
        self.template_combo.blockSignals(True)
        self.template_combo.setCurrentIndex(max(self.template_combo.findData(cfg.key), 0))
        self.template_combo.blockSignals(False)
        self._template_defaults = self._defaults_for(cfg.key)
        self._template_reviewer = cfg.reviewer
        self._set_reviewer(first.reviewer or cfg.reviewer)
        self._bulk_data = None
        self.mode_tabs.setCurrentWidget(self.bulk_panel)
        self.bulk_panel.restore(reports, cfg.key, self.reviewer_combo.currentData() or "")

    @staticmethod
    def _defaults_for(key: str) -> dict[str, str]:
        cfg = get_template(key)
        return {"hospital": cfg.hospital, "referring_clinician": cfg.referring_clinician,
                "specimen": cfg.specimen}

    def _bulk_upload(self, path: str = ""):
        """Pick an Excel file: its patients are looked up and listed in the
        Bulk Upload tab; only the ones the user ticks are exported."""
        if self.bulk_panel.busy:
            return
        if not path:
            path, _ = QFileDialog.getOpenFileName(self, "Choose the Excel file", "",
                                                  "Excel files (*.xlsx *.xlsm)")
        if not path:
            return
        try:
            rows = bulk.read_rows(path)
        except bulk.BulkError as e:
            QMessageBox.warning(self, "Bulk upload", str(e))
            return
        self._bulk_data = None
        self.bulk_panel.load(rows, self.template_combo.currentData() or DEFAULT_TEMPLATE,
                             self.reviewer_combo.currentData() or "",
                             self.reviewer_combo.currentText(),
                             self.fields["report_date"].text().strip())

    def _reviewer_changed(self, *_):
        """The signature follows the chosen reviewer in the bulk list too."""
        key = self.reviewer_combo.currentData() or ""
        self.bulk_panel.set_reviewer(key)
        if self._in_bulk():
            self._preview()

    def _bulk_stale(self, *_):
        self.bulk_panel.check_current(self.template_combo.currentData() or DEFAULT_TEMPLATE,
                                      self.reviewer_combo.currentData() or "")

    def _bulk_show(self, data):
        self._bulk_data = data
        if self.mode_tabs.currentWidget() is self.bulk_panel:
            self._preview()

    def _in_bulk(self) -> bool:
        return self.mode_tabs.currentWidget() is self.bulk_panel

    def _choose_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Choose export folder", self.folder_edit.text())
        if folder:
            folder = os.path.normpath(folder)
            self.folder_edit.setText(folder)
            self.settings.setValue("out_dir", folder)

    # -- preview -----------------------------------------------------------
    def _preview(self, force: bool = False):
        if self._preview_worker is not None and self._preview_worker.isRunning():
            self._pending = True           # re-render once this one finishes
            return
        self._pending = False
        data = (self._bulk_data if self._in_bulk() else None) or self._collect()
        snapshot = data.to_dict()
        if (not force and snapshot == self._preview_snapshot and self._preview_pdf
                and os.path.exists(self._preview_pdf)):
            return                         # nothing changed since the last render
        self._preview_snapshot = snapshot
        self._preview_seq += 1
        out = os.path.join(self._preview_dir, f"preview_{self._preview_seq}.pdf")
        self.preview_status.setText("Rendering preview…")
        self._preview_worker = RenderWorker(data, out, "pdf")
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
        self._preview_snapshot = None       # retry on the next change
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
        width = max(self.scroll.viewport().width() - 12, 240)
        doc = pdfium.PdfDocument(self._preview_pdf)
        try:
            count = len(doc)
            while self.pages.count() > count:
                self.pages.takeAt(self.pages.count() - 1).widget().deleteLater()
            while self.pages.count() < count:
                lbl = QLabel(alignment=Qt.AlignHCenter)
                lbl.setStyleSheet("background:white;margin:8px 0;border:1px solid #d1d5db;")
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
    _file_name_part = staticmethod(bulk.file_name_part)

    @staticmethod
    def _file_stem(data: ReportData) -> str:
        return bulk.file_stem(data)

    def _export(self, kind: str):
        if self._in_bulk():
            self.bulk_panel.export(kind, self.folder_edit)
            return
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
        self.bulk_panel.shutdown()
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
    style.apply(app)
    win = MainWindow()
    win.showMaximized()
    sys.exit(app.exec())
