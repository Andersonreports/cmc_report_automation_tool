"""Look of the app: a light blue theme with orange highlights, cards for the form
sections, slim scroll bars and flat tabs. Applied once from main()."""
from __future__ import annotations

import os
import tempfile

from PySide6.QtGui import QColor, QFont, QPainter, QPalette, QPen, QPixmap
from PySide6.QtCore import QPoint, Qt

ACCENT = "#0f4c5c"
ACCENT_DARK = "#0a3642"
ORANGE = "#e09f3e"
ORANGE_DARK = "#c98626"
SURFACE = "#ffffff"
BG = "#f4f8f8"
SIDEBAR = "#0b2545"
TINT = "#fdf0dc"
LINE = "#c9d8ea"
MUTED = "#6b7280"
INK = "#1f2937"
OK, WARN, BAD = "#15803d", "#b35c00", "#c00000"


def _arrow_png() -> str:
    """A small down chevron for the dropdown boxes (a styled QComboBox has no
    native arrow), drawn here so no image file has to ship with the app."""
    folder = os.path.join(tempfile.gettempdir(), "cmc_report_ui")
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, "chevron_down.png")
    px = QPixmap(24, 24)
    px.fill(Qt.transparent)
    p = QPainter(px)
    p.setRenderHint(QPainter.Antialiasing)
    p.setPen(QPen(QColor(MUTED), 2.6, c=Qt.RoundCap, j=Qt.RoundJoin))
    p.drawPolyline([QPoint(6, 9), QPoint(12, 15), QPoint(18, 9)])
    p.end()
    px.save(path, "PNG")
    return path.replace("\\", "/")


def stylesheet() -> str:
    return f"""
* {{ font-family: "Segoe UI", "Noto Sans", sans-serif; font-size: 10pt; color: {INK}; }}
QMainWindow, QDialog, QMessageBox, QProgressDialog {{ background: {BG}; }}
QToolTip {{ background: {INK}; color: white; border: none; padding: 5px 8px; }}
QLabel:disabled {{ color: #9ca3af; }}
QLabel#caption {{ color: {MUTED}; font-weight: 600; }}
QLabel#hint {{ color: {MUTED}; }}
QLabel#chip {{ background: {TINT}; color: {ACCENT}; border-radius: 11px;
              padding: 3px 12px; font-weight: 600; }}

QLineEdit, QPlainTextEdit, QComboBox {{ background: {SURFACE}; border: 1px solid {LINE};
    border-radius: 8px; padding: 7px 10px; selection-background-color: #fbe3bd;
    selection-color: #111827; }}
QLineEdit:focus, QPlainTextEdit:focus, QComboBox:focus {{ border: 2px solid {ACCENT}; padding: 6px 9px; }}
QLineEdit:read-only {{ background: #eef3fa; color: #4b5563; }}
QLineEdit:hover, QPlainTextEdit:hover, QComboBox:hover {{ border: 1px solid #8fb4bc; }}
QComboBox {{ padding-right: 30px; combobox-popup: 0; }}
QComboBox::drop-down {{ subcontrol-origin: padding; subcontrol-position: center right;
    width: 28px; border: none; }}
QComboBox::down-arrow {{ image: url({_arrow_png()}); width: 12px; height: 12px; }}
QComboBox QLineEdit {{ border: none; padding: 0; background: transparent; }}
QComboBox QAbstractItemView {{ background: {SURFACE}; border: 1px solid {LINE}; outline: 0;
    padding: 4px; selection-background-color: {TINT}; selection-color: #111827; }}
QComboBox QAbstractItemView::item {{ min-height: 30px; padding: 2px 8px; border-radius: 6px; }}

QPushButton {{ background: #edf4f5; border: none; border-radius: 10px;
    padding: 8px 18px; color: {ACCENT_DARK}; font-weight: 600; }}
QPushButton:hover {{ background: #d9e8ea; }}
QPushButton:pressed {{ background: #c3dadd; }}
QPushButton:disabled {{ color: #a3acb8; background: #f1f3f6; }}
QPushButton[solid="true"] {{ background: {ACCENT}; color: white; border: none; }}
QPushButton[solid="true"]:hover {{ background: #1b6377; }}
QPushButton[solid="true"]:pressed {{ background: {ACCENT_DARK}; }}
QPushButton[solid="true"]:disabled {{ background: #a9c4cb; color: white; }}
QPushButton[primary="true"] {{ background: {ORANGE}; color: white; border: none;
    font-weight: 700; }}
QPushButton[primary="true"]:hover {{ background: #e8ad57; }}
QPushButton[primary="true"]:pressed {{ background: {ORANGE_DARK}; }}
QPushButton[primary="true"]:disabled {{ background: #f0d6a8; color: white; }}
QPushButton[link="true"] {{ background: transparent; border: none; color: {ACCENT};
    padding: 7px 8px; font-weight: 600; }}
QPushButton[link="true"]:hover {{ background: {TINT}; }}
QPushButton[link="true"]:disabled {{ color: #9ca3af; background: transparent; }}

QFrame#dropzone {{ background: #e9f7ee; border: 1.5px dashed #4fae75; border-radius: 9px; }}
QFrame#dropzone[hover="true"] {{ border-color: #1f8a4c; background: #d8f1e2; }}
QFrame#dropzone QLabel {{ border: none; background: transparent; color: #1b6b3a; font-weight: 600; }}

QTabWidget::pane {{ border: none; background: transparent; top: 6px; }}
QTabBar::tab {{ background: #e1eced; color: #475569; padding: 8px 20px; margin: 0 6px 0 0;
    border: none; border-radius: 15px; font-weight: 600; }}
QTabBar::tab:selected {{ color: white;
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #1b6377, stop:1 #0f4c5c); }}
QTabBar::tab:hover:!selected {{ background: #cfe2e4; color: {ACCENT}; }}

QGroupBox {{ background: {SURFACE}; border: 1px solid #dbe5f1; border-radius: 12px;
    margin-top: 24px; padding: 16px 12px 12px 12px; }}
QGroupBox::title {{ subcontrol-origin: margin; left: 4px; top: 0px; padding: 0 2px;
    color: {ACCENT}; font-weight: 700; }}
QFrame#card {{ background: {SURFACE}; border: 1px solid #dbe5f1; border-radius: 12px; }}
QFrame#card QLabel {{ border: none; }}

QTableWidget {{ background: {SURFACE}; border: 1px solid #e5e7eb; border-radius: 8px;
    gridline-color: transparent; outline: 0; }}
QTableWidget::item {{ padding: 4px 8px; border-bottom: 1px solid #f1f2f4; }}
QTableWidget::item:selected {{ background: {TINT}; color: #111827; }}
QHeaderView::section {{ background: #eaf2fb; border: none; border-bottom: 2px solid {ORANGE};
    padding: 8px; font-weight: 600; color: {ACCENT}; }}

QScrollArea {{ background: transparent; }}
QScrollBar:vertical {{ background: transparent; width: 11px; margin: 2px; }}
QScrollBar::handle:vertical {{ background: #cbd5e1; border-radius: 4px; min-height: 30px; }}
QScrollBar::handle:vertical:hover {{ background: #94a3b8; }}
QScrollBar:horizontal {{ background: transparent; height: 11px; margin: 2px; }}
QScrollBar::handle:horizontal {{ background: #cbd5e1; border-radius: 4px; min-width: 30px; }}
QScrollBar::handle:horizontal:hover {{ background: #94a3b8; }}
QScrollBar::add-line, QScrollBar::sub-line {{ width: 0; height: 0; }}
QScrollBar::add-page, QScrollBar::sub-page {{ background: transparent; }}
QSplitter::handle {{ background: transparent; }}
"""


def apply(app) -> None:
    app.setFont(QFont("Segoe UI", 10))
    pal = QPalette()
    pal.setColor(QPalette.Window, QColor(BG))
    pal.setColor(QPalette.Base, QColor(SURFACE))
    pal.setColor(QPalette.Highlight, QColor(ACCENT))
    pal.setColor(QPalette.HighlightedText, QColor("white"))
    app.setPalette(pal)
    app.setStyleSheet(stylesheet())
