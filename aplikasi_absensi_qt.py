import csv
import os
import sqlite3
import sys
from typing import Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PySide6.QtCore import Qt, QDate, QLocale, QEasingCurve, QPropertyAnimation, QSize
from PySide6.QtGui import QColor, QFont, QFontDatabase, QIcon, QPainter, QPixmap
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QButtonGroup,
    QCalendarWidget,
    QComboBox,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFileDialog,
    QFrame,
    QGraphicsDropShadowEffect,
    QGraphicsOpacityEffect,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_NAME = os.path.join(BASE_DIR, "absensi.db")

LOGO_PATH = os.path.join(BASE_DIR, "img", "logo.png")
STATUS_LIST = ["Hadir", "Alpa", "Izin", "Sakit"]
STATUS_WARNA = {
    "Hadir": "#059669",
    "Alpa": "#DC2626",
    "Izin": "#2563EB",
    "Sakit": "#D97706",
}
JENIS_NILAI = ["Latihan", "Ulangan", "UTS", "UAS"]
BOBOT_NILAI = {"Latihan": 20, "Ulangan": 30, "UTS": 20, "UAS": 30}
JENIS_NILAI_WARNA = {
    "Latihan": "#0EA5E9",
    "Ulangan": "#8B5CF6",
    "UTS": "#F59E0B",
    "UAS": "#EF4444",
}

# Hari Libur Nasional Indonesia
# Libur tetap (tanggal sama setiap tahun)
LIBUR_TETAP = {
    (1, 1): "Tahun Baru Masehi",
    (5, 1): "Hari Buruh Internasional",
    (6, 1): "Hari Lahir Pancasila",
    (17, 8): "Hari Kemerdekaan RI",
    (25, 12): "Hari Raya Natal",
}

# Libur dinamis (berubah tiap tahun berdasarkan SKB 3 Menteri)
LIBUR_DINAMIS = {
    2025: {
        (1, 17): "Isra Miraj Nabi Muhammad SAW",
        (1, 29): "Tahun Baru Imlek 2576",
        (3, 29): "Hari Raya Nyepi",
        (3, 30): "Hari Raya Idul Fitri",
        (3, 31): "Hari Raya Idul Fitri",
        (4, 18): "Wafat Isa Almasih",
        (4, 20): "Kebangkitan Yesus Kristus",
        (5, 1): "Hari Raya Waisak 2569",
        (5, 29): "Kenaikan Isa Almasih",
        (6, 7): "Hari Raya Idul Adha",
        (6, 27): "Tahun Baru Islam 1447 H",
        (9, 5): "Maulid Nabi Muhammad SAW",
    },
    2026: {
        (1, 16): "Isra Miraj Nabi Muhammad SAW",
        (2, 17): "Tahun Baru Imlek 2577",
        (3, 19): "Hari Raya Nyepi",
        (3, 21): "Hari Raya Idul Fitri 1447 H",
        (3, 22): "Hari Raya Idul Fitri 1447 H",
        (4, 3): "Wafat Isa Almasih",
        (4, 5): "Kebangkitan Yesus Kristus",
        (5, 31): "Hari Raya Waisak 2570",
        (5, 14): "Kenaikan Isa Almasih",
        (5, 27): "Hari Raya Idul Adha 1447 H",
        (6, 16): "Tahun Baru Islam 1448 H",
    },
    2027: {
        (2, 6): "Tahun Baru Imlek 2578",
        (3, 9): "Hari Raya Nyepi",
        (3, 10): "Hari Raya Idul Fitri 1448 H",
        (3, 11): "Hari Raya Idul Fitri 1448 H",
        (4, 2): "Wafat Isa Almasih",
        (4, 4): "Kebangkitan Yesus Kristus",
        (5, 20): "Kenaikan Isa Almasih",
        (5, 21): "Hari Raya Waisak 2571",
        (5, 27): "Hari Raya Idul Adha 1448 H",
        (6, 16): "Tahun Baru Islam 1449 H",
        (1, 16): "Isra Miraj Nabi Muhammad SAW",
        (9, 14): "Maulid Nabi Muhammad SAW",
    },
    2028: {
        (1, 5): "Isra Miraj Nabi Muhammad SAW",
        (1, 26): "Tahun Baru Imlek 2579",
        (2, 26): "Hari Raya Nyepi",
        (2, 27): "Hari Raya Idul Fitri 1449 H",
        (3, 17): "Wafat Isa Almasih",
        (3, 19): "Kebangkitan Yesus Kristus",
        (5, 8): "Kenaikan Isa Almasih",
        (5, 10): "Hari Raya Waisak 2572",
        (5, 16): "Hari Raya Idul Adha 1449 H",
        (6, 5): "Tahun Baru Islam 1450 H",
        (9, 3): "Maulid Nabi Muhammad SAW",
    },
    2029: {
        (1, 25): "Tahun Baru Imlek 2580",
        (2, 16): "Hari Raya Nyepi",
        (2, 17): "Hari Raya Idul Fitri 1450 H",
        (3, 6): "Wafat Isa Almasih",
        (3, 8): "Kebangkitan Yesus Kristus",
        (4, 26): "Kenaikan Isa Almasih",
        (4, 29): "Hari Raya Waisak 2573",
        (5, 6): "Hari Raya Idul Adha 1450 H",
        (5, 25): "Tahun Baru Islam 1451 H",
        (12, 24): "Maulid Nabi Muhammad SAW",
    },
}

def apakah_libur(tanggal: QDate) -> Optional[str]:
    bulan = tanggal.month()
    hari = tanggal.day()
    tahun = tanggal.year()
    key = (bulan, hari)
    if key in LIBUR_TETAP:
        return LIBUR_TETAP[key]
    tahun_data = LIBUR_DINAMIS.get(tahun, {})
    if key in tahun_data:
        return tahun_data[key]
    return None


class KalenderMerah(QCalendarWidget):
    """Calendar widget yang menandai tanggal merah (hari libur nasional) dengan warna merah."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setLocale(QLocale(QLocale.Language.Indonesian, QLocale.Country.Indonesia))
        self.setVerticalHeaderFormat(QCalendarWidget.VerticalHeaderFormat.NoVerticalHeader)
        self.setFirstDayOfWeek(Qt.DayOfWeek.Monday)
        self._cache_libur = {}
        self._connect_signals()

    def _connect_signals(self):
        self.currentPageChanged.connect(self._invalidate_cache)

    def _invalidate_cache(self):
        self._cache_libur.clear()

    def _is_libur(self, date: QDate) -> bool:
        return apakah_libur(date) is not None

    def paintCell(self, painter: QPainter, rect, date):
        libur = self._is_libur(date)

        if not libur:
            super().paintCell(painter, rect, date)
            return

        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        bg = QColor(254, 226, 226)
        painter.fillRect(rect, bg)

        painter.setPen(QColor(153, 27, 27))
        font = painter.font()
        font.setBold(True)
        painter.setFont(font)

        text_rect = rect.adjusted(0, 0, 0, -4)
        painter.drawText(text_rect, Qt.AlignmentFlag.AlignCenter, str(date.day()))

        small_font = painter.font()
        small_font.setPixelSize(7)
        small_font.setBold(True)
        painter.setFont(small_font)
        painter.setPen(QColor(185, 28, 28, 200))
        label_rect = rect.adjusted(0, rect.height() - 14, 0, 0)
        painter.drawText(label_rect, Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignBottom, "Libur")

        painter.restore()


class KalenderMerahPopup(QDateEdit):
    """QDateEdit yang menggunakan KalenderMerah sebagai popup."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCalendarPopup(True)
        self._kalender = KalenderMerah()
        self.setCalendarWidget(self._kalender)

STYLESHEET = """
* { font-family: 'Poppins', 'Segoe UI'; }
QMainWindow { background-color: #EFF7F1; }

QWidget#sidebar {
    background-color: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #0F8A43, stop:1 #0B5A2E);
    border-right: 1px solid rgba(0, 0, 0, 0.08);
}
QLabel#brand { font-size: 19px; font-weight: 700; color: #FFFFFF; }
QLabel#brand_small { font-size: 15px; font-weight: 800; color: #FFFFFF; letter-spacing: 0.2px; }
QLabel#brand_sub { font-size: 11px; color: rgba(255, 255, 255, 0.68); font-weight: 500; }
QLabel#logo {
    background-color: #FFFFFF; color: #15803D; border-radius: 12px;
    font-size: 17px; font-weight: 800;
}
QLabel#logo_gambar {
    background: rgba(255, 255, 255, 0.14);
    border: none; border-radius: 12px;
    padding: 4px;
}
QPushButton#nav {
    color: rgba(255, 255, 255, 0.78); background: transparent; border: none;
    text-align: left; padding: 11px 14px;
    border-radius: 10px;
    font-size: 13.5px; font-weight: 600;
    letter-spacing: 0.2px;
}
QPushButton#nav::hover { background-color: rgba(255, 255, 255, 0.12); color: #FFFFFF; }
QPushButton#nav:pressed { background-color: rgba(255, 255, 255, 0.20); }
QPushButton#nav:checked { background-color: #FFFFFF; color: #0B5A2E; font-weight: 700; }
QLabel#seksi {
    color: rgba(255, 255, 255, 0.50);
    font-size: 10.5px; font-weight: 700;
    letter-spacing: 1.5px;
    padding: 18px 14px 4px 14px;
    background: transparent;
}
QLabel#versi { color: rgba(255, 255, 255, 0.55); font-size: 10.5px; font-weight: 500; }

QLabel#judul_halaman { font-size: 24px; font-weight: 800; color: #123520; }
QLabel#subjudul_halaman { font-size: 13.5px; color: #5F7D6B; font-weight: 500; }
QLabel#bagian { font-size: 14px; font-weight: 700; color: #31513D; }
QLabel#total { font-size: 13px; color: #5F7D6B; }
QLabel#kosong { color: #87AB97; font-size: 14px; background: transparent; }

QFrame#kartu {
    background-color: #FFFFFF;
    border: 1px solid #DCECE2;
    border-radius: 16px;
}
QFrame#kartu_siswa {
    background-color: #FFFFFF;
    border: 1px solid #DCECE2;
    border-radius: 13px;
}
QFrame#kartu_siswa:hover { border: 1px solid #A9DEC1; }

QLineEdit, QPlainTextEdit {
    border: 1px solid #C9DFD2;
    border-radius: 11px;
    padding: 11px 16px;
    font-size: 14px;
    background-color: #FFFFFF;
    color: #13301F;
    selection-background-color: #A9DEC1;
}
QLineEdit { padding: 0px 16px; }
QLineEdit:focus, QPlainTextEdit:focus {
    border: 1.5px solid #16A34A;
    background-color: #FBFFFC;
}
QLineEdit::placeholder, QPlainTextEdit::placeholder { color: #9DBBA9; }

QComboBox {
    border: 1px solid #C9DFD2;
    border-radius: 11px;
    padding: 0px 8px 0px 16px;
    font-size: 14px;
    background-color: #FFFFFF;
    color: #13301F;
}
QComboBox:hover { border-color: #87AB97; }
QComboBox:focus { border: 1.5px solid #16A34A; }
QComboBox::drop-down { border: none; width: 30px; }
QComboBox::down-arrow {
    image: none;
    width: 0px;
    height: 0px;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 6px solid #87AB97;
    margin-right: 10px;
}
QComboBox QAbstractItemView {
    background-color: #FFFFFF;
    color: #13301F;
    border: 1px solid #DCECE2;
    border-radius: 10px;
    padding: 6px;
    outline: none;
    selection-background-color: #DCFCE7;
    selection-color: #16A34A;
}
QComboBox QAbstractItemView::item {
    min-height: 36px;
    padding: 2px 12px;
    margin: 2px 0px;
    border-radius: 8px;
}
QComboBox QAbstractItemView::item:hover { background-color: #EDF5EF; }

QPushButton {
    border: none;
    border-radius: 11px;
    padding: 0px 20px;
    font-size: 14px;
    font-weight: 600;
}
QPushButton#utama { background-color: #16A34A; color: #FFFFFF; }
QPushButton#utama::hover { background-color: #15803D; }
QPushButton#utama:pressed { background-color: #116933; }
QPushButton#masuk { background-color: #10B981; color: #FFFFFF; }
QPushButton#masuk::hover { background-color: #059669; }
QPushButton#netral {
    background-color: #FFFFFF; color: #31513D; border: 1px solid #C9DFD2;
}
QPushButton#netral::hover { background-color: #EDF5EF; }
QPushButton#bahaya { background-color: #EF4444; color: #FFFFFF; }
QPushButton#bahaya::hover { background-color: #DC2626; }

QPushButton[jenis] {
    background-color: #FFFFFF; color: #5F7D6B;
    border: 1px solid #C9DFD2; border-radius: 9px;
    padding: 8px 16px; font-size: 12px; font-weight: 700;
}
QPushButton[jenis]:hover { background-color: #F3FAF6; }
QPushButton[jenis="Hadir"]:checked { background-color: #10B981; border-color: #10B981; color: #FFFFFF; }
QPushButton[jenis="Alpa"]:checked { background-color: #EF4444; border-color: #EF4444; color: #FFFFFF; }
QPushButton[jenis="Izin"]:checked { background-color: #3B82F6; border-color: #3B82F6; color: #FFFFFF; }
QPushButton[jenis="Sakit"]:checked { background-color: #F59E0B; border-color: #F59E0B; color: #FFFFFF; }

QLabel[chip] {
    border-radius: 16px; padding: 7px 16px;
    font-size: 12px; font-weight: 700;
}
QLabel[chip="Hadir"] { background-color: #D1FAE5; color: #065F46; }
QLabel[chip="Alpa"] { background-color: #FEE2E2; color: #991B1B; }
QLabel[chip="Izin"] { background-color: #DBEAFE; color: #1E40AF; }
QLabel[chip="Sakit"] { background-color: #FEF3C7; color: #92400E; }
QLabel[chip="Belum"] { background-color: #EDF5EF; color: #5F7D6B; }

QDateEdit {
    border: 1px solid #C9DFD2;
    border-radius: 11px;
    padding: 0px 6px 0px 14px;
    font-size: 13.5px;
    background-color: #FFFFFF;
    color: #13301F;
}
QDateEdit:hover { border-color: #87AB97; }
QDateEdit:focus { border: 1.5px solid #16A34A; }
QDateEdit::up-button, QDateEdit::down-button {
    width: 0px;
    height: 0px;
    border: none;
}
QDateEdit::up-arrow, QDateEdit::down-arrow {
    image: none;
    width: 0px;
    height: 0px;
    border: none;
}
QDateEdit::drop-down { border: none; width: 30px; }
QDateEdit::drop-down:hover {
    background-color: transparent;
}
QDateEdit::down-arrow {
    image: none;
    width: 0px;
    height: 0px;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 6px solid #87AB97;
    margin-right: 10px;
}

QMenu {
    background-color: #FFFFFF;
    color: #13301F;
    border: 1px solid #DCECE2;
    border-radius: 10px;
    padding: 6px;
}
QMenu::item {
    padding: 9px 26px;
    border-radius: 8px;
    font-size: 13.5px;
}
QMenu::item:selected { background-color: #DCFCE7; color: #16A34A; }

QCalendarWidget QWidget { alternate-background-color: #F3FAF6; }
QCalendarWidget #qt_calendar_navigationbar {
    background-color: #FFFFFF;
    border-bottom: 1px solid #DCECE2;
    padding: 8px;
}
QCalendarWidget QToolButton {
    background-color: #FFFFFF;
    color: #31513D;
    border: none;
    border-radius: 8px;
    padding: 7px 13px;
    font-size: 13.5px;
    font-weight: 700;
}
QCalendarWidget QToolButton:hover { background-color: #DCFCE7; color: #16A34A; }
QCalendarWidget QToolButton::menu-indicator { image: none; }
QCalendarWidget QSpinBox {
    background-color: #FFFFFF;
    color: #31513D;
    border: 1px solid #C9DFD2;
    border-radius: 8px;
    padding: 2px 6px;
    font-size: 13.5px;
    font-weight: 700;
}
QCalendarWidget QAbstractItemView {
    background-color: #FFFFFF;
    color: #13301F;
    selection-background-color: #16A34A;
    selection-color: #FFFFFF;
    outline: none;
    font-size: 12px;
}
QCalendarWidget QTableView { alternate-background-color: #FFFFFF; }
QCalendarWidget QAbstractItemView:enabled { color: #13301F; }

QListWidget {
    border: 1px solid #DCECE2;
    border-radius: 12px;
    background-color: #FFFFFF;
    color: #13301F;
    font-size: 14px;
    padding: 6px;
    outline: none;
}
QListWidget::item { padding: 10px 12px; border-radius: 8px; }
QListWidget::item:hover { background-color: #F3FAF6; }
QListWidget::item:selected { background-color: #DCFCE7; color: #16A34A; font-weight: 700; }

QTableWidget {
    border: none;
    background-color: #FFFFFF;
    gridline-color: transparent;
    alternate-background-color: #EDF8F2;
    selection-background-color: #CDEFDD;
    selection-color: #123A26;
    font-size: 13px;
    color: #13301F;
}
QHeaderView::section {
    background-color: #D5F0E0;
    color: #1F5B3A;
    font-size: 12px;
    font-weight: 700;
    border: none;
    border-bottom: 1px solid #BFE2CE;
    padding: 11px 9px;
}
QTableCornerButton::section { background-color: #D5F0E0; border: none; }

QFrame#card_header {
    background-color: transparent;
    border: none;
}
QLabel#card_judul { font-size: 15px; font-weight: 700; color: #294A37; background: transparent; }
QLabel#card_badge {
    background-color: #DCFCE7; color: #15803D;
    border-radius: 13px;
    padding: 5px 15px;
    font-size: 12px; font-weight: 700;
}
QFrame#search_frame {
    background-color: #F3FAF6;
    border: 1px solid #DCECE2;
    border-radius: 12px;
}
QLabel#search_label { font-size: 13px; font-weight: 600; color: #5F7D6B; background: transparent; }

QFrame#nilai_kartu {
    background-color: #FFFFFF;
    border: 1px solid #DCECE2;
    border-radius: 14px;
}
QLabel#nilai_jenis {
    font-size: 15px; font-weight: 800; color: #13301F;
    border-radius: 10px;
    padding: 6px 14px;
}
QLabel#jadwal_kartu {
    font-size: 13px; font-weight: 600; color: #31513D;
    background: transparent;
}
QFrame#jadwal_badge {
    background-color: #DCFCE7; color: #15803D;
    border-radius: 10px; padding: 4px 12px;
}
QLabel#val_akhir {
    font-size: 17px; font-weight: 800; color: #15803D;
    background: transparent;
}
QLabel#label_rapor { font-size: 12px; font-weight: 600; color: #5F7D6B; background: transparent; }
QDoubleSpinBox {
    border: 1px solid #C9DFD2;
    border-radius: 11px;
    padding: 0px 12px;
    font-size: 14px;
    background-color: #FFFFFF;
    color: #13301F;
}
QDoubleSpinBox:focus { border: 1.5px solid #16A34A; }
QFrame#kosong_jadwal {
    background-color: #FFFFFF;
    border: 1px dashed #C9DFD2;
    border-radius: 14px;
}

QScrollArea { border: none; background: transparent; }
QWidget#konten_absen { background: transparent; }

QScrollBar:vertical { background: transparent; width: 9px; margin: 2px; }
QScrollBar::handle:vertical { background: #C9DFD2; border-radius: 4px; min-height: 30px; }
QScrollBar::handle:vertical:hover { background: #87AB97; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }
QScrollBar:horizontal { background: transparent; height: 9px; margin: 2px; }
QScrollBar::handle:horizontal { background: #C9DFD2; border-radius: 4px; min-width: 30px; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0px; }
"""


def get_conn():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    return conn


def init_db():
    conn = sqlite3.connect(DB_NAME)
    try:
        conn.execute("PRAGMA foreign_keys = ON")
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS kelas (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                nama_kelas TEXT NOT NULL UNIQUE
            );

            CREATE TABLE IF NOT EXISTS siswa (
                id       INTEGER PRIMARY KEY AUTOINCREMENT,
                kelas_id INTEGER NOT NULL REFERENCES kelas(id) ON DELETE CASCADE,
                nama     TEXT NOT NULL,
                UNIQUE (kelas_id, nama)
            );

            CREATE TABLE IF NOT EXISTS presensi (
                id       INTEGER PRIMARY KEY AUTOINCREMENT,
                tanggal  TEXT NOT NULL,
                siswa_id INTEGER NOT NULL REFERENCES siswa(id) ON DELETE CASCADE,
                status   TEXT NOT NULL CHECK (status IN ('Hadir','Alpa','Izin','Sakit')),
                UNIQUE (tanggal, siswa_id)
            );

            CREATE TABLE IF NOT EXISTS jadwal_ujian (
                id       INTEGER PRIMARY KEY AUTOINCREMENT,
                kelas_id INTEGER NOT NULL REFERENCES kelas(id) ON DELETE CASCADE,
                tanggal  TEXT NOT NULL,
                jenis    TEXT NOT NULL CHECK (jenis IN ('Latihan','Ulangan','UTS','UAS')),
                keterangan TEXT
            );

            CREATE TABLE IF NOT EXISTS nilai (
                id       INTEGER PRIMARY KEY AUTOINCREMENT,
                siswa_id INTEGER NOT NULL REFERENCES siswa(id) ON DELETE CASCADE,
                jenis    TEXT NOT NULL CHECK (jenis IN ('Latihan','Ulangan','UTS','UAS')),
                tanggal  TEXT NOT NULL,
                nilai    REAL NOT NULL
            );
            """
        )
        legacy = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='absensi'"
        ).fetchone()
        if legacy and conn.execute("SELECT COUNT(*) FROM absensi").fetchone()[0] == 0:
            conn.execute("DROP TABLE absensi")
        conn.commit()
    finally:
        conn.close()


def ambil_daftar_kelas():
    conn = get_conn()
    try:
        rows = conn.execute(
            "SELECT id, nama_kelas FROM kelas ORDER BY nama_kelas"
        ).fetchall()
        return [(r["id"], r["nama_kelas"]) for r in rows]
    finally:
        conn.close()


def ambil_status(tanggal, siswa_id):
    conn = get_conn()
    try:
        r = conn.execute(
            "SELECT status FROM presensi WHERE tanggal = ? AND siswa_id = ?",
            (tanggal, siswa_id),
        ).fetchone()
        return r["status"] if r else None
    finally:
        conn.close()


def simpan_nilai_batch(data):
    conn = get_conn()
    try:
        conn.execute("BEGIN")
        for siswa_id, jenis, tanggal, nilai in data:
            conn.execute(
                "DELETE FROM nilai WHERE siswa_id = ? AND jenis = ? AND tanggal = ?",
                (siswa_id, jenis, tanggal),
            )
            conn.execute(
                "INSERT INTO nilai (siswa_id, jenis, tanggal, nilai) VALUES (?, ?, ?, ?)",
                (siswa_id, jenis, tanggal, nilai),
            )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def hapus_tanggal_nilai(kelas_id, jenis, tanggal):
    conn = get_conn()
    try:
        conn.execute(
            """
            DELETE FROM nilai
            WHERE tanggal = ? AND jenis = ? AND siswa_id IN
                (SELECT id FROM siswa WHERE kelas_id = ?)
            """,
            (tanggal, jenis, kelas_id),
        )
        conn.commit()
    finally:
        conn.close()


def ambil_nilai_tanggal(kelas_id, jenis, tanggal):
    conn = get_conn()
    try:
        rows = conn.execute(
            """
            SELECT s.id AS siswa_id, s.nama, n.nilai
            FROM siswa s
            LEFT JOIN nilai n
                ON n.siswa_id = s.id AND n.jenis = ? AND n.tanggal = ?
            WHERE s.kelas_id = ?
            ORDER BY s.nama
            """,
            (jenis, tanggal, kelas_id),
        ).fetchall()
        return [{"id": r["siswa_id"], "nama": r["nama"], "nilai": r["nilai"]} for r in rows]
    finally:
        conn.close()


def ambil_jadwal(kelas_id):
    conn = get_conn()
    try:
        rows = conn.execute(
            """
            SELECT id, tanggal, jenis, keterangan
            FROM jadwal_ujian
            WHERE kelas_id = ?
            ORDER BY tanggal
            """,
            (kelas_id,),
        ).fetchall()
        return [
            {"id": r["id"], "tanggal": r["tanggal"], "jenis": r["jenis"], "keterangan": r["keterangan"]}
            for r in rows
        ]
    finally:
        conn.close()


def tambah_jadwal(kelas_id, tanggal, jenis, keterangan):
    conn = get_conn()
    try:
        conn.execute(
            "INSERT INTO jadwal_ujian (kelas_id, tanggal, jenis, keterangan) VALUES (?, ?, ?, ?)",
            (kelas_id, tanggal, jenis, keterangan),
        )
        conn.commit()
    finally:
        conn.close()


def hapus_jadwal(jadwal_id):
    conn = get_conn()
    try:
        conn.execute("DELETE FROM jadwal_ujian WHERE id = ?", (jadwal_id,))
        conn.commit()
    finally:
        conn.close()


def hapus_semua_presensi():
    conn = get_conn()
    try:
        conn.execute("DELETE FROM presensi")
        conn.commit()
    finally:
        conn.close()


def hapus_semua_nilai():
    conn = get_conn()
    try:
        conn.execute("DELETE FROM nilai")
        conn.commit()
    finally:
        conn.close()


def hitung_nilai_siswa(kelas_id):
    conn = get_conn()
    try:
        siswa_rows = conn.execute(
            """
            SELECT s.id, s.nama
            FROM siswa s
            WHERE s.kelas_id = ?
              AND EXISTS (SELECT 1 FROM nilai n WHERE n.siswa_id = s.id)
            ORDER BY s.nama
            """,
            (kelas_id,),
        ).fetchall()
        hasil = []
        for s in siswa_rows:
            rata = {}
            for jenis in JENIS_NILAI:
                rows = conn.execute(
                    "SELECT nilai FROM nilai WHERE siswa_id = ? AND jenis = ?",
                    (s["id"], jenis),
                ).fetchall()
                if rows:
                    rata[jenis] = sum(r["nilai"] for r in rows) / len(rows)
                else:
                    rata[jenis] = None
            nilai_akhir = 0
            total_bobot = 0
            for jenis in JENIS_NILAI:
                if rata[jenis] is not None:
                    nilai_akhir += rata[jenis] * BOBOT_NILAI[jenis]
                    total_bobot += BOBOT_NILAI[jenis]
            if not total_bobot:
                nilai_akhir = None
            else:
                nilai_akhir = nilai_akhir / total_bobot
            hasil.append({"nama": s["nama"], "rata": rata, "akhir": nilai_akhir})
        return hasil
    finally:
        conn.close()


def tinggi_seragam(*widgets, tinggi=40):
    for w in widgets:
        w.setFixedHeight(tinggi)


def pasang_shadow(widget, blur=34, alfa=55, offset_y=5):
    shadow = QGraphicsDropShadowEffect(widget)
    shadow.setBlurRadius(blur)
    shadow.setColor(QColor(16, 92, 45, alfa))
    shadow.setOffset(0, offset_y)
    widget.setGraphicsEffect(shadow)


def muat_font():
    folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")
    if not os.path.isdir(folder):
        return
    for nama_file in sorted(os.listdir(folder)):
        if nama_file.lower().endswith((".ttf", ".otf")):
            QFontDatabase.addApplicationFont(os.path.join(folder, nama_file))


class DialogTambahBanyak(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Tambah Banyak Siswa")
        self.setMinimumSize(420, 380)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)
        info = QLabel("Tempel atau ketik daftar nama,\nsatu nama per baris:")
        info.setObjectName("subjudul_halaman")
        self.editor = QPlainTextEdit()
        self.editor.setPlaceholderText("Budi Santoso\nSiti Aminah\nAhmad Fauzi")
        tombol = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        tombol.button(QDialogButtonBox.StandardButton.Ok).setText("Tambahkan")
        tombol.accepted.connect(self.accept)
        tombol.rejected.connect(self.reject)
        layout.addWidget(info)
        layout.addWidget(self.editor, 1)
        layout.addWidget(tombol)

    def daftar_nama(self):
        return [b.strip() for b in self.editor.toPlainText().splitlines() if b.strip()]


class HalamanAbsensi(QWidget):
    def __init__(self, utama):
        super().__init__()
        self.utama = utama
        self.peta_tombol = {}
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 20)
        root.setSpacing(14)

        kepala = QVBoxLayout()
        judul = QLabel("Absensi Harian")
        judul.setObjectName("judul_halaman")
        sub = QLabel("Pilih kelas dan tanggal, lalu tandai status tiap siswa.")
        sub.setObjectName("subjudul_halaman")
        kepala.addWidget(judul)
        kepala.addWidget(sub)
        root.addLayout(kepala)

        kartu_alat = QFrame()
        kartu_alat.setObjectName("kartu")
        baris_alat = QHBoxLayout(kartu_alat)
        baris_alat.setContentsMargins(18, 14, 18, 14)
        baris_alat.setSpacing(10)
        label_kelas = QLabel("Kelas")
        label_kelas.setObjectName("bagian")
        self.combo_kelas = QComboBox()
        self.combo_kelas.setMinimumWidth(190)
        self.date_edit = KalenderMerahPopup()
        self.date_edit.setDisplayFormat("yyyy-MM-dd")
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setMinimumWidth(158)
        btn_semua_hadir = QPushButton("Tandai Semua Hadir")
        btn_semua_hadir.setObjectName("masuk")
        btn_semua_hadir.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_semua_hadir.clicked.connect(self.tandai_semua_hadir)
        btn_segarkan = QPushButton("Segarkan")
        btn_segarkan.setObjectName("netral")
        btn_segarkan.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_segarkan.clicked.connect(self.muat_presensi)
        baris_alat.addWidget(label_kelas)
        baris_alat.addWidget(self.combo_kelas)
        baris_alat.addWidget(self.date_edit)
        baris_alat.addStretch()
        baris_alat.addWidget(btn_semua_hadir)
        baris_alat.addWidget(btn_segarkan)
        root.addWidget(kartu_alat)

        baris_chip = QHBoxLayout()
        baris_chip.setSpacing(8)
        self.chips = {}
        for jenis in STATUS_LIST + ["Belum"]:
            chip = QLabel(f"{jenis}: 0")
            chip.setProperty("chip", jenis)
            baris_chip.addWidget(chip)
            self.chips[jenis] = chip
        baris_chip.addStretch()
        root.addLayout(baris_chip)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.konten = QWidget()
        self.konten.setObjectName("konten_absen")
        self.layout_kartu = QVBoxLayout(self.konten)
        self.layout_kartu.setContentsMargins(0, 0, 8, 0)
        self.layout_kartu.setSpacing(8)
        self.layout_kartu.addStretch()
        self.scroll.setWidget(self.konten)
        root.addWidget(self.scroll, 1)

        self.combo_kelas.currentIndexChanged.connect(lambda _: self.muat_presensi())
        self.date_edit.dateChanged.connect(lambda _: self.muat_presensi())
        tinggi_seragam(
            self.combo_kelas, self.date_edit,
            btn_semua_hadir, btn_segarkan,
        )

    def kelas_id_aktif(self):
        return self.combo_kelas.currentData()

    def tanggal_aktif(self):
        return self.date_edit.date().toString("yyyy-MM-dd")

    def reset_tanggal(self):
        self.date_edit.blockSignals(True)
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.blockSignals(False)
        self.muat_presensi()

    def _peringatan_hari(self, qdate):
        alasan = None
        if qdate.dayOfWeek() == Qt.DayOfWeek.Sunday:
            alasan = "hari Minggu"
        else:
            libur = apakah_libur(qdate)
            if libur:
                alasan = f"hari libur ({libur})"
        if alasan is None:
            return True
        jawab = QMessageBox.question(
            self,
            "Peringatan",
            f"Tanggal tersebut adalah {alasan}.\n"
            "Apakah Anda tetap ingin mengisi absensi pada tanggal ini?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        return jawab == QMessageBox.StandardButton.Yes

    def muat_combo(self, kelas_list):
        self.combo_kelas.blockSignals(True)
        self.combo_kelas.clear()
        self.combo_kelas.addItem("-- Pilih Kelas --", None)
        for kid, nama in kelas_list:
            self.combo_kelas.addItem(nama, kid)
        self.combo_kelas.blockSignals(False)
        self.muat_presensi()

    def muat_presensi(self):
        self.peta_tombol.clear()
        while self.layout_kartu.count() > 1:
            item = self.layout_kartu.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        kid = self.kelas_id_aktif()
        if kid is None:
            self._tampil_kosong("Belum ada kelas dipilih.\nTambahkan kelas & siswa lewat menu Kelola Data.")
            self.perbarui_chip({})
            return

        tanggal = self.tanggal_aktif()
        conn = get_conn()
        try:
            rows = conn.execute(
                """
                SELECT s.id, s.nama, p.status
                FROM siswa s
                LEFT JOIN presensi p
                    ON p.siswa_id = s.id AND p.tanggal = ?
                WHERE s.kelas_id = ?
                ORDER BY s.nama
                """,
                (tanggal, kid),
            ).fetchall()
        finally:
            conn.close()

        if not rows:
            self._tampil_kosong("Kelas ini masih kosong.\nTambahkan siswa di menu Kelola Data.")
            self.perbarui_chip({})
            return

        for data in rows:
            kartu = self._buat_kartu(data["id"], data["nama"], data["status"])
            self.layout_kartu.insertWidget(self.layout_kartu.count() - 1, kartu)
        self.perbarui_chip(self._hitung_status(rows))

    def _tampil_kosong(self, teks):
        label = QLabel(teks)
        label.setObjectName("kosong")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.layout_kartu.insertWidget(0, label, 1, Qt.AlignmentFlag.AlignHCenter)

    def _buat_kartu(self, siswa_id, nama, status):
        kartu = QFrame()
        kartu.setObjectName("kartu_siswa")
        baris = QHBoxLayout(kartu)
        baris.setContentsMargins(16, 10, 16, 10)
        baris.setSpacing(8)

        inisial = "".join(k[0] for k in nama.split()[:2]).upper()
        avatar = QLabel(inisial or "?")
        avatar.setFixedSize(36, 36)
        avatar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        avatar.setStyleSheet(
            "background:#DCFCE7;color:#16A34A;border-radius:18px;"
            "font-weight:800;font-size:13px;"
        )

        label_nama = QLabel(nama)
        label_nama.setStyleSheet("font-size:14px;font-weight:600;color:#13301F;background:transparent;border:none;")

        baris.addWidget(avatar)
        baris.addWidget(label_nama)
        baris.addStretch()

        tombol_peta = {}
        for jenis in STATUS_LIST:
            b = QPushButton(jenis)
            b.setProperty("jenis", jenis)
            b.setCheckable(True)
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            b.setChecked(status == jenis)
            b.clicked.connect(
                lambda _, sid=siswa_id, j=jenis: self.klik_status(sid, j)
            )
            tombol_peta[jenis] = b
            baris.addWidget(b)
        self.peta_tombol[siswa_id] = tombol_peta
        return kartu

    def klik_status(self, siswa_id, jenis):
        tanggal = self.tanggal_aktif()
        qdate = self.date_edit.date()
        if not self._peringatan_hari(qdate):
            return
        lama = ambil_status(tanggal, siswa_id)
        tombol_peta = self.peta_tombol.get(siswa_id, {})
        conn = get_conn()
        try:
            if lama == jenis:
                conn.execute(
                    "DELETE FROM presensi WHERE tanggal = ? AND siswa_id = ?",
                    (tanggal, siswa_id),
                )
                baru = None
            else:
                conn.execute(
                    """
                    INSERT INTO presensi (tanggal, siswa_id, status) VALUES (?, ?, ?)
                    ON CONFLICT(tanggal, siswa_id) DO UPDATE SET status = excluded.status
                    """,
                    (tanggal, siswa_id, jenis),
                )
                baru = jenis
            conn.commit()
        except sqlite3.Error as e:
            QMessageBox.critical(
                self, "Gagal", f"Terjadi kesalahan saat menyimpan presensi:\n{e}"
            )
            return
        finally:
            conn.close()

        for j, b in tombol_peta.items():
            b.setChecked(baru == j)
        self.perbarui_chip(self._hitung_db())

    def _hitung_status(self, rows):
        hitung = {j: 0 for j in STATUS_LIST}
        belum = 0
        for r in rows:
            if r["status"] in hitung:
                hitung[r["status"]] += 1
            else:
                belum += 1
        hasil = dict(hitung)
        hasil["Belum"] = belum
        return hasil

    def _hitung_db(self):
        kid = self.kelas_id_aktif()
        if kid is None:
            return {}
        tanggal = self.tanggal_aktif()
        conn = get_conn()
        try:
            total_siswa = conn.execute(
                "SELECT COUNT(*) AS n FROM siswa WHERE kelas_id = ?", (kid,)
            ).fetchone()["n"]
            terisi = {r["status"]: r["n"] for r in conn.execute(
                """
                SELECT p.status, COUNT(*) AS n
                FROM presensi p JOIN siswa s ON s.id = p.siswa_id
                WHERE p.tanggal = ? AND s.kelas_id = ?
                GROUP BY p.status
                """,
                (tanggal, kid),
            ).fetchall()}
        finally:
            conn.close()
        hasil = {j: terisi.get(j, 0) for j in STATUS_LIST}
        hasil["Belum"] = max(total_siswa - sum(hasil.values()), 0)
        return hasil

    def perbarui_chip(self, hitung):
        for jenis, chip in self.chips.items():
            chip.setText(f"{jenis}: {hitung.get(jenis, 0)}")

    def tandai_semua_hadir(self):
        kid = self.kelas_id_aktif()
        if kid is None:
            QMessageBox.warning(self, "Peringatan", "Pilih kelas dulu!")
            return
        if not self._peringatan_hari(self.date_edit.date()):
            return
        tanggal = self.tanggal_aktif()
        jawab = QMessageBox.question(
            self,
            "Konfirmasi",
            f"Tandai SEMUA siswa Hadir untuk {tanggal}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if jawab != QMessageBox.StandardButton.Yes:
            return
        conn = get_conn()
        try:
            ids = [r["id"] for r in conn.execute(
                "SELECT id FROM siswa WHERE kelas_id = ?", (kid,)
            ).fetchall()]
            conn.executemany(
                """
                INSERT INTO presensi (tanggal, siswa_id, status) VALUES (?, ?, 'Hadir')
                ON CONFLICT(tanggal, siswa_id) DO UPDATE SET status = 'Hadir'
                """,
                [(tanggal, i) for i in ids],
            )
            conn.commit()
        except sqlite3.Error as e:
            QMessageBox.critical(
                self, "Gagal", f"Terjadi kesalahan saat menandai hadir:\n{e}"
            )
            return
        finally:
            conn.close()
        self.muat_presensi()


class HalamanKelola(QWidget):
    def __init__(self, utama):
        super().__init__()
        self.utama = utama
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 20)
        root.setSpacing(14)

        kepala = QVBoxLayout()
        judul = QLabel("Kelola Data")
        judul.setObjectName("judul_halaman")
        sub = QLabel("Atur daftar kelas dan siswa di masing-masing kelas.")
        sub.setObjectName("subjudul_halaman")
        kepala.addWidget(judul)
        kepala.addWidget(sub)
        root.addLayout(kepala)

        kolom = QHBoxLayout()
        kolom.setSpacing(14)
        kolom.addWidget(self._panel_kelas(), 1)
        kolom.addWidget(self._panel_siswa(), 2)
        root.addLayout(kolom, 1)
        tinggi_seragam(
            self.input_kelas, self.btn_tambah_kelas,
            self.btn_ubah_kelas, self.btn_hapus_kelas,
            self.input_siswa, self.btn_tambah_siswa, self.btn_banyak,
            self.btn_ubah_siswa, self.btn_hapus_siswa,
        )

    def _panel_kelas(self):
        kartu = QFrame()
        kartu.setObjectName("kartu")
        kolom = QVBoxLayout(kartu)
        kolom.setContentsMargins(18, 16, 18, 16)
        kolom.setSpacing(10)
        label = QLabel("Daftar Kelas")
        label.setObjectName("bagian")
        baris_input = QHBoxLayout()
        baris_input.setSpacing(8)
        self.input_kelas = QLineEdit()
        self.input_kelas.setPlaceholderText("Nama kelas, misal: X RPL 1")
        self.input_kelas.returnPressed.connect(self.tambah_kelas)
        self.btn_tambah_kelas = QPushButton("Tambah")
        self.btn_tambah_kelas.setObjectName("utama")
        self.btn_tambah_kelas.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_tambah_kelas.clicked.connect(self.tambah_kelas)
        baris_input.addWidget(self.input_kelas, 1)
        baris_input.addWidget(self.btn_tambah_kelas)
        self.list_kelas = QListWidget()
        self.btn_ubah_kelas = QPushButton("Ubah Nama")
        self.btn_ubah_kelas.setObjectName("netral")
        self.btn_ubah_kelas.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_ubah_kelas.clicked.connect(self.ubah_kelas)
        self.btn_hapus_kelas = QPushButton("Hapus Kelas")
        self.btn_hapus_kelas.setObjectName("bahaya")
        self.btn_hapus_kelas.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_hapus_kelas.clicked.connect(self.hapus_kelas)
        baris_tombol_kelas = QHBoxLayout()
        baris_tombol_kelas.setSpacing(8)
        baris_tombol_kelas.addWidget(self.btn_ubah_kelas, 1)
        baris_tombol_kelas.addWidget(self.btn_hapus_kelas, 1)
        kolom.addWidget(label)
        kolom.addLayout(baris_input)
        kolom.addWidget(self.list_kelas, 1)
        kolom.addLayout(baris_tombol_kelas)
        self.list_kelas.currentRowChanged.connect(lambda _: self.muat_siswa())
        return kartu

    def _panel_siswa(self):
        kartu = QFrame()
        kartu.setObjectName("kartu")
        kolom = QVBoxLayout(kartu)
        kolom.setContentsMargins(18, 16, 18, 16)
        kolom.setSpacing(10)
        self.label_siswa_judul = QLabel("Siswa")
        self.label_siswa_judul.setObjectName("bagian")
        baris_input = QHBoxLayout()
        baris_input.setSpacing(8)
        self.input_siswa = QLineEdit()
        self.input_siswa.setPlaceholderText("Nama siswa...")
        self.input_siswa.returnPressed.connect(self.tambah_siswa)
        self.btn_tambah_siswa = QPushButton("Tambah")
        self.btn_tambah_siswa.setObjectName("utama")
        self.btn_tambah_siswa.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_tambah_siswa.clicked.connect(self.tambah_siswa)
        self.btn_banyak = QPushButton("+ Banyak")
        self.btn_banyak.setObjectName("netral")
        self.btn_banyak.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_banyak.clicked.connect(self.tambah_banyak)
        baris_input.addWidget(self.input_siswa, 1)
        baris_input.addWidget(self.btn_tambah_siswa)
        baris_input.addWidget(self.btn_banyak)
        self.list_siswa = QListWidget()
        self.btn_ubah_siswa = QPushButton("Ubah Nama")
        self.btn_ubah_siswa.setObjectName("netral")
        self.btn_ubah_siswa.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_ubah_siswa.clicked.connect(self.ubah_siswa)
        self.btn_hapus_siswa = QPushButton("Hapus Siswa")
        self.btn_hapus_siswa.setObjectName("bahaya")
        self.btn_hapus_siswa.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_hapus_siswa.clicked.connect(self.hapus_siswa)
        baris_tombol_siswa = QHBoxLayout()
        baris_tombol_siswa.setSpacing(8)
        baris_tombol_siswa.addWidget(self.btn_ubah_siswa, 1)
        baris_tombol_siswa.addWidget(self.btn_hapus_siswa, 1)
        kolom.addWidget(self.label_siswa_judul)
        kolom.addLayout(baris_input)
        kolom.addWidget(self.list_siswa, 1)
        kolom.addLayout(baris_tombol_siswa)
        return kartu

    def kelas_terpilih(self):
        item = self.list_kelas.currentItem()
        if not item:
            return None
        return item.data(Qt.ItemDataRole.UserRole)

    def muat_kelas(self, kelas_list):
        terpilih = self.kelas_terpilih()
        self.list_kelas.clear()
        for kid, nama in kelas_list:
            item = QListWidgetItem(nama)
            item.setData(Qt.ItemDataRole.UserRole, kid)
            self.list_kelas.addItem(item)
            if kid == terpilih:
                self.list_kelas.setCurrentItem(item)
        if self.list_kelas.currentRow() < 0 and self.list_kelas.count() > 0:
            self.list_kelas.setCurrentRow(0)
        self.muat_siswa()

    def muat_siswa(self):
        kid = self.kelas_terpilih()
        self.list_siswa.clear()
        item = self.list_kelas.currentItem()
        nama_kelas = item.text() if item else "-"
        self.label_siswa_judul.setText(f"Siswa — {nama_kelas}")
        if kid is None:
            return
        conn = get_conn()
        try:
            rows = conn.execute(
                "SELECT id, nama FROM siswa WHERE kelas_id = ? ORDER BY nama",
                (kid,),
            ).fetchall()
        finally:
            conn.close()
        for r in rows:
            item = QListWidgetItem(r["nama"])
            item.setData(Qt.ItemDataRole.UserRole, r["id"])
            self.list_siswa.addItem(item)

    def tambah_kelas(self):
        nama = self.input_kelas.text().strip()
        if not nama:
            QMessageBox.warning(self, "Peringatan", "Nama kelas tidak boleh kosong!")
            return
        conn = get_conn()
        kid_baru = None
        try:
            ada = conn.execute(
                "SELECT 1 FROM kelas WHERE nama_kelas = ?", (nama,)
            ).fetchone()
            if not ada:
                cur = conn.execute("INSERT INTO kelas (nama_kelas) VALUES (?)", (nama,))
                kid_baru = cur.lastrowid
                conn.commit()
        except sqlite3.Error:
            QMessageBox.critical(self, "Gagal", "Terjadi kesalahan saat menambah kelas.")
            return
        finally:
            conn.close()
        if kid_baru is None:
            QMessageBox.warning(self, "Peringatan", f"Kelas '{nama}' sudah ada!")
            return
        self.input_kelas.clear()
        self.utama.segarkan_kelas()
        for i in range(self.list_kelas.count()):
            if self.list_kelas.item(i).data(Qt.ItemDataRole.UserRole) == kid_baru:
                self.list_kelas.setCurrentRow(i)
                break

    def hapus_kelas(self):
        kid = self.kelas_terpilih()
        if kid is None:
            QMessageBox.warning(self, "Peringatan", "Pilih kelas yang mau dihapus!")
            return
        nama = self.list_kelas.currentItem().text()
        conn = get_conn()
        try:
            n = conn.execute(
                "SELECT COUNT(*) AS n FROM siswa WHERE kelas_id = ?", (kid,)
            ).fetchone()["n"]
        except sqlite3.Error:
            QMessageBox.critical(self, "Gagal", "Terjadi kesalahan saat menghapus kelas.")
            return
        finally:
            conn.close()
        pesan = (
            f"Hapus kelas '{nama}' beserta {n} siswa di dalamnya\ndan seluruh data presensinya?"
            if n
            else f"Hapus kelas '{nama}'?"
        )
        jawab = QMessageBox.question(
            self,
            "Konfirmasi",
            pesan,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if jawab != QMessageBox.StandardButton.Yes:
            return
        conn = get_conn()
        try:
            conn.execute("DELETE FROM kelas WHERE id = ?", (kid,))
            conn.commit()
        except sqlite3.Error:
            QMessageBox.critical(self, "Gagal", "Terjadi kesalahan saat menghapus kelas.")
        finally:
            conn.close()
        self.utama.segarkan_kelas()

    def ubah_kelas(self):
        item = self.list_kelas.currentItem()
        if not item:
            QMessageBox.warning(self, "Peringatan", "Pilih kelas yang mau diubah!")
            return
        kid = item.data(Qt.ItemDataRole.UserRole)
        lama = item.text()
        baru, ok = QInputDialog.getText(
            self, "Ubah Nama Kelas", "Nama kelas baru:", text=lama
        )
        baru = baru.strip() if ok else ""
        if not ok or not baru or baru == lama:
            return
        conn = get_conn()
        try:
            ada = conn.execute(
                "SELECT 1 FROM kelas WHERE nama_kelas = ?", (baru,)
            ).fetchone()
            if ada:
                QMessageBox.warning(self, "Peringatan", f"Kelas '{baru}' sudah ada!")
                return
            conn.execute("UPDATE kelas SET nama_kelas = ? WHERE id = ?", (baru, kid))
            conn.commit()
        except sqlite3.Error:
            QMessageBox.critical(self, "Gagal", "Terjadi kesalahan saat mengubah kelas.")
            return
        finally:
            conn.close()
        self.utama.segarkan_kelas()
        for i in range(self.list_kelas.count()):
            if self.list_kelas.item(i).data(Qt.ItemDataRole.UserRole) == kid:
                self.list_kelas.setCurrentRow(i)
                break

    def tambah_siswa(self):
        kid = self.kelas_terpilih()
        if kid is None:
            QMessageBox.warning(self, "Peringatan", "Pilih kelas dulu!")
            return
        nama = self.input_siswa.text().strip()
        if not nama:
            QMessageBox.warning(self, "Peringatan", "Nama siswa tidak boleh kosong!")
            return
        conn = get_conn()
        try:
            ada = conn.execute(
                "SELECT 1 FROM siswa WHERE kelas_id = ? AND nama = ?", (kid, nama)
            ).fetchone()
            if not ada:
                conn.execute(
                    "INSERT INTO siswa (kelas_id, nama) VALUES (?, ?)", (kid, nama)
                )
                conn.commit()
        except sqlite3.Error:
            QMessageBox.critical(self, "Gagal", "Terjadi kesalahan saat menambah siswa.")
            return
        finally:
            conn.close()
        if ada:
            QMessageBox.warning(
                self, "Peringatan", f"'{nama}' sudah ada di kelas ini!"
            )
            return
        self.input_siswa.clear()
        self.muat_siswa()
        self.utama.halaman_absensi.muat_presensi()

    def tambah_banyak(self):
        kid = self.kelas_terpilih()
        if kid is None:
            QMessageBox.warning(self, "Peringatan", "Pilih kelas dulu!")
            return
        dialog = DialogTambahBanyak(self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        nama_list = dialog.daftar_nama()
        if not nama_list:
            return
        conn = get_conn()
        try:
            sudah = {
                r["nama"]
                for r in conn.execute(
                    "SELECT nama FROM siswa WHERE kelas_id = ?", (kid,)
                ).fetchall()
            }
            baru = []
            duplikat = 0
            for nama in nama_list:
                if nama in sudah:
                    duplikat += 1
                    continue
                sudah.add(nama)
                baru.append((kid, nama))
            conn.executemany(
                "INSERT INTO siswa (kelas_id, nama) VALUES (?, ?)", baru
            )
            conn.commit()
        finally:
            conn.close()
        self.muat_siswa()
        self.utama.halaman_absensi.muat_presensi()
        pesan = f"{len(baru)} siswa ditambahkan."
        if duplikat:
            pesan += f"\n{duplikat} nama dilewati karena sudah ada."
        QMessageBox.information(self, "Sukses", pesan)

    def hapus_siswa(self):
        item = self.list_siswa.currentItem()
        if not item:
            QMessageBox.warning(self, "Peringatan", "Pilih siswa yang mau dihapus!")
            return
        sid = item.data(Qt.ItemDataRole.UserRole)
        nama = item.text()
        jawab = QMessageBox.question(
            self,
            "Konfirmasi",
            f"Hapus siswa '{nama}' beserta semua data presensinya?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if jawab != QMessageBox.StandardButton.Yes:
            return
        conn = get_conn()
        try:
            conn.execute("DELETE FROM siswa WHERE id = ?", (sid,))
            conn.commit()
        finally:
            conn.close()
        self.muat_siswa()
        self.utama.halaman_absensi.muat_presensi()

    def ubah_siswa(self):
        item = self.list_siswa.currentItem()
        if not item:
            QMessageBox.warning(self, "Peringatan", "Pilih siswa yang mau diubah!")
            return
        sid = item.data(Qt.ItemDataRole.UserRole)
        lama = item.text()
        baru, ok = QInputDialog.getText(
            self, "Ubah Nama Siswa", "Nama siswa baru:", text=lama
        )
        baru = baru.strip() if ok else ""
        if not ok or not baru or baru == lama:
            return
        kid = self.kelas_terpilih()
        conn = get_conn()
        try:
            ada = conn.execute(
                "SELECT 1 FROM siswa WHERE kelas_id = ? AND nama = ? AND id != ?",
                (kid, baru, sid),
            ).fetchone()
            if not ada:
                conn.execute("UPDATE siswa SET nama = ? WHERE id = ?", (baru, sid))
                conn.commit()
        except sqlite3.Error:
            QMessageBox.critical(self, "Gagal", "Terjadi kesalahan saat mengubah siswa.")
            return
        finally:
            conn.close()
        if ada:
            QMessageBox.warning(
                self, "Peringatan", f"'{baru}' sudah ada di kelas ini!"
            )
            return
        self.muat_siswa()
        self.utama.halaman_absensi.muat_presensi()


class HalamanRekap(QWidget):
    def __init__(self, utama):
        super().__init__()
        self.utama = utama
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 20)
        root.setSpacing(14)

        kepala = QVBoxLayout()
        judul = QLabel("Rekap Absensi")
        judul.setObjectName("judul_halaman")
        sub = QLabel("Lihat riwayat presensi, filter sesuai kebutuhan, lalu export ke Excel/CSV.")
        sub.setObjectName("subjudul_halaman")
        kepala.addWidget(judul)
        kepala.addWidget(sub)
        root.addLayout(kepala)

        kartu_alat = QFrame()
        kartu_alat.setObjectName("kartu")
        baris = QHBoxLayout(kartu_alat)
        baris.setContentsMargins(18, 14, 18, 14)
        baris.setSpacing(10)
        label_kelas = QLabel("Kelas")
        label_kelas.setObjectName("bagian")
        self.combo_kelas = QComboBox()
        self.combo_kelas.setMinimumWidth(170)
        label_dari = QLabel("Dari")
        label_dari.setObjectName("bagian")
        self.dari = KalenderMerahPopup()
        self.dari.setDisplayFormat("yyyy-MM-dd")
        self.dari.setDate(QDate.currentDate().addDays(-30))
        self.dari.setMinimumWidth(158)
        label_sampai = QLabel("Sampai")
        label_sampai.setObjectName("bagian")
        self.sampai = KalenderMerahPopup()
        self.sampai.setDisplayFormat("yyyy-MM-dd")
        self.sampai.setDate(QDate.currentDate())
        self.sampai.setMinimumWidth(158)
        label_status = QLabel("Status")
        label_status.setObjectName("bagian")
        self.combo_status = QComboBox()
        self.combo_status.addItem("Semua Status", None)
        for jenis in STATUS_LIST:
            self.combo_status.addItem(jenis, jenis)
        self.combo_status.setMinimumWidth(150)
        btn_export = QPushButton("Export Excel")
        btn_export.setObjectName("utama")
        btn_export.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_export.clicked.connect(self.export_excel)
        btn_hapus_semua = QPushButton("Hapus Semua")
        btn_hapus_semua.setObjectName("bahaya")
        btn_hapus_semua.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_hapus_semua.clicked.connect(self.hapus_semua_data)
        for w in (
            label_kelas, self.combo_kelas, label_dari, self.dari,
            label_sampai, self.sampai, label_status, self.combo_status,
        ):
            baris.addWidget(w)
        baris.addStretch()
        baris.addWidget(btn_export)
        baris.addWidget(btn_hapus_semua)
        root.addWidget(kartu_alat)

        kartu_data = QFrame()
        kartu_data.setObjectName("kartu")
        kolom = QVBoxLayout(kartu_data)
        kolom.setContentsMargins(18, 16, 18, 16)
        kolom.setSpacing(12)

        header_data = QFrame()
        header_data.setObjectName("card_header")
        baris_header_data = QHBoxLayout(header_data)
        baris_header_data.setContentsMargins(0, 0, 0, 0)
        baris_header_data.setSpacing(8)
        judul_data = QLabel("Detail Presensi")
        judul_data.setObjectName("card_judul")
        self.badge_data = QLabel("0 catatan")
        self.badge_data.setObjectName("card_badge")
        baris_header_data.addWidget(judul_data)
        baris_header_data.addStretch()
        baris_header_data.addWidget(self.badge_data)
        kolom.addWidget(header_data)

        frame_cari = QFrame()
        frame_cari.setObjectName("search_frame")
        baris_cari = QHBoxLayout(frame_cari)
        baris_cari.setContentsMargins(12, 6, 12, 6)
        baris_cari.setSpacing(8)
        label_cari = QLabel("Cari nama")
        label_cari.setObjectName("search_label")
        self.input_cari = QLineEdit()
        self.input_cari.setPlaceholderText("Ketik nama siswa...")
        self.input_cari.setClearButtonEnabled(True)
        self.input_cari.textChanged.connect(lambda _: self.muat_rekap())
        baris_cari.addWidget(label_cari)
        baris_cari.addWidget(self.input_cari, 1)
        kolom.addWidget(frame_cari)

        self.tabel = QTableWidget()
        self.tabel.setColumnCount(4)
        self.tabel.setHorizontalHeaderLabels(["Tanggal", "Kelas", "Nama", "Status"])
        self.tabel.verticalHeader().setVisible(False)
        self.tabel.verticalHeader().setDefaultSectionSize(36)
        self.tabel.setAlternatingRowColors(True)
        self.tabel.setShowGrid(False)
        self.tabel.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.tabel.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tabel.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tabel.horizontalHeader().setHighlightSections(False)
        self.tabel.setColumnWidth(2, 260)
        kolom.addWidget(self.tabel, 1)
        root.addWidget(kartu_data, 3)

        self.combo_kelas.currentIndexChanged.connect(lambda _: self.muat_rekap())
        self.combo_status.currentIndexChanged.connect(lambda _: self.muat_rekap())
        self.dari.dateChanged.connect(lambda _: self.muat_rekap())
        self.sampai.dateChanged.connect(lambda _: self.muat_rekap())
        tinggi_seragam(
            self.combo_kelas, self.dari, self.sampai, self.combo_status,
            btn_export, btn_hapus_semua,
        )

    def muat_combo(self, kelas_list):
        self.combo_kelas.blockSignals(True)
        self.combo_kelas.clear()
        self.combo_kelas.addItem("Semua Kelas", None)
        for kid, nama in kelas_list:
            self.combo_kelas.addItem(nama, kid)
        self.combo_kelas.blockSignals(False)
        self.muat_rekap()

    def muat_rekap(self):
        kelas_id = self.combo_kelas.currentData()
        status = self.combo_status.currentData()
        tgl_awal = self.dari.date().toString("yyyy-MM-dd")
        tgl_akhir = self.sampai.date().toString("yyyy-MM-dd")

        if tgl_awal > tgl_akhir:
            self.tabel.setRowCount(0)
            self.badge_data.setText("0 catatan")
            return

        kata = self.input_cari.text().strip()

        kondisi = ["p.tanggal BETWEEN ? AND ?"]
        params = [tgl_awal, tgl_akhir]
        if kelas_id is not None:
            kondisi.append("s.kelas_id = ?")
            params.append(kelas_id)
        if status is not None:
            kondisi.append("p.status = ?")
            params.append(status)
        if kata:
            kondisi.append("s.nama LIKE ?")
            params.append(f"%{kata}%")
        where = " AND ".join(kondisi)

        dasar = """
            FROM presensi p
            JOIN siswa s ON s.id = p.siswa_id
            JOIN kelas k ON k.id = s.kelas_id
            WHERE
        """
        conn = get_conn()
        try:
            total = conn.execute(
                "SELECT COUNT(*) AS n " + dasar + where, params
            ).fetchone()["n"]
            rows = conn.execute(
                "SELECT p.tanggal, k.nama_kelas, s.nama, p.status "
                + dasar + where
                + " ORDER BY p.tanggal DESC, k.nama_kelas, s.nama",
                params,
            ).fetchall()
        finally:
            conn.close()

        self.tabel.setRowCount(0)
        for data in rows:
            r = self.tabel.rowCount()
            self.tabel.insertRow(r)
            nilai = [data["tanggal"], data["nama_kelas"], data["nama"], data["status"]]
            for c, teks in enumerate(nilai):
                item = QTableWidgetItem(teks)
                if c == 3:
                    item.setForeground(QColor(STATUS_WARNA.get(teks, "#31513D")))
                elif c != 2:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tabel.setItem(r, c, item)
        self.badge_data.setText(f"{total} catatan")
    def hapus_semua_data(self):
        jawab = QMessageBox.question(
            self,
            "Konfirmasi",
            "HAPUS SEMUA data presensi?\n\n"
            "Seluruh riwayat absensi akan dihapus permanen.\n"
            "Tindakan ini tidak dapat dibatalkan!",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if jawab != QMessageBox.StandardButton.Yes:
            return
        hapus_semua_presensi()
        self.utama.segarkan_kelas()
        QMessageBox.information(self, "Sukses", "Semua data presensi berhasil dihapus.")

    def export_excel(self):
        if self.tabel.rowCount() == 0:
            QMessageBox.information(self, "Info", "Tidak ada data untuk diexport.")
            return
        path, terpilih = QFileDialog.getSaveFileName(
            self,
            "Export Excel",
            "rekap_presensi.xlsx",
            "File Excel (*.xlsx);;File CSV (*.csv)",
        )
        if not path:
            return

        data = [
            [self.tabel.item(r, c).text() for c in range(4)]
            for r in range(self.tabel.rowCount())
        ]

        if path.lower().endswith(".csv"):
            with open(path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["Tanggal", "Kelas", "Nama", "Status"])
                writer.writerows(data)
        else:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            self._tulis_excel(path, data)

        pesan = f"Data berhasil diexport ({len(data)} baris) ke:\n{path}"
        QMessageBox.information(self, "Sukses", pesan)

    def _tulis_excel(self, path, data):
        wb = Workbook()
        ws = wb.active
        ws.title = "Rekap Presensi"

        lebar_kolom = [13, 20, 30, 12]
        for i, lebar in enumerate(lebar_kolom, start=1):
            ws.column_dimensions[get_column_letter(i)].width = lebar

        tipis = Side(style="thin", color="BFE2CE")
        garis = Border(left=tipis, right=tipis, top=tipis, bottom=tipis)
        tengah = Alignment(horizontal="center", vertical="center")
        kiri = Alignment(horizontal="left", vertical="center")

        header_isi = ["Tanggal", "Kelas", "Nama", "Status"]
        header_fill = PatternFill("solid", fgColor="16A34A")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        for c, judul in enumerate(header_isi, start=1):
            sel = ws.cell(row=1, column=c, value=judul)
            sel.fill = header_fill
            sel.font = header_font
            sel.border = garis
            sel.alignment = tengah
        ws.row_dimensions[1].height = 24

        gaya_status = {
            "Hadir": ("D1FAE5", "065F46"),
            "Alpa": ("FEE2E2", "991B1B"),
            "Izin": ("DBEAFE", "1E40AF"),
            "Sakit": ("FEF3C7", "92400E"),
        }
        strip_fill = PatternFill("solid", fgColor="EDF8F2")

        for r, baris_data in enumerate(data, start=2):
            berstrip = (r % 2) == 0
            for c, nilai in enumerate(baris_data, start=1):
                sel = ws.cell(row=r, column=c, value=nilai)
                sel.border = garis
                sel.alignment = tengah if c in (1, 4) else kiri
                if berstrip and c != 4:
                    sel.fill = strip_fill
            status = baris_data[3]
            if status in gaya_status:
                warna_bg, warna_teks = gaya_status[status]
                sel_status = ws.cell(row=r, column=4)
                sel_status.fill = PatternFill("solid", fgColor=warna_bg)
                sel_status.font = Font(bold=True, color=warna_teks)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:D{max(len(data) + 1, 2)}"

        wb.save(path)


class HalamanKesimpulanAbsensi(QWidget):
    def __init__(self, utama):
        super().__init__()
        self.utama = utama
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 20)
        root.setSpacing(14)

        kepala = QVBoxLayout()
        judul = QLabel("Kesimpulan Absensi")
        judul.setObjectName("judul_halaman")
        sub = QLabel("Rekap jumlah kehadiran tiap siswa (Hadir, Alpa, Izin, Sakit).")
        sub.setObjectName("subjudul_halaman")
        kepala.addWidget(judul)
        kepala.addWidget(sub)
        root.addLayout(kepala)

        kartu_alat = QFrame()
        kartu_alat.setObjectName("kartu")
        baris = QHBoxLayout(kartu_alat)
        baris.setContentsMargins(18, 14, 18, 14)
        baris.setSpacing(10)
        label_kelas = QLabel("Kelas")
        label_kelas.setObjectName("bagian")
        self.combo_kelas = QComboBox()
        self.combo_kelas.setMinimumWidth(170)
        label_dari = QLabel("Dari")
        label_dari.setObjectName("bagian")
        self.dari = KalenderMerahPopup()
        self.dari.setDisplayFormat("yyyy-MM-dd")
        self.dari.setDate(QDate.currentDate().addDays(-30))
        self.dari.setMinimumWidth(158)
        label_sampai = QLabel("Sampai")
        label_sampai.setObjectName("bagian")
        self.sampai = KalenderMerahPopup()
        self.sampai.setDisplayFormat("yyyy-MM-dd")
        self.sampai.setDate(QDate.currentDate())
        self.sampai.setMinimumWidth(158)
        label_status = QLabel("Status")
        label_status.setObjectName("bagian")
        self.combo_status = QComboBox()
        self.combo_status.addItem("Semua Status", None)
        for jenis in STATUS_LIST:
            self.combo_status.addItem(jenis, jenis)
        self.combo_status.setMinimumWidth(150)
        btn_export = QPushButton("Export Excel")
        btn_export.setObjectName("utama")
        btn_export.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_export.clicked.connect(self.export_excel)
        for w in (
            label_kelas, self.combo_kelas, label_dari, self.dari,
            label_sampai, self.sampai, label_status, self.combo_status,
        ):
            baris.addWidget(w)
        baris.addStretch()
        baris.addWidget(btn_export)
        root.addWidget(kartu_alat)

        self.tabel = QTableWidget()
        self.tabel.setColumnCount(6)
        self.tabel.setHorizontalHeaderLabels(
            ["Kelas", "Nama", "Hadir", "Alpa", "Izin", "Sakit"]
        )
        self.tabel.verticalHeader().setVisible(False)
        self.tabel.verticalHeader().setDefaultSectionSize(36)
        self.tabel.setAlternatingRowColors(True)
        self.tabel.setShowGrid(False)
        self.tabel.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.tabel.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tabel.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tabel.horizontalHeader().setHighlightSections(False)
        self.tabel.setColumnWidth(0, 180)
        self.tabel.setColumnWidth(1, 260)
        root.addWidget(self.tabel, 1)

        baris_bawah = QHBoxLayout()
        self.label_total = QLabel("0 siswa")
        self.label_total.setObjectName("total")
        baris_bawah.addWidget(self.label_total)
        baris_bawah.addStretch()
        root.addLayout(baris_bawah)

        self.combo_kelas.currentIndexChanged.connect(lambda _: self.muat_rekap())
        self.combo_status.currentIndexChanged.connect(lambda _: self.muat_rekap())
        self.dari.dateChanged.connect(lambda _: self.muat_rekap())
        self.sampai.dateChanged.connect(lambda _: self.muat_rekap())
        tinggi_seragam(
            self.combo_kelas, self.dari, self.sampai, self.combo_status, btn_export,
        )

    def muat_combo(self, kelas_list):
        self.combo_kelas.blockSignals(True)
        self.combo_kelas.clear()
        self.combo_kelas.addItem("Semua Kelas", None)
        for kid, nama in kelas_list:
            self.combo_kelas.addItem(nama, kid)
        self.combo_kelas.blockSignals(False)
        self.muat_rekap()

    def reset_tanggal(self):
        self.dari.blockSignals(True)
        self.sampai.blockSignals(True)
        self.dari.setDate(QDate.currentDate().addDays(-30))
        self.sampai.setDate(QDate.currentDate())
        self.dari.blockSignals(False)
        self.sampai.blockSignals(False)
        self.muat_rekap()

    def muat_rekap(self):
        kelas_id = self.combo_kelas.currentData()
        status = self.combo_status.currentData()
        tgl_awal = self.dari.date().toString("yyyy-MM-dd")
        tgl_akhir = self.sampai.date().toString("yyyy-MM-dd")

        if tgl_awal > tgl_akhir:
            self.tabel.setRowCount(0)
            self.label_total.setText("0 siswa")
            return

        kondisi = ["p.tanggal BETWEEN ? AND ?"]
        params = [tgl_awal, tgl_akhir]
        if kelas_id is not None:
            kondisi.append("s.kelas_id = ?")
            params.append(kelas_id)
        if status is not None:
            kondisi.append("p.status = ?")
            params.append(status)
        where = " AND ".join(kondisi)

        sql = (
            """
            SELECT k.nama_kelas, s.nama, p.status
            FROM presensi p
            JOIN siswa s ON s.id = p.siswa_id
            JOIN kelas k ON k.id = s.kelas_id
            WHERE
            """
            + where
            + " ORDER BY k.nama_kelas, s.nama"
        )
        conn = get_conn()
        try:
            rows = conn.execute(sql, params).fetchall()
        finally:
            conn.close()

        ringkas = {}
        for r in rows:
            key = (r["nama_kelas"], r["nama"])
            ringkas.setdefault(key, {j: 0 for j in STATUS_LIST})[r["status"]] += 1

        self.tabel.setRowCount(0)
        for (nama_kelas, nama), hitung in ringkas.items():
            row = self.tabel.rowCount()
            self.tabel.insertRow(row)
            nilai = [nama_kelas, nama]
            nilai += [hitung[j] for j in STATUS_LIST]
            for c, teks in enumerate(nilai):
                item = QTableWidgetItem(str(teks))
                if c >= 2:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    item.setForeground(QColor(STATUS_WARNA.get(STATUS_LIST[c - 2], "#31513D")))
                elif c != 1:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tabel.setItem(row, c, item)
        self.label_total.setText(f"{len(ringkas)} siswa")

    def export_excel(self):
        if self.tabel.rowCount() == 0:
            QMessageBox.information(self, "Info", "Tidak ada data untuk diexport.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Export Excel", "kesimpulan_absensi.xlsx",
            "File Excel (*.xlsx);;File CSV (*.csv)",
        )
        if not path:
            return
        data = [
            [self.tabel.item(r, c).text() for c in range(6)]
            for r in range(self.tabel.rowCount())
        ]
        header = ["Kelas", "Nama", "Hadir", "Alpa", "Izin", "Sakit"]
        if path.lower().endswith(".csv"):
            with open(path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(header)
                writer.writerows(data)
        else:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Kesimpulan Absensi"
            for i, w in enumerate([20, 30, 10, 10, 10, 10], start=1):
                ws.column_dimensions[get_column_letter(i)].width = w

            tipis = Side(style="thin", color="BFE2CE")
            garis = Border(left=tipis, right=tipis, top=tipis, bottom=tipis)
            tengah = Alignment(horizontal="center", vertical="center")
            kiri = Alignment(horizontal="left", vertical="center")

            sel = ws.cell(row=1, column=1, value="KESIMPULAN ABSENSI PER SISWA")
            sel.font = Font(bold=True, size=13, color="065F46")
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
            for c, h in enumerate(header, start=1):
                cell = ws.cell(row=2, column=c, value=h)
                cell.fill = PatternFill("solid", fgColor="16A34A")
                cell.font = Font(bold=True, color="FFFFFF")
                cell.border = garis
                cell.alignment = tengah if c != 2 else kiri
            ws.row_dimensions[2].height = 24

            gaya_status = {
                "Hadir": ("D1FAE5", "065F46"),
                "Alpa": ("FEE2E2", "991B1B"),
                "Izin": ("DBEAFE", "1E40AF"),
                "Sakit": ("FEF3C7", "92400E"),
            }
            strip_fill = PatternFill("solid", fgColor="EDF8F2")
            for r, d in enumerate(data, start=3):
                berstrip = (r % 2) == 0
                for c, v in enumerate(d, start=1):
                    cell = ws.cell(row=r, column=c, value=v)
                    cell.border = garis
                    cell.alignment = tengah if c != 2 else kiri
                    if berstrip:
                        cell.fill = strip_fill
                for i, status in enumerate(STATUS_LIST, start=3):
                    if status in gaya_status:
                        warna_bg, warna_teks = gaya_status[status]
                        sel_status = ws.cell(row=r, column=i)
                        sel_status.fill = PatternFill("solid", fgColor=warna_bg)
                        sel_status.font = Font(bold=True, color=warna_teks)
            ws.freeze_panes = "A3"
            ws.auto_filter.ref = f"A2:F{max(len(data) + 2, 3)}"
            wb.save(path)
        QMessageBox.information(
            self, "Sukses",
            f"Kesimpulan absensi ({len(data)} siswa) berhasil diexport ke:\n{path}",
        )


class HalamanNilai(QWidget):
    def __init__(self, utama):
        super().__init__()
        self.utama = utama
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 20)
        root.setSpacing(14)

        kepala = QVBoxLayout()
        judul = QLabel("Input Nilai")
        judul.setObjectName("judul_halaman")
        sub = QLabel(
            "Pilih kelas, jenis nilai, dan tanggal sesuai keinginan. "
            "Isi nilai tiap siswa lalu klik \"Simpan Nilai\" - data otomatis masuk ke Rekap Nilai & Kesimpulan Nilai."
        )
        sub.setObjectName("subjudul_halaman")
        kepala.addWidget(judul)
        kepala.addWidget(sub)
        root.addLayout(kepala)

        kartu_alat = QFrame()
        kartu_alat.setObjectName("kartu")
        baris = QHBoxLayout(kartu_alat)
        baris.setContentsMargins(18, 14, 18, 14)
        baris.setSpacing(10)
        label_kelas = QLabel("Kelas")
        label_kelas.setObjectName("bagian")
        self.combo_kelas = QComboBox()
        self.combo_kelas.setMinimumWidth(170)
        label_jenis = QLabel("Jenis")
        label_jenis.setObjectName("bagian")
        self.combo_jenis = QComboBox()
        for j in JENIS_NILAI:
            self.combo_jenis.addItem(j)
        self.combo_jenis.setMinimumWidth(130)
        label_tanggal = QLabel("Tanggal")
        label_tanggal.setObjectName("bagian")
        self.date_tanggal = KalenderMerahPopup()
        self.date_tanggal.setDisplayFormat("yyyy-MM-dd")
        self.date_tanggal.setCalendarPopup(True)
        self.date_tanggal.setDate(QDate.currentDate())
        self.date_tanggal.setMinimumWidth(158)
        btn_simpan = QPushButton("Simpan Nilai")
        btn_simpan.setObjectName("utama")
        btn_simpan.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_simpan.clicked.connect(self.simpan_nilai)
        for w in (
            label_kelas, self.combo_kelas, label_jenis, self.combo_jenis,
            label_tanggal, self.date_tanggal,
        ):
            baris.addWidget(w)
        baris.addStretch()
        baris.addWidget(btn_simpan)
        root.addWidget(kartu_alat)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.konten = QWidget()
        self.konten.setObjectName("konten_absen")
        self.layout_kartu = QVBoxLayout(self.konten)
        self.layout_kartu.setContentsMargins(0, 0, 8, 0)
        self.layout_kartu.setSpacing(10)
        self.layout_kartu.addStretch()
        self.scroll.setWidget(self.konten)
        root.addWidget(self.scroll, 1)

        self.combo_kelas.currentIndexChanged.connect(lambda _: self.muat_nilai())
        self.combo_jenis.currentIndexChanged.connect(lambda _: self.muat_nilai())
        self.date_tanggal.dateChanged.connect(lambda _: self.muat_nilai())
        self.peta_spin = {}
        tinggi_seragam(
            self.combo_kelas, self.combo_jenis,
            self.date_tanggal, btn_simpan,
        )

    def kelas_id_aktif(self):
        return self.combo_kelas.currentData()

    def muat_combo(self, kelas_list):
        self.combo_kelas.blockSignals(True)
        self.combo_kelas.clear()
        self.combo_kelas.addItem("-- Pilih Kelas --", None)
        for kid, nama in kelas_list:
            self.combo_kelas.addItem(nama, kid)
        self.combo_kelas.blockSignals(False)
        self.muat_nilai()

    def reset_tanggal(self):
        self.date_tanggal.blockSignals(True)
        self.date_tanggal.setDate(QDate.currentDate())
        self.date_tanggal.blockSignals(False)
        self.muat_nilai()

    def muat_nilai(self):
        while self.layout_kartu.count() > 1:
            item = self.layout_kartu.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        kid = self.kelas_id_aktif()
        if kid is None:
            self._tampil_kosong("Pilih kelas terlebih dahulu.")
            return
        jenis = self.combo_jenis.currentText()
        self._muat_isi(kid, jenis)

    def _muat_isi(self, kid, jenis):
        tanggal = self.date_tanggal.date().toString("yyyy-MM-dd")
        siswa = ambil_nilai_tanggal(kid, jenis, tanggal)
        if not siswa:
            self._tampil_kosong("Kelas ini masih kosong atau belum ada data.")
            return

        kartu = QFrame()
        kartu.setObjectName("nilai_kartu")
        pasang_shadow(kartu, blur=24, alfa=35, offset_y=3)
        kolom = QVBoxLayout(kartu)
        kolom.setContentsMargins(18, 16, 18, 16)
        kolom.setSpacing(10)

        header = QHBoxLayout()
        badge = QLabel(jenis)
        badge.setProperty("jenis", jenis)
        badge.setObjectName("nilai_jenis")
        badge.setStyleSheet(
            f"background: {JENIS_NILAI_WARNA[jenis]}; color: #FFFFFF;"
        )
        lab_tgl = QLabel(f"Nilai {jenis} - {tanggal}")
        lab_tgl.setObjectName("card_judul")
        header.addWidget(badge)
        header.addWidget(lab_tgl)
        header.addStretch()
        kolom.addLayout(header)

        jumlah_terisi = sum(1 for d in siswa if d["nilai"] is not None)
        info = QLabel(
            f"{jumlah_terisi} dari {len(siswa)} siswa sudah punya nilai di tanggal ini. "
            "Nilai yang tersimpan otomatis termuat dan bisa diubah."
            if jumlah_terisi
            else f"Belum ada nilai untuk tanggal {tanggal}. "
                 "Isi nilai tiap siswa di bawah lalu klik \"Simpan Nilai\"."
        )
        info.setObjectName("total")
        info.setWordWrap(True)
        kolom.addWidget(info)

        self.peta_spin = {}
        for data in siswa:
            baris_s = QHBoxLayout()
            baris_s.setSpacing(10)
            inisial = "".join(k[0] for k in data["nama"].split()[:2]).upper()
            avatar = QLabel(inisial or "?")
            avatar.setFixedSize(34, 34)
            avatar.setAlignment(Qt.AlignmentFlag.AlignCenter)
            avatar.setStyleSheet(
                "background:#DCFCE7;color:#16A34A;border-radius:17px;"
                "font-weight:800;font-size:12px;"
            )
            nama_label = QLabel(data["nama"])
            nama_label.setStyleSheet("font-size:14px;font-weight:600;color:#13301F;background:transparent;border:none;")
            spin = QDoubleSpinBox()
            spin.setRange(0, 100)
            spin.setDecimals(0)
            spin.setSuffix("  ")
            spin.setFixedWidth(110)
            if data["nilai"] is not None:
                spin.setValue(float(data["nilai"]))
            self.peta_spin[data["id"]] = spin
            baris_s.addWidget(avatar)
            baris_s.addWidget(nama_label)
            baris_s.addStretch()
            baris_s.addWidget(spin)
            kolom.addLayout(baris_s)

        self.layout_kartu.insertWidget(self.layout_kartu.count() - 1, kartu)

    def _tampil_kosong(self, teks):
        label = QLabel(teks)
        label.setObjectName("kosong")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.layout_kartu.insertWidget(0, label, 1, Qt.AlignmentFlag.AlignHCenter)

    def simpan_nilai(self):
        kid = self.kelas_id_aktif()
        if kid is None:
            QMessageBox.warning(self, "Peringatan", "Pilih kelas dulu!")
            return
        jenis = self.combo_jenis.currentText()
        tanggal = self.date_tanggal.date().toString("yyyy-MM-dd")
        if not self.peta_spin:
            return
        try:
            simpan_nilai_batch(
                (sid, jenis, tanggal, spin.value())
                for sid, spin in self.peta_spin.items()
            )
        except sqlite3.Error as e:
            QMessageBox.critical(
                self, "Gagal",
                f"Terjadi kesalahan saat menyimpan nilai:\n{e}",
            )
            return
        QMessageBox.information(
            self, "Sukses",
            f"Nilai {jenis} ({tanggal}) berhasil disimpan untuk kelas terpilih.",
        )
        self.muat_nilai()


class HalamanJadwal(QWidget):
    def __init__(self, utama):
        super().__init__()
        self.utama = utama
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 20)
        root.setSpacing(14)

        kepala = QVBoxLayout()
        judul = QLabel("Jadwal UTS & Ujian")
        judul.setObjectName("judul_halaman")
        sub = QLabel("Atur jadwal Latihan, Ulangan, UTS, dan UAS untuk tiap kelas.")
        sub.setObjectName("subjudul_halaman")
        kepala.addWidget(judul)
        kepala.addWidget(sub)
        root.addLayout(kepala)

        kartu_alat = QFrame()
        kartu_alat.setObjectName("kartu")
        baris = QHBoxLayout(kartu_alat)
        baris.setContentsMargins(18, 14, 18, 14)
        baris.setSpacing(10)
        label_kelas = QLabel("Kelas")
        label_kelas.setObjectName("bagian")
        self.combo_kelas = QComboBox()
        self.combo_kelas.setMinimumWidth(160)
        label_tanggal = QLabel("Tanggal")
        label_tanggal.setObjectName("bagian")
        self.date_tgl = KalenderMerahPopup()
        self.date_tgl.setDisplayFormat("yyyy-MM-dd")
        self.date_tgl.setDate(QDate.currentDate())
        self.date_tgl.setMinimumWidth(158)
        label_jenis = QLabel("Jenis")
        label_jenis.setObjectName("bagian")
        self.combo_jenis = QComboBox()
        for j in JENIS_NILAI:
            self.combo_jenis.addItem(j)
        self.combo_jenis.setMinimumWidth(110)
        label_ket = QLabel("Keterangan")
        label_ket.setObjectName("bagian")
        self.input_ket = QLineEdit()
        self.input_ket.setPlaceholderText("Opsional, misal: Ulangan Harian 1")
        self.input_ket.setMinimumWidth(180)
        btn_tambah = QPushButton("Tambah Jadwal")
        btn_tambah.setObjectName("utama")
        btn_tambah.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_tambah.clicked.connect(self.tambah_jadwal)
        for w in (
            label_kelas, self.combo_kelas, label_tanggal, self.date_tgl,
            label_jenis, self.combo_jenis, label_ket, self.input_ket,
        ):
            baris.addWidget(w)
        baris.addStretch()
        baris.addWidget(btn_tambah)
        root.addWidget(kartu_alat)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.konten = QWidget()
        self.konten.setObjectName("konten_absen")
        self.layout_kartu = QVBoxLayout(self.konten)
        self.layout_kartu.setContentsMargins(0, 0, 8, 0)
        self.layout_kartu.setSpacing(10)
        self.layout_kartu.addStretch()
        self.scroll.setWidget(self.konten)
        root.addWidget(self.scroll, 1)

        self.combo_kelas.currentIndexChanged.connect(lambda _: self.muat_jadwal())
        tinggi_seragam(
            self.combo_kelas, self.date_tgl, self.combo_jenis,
            self.input_ket, btn_tambah,
        )

    def kelas_id_aktif(self):
        return self.combo_kelas.currentData()

    def muat_combo(self, kelas_list):
        self.combo_kelas.blockSignals(True)
        self.combo_kelas.clear()
        self.combo_kelas.addItem("-- Pilih Kelas --", None)
        for kid, nama in kelas_list:
            self.combo_kelas.addItem(nama, kid)
        self.combo_kelas.blockSignals(False)
        self.muat_jadwal()

    def reset_tanggal(self):
        self.date_tgl.blockSignals(True)
        self.date_tgl.setDate(QDate.currentDate())
        self.date_tgl.blockSignals(False)
        self.muat_jadwal()

    def muat_jadwal(self):
        while self.layout_kartu.count() > 1:
            item = self.layout_kartu.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        kid = self.kelas_id_aktif()
        if kid is None:
            label = QLabel("Pilih kelas untuk melihat jadwal.")
            label.setObjectName("kosong")
            label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.layout_kartu.insertWidget(0, label, 1, Qt.AlignmentFlag.AlignHCenter)
            return
        jadwal = ambil_jadwal(kid)
        if not jadwal:
            kosong = QFrame()
            kosong.setObjectName("kosong_jadwal")
            lab = QLabel("Belum ada jadwal untuk kelas ini.\nGunakan form di atas untuk menambahkan.")
            lab.setObjectName("kosong")
            lab.setAlignment(Qt.AlignmentFlag.AlignCenter)
            kosong_layout = QVBoxLayout(kosong)
            kosong_layout.addWidget(lab)
            kosong.setMinimumHeight(120)
            self.layout_kartu.insertWidget(0, kosong)
            return

        for j in jadwal:
            kartu = QFrame()
            kartu.setObjectName("nilai_kartu")
            pasang_shadow(kartu, blur=24, alfa=35, offset_y=3)
            baris = QHBoxLayout(kartu)
            baris.setContentsMargins(16, 12, 16, 12)
            baris.setSpacing(10)
            badge = QLabel(j["jenis"])
            badge.setProperty("jenis", j["jenis"])
            badge.setObjectName("nilai_jenis")
            badge.setStyleSheet(f"background: {JENIS_NILAI_WARNA[j['jenis']]}; color: #FFFFFF;")
            tgl_label = QLabel(j["tanggal"])
            tgl_label.setStyleSheet("font-size:13px;font-weight:800;color:#13301F;")
            ket_label = QLabel(j["keterangan"] or "")
            ket_label.setObjectName("jadwal_kartu")
            btn_hapus = QPushButton("Hapus")
            btn_hapus.setObjectName("bahaya")
            btn_hapus.setCursor(Qt.CursorShape.PointingHandCursor)
            btn_hapus.clicked.connect(lambda _, jid=j["id"]: self.hapus_jadwal(jid))
            baris.addWidget(badge)
            baris.addWidget(tgl_label)
            baris.addSpacing(6)
            baris.addWidget(ket_label, 1)
            baris.addStretch()
            baris.addWidget(btn_hapus)
            self.layout_kartu.insertWidget(self.layout_kartu.count() - 1, kartu)

    def tambah_jadwal(self):
        kid = self.kelas_id_aktif()
        if kid is None:
            QMessageBox.warning(self, "Peringatan", "Pilih kelas dulu!")
            return
        tanggal = self.date_tgl.date().toString("yyyy-MM-dd")
        jenis = self.combo_jenis.currentText()
        keterangan = self.input_ket.text().strip()
        tambah_jadwal(kid, tanggal, jenis, keterangan)
        self.input_ket.clear()
        self.muat_jadwal()

    def hapus_jadwal(self, jid):
        jawab = QMessageBox.question(
            self,
            "Konfirmasi",
            "Hapus jadwal ini?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if jawab != QMessageBox.StandardButton.Yes:
            return
        hapus_jadwal(jid)
        self.muat_jadwal()


class HalamanRekapNilai(QWidget):
    def __init__(self, utama):
        super().__init__()
        self.utama = utama
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 20)
        root.setSpacing(14)

        kepala = QVBoxLayout()
        judul = QLabel("Rekap Nilai")
        judul.setObjectName("judul_halaman")
        sub = QLabel(
            "Detail nilai tiap siswa per tanggal (Latihan, Ulangan, UTS, UAS). "
            "Setiap nilai yang disimpan menampilkan jenis dan tanggal inputnya."
        )
        sub.setObjectName("subjudul_halaman")
        kepala.addWidget(judul)
        kepala.addWidget(sub)
        root.addLayout(kepala)

        kartu_alat = QFrame()
        kartu_alat.setObjectName("kartu")
        baris = QHBoxLayout(kartu_alat)
        baris.setContentsMargins(18, 14, 18, 14)
        baris.setSpacing(10)
        label_kelas = QLabel("Kelas")
        label_kelas.setObjectName("bagian")
        self.combo_kelas = QComboBox()
        self.combo_kelas.setMinimumWidth(170)
        label_jenis = QLabel("Jenis")
        label_jenis.setObjectName("bagian")
        self.combo_jenis = QComboBox()
        self.combo_jenis.addItem("Semua Jenis", None)
        for j in JENIS_NILAI:
            self.combo_jenis.addItem(j, j)
        self.combo_jenis.setMinimumWidth(140)
        label_dari = QLabel("Dari")
        label_dari.setObjectName("bagian")
        self.dari = KalenderMerahPopup()
        self.dari.setDisplayFormat("yyyy-MM-dd")
        self.dari.setDate(QDate.currentDate().addDays(-365))
        self.dari.setMinimumWidth(158)
        label_sampai = QLabel("Sampai")
        label_sampai.setObjectName("bagian")
        self.sampai = KalenderMerahPopup()
        self.sampai.setDisplayFormat("yyyy-MM-dd")
        self.sampai.setDate(QDate.currentDate())
        self.sampai.setMinimumWidth(158)
        btn_export = QPushButton("Export Excel")
        btn_export.setObjectName("utama")
        btn_export.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_export.clicked.connect(self.export_excel)
        btn_hapus_semua = QPushButton("Hapus Semua")
        btn_hapus_semua.setObjectName("bahaya")
        btn_hapus_semua.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_hapus_semua.clicked.connect(self.hapus_semua_data)
        for w in (
            label_kelas, self.combo_kelas, label_jenis, self.combo_jenis,
            label_dari, self.dari, label_sampai, self.sampai,
        ):
            baris.addWidget(w)
        baris.addStretch()
        baris.addWidget(btn_export)
        baris.addWidget(btn_hapus_semua)
        root.addWidget(kartu_alat)

        kartu_data = QFrame()
        kartu_data.setObjectName("kartu")
        kolom = QVBoxLayout(kartu_data)
        kolom.setContentsMargins(18, 16, 18, 16)
        kolom.setSpacing(12)

        header_data = QFrame()
        header_data.setObjectName("card_header")
        baris_header_data = QHBoxLayout(header_data)
        baris_header_data.setContentsMargins(0, 0, 0, 0)
        baris_header_data.setSpacing(8)
        judul_data = QLabel("Detail Nilai")
        judul_data.setObjectName("card_judul")
        self.badge_data = QLabel("0 catatan")
        self.badge_data.setObjectName("card_badge")
        baris_header_data.addWidget(judul_data)
        baris_header_data.addStretch()
        baris_header_data.addWidget(self.badge_data)
        kolom.addWidget(header_data)

        frame_cari = QFrame()
        frame_cari.setObjectName("search_frame")
        baris_cari = QHBoxLayout(frame_cari)
        baris_cari.setContentsMargins(12, 6, 12, 6)
        baris_cari.setSpacing(8)
        label_cari = QLabel("Cari nama")
        label_cari.setObjectName("search_label")
        self.input_cari = QLineEdit()
        self.input_cari.setPlaceholderText("Ketik nama siswa...")
        self.input_cari.setClearButtonEnabled(True)
        self.input_cari.textChanged.connect(lambda _: self.muat_rekap())
        baris_cari.addWidget(label_cari)
        baris_cari.addWidget(self.input_cari, 1)
        kolom.addWidget(frame_cari)

        self.tabel = QTableWidget()
        self.tabel.setColumnCount(5)
        self.tabel.setHorizontalHeaderLabels(
            ["Tanggal", "Kelas", "Nama", "Jenis", "Nilai"]
        )
        self.tabel.verticalHeader().setVisible(False)
        self.tabel.verticalHeader().setDefaultSectionSize(36)
        self.tabel.setAlternatingRowColors(True)
        self.tabel.setShowGrid(False)
        self.tabel.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.tabel.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tabel.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tabel.horizontalHeader().setHighlightSections(False)
        self.tabel.setColumnWidth(2, 260)
        kolom.addWidget(self.tabel, 1)
        root.addWidget(kartu_data, 1)

        baris_bawah = QHBoxLayout()
        self.label_total = QLabel("0 catatan")
        self.label_total.setObjectName("total")
        baris_bawah.addWidget(self.label_total)
        baris_bawah.addStretch()
        root.addLayout(baris_bawah)

        self.combo_kelas.currentIndexChanged.connect(lambda _: self.muat_rekap())
        self.combo_jenis.currentIndexChanged.connect(lambda _: self.muat_rekap())
        self.dari.dateChanged.connect(lambda _: self.muat_rekap())
        self.sampai.dateChanged.connect(lambda _: self.muat_rekap())
        tinggi_seragam(
            self.combo_kelas, self.combo_jenis, self.dari, self.sampai,
            btn_export, btn_hapus_semua,
        )

    def muat_combo(self, kelas_list):
        self.combo_kelas.blockSignals(True)
        self.combo_kelas.clear()
        self.combo_kelas.addItem("Semua Kelas", None)
        for kid, nama in kelas_list:
            self.combo_kelas.addItem(nama, kid)
        self.combo_kelas.blockSignals(False)
        self.muat_rekap()

    def reset_tanggal(self):
        self.dari.blockSignals(True)
        self.sampai.blockSignals(True)
        self.dari.setDate(QDate.currentDate().addDays(-365))
        self.sampai.setDate(QDate.currentDate())
        self.dari.blockSignals(False)
        self.sampai.blockSignals(False)
        self.muat_rekap()

    def muat_rekap(self):
        kelas_id = self.combo_kelas.currentData()
        jenis = self.combo_jenis.currentData()
        tgl_awal = self.dari.date().toString("yyyy-MM-dd")
        tgl_akhir = self.sampai.date().toString("yyyy-MM-dd")

        if tgl_awal > tgl_akhir:
            self.tabel.setRowCount(0)
            self.label_total.setText("0 catatan")
            self.badge_data.setText("0 catatan")
            return

        kata = self.input_cari.text().strip()

        kondisi = ["n.tanggal BETWEEN ? AND ?"]
        params = [tgl_awal, tgl_akhir]
        if kelas_id is not None:
            kondisi.append("s.kelas_id = ?")
            params.append(kelas_id)
        if jenis is not None:
            kondisi.append("n.jenis = ?")
            params.append(jenis)
        if kata:
            kondisi.append("s.nama LIKE ?")
            params.append(f"%{kata}%")
        where = " AND ".join(kondisi)

        dasar = """
            FROM nilai n
            JOIN siswa s ON s.id = n.siswa_id
            JOIN kelas k ON k.id = s.kelas_id
            WHERE
        """
        conn = get_conn()
        try:
            total = conn.execute(
                "SELECT COUNT(*) AS n " + dasar + where, params
            ).fetchone()["n"]
            rows = conn.execute(
                "SELECT n.tanggal, k.nama_kelas, s.nama, n.jenis, n.nilai "
                + dasar + where
                + " ORDER BY n.tanggal DESC, k.nama_kelas, s.nama",
                params,
            ).fetchall()
        finally:
            conn.close()

        self.tabel.setRowCount(0)
        for data in rows:
            r = self.tabel.rowCount()
            self.tabel.insertRow(r)
            nilai = [
                data["tanggal"], data["nama_kelas"], data["nama"],
                data["jenis"], f"{data['nilai']:.0f}",
            ]
            for c, teks in enumerate(nilai):
                item = QTableWidgetItem(teks)
                if c == 3:
                    item.setForeground(QColor(JENIS_NILAI_WARNA.get(teks, "#31513D")))
                elif c in (0, 4):
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tabel.setItem(r, c, item)
        self.badge_data.setText(f"{total} catatan")
        self.label_total.setText(f"{len(rows)} catatan")

    def hapus_semua_data(self):
        jawab = QMessageBox.question(
            self,
            "Konfirmasi",
            "HAPUS SEMUA data nilai?\n\n"
            "Seluruh nilai Latihan, Ulangan, UTS, dan UAS akan dihapus permanen.\n"
            "Tindakan ini tidak dapat dibatalkan!",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if jawab != QMessageBox.StandardButton.Yes:
            return
        hapus_semua_nilai()
        self.utama.segarkan_kelas()
        QMessageBox.information(self, "Sukses", "Semua data nilai berhasil dihapus.")

    def export_excel(self):
        if self.tabel.rowCount() == 0:
            QMessageBox.information(self, "Info", "Tidak ada data untuk diexport.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Export Excel", "rekap_nilai.xlsx",
            "File Excel (*.xlsx);;File CSV (*.csv)",
        )
        if not path:
            return
        data = [
            [self.tabel.item(r, c).text() for c in range(5)]
            for r in range(self.tabel.rowCount())
        ]
        header = ["Tanggal", "Kelas", "Nama", "Jenis", "Nilai"]
        if path.lower().endswith(".csv"):
            with open(path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(header)
                writer.writerows(data)
        else:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Rekap Nilai"
            for i, w in enumerate([13, 20, 30, 14, 12], start=1):
                ws.column_dimensions[get_column_letter(i)].width = w
            for c, h in enumerate(header, start=1):
                cell = ws.cell(row=1, column=c, value=h)
                cell.fill = PatternFill("solid", fgColor="16A34A")
                cell.font = Font(bold=True, color="FFFFFF")
            for r, d in enumerate(data, start=2):
                for c, v in enumerate(d, start=1):
                    ws.cell(row=r, column=c, value=v)
            ws.freeze_panes = "A2"
            wb.save(path)
        QMessageBox.information(
            self, "Sukses", f"Rekap nilai ({len(data)} catatan) berhasil diexport ke:\n{path}"
        )


class HalamanKesimpulanNilai(QWidget):
    def __init__(self, utama):
        super().__init__()
        self.utama = utama
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 20)
        root.setSpacing(14)

        kepala = QVBoxLayout()
        judul = QLabel("Kesimpulan Nilai")
        judul.setObjectName("judul_halaman")
        sub = QLabel("Rekap rata-rata nilai dan nilai akhir tiap siswa berdasarkan bobot.")
        sub.setObjectName("subjudul_halaman")
        kepala.addWidget(judul)
        kepala.addWidget(sub)
        root.addLayout(kepala)

        kartu_alat = QFrame()
        kartu_alat.setObjectName("kartu")
        baris = QHBoxLayout(kartu_alat)
        baris.setContentsMargins(18, 14, 18, 14)
        baris.setSpacing(10)
        label_kelas = QLabel("Kelas")
        label_kelas.setObjectName("bagian")
        self.combo_kelas = QComboBox()
        self.combo_kelas.setMinimumWidth(170)
        for w in (label_kelas, self.combo_kelas):
            baris.addWidget(w)
        self.label_bobot = QLabel(
            "Bobot: Latihan 20%  •  Ulangan 30%  •  UTS 20%  •  UAS 30%"
        )
        self.label_bobot.setObjectName("total")
        baris.addSpacing(10)
        baris.addWidget(self.label_bobot)
        baris.addStretch()
        btn_export = QPushButton("Export Excel")
        btn_export.setObjectName("utama")
        btn_export.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_export.clicked.connect(self.export_excel)
        baris.addWidget(btn_export)
        root.addWidget(kartu_alat)

        self.tabel = QTableWidget()
        self.tabel.setColumnCount(7)
        self.tabel.setHorizontalHeaderLabels(
            ["Nama", "Latihan", "Ulangan", "UTS", "UAS", "Nilai Akhir", "Predikat"]
        )
        self.tabel.verticalHeader().setVisible(False)
        self.tabel.verticalHeader().setDefaultSectionSize(36)
        self.tabel.setAlternatingRowColors(True)
        self.tabel.setShowGrid(False)
        self.tabel.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.tabel.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tabel.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tabel.horizontalHeader().setHighlightSections(False)
        self.tabel.setColumnWidth(0, 260)
        root.addWidget(self.tabel, 1)

        self.combo_kelas.currentIndexChanged.connect(lambda _: self.muat_rekap())
        tinggi_seragam(self.combo_kelas, btn_export)

    def muat_combo(self, kelas_list):
        self.combo_kelas.blockSignals(True)
        self.combo_kelas.clear()
        self.combo_kelas.addItem("-- Pilih Kelas --", None)
        for kid, nama in kelas_list:
            self.combo_kelas.addItem(nama, kid)
        self.combo_kelas.blockSignals(False)
        self.muat_rekap()

    def predikat(self, nilai):
        if nilai is None:
            return "-"
        if nilai >= 90:
            return {"teks": "A", "warna": "#059669"}
        if nilai >= 80:
            return {"teks": "B", "warna": "#16A34A"}
        if nilai >= 70:
            return {"teks": "C", "warna": "#F59E0B"}
        if nilai >= 60:
            return {"teks": "D", "warna": "#F97316"}
        return {"teks": "E", "warna": "#DC2626"}

    def muat_rekap(self):
        kid = self.combo_kelas.currentData()
        if kid is None:
            self.tabel.setRowCount(0)
            return
        data = hitung_nilai_siswa(kid)
        self.tabel.setRowCount(0)
        for siswa in data:
            r = self.tabel.rowCount()
            self.tabel.insertRow(r)
            self.tabel.setItem(r, 0, QTableWidgetItem(siswa["nama"]))
            for i, jenis in enumerate(JENIS_NILAI, start=1):
                v = siswa["rata"][jenis]
                item = QTableWidgetItem("" if v is None else f"{v:.0f}")
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if v is not None:
                    item.setForeground(QColor(JENIS_NILAI_WARNA[jenis]))
                self.tabel.setItem(r, i, item)
            akhir = siswa["akhir"]
            item_akhir = QTableWidgetItem("" if akhir is None else f"{akhir:.0f}")
            item_akhir.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if akhir is not None:
                item_akhir.setForeground(QColor("#15803D"))
            self.tabel.setItem(r, 5, item_akhir)
            p = self.predikat(akhir)
            item_pred = QTableWidgetItem(p["teks"])
            item_pred.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item_pred.setForeground(QColor(p["warna"]))
            font = item_pred.font()
            font.setBold(True)
            item_pred.setFont(font)
            self.tabel.setItem(r, 6, item_pred)

    def export_excel(self):
        if self.tabel.rowCount() == 0:
            QMessageBox.information(self, "Info", "Tidak ada data untuk diexport.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Export Nilai", "rekap_nilai.xlsx", "File Excel (*.xlsx)"
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "Kesimpulan Nilai"
        lebar = [30, 12, 12, 12, 12, 14, 12]
        for i, w in enumerate(lebar, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        header = ["Nama", "Latihan", "Ulangan", "UTS", "UAS", "Nilai Akhir", "Predikat"]
        for c, h in enumerate(header, start=1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.fill = PatternFill("solid", fgColor="16A34A")
            cell.font = Font(bold=True, color="FFFFFF")
        for r in range(self.tabel.rowCount()):
            for c in range(7):
                item = self.tabel.item(r, c)
                ws.cell(row=r + 2, column=c + 1, value=item.text() if item else "")
        ws.freeze_panes = "A2"
        wb.save(path)
        QMessageBox.information(self, "Sukses", f"Nilai berhasil diexport ke:\n{path}")


class JendelaUtama(QMainWindow):
    NAV = [
        ("Absensi Harian", 0, "Absensi", "absensi"),
        ("Rekap Absensi", 4, "Absensi", "rekap"),
        ("Kesimpulan Absensi", 6, "Absensi", "kesimpulan_absensi"),
        ("Input Nilai", 1, "Nilai", "nilai"),
        ("Rekap Nilai", 7, "Nilai", "rekap_nilai"),
        ("Kesimpulan Nilai", 3, "Nilai", "kesimpulan_nilai"),
        ("Jadwal UTS & Ujian", 2, "Nilai", "jadwal"),
        ("Kelola Data", 5, "Data", "kelola"),
    ]

    def _ikon_nav(self, nama):
        path_putih = os.path.join(BASE_DIR, "img", "icons", f"{nama}.svg")
        path_aktif = os.path.join(BASE_DIR, "img", "icons", f"{nama}_active.svg")
        ikon = QIcon()
        if os.path.exists(path_putih):
            ikon.addFile(path_putih, QSize(20, 20), QIcon.Mode.Normal, QIcon.State.Off)
        if os.path.exists(path_aktif):
            ikon.addFile(path_aktif, QSize(20, 20), QIcon.Mode.Normal, QIcon.State.On)
        return ikon

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Alfatih SmartSchool")
        self.resize(1120, 720)
        self.setMinimumSize(960, 640)
        if os.path.exists(LOGO_PATH):
            self.setWindowIcon(QIcon(LOGO_PATH))
        self.buat_sidebar()

        self.stack = QStackedWidget()
        self.halaman_absensi = HalamanAbsensi(self)
        self.halaman_nilai = HalamanNilai(self)
        self.halaman_jadwal = HalamanJadwal(self)
        self.halaman_kesimpulan_nilai = HalamanKesimpulanNilai(self)
        self.halaman_rekap = HalamanRekap(self)
        self.halaman_kelola = HalamanKelola(self)
        self.halaman_kesimpulan_absensi = HalamanKesimpulanAbsensi(self)
        self.halaman_rekap_nilai = HalamanRekapNilai(self)
        for halaman in (
            self.halaman_absensi,
            self.halaman_nilai,
            self.halaman_jadwal,
            self.halaman_kesimpulan_nilai,
            self.halaman_rekap,
            self.halaman_kelola,
            self.halaman_kesimpulan_absensi,
            self.halaman_rekap_nilai,
        ):
            self.stack.addWidget(halaman)

        wrapper = QWidget()
        layout_utama = QHBoxLayout(wrapper)
        layout_utama.setContentsMargins(0, 0, 0, 0)
        layout_utama.setSpacing(0)
        layout_utama.addWidget(self.sidebar)
        layout_utama.addWidget(self.stack, 1)
        self.setCentralWidget(wrapper)

        self.terapkan_shadow_kartu()

        self.grup_nav.idClicked.connect(self.pindah_halaman)
        self._animasi = None
        self.segarkan_kelas()
        self.pindah_halaman(0)
        self.showMaximized()

    def terapkan_shadow_kartu(self):
        for w in self.findChildren(QFrame, "kartu"):
            if not w.isVisibleTo(self):
                continue
            pasang_shadow(w, blur=32, alfa=45, offset_y=4)

    def _fade_halaman(self, halaman):
        if self._animasi is not None:
            self._animasi.stop()
            self._animasi.deleteLater()
            self._animasi = None
        efek = QGraphicsOpacityEffect(halaman)
        halaman.setGraphicsEffect(efek)
        self._animasi = QPropertyAnimation(efek, b"opacity", self)
        self._animasi.setDuration(180)
        self._animasi.setStartValue(0.25)
        self._animasi.setEndValue(1.0)
        self._animasi.setEasingCurve(QEasingCurve.Type.OutCubic)
        self._animasi.finished.connect(lambda: halaman.setGraphicsEffect(None))
        self._animasi.start()

    def buat_sidebar(self):
        self.sidebar = QWidget()
        self.sidebar.setObjectName("sidebar")
        self.sidebar.setFixedWidth(258)
        kolom = QVBoxLayout(self.sidebar)
        kolom.setContentsMargins(14, 24, 14, 16)
        kolom.setSpacing(5)

        brand_row = QHBoxLayout()
        brand_row.setSpacing(12)
        logo = QLabel()
        logo.setObjectName("logo_gambar")
        logo.setFixedSize(42, 42)
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        if os.path.exists(LOGO_PATH):
            pixmap = QPixmap(LOGO_PATH)
            logo.setPixmap(
                pixmap.scaled(
                    42, 42,
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation,
                )
            )
        brand_text = QVBoxLayout()
        brand_text.setSpacing(1)
        brand = QLabel("Alfatih SmartSchool")
        brand.setObjectName("brand_small")
        brand_sub = QLabel("Sistem Absensi dan Nilai")
        brand_sub.setObjectName("brand_sub")
        brand_text.addWidget(brand)
        brand_text.addWidget(brand_sub)
        brand_row.addWidget(logo)
        brand_row.addLayout(brand_text)
        brand_row.addStretch()
        kolom.addLayout(brand_row)
        kolom.addSpacing(20)

        self.grup_nav = QButtonGroup(self)
        self.grup_nav.setExclusive(True)
        seksi_sekarang = None
        for idx, (teks, _, seksi, nama_ikon) in enumerate(self.NAV):
            if seksi != seksi_sekarang:
                seksi_sekarang = seksi
                label_seksi = QLabel(seksi.upper())
                label_seksi.setObjectName("seksi")
                kolom.addWidget(label_seksi)
            btn = QPushButton(teks)
            btn.setObjectName("nav")
            btn.setCheckable(True)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setToolTip(teks)
            ikon = self._ikon_nav(nama_ikon)
            if not ikon.isNull():
                btn.setIcon(ikon)
                btn.setIconSize(QSize(20, 20))
            self.grup_nav.addButton(btn, idx)
            kolom.addWidget(btn)
        kolom.addStretch()

        versi = QLabel("v2.0   •   Alfatih SmartSchool")
        versi.setObjectName("versi")
        versi.setAlignment(Qt.AlignmentFlag.AlignCenter)
        kolom.addWidget(versi)

    def pindah_halaman(self, idx):
        stack_idx = self.NAV[idx][1]
        self.stack.setCurrentIndex(stack_idx)
        self._fade_halaman(self.stack.currentWidget())

    def segarkan_kelas(self):
        kelas_list = ambil_daftar_kelas()
        self.halaman_absensi.muat_combo(kelas_list)
        self.halaman_kelola.muat_kelas(kelas_list)
        self.halaman_rekap.muat_combo(kelas_list)
        self.halaman_nilai.muat_combo(kelas_list)
        self.halaman_jadwal.muat_combo(kelas_list)
        self.halaman_kesimpulan_nilai.muat_combo(kelas_list)
        self.halaman_kesimpulan_absensi.muat_combo(kelas_list)
        self.halaman_rekap_nilai.muat_combo(kelas_list)


def main():
    init_db()
    app = QApplication(sys.argv)
    QLocale.setDefault(QLocale(QLocale.Language.Indonesian, QLocale.Country.Indonesia))
    muat_font()
    app.setFont(QFont("Poppins"))
    app.setStyleSheet(STYLESHEET)
    jendela = JendelaUtama()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
