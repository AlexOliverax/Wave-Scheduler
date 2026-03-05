import os
import sys
import re
import logging
from datetime import datetime
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QDateEdit, QCheckBox, QComboBox, QFileDialog, QMessageBox,
    QGroupBox, QFormLayout, QSpinBox, QDoubleSpinBox, QTimeEdit,
    QDialog, QListWidget, QAbstractItemView, QTextEdit, QScrollArea,
    QProgressBar, QSizePolicy, QFrame
)
from PyQt5.QtCore import Qt, QDate, QTime, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon, QPixmap

from app.core import WavesScheduler
from app.utils import (save_config, get_timezone_list, get_country_holidays_dict,
                      get_supported_countries, is_holiday, get_translation,
                      get_supported_languages, get_supported_countries_translated,
                      translate_holiday_name)
from app.outlook import create_outlook_events, is_outlook_available
from app.dialogs import WavePreviewDialog, HistoryDialog
from app.history import add_entry
from app.updater import check_for_update_async, download_and_install, get_current_version

logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# STYLESHEET PREMIUM – Dark Theme + Glassmorphism
# ─────────────────────────────────────────────────────────────────────────────
DARK_STYLESHEET = """
/* ── Base ─────────────────────────────────────────────────────────────── */
QMainWindow, QDialog, QWidget {
    background-color: #0f1117;
    color: #e2e8f0;
    font-family: "Segoe UI", Arial, sans-serif;
    font-size: 10pt;
}

QScrollArea {
    border: none;
    background-color: transparent;
}

QScrollBar:vertical {
    background: #1a1d2e;
    width: 8px;
    border-radius: 4px;
}
QScrollBar::handle:vertical {
    background: #4f76ff;
    border-radius: 4px;
    min-height: 20px;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }

/* ── GroupBox ──────────────────────────────────────────────────────────── */
QGroupBox {
    background-color: rgba(26, 29, 46, 0.85);
    border: 1px solid rgba(79, 118, 255, 0.25);
    border-radius: 10px;
    margin-top: 14px;
    padding: 12px 10px 10px 10px;
    font-weight: bold;
    color: #7bb3ff;
    font-size: 10pt;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 2px 10px;
    background-color: rgba(79, 118, 255, 0.18);
    border-radius: 5px;
    color: #93b8ff;
}

/* ── Labels ────────────────────────────────────────────────────────────── */
QLabel {
    color: #cbd5e1;
    background-color: transparent;
}

/* ── Inputs ────────────────────────────────────────────────────────────── */
QLineEdit, QSpinBox, QDoubleSpinBox, QTimeEdit, QDateEdit, QComboBox, QTextEdit, QListWidget {
    background-color: #1e2235;
    border: 1px solid rgba(79, 118, 255, 0.3);
    border-radius: 6px;
    padding: 5px 8px;
    color: #e2e8f0;
    selection-background-color: #4f76ff;
}
QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus,
QTimeEdit:focus, QDateEdit:focus, QComboBox:focus {
    border: 1px solid #4f76ff;
    background-color: #232740;
}
QLineEdit::placeholder { color: #64748b; }

QSpinBox::up-button, QSpinBox::down-button,
QDoubleSpinBox::up-button, QDoubleSpinBox::down-button,
QTimeEdit::up-button, QTimeEdit::down-button,
QDateEdit::up-button, QDateEdit::down-button {
    background-color: #2d3155;
    border: none;
    border-radius: 3px;
    width: 14px;
}
QSpinBox::up-button:hover, QSpinBox::down-button:hover,
QDoubleSpinBox::up-button:hover, QDoubleSpinBox::down-button:hover {
    background-color: #4f76ff;
}

/* ── ComboBox ──────────────────────────────────────────────────────────── */
QComboBox QAbstractItemView {
    background-color: #1e2235;
    border: 1px solid #4f76ff;
    selection-background-color: #4f76ff;
    color: #e2e8f0;
    outline: none;
}
QComboBox::drop-down {
    border: none;
    background-color: #2d3155;
    border-radius: 0 5px 5px 0;
    width: 20px;
}

/* ── Buttons ────────────────────────────────────────────────────────────── */
QPushButton {
    background-color: #2d3155;
    color: #93b8ff;
    border: 1px solid rgba(79, 118, 255, 0.3);
    border-radius: 7px;
    padding: 7px 16px;
    font-weight: 600;
    min-height: 28px;
}
QPushButton:hover {
    background-color: #3b4080;
    border-color: #4f76ff;
    color: #ffffff;
}
QPushButton:pressed { background-color: #4f76ff; color: #ffffff; }
QPushButton:disabled { background-color: #1a1d2e; color: #3d4461; border-color: #2d3155; }

/* Botão principal de gerar waves */
QPushButton#generateBtn {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #4f76ff, stop:1 #7c3aed);
    color: white;
    border: none;
    border-radius: 10px;
    padding: 12px;
    font-size: 12pt;
    font-weight: bold;
    min-height: 44px;
}
QPushButton#generateBtn:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #6b8fff, stop:1 #9d54ff);
}
QPushButton#generateBtn:pressed {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #3b5bd6, stop:1 #6929d4);
}
QPushButton#generateBtn:disabled {
    background: #2d3155;
    color: #3d4461;
}

/* ── Checkbox ───────────────────────────────────────────────────────────── */
QCheckBox {
    color: #cbd5e1;
    spacing: 8px;
}
QCheckBox::indicator {
    width: 16px;
    height: 16px;
    border: 2px solid rgba(79, 118, 255, 0.5);
    border-radius: 4px;
    background-color: #1e2235;
}
QCheckBox::indicator:checked {
    background-color: #4f76ff;
    border-color: #4f76ff;
    image: url(none);
}
QCheckBox::indicator:hover { border-color: #4f76ff; }

/* ── StatusBar ──────────────────────────────────────────────────────────── */
QStatusBar {
    background-color: #0b0e18;
    color: #64748b;
    border-top: 1px solid rgba(79, 118, 255, 0.2);
    font-size: 9pt;
}

/* ── ListWidget ─────────────────────────────────────────────────────────── */
QListWidget {
    alternate-background-color: #232740;
    outline: none;
}
QListWidget::item:selected {
    background-color: #4f76ff;
    color: white;
    border-radius: 4px;
}
QListWidget::item:hover { background-color: #2d3155; }

/* ── TableWidget / TableView ─────────────────────────────────────────────── */
QTableWidget, QTableView {
    background-color: #131726;
    alternate-background-color: #1e2235;
    gridline-color: transparent;
    border: none;
    outline: none;
    color: #e2e8f0;
    selection-background-color: rgba(79, 118, 255, 0.35);
    selection-color: #ffffff;
}
QTableWidget::item, QTableView::item {
    padding: 4px 8px;
    border: none;
}
QTableWidget::item:selected, QTableView::item:selected {
    background-color: rgba(79, 118, 255, 0.35);
    color: #ffffff;
}
QHeaderView::section {
    background-color: #1a1d2e;
    color: #7bb3ff;
    font-weight: bold;
    padding: 6px 8px;
    border: none;
    border-bottom: 2px solid rgba(79, 118, 255, 0.4);
    font-size: 9pt;
}
QHeaderView::section:hover {
    background-color: #232740;
}

/* ── Separators ─────────────────────────────────────────────────────────── */
QFrame[frameShape="4"] { /* HLine */
    background-color: rgba(79, 118, 255, 0.2);
    min-height: 1px;
    max-height: 1px;
    border: none;
}
"""


class GenerateWavesThread(QThread):
    """Thread para geração de waves sem travar a UI, com sinais de progresso."""
    progress = pyqtSignal(int, str)   # (percent, step_label)
    finished = pyqtSignal(object, object, object)  # (wave_distribution, wave_labels, error)

    def __init__(self, scheduler, num_waves, devices_per_wave, start_date, avoid_holidays, country_code):
        super().__init__()
        self.scheduler = scheduler
        self.num_waves = num_waves
        self.devices_per_wave = devices_per_wave
        self.start_date = start_date
        self.avoid_holidays = avoid_holidays
        self.country_code = country_code

    def run(self):
        try:
            self.progress.emit(10, "step_holidays")
            # Pré-aquecer cache de feriados
            from app.utils import get_country_holidays_dict
            get_country_holidays_dict(self.start_date.year, self.country_code)
            get_country_holidays_dict(self.start_date.year + 1, self.country_code)

            self.progress.emit(30, "step_labels")
            wave_labels = self.scheduler.generate_wave_labels(
                self.start_date, self.num_waves,
                avoid_holidays=self.avoid_holidays,
                country_code=self.country_code
            )

            self.progress.emit(60, "step_devices")
            wave_distribution = self.scheduler.distribute_devices(
                self.num_waves, self.devices_per_wave
            )

            self.progress.emit(90, "step_excel")
            self.finished.emit(wave_distribution, wave_labels, None)
        except Exception as e:
            self.finished.emit(None, None, str(e))


# ─────────────────────────────────────────────────────────────────────────────
# ColumnSelectionDialog
# ─────────────────────────────────────────────────────────────────────────────
class ColumnSelectionDialog(QDialog):
    """Dialog for selecting columns from the data file."""

    def __init__(self, available_columns, required_columns, language="pt-BR", parent=None):
        super().__init__(parent)
        self.language = language
        self.setWindowTitle(get_translation("select_columns", self.language))
        self.setMinimumSize(450, 400)

        self.available_columns = available_columns
        self.required_columns = required_columns
        self.selected_columns = required_columns.copy()

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(10)

        instructions = QLabel(get_translation("column_instructions", self.language))
        instructions.setWordWrap(True)
        layout.addWidget(instructions)

        required_note = QLabel(
            get_translation("required_columns", self.language).format(', '.join(self.required_columns))
            if self.required_columns else ""
        )
        required_note.setStyleSheet("color: #f87171; font-size: 9pt;")
        if self.required_columns:
            layout.addWidget(required_note)

        self.column_list = QListWidget()
        self.column_list.setSelectionMode(QAbstractItemView.MultiSelection)
        self.column_list.setAlternatingRowColors(True)

        for column in self.available_columns:
            self.column_list.addItem(column)
            if column in self.required_columns:
                self.column_list.item(self.available_columns.index(column)).setSelected(True)

        layout.addWidget(self.column_list)

        button_layout = QHBoxLayout()
        select_all_button = QPushButton(get_translation("select_all", self.language))
        select_all_button.clicked.connect(self.select_all)
        button_layout.addWidget(select_all_button)

        clear_button = QPushButton(get_translation("clear_selection", self.language))
        clear_button.clicked.connect(self.clear_selection)
        button_layout.addWidget(clear_button)

        button_layout.addStretch()

        ok_button = QPushButton(get_translation("ok", self.language))
        ok_button.setDefault(True)
        ok_button.clicked.connect(self.accept)
        button_layout.addWidget(ok_button)

        cancel_button = QPushButton(get_translation("cancel", self.language))
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def select_all(self):
        for i in range(self.column_list.count()):
            self.column_list.item(i).setSelected(True)

    def clear_selection(self):
        for i in range(self.column_list.count()):
            item = self.column_list.item(i)
            item.setSelected(item.text() in self.required_columns)

    def get_selected_columns(self):
        selected = []
        for i in range(self.column_list.count()):
            if self.column_list.item(i).isSelected():
                selected.append(self.column_list.item(i).text())
        for col in self.required_columns:
            if col not in selected:
                selected.append(col)
        return selected


# ─────────────────────────────────────────────────────────────────────────────
# TimezoneSelectionDialog
# ─────────────────────────────────────────────────────────────────────────────
class TimezoneSelectionDialog(QDialog):
    """Advanced timezone selection dialog with search."""

    def __init__(self, current_timezone, language="pt-BR", parent=None):
        super().__init__(parent)
        self.language = language
        self.setWindowTitle(get_translation("timezone_selection", self.language))
        self.setMinimumSize(520, 420)
        self.current_timezone = current_timezone
        self.selected_timezone = current_timezone
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(10)

        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel(get_translation("search_timezone", self.language)))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText(get_translation("search_timezone", self.language))
        self.search_input.textChanged.connect(self.filter_timezones)
        search_layout.addWidget(self.search_input)
        layout.addLayout(search_layout)

        self.timezone_list = QListWidget()
        self.timezone_list.setAlternatingRowColors(True)
        timezones = get_timezone_list()
        for tz in timezones:
            self.timezone_list.addItem(tz)

        items = self.timezone_list.findItems(self.current_timezone, Qt.MatchExactly)
        if items:
            self.timezone_list.setCurrentItem(items[0])
            self.timezone_list.scrollToItem(items[0])

        self.timezone_list.itemDoubleClicked.connect(self.accept)
        layout.addWidget(self.timezone_list)

        button_layout = QHBoxLayout()
        ok_button = QPushButton(get_translation("ok", self.language))
        ok_button.setDefault(True)
        ok_button.clicked.connect(self.accept)
        cancel_button = QPushButton(get_translation("cancel", self.language))
        cancel_button.clicked.connect(self.reject)
        button_layout.addStretch()
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def filter_timezones(self, text):
        for i in range(self.timezone_list.count()):
            item = self.timezone_list.item(i)
            item.setHidden(text.lower() not in item.text().lower())

    def get_selected_timezone(self):
        current_item = self.timezone_list.currentItem()
        return current_item.text() if current_item else self.current_timezone


# ─────────────────────────────────────────────────────────────────────────────
# HolidaysViewDialog
# ─────────────────────────────────────────────────────────────────────────────
class HolidaysViewDialog(QDialog):
    """Dialog to view holidays for a specific year and location."""

    def __init__(self, year, country, state, language="pt-BR", parent=None):
        super().__init__(parent)
        self.year = year
        self.country = country
        self.state = state
        self.language = language

        window_title = get_translation("holidays_view", self.language).format(self.year, self.country)
        self.setWindowTitle(window_title)
        self.setModal(True)
        self.resize(520, 440)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(12)

        title_text = get_translation("national_holidays", self.language) + f" — {self.year}"
        title_label = QLabel(title_text)
        title_label.setFont(QFont("Segoe UI", 13, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("color: #7bb3ff; padding: 8px;")
        layout.addWidget(title_label)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        layout.addWidget(sep)

        holidays_text = QTextEdit()
        holidays_text.setReadOnly(True)
        holidays_text.setFont(QFont("Segoe UI", 9))

        holidays_dict = get_country_holidays_dict(self.year, self.country)

        if holidays_dict:
            holidays_list = []
            for date, name in sorted(holidays_dict.items()):
                day_name = date.strftime("%A")
                day_names_keys = {
                    "Monday": "monday", "Tuesday": "tuesday", "Wednesday": "wednesday",
                    "Thursday": "thursday", "Friday": "friday", "Saturday": "saturday",
                    "Sunday": "sunday"
                }
                day_key = day_names_keys.get(day_name, "monday")
                day_name_translated = get_translation(day_key, self.language)
                translated_holiday_name = translate_holiday_name(name, self.language)
                date_str = date.strftime("%d/%m/%Y")
                holidays_list.append(f"📅  {date_str}  ({day_name_translated})  —  {translated_holiday_name}")

            holidays_text.setPlainText("\n".join(holidays_list))
        else:
            holidays_text.setPlainText(get_translation("no_holidays_found", self.language))

        layout.addWidget(holidays_text)

        close_button = QPushButton(get_translation("close", self.language))
        close_button.setDefault(True)
        close_button.clicked.connect(self.accept)
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(close_button)
        layout.addLayout(btn_layout)

        self.setLayout(layout)


# ─────────────────────────────────────────────────────────────────────────────
# MainWindow
# ─────────────────────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    """Main window for the Waves Scheduler application."""

    def __init__(self, config):
        super().__init__()
        self.config = config
        self.current_language = config.get("language", "pt-BR")
        self.scheduler = WavesScheduler()
        self.csv_file_path = None
        self._gen_thread = None

        self.init_ui()
        self.setStyleSheet(DARK_STYLESHEET)

    # ── UI Setup ──────────────────────────────────────────────────────────────
    def init_ui(self):
        """Initialize the user interface."""
        from app.updater import get_current_version
        version = get_current_version()
        self.setWindowTitle(f"Waves Scheduler v{version}")
        self.setMinimumSize(650, 750)
        self.resize(750, 860)

        if "app_icon" in self.config and os.path.exists(self.config["app_icon"]):
            self.setWindowIcon(QIcon(self.config["app_icon"]))

        # ── Central scroll area ──────────────────────────────────────────────
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        self.setCentralWidget(scroll)

        container = QWidget()
        main_layout = QVBoxLayout(container)
        main_layout.setSpacing(14)
        main_layout.setContentsMargins(18, 18, 18, 18)
        scroll.setWidget(container)

        # ── Header / Logo ────────────────────────────────────────────────────
        header_layout = QHBoxLayout()
        logo_path = self.config.get("logo_image", "")
        if logo_path and os.path.exists(logo_path):
            logo_label = QLabel()
            pixmap = QPixmap(logo_path)
            logo_label.setPixmap(pixmap.scaled(180, 80, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            header_layout.addWidget(logo_label)
        else:
            from app.updater import get_current_version
            version = get_current_version()
            title_label = QLabel(f"🌊 Waves Scheduler v{version}")
            title_label.setFont(QFont("Segoe UI", 18, QFont.Bold))
            title_label.setStyleSheet("color: #4f76ff; letter-spacing: 1px;")
            header_layout.addWidget(title_label)

        header_layout.addStretch()

        # Botão Atualizar
        self.update_btn = QPushButton("🔄  " + get_translation("check_updates", self.current_language))
        self.update_btn.setMaximumWidth(135)
        self.update_btn.setToolTip(get_translation("check_updates_tip", self.current_language))
        self.update_btn.clicked.connect(self.check_for_updates)
        header_layout.addWidget(self.update_btn)

        # Botão Histórico
        history_btn = QPushButton("📋  " + get_translation("history", self.current_language))
        history_btn.setMaximumWidth(120)
        history_btn.clicked.connect(self.show_history)
        header_layout.addWidget(history_btn)

        # Status badge
        self.status_badge = QLabel("● " + get_translation("ready", self.current_language))
        self.status_badge.setStyleSheet("color: #4ade80; font-size: 9pt; font-weight: bold; margin-left:6px;")
        header_layout.addWidget(self.status_badge)
        main_layout.addLayout(header_layout)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        main_layout.addWidget(sep)

        # ── Seção 1: Arquivo de Dados ────────────────────────────────────────
        csv_group = QGroupBox(get_translation("data_file_selection", self.current_language))
        csv_layout = QVBoxLayout()
        csv_layout.setSpacing(8)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        csv_button = QPushButton("📂  " + get_translation("select_data_file", self.current_language))
        csv_button.clicked.connect(self.select_csv)
        csv_button.setMinimumWidth(280)
        btn_row.addWidget(csv_button)
        btn_row.addStretch()
        csv_layout.addLayout(btn_row)

        self.csv_path_label = QLabel(get_translation("no_file_selected", self.current_language))
        self.csv_path_label.setAlignment(Qt.AlignCenter)
        self.csv_path_label.setStyleSheet("color: #64748b; font-size: 9pt; font-style: italic;")
        csv_layout.addWidget(self.csv_path_label)
        csv_group.setLayout(csv_layout)
        main_layout.addWidget(csv_group)

        # ── Seção 2: Informações / Recomendações ─────────────────────────────
        info_group = QGroupBox(get_translation("device_information", self.current_language))
        info_layout = QVBoxLayout()
        info_layout.setSpacing(6)

        self.total_devices_label = QLabel(
            get_translation("total_devices", self.current_language).format(0)
        )
        self.total_devices_label.setFont(QFont("Segoe UI", 11, QFont.Bold))
        self.total_devices_label.setStyleSheet("color: #7bb3ff;")
        info_layout.addWidget(self.total_devices_label)

        rec_group = QGroupBox(get_translation("recommendations", self.current_language))
        rec_layout = QFormLayout()
        rec_layout.setSpacing(8)

        self.ideal_bandwidth_label = QLabel("-")
        self.ideal_waves_label = QLabel("-")
        self.devices_per_wave_label = QLabel("-")

        for lbl in [self.ideal_bandwidth_label, self.ideal_waves_label, self.devices_per_wave_label]:
            lbl.setFont(QFont("Segoe UI", 10, QFont.Bold))

        rec_layout.addRow(get_translation("ideal_bandwidth", self.current_language), self.ideal_bandwidth_label)
        rec_layout.addRow(get_translation("ideal_waves", self.current_language), self.ideal_waves_label)
        rec_layout.addRow(get_translation("devices_per_wave", self.current_language), self.devices_per_wave_label)

        rec_group.setLayout(rec_layout)
        info_layout.addWidget(rec_group)
        info_group.setLayout(info_layout)
        main_layout.addWidget(info_group)

        # ── Seção 3: Configuração ────────────────────────────────────────────
        config_group = QGroupBox(get_translation("wave_configuration", self.current_language))
        config_layout = QFormLayout()
        config_layout.setSpacing(10)

        # RFC
        self.rfc_input = QLineEdit()
        self.rfc_input.setPlaceholderText("Ex: RFC1234567 - Descrição")
        self.rfc_hint = QLabel("")
        self.rfc_hint.setStyleSheet("color: #64748b; font-size: 8pt;")
        self.rfc_input.textChanged.connect(self._validate_rfc)
        rfc_layout = QVBoxLayout()
        rfc_layout.setContentsMargins(0, 0, 0, 0)
        rfc_layout.setSpacing(2)
        rfc_layout.addWidget(self.rfc_input)
        rfc_layout.addWidget(self.rfc_hint)
        rfc_widget = QWidget()
        rfc_widget.setLayout(rfc_layout)
        config_layout.addRow(get_translation("rfc", self.current_language), rfc_widget)

        # Bandwidth
        self.bandwidth_input = QDoubleSpinBox()
        self.bandwidth_input.setRange(0.1, 100000)
        self.bandwidth_input.setValue(100)
        self.bandwidth_input.setSuffix(" MB")
        self.bandwidth_input.valueChanged.connect(self.update_recommendations)
        config_layout.addRow(get_translation("bandwidth", self.current_language), self.bandwidth_input)

        # MB per Device
        self.mb_per_device_input = QDoubleSpinBox()
        self.mb_per_device_input.setRange(0.01, 100.0)
        self.mb_per_device_input.setValue(self.config.get("mb_per_device", 0.5))
        self.mb_per_device_input.setSuffix(" MB")
        self.mb_per_device_input.setSingleStep(0.1)
        self.mb_per_device_input.setDecimals(2)
        self.mb_per_device_input.valueChanged.connect(self.update_recommendations)
        config_layout.addRow(get_translation("mb_per_device", self.current_language), self.mb_per_device_input)

        # Número de Waves
        self.waves_input = QSpinBox()
        self.waves_input.setRange(1, 365)
        self.waves_input.setValue(5)
        config_layout.addRow(get_translation("number_of_waves", self.current_language), self.waves_input)

        # Data de Início
        self.start_date_input = QDateEdit()
        self.start_date_input.setCalendarPopup(True)
        self.start_date_input.setDate(QDate.currentDate())
        self.start_date_input.setDisplayFormat("dd/MM/yyyy")
        self.start_date_input.dateChanged.connect(self.check_holiday_date)
        config_layout.addRow(get_translation("start_date", self.current_language), self.start_date_input)

        # Fuso Horário
        tz_layout = QHBoxLayout()
        tz_layout.setContentsMargins(0, 0, 0, 0)
        self.timezone_combo = QComboBox()
        self.timezone_combo.addItems(get_timezone_list())
        current_timezone = self.config.get("timezone", "America/Sao_Paulo")
        idx = self.timezone_combo.findText(current_timezone)
        if idx >= 0:
            self.timezone_combo.setCurrentIndex(idx)
        tz_layout.addWidget(self.timezone_combo)

        self.timezone_button = QPushButton(get_translation("timezone_select", self.current_language))
        self.timezone_button.setMaximumWidth(120)
        self.timezone_button.clicked.connect(self.open_timezone_selector)
        tz_layout.addWidget(self.timezone_button)

        tz_widget = QWidget()
        tz_widget.setLayout(tz_layout)
        config_layout.addRow(get_translation("timezone", self.current_language), tz_widget)

        # Idioma
        self.language_combo = QComboBox()
        for code, name in get_supported_languages(self.current_language).items():
            self.language_combo.addItem(name, code)
        lang_idx = self.language_combo.findData(self.current_language)
        if lang_idx >= 0:
            self.language_combo.setCurrentIndex(lang_idx)
        self.language_combo.currentTextChanged.connect(self.on_language_changed)
        config_layout.addRow(get_translation("language", self.current_language), self.language_combo)

        # País
        self.country_combo = QComboBox()
        for code, name in get_supported_countries_translated(self.current_language).items():
            self.country_combo.addItem(f"{code} — {name}", code)
        default_country = self.config.get("country", "BR")
        country_idx = self.country_combo.findData(default_country)
        if country_idx >= 0:
            self.country_combo.setCurrentIndex(country_idx)
        self.country_combo.currentTextChanged.connect(self.on_country_changed)
        config_layout.addRow(get_translation("country", self.current_language), self.country_combo)

        # Evitar feriados
        self.avoid_holidays_checkbox = QCheckBox(get_translation("avoid_holidays", self.current_language))
        self.avoid_holidays_checkbox.setChecked(True)
        config_layout.addRow("", self.avoid_holidays_checkbox)

        # Ver feriados
        self.view_holidays_button = QPushButton(get_translation("view_holidays", self.current_language))
        self.view_holidays_button.clicked.connect(self.show_holidays)
        config_layout.addRow("", self.view_holidays_button)

        sep2 = QFrame()
        sep2.setFrameShape(QFrame.HLine)
        config_layout.addRow(sep2)

        # Integração Outlook
        self.outlook_checkbox = QCheckBox(get_translation("outlook_integration", self.current_language))
        self.outlook_checkbox.setChecked(False)
        self.outlook_checkbox.stateChanged.connect(self.toggle_outlook_fields)
        config_layout.addRow("", self.outlook_checkbox)

        # Hora de Início
        self.start_time_input = QTimeEdit()
        self.start_time_input.setTime(QTime(9, 0))
        self.start_time_input.setDisplayFormat("HH:mm")
        config_layout.addRow(get_translation("start_time", self.current_language), self.start_time_input)

        # Hora de Fim
        self.end_time_input = QTimeEdit()
        self.end_time_input.setTime(QTime(10, 0))
        self.end_time_input.setDisplayFormat("HH:mm")
        config_layout.addRow(get_translation("end_time", self.current_language), self.end_time_input)

        # Evento de dia inteiro
        self.all_day_checkbox = QCheckBox()
        self.all_day_checkbox.stateChanged.connect(self.toggle_time_inputs)
        config_layout.addRow(get_translation("all_day_event", self.current_language), self.all_day_checkbox)

        # Local
        self.location_input = QLineEdit()
        self.location_input.setPlaceholderText(get_translation("location_placeholder", self.current_language))
        config_layout.addRow(get_translation("location", self.current_language), self.location_input)

        # Participantes Obrigatórios
        self.required_participants_input = QLineEdit()
        self.required_participants_input.setPlaceholderText(
            get_translation("participants_placeholder", self.current_language)
        )
        config_layout.addRow(
            get_translation("required_participants", self.current_language),
            self.required_participants_input
        )

        # Participantes Opcionais
        self.optional_participants_input = QLineEdit()
        self.optional_participants_input.setPlaceholderText(
            get_translation("participants_placeholder", self.current_language)
        )
        config_layout.addRow(
            get_translation("optional_participants", self.current_language),
            self.optional_participants_input
        )

        # Corpo do email Outlook
        self.email_body_input = QTextEdit()
        self.email_body_input.setPlaceholderText(
            get_translation("email_body_placeholder", self.current_language)
        )
        self.email_body_input.setMaximumHeight(90)
        self.email_body_input.setFont(QFont("Consolas", 9))
        config_layout.addRow(
            get_translation("email_body_label", self.current_language),
            self.email_body_input
        )

        # Desabilitar campos de Outlook por padrão
        self._outlook_fields = [
            self.start_time_input, self.end_time_input, self.all_day_checkbox,
            self.location_input, self.required_participants_input,
            self.optional_participants_input, self.email_body_input
        ]
        for w in self._outlook_fields:
            w.setEnabled(False)

        config_group.setLayout(config_layout)
        main_layout.addWidget(config_group)

        # ── Botão Gerar ──────────────────────────────────────────────────────
        self.generate_button = QPushButton("🌊  " + get_translation("generate_waves", self.current_language))
        self.generate_button.setObjectName("generateBtn")
        self.generate_button.clicked.connect(self.generate_waves)
        self.generate_button.setEnabled(False)
        main_layout.addWidget(self.generate_button)

        # ── Exportar CSV também ───────────────────────────────────────────────
        self.export_csv_checkbox = QCheckBox(get_translation("export_csv", self.current_language))
        self.export_csv_checkbox.setChecked(False)
        main_layout.addWidget(self.export_csv_checkbox)

        # ── Barra de Progresso ────────────────────────────────────────────────
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setMaximumHeight(12)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                background-color: #1e2235;
                border: 1px solid rgba(79,118,255,0.2);
                border-radius: 6px;
                font-size: 8pt; color: transparent;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 #4f76ff, stop:1 #7c3aed);
                border-radius: 6px;
            }
        """)
        self.progress_bar.hide()
        main_layout.addWidget(self.progress_bar)
        self.progress_label = QLabel("")
        self.progress_label.setAlignment(Qt.AlignCenter)
        self.progress_label.setStyleSheet("color: #64748b; font-size: 9pt;")
        self.progress_label.hide()
        main_layout.addWidget(self.progress_label)

        main_layout.addStretch()

        # ── Status Bar ───────────────────────────────────────────────────────
        self.statusBar().showMessage("⚡ " + get_translation("ready", self.current_language))

        self.update_recommendations()

        # Verificar atualizações em segundo plano ao iniciar
        def _init_done(info):
            from PyQt5.QtCore import QTimer
            QTimer.singleShot(0, lambda: self._on_update_check_done(info, silent=True))

        check_for_update_async(_init_done)

    # ── Helpers ────────────────────────────────────────────────────────────────
    def _get_country_code(self):
        data = self.country_combo.currentData()
        return data if data else "BR"

    def _validate_rfc(self, text):
        """Valida o RFC em formato livre (pelo menos 3 caracteres)."""
        if not text:
            self.rfc_input.setStyleSheet("")
            self.rfc_hint.setText("")
            return
            
        if len(text.strip()) >= 3:
            self.rfc_input.setStyleSheet("border: 1px solid #4ade80;")
            self.rfc_hint.setText("✓ " + get_translation("rfc_valid", self.current_language))
            self.rfc_hint.setStyleSheet("color: #4ade80; font-size: 8pt;")
        else:
            self.rfc_input.setStyleSheet("border: 1px solid #f87171;")
            self.rfc_hint.setText("⚠ " + get_translation("rfc_invalid", self.current_language))
            self.rfc_hint.setStyleSheet("color: #f87171; font-size: 8pt;")

    def check_for_updates(self, silent=False):
        """Verifica atualizações manualmente (com feedback visual)."""
        self.update_btn.setEnabled(False)
        self.update_btn.setText("🔄  ...")
        self.statusBar().showMessage("🔍 " + get_translation("checking_updates", self.current_language))

        # Guard: restaura o botão após 15s mesmo se o callback nunca vier
        from PyQt5.QtCore import QTimer
        self._update_guard = QTimer(self)
        self._update_guard.setSingleShot(True)
        self._update_guard.timeout.connect(self._restore_update_btn)
        self._update_guard.start(15000)

        def _done(info):
            # Callback rodando na thread — usar QTimer para voltar à UI thread
            QTimer.singleShot(0, lambda: self._on_update_check_done(info, silent=silent))

        check_for_update_async(_done)

    def _restore_update_btn(self):
        """Restaura o botão de update ao estado original (chamado pelo guard timer ou pelo callback)."""
        lang = self.current_language
        self.update_btn.setEnabled(True)
        self.update_btn.setText("🔄  " + get_translation("check_updates", lang))

    def _on_update_check_done(self, update_info, silent=False):
        """Callback chamado quando a verificação de updates termina."""
        # Cancelar o guard timer (se ainda estiver rodando)
        if hasattr(self, "_update_guard") and self._update_guard.isActive():
            self._update_guard.stop()

        lang = self.current_language
        self._restore_update_btn()

        if update_info is None:
            # Sem update
            if not silent:
                self.statusBar().showMessage("✅ " + get_translation("up_to_date", lang))
                QMessageBox.information(
                    self,
                    get_translation("check_updates", lang),
                    get_translation("up_to_date", lang)
                )
            else:
                self.statusBar().showMessage("✅ v" + get_current_version())
            return

        # Update disponível
        current = update_info.get("current_version", "?")
        latest  = update_info.get("tag_name", "?")
        name    = update_info.get("name", latest)
        notes   = update_info.get("body", "")[:400] or get_translation("no_release_notes", lang)

        msg = (
            f"🆕 {get_translation('update_available', lang)}\n\n"
            f"  {get_translation('current_version_label', lang)}: v{current}\n"
            f"  {get_translation('new_version_label', lang)}: {latest}\n\n"
            f"📋 {notes}\n\n"
            f"{get_translation('update_prompt', lang)}"
        )

        reply = QMessageBox.question(
            self,
            f"🌊 {name}",
            msg,
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )

        if reply != QMessageBox.Yes:
            self.statusBar().showMessage(
                f"⏭️ {get_translation('update_skipped', lang)}"
            )
            return

        download_url = update_info.get("download_url")
        if not download_url:
            QMessageBox.warning(
                self,
                "Download",
                get_translation("update_no_asset", lang) + "\n" + update_info.get("html_url", "")
            )
            return

        self.statusBar().showMessage("⬇️ " + get_translation("downloading_update", lang))
        self.update_btn.setEnabled(False)

        def _download_done():
            ok = download_and_install(
                download_url,
                progress_callback=lambda done, total: self.statusBar().showMessage(
                    f"⬇️ {done//1024:,} KB / {total//1024:,} KB"
                )
            )
            from PyQt5.QtCore import QTimer
            if ok:
                QTimer.singleShot(0, lambda: (
                    QMessageBox.information(
                        self,
                        get_translation("success", lang),
                        get_translation("update_installing", lang)
                    ),
                    self.close()
                ))
            else:
                QTimer.singleShot(0, lambda: (
                    self.update_btn.setEnabled(True),
                    QMessageBox.critical(
                        self,
                        get_translation("error", lang),
                        get_translation("update_download_error", lang)
                    )
                ))

        import threading
        threading.Thread(target=_download_done, daemon=True, name="updater-download").start()

    def show_history(self):
        """Abre o diálogo de histórico de schedules."""
        dialog = HistoryDialog(self.current_language, self)
        dialog.exec_()

    def _on_progress(self, percent, step_key):
        """Recebe sinal de progresso do thread e atualiza a UI."""
        self.progress_bar.setValue(percent)
        label = get_translation(step_key, self.current_language)
        self.progress_label.setText(label)

    def toggle_outlook_fields(self, state):
        enabled = bool(state)
        for w in self._outlook_fields:
            w.setEnabled(enabled)

    def toggle_time_inputs(self, state):
        enabled = not bool(state)
        self.start_time_input.setEnabled(enabled)
        self.end_time_input.setEnabled(enabled)

    # ── File Selection ─────────────────────────────────────────────────────────
    def select_csv(self):
        """Open file dialog to select a data file (CSV or Excel)."""
        start_dir = self.config.get("last_directory", "")
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            get_translation("select_data_file", self.current_language),
            start_dir,
            "Data Files (*.csv *.xlsx *.xls *.xlsm);;CSV Files (*.csv);;Excel Files (*.xlsx *.xls *.xlsm);;All Files (*)"
        )

        if not file_path:
            return

        self.csv_file_path = file_path
        self.csv_path_label.setText(f"📄  {os.path.basename(file_path)}")
        self.csv_path_label.setStyleSheet("color: #4ade80; font-size: 9pt;")

        self.config["last_directory"] = os.path.dirname(file_path)
        save_config(self.config)

        if not self.scheduler.load_file(file_path, preview_only=True):
            QMessageBox.critical(self, "Error",
                "Failed to load file. Please check the file format and try again.")
            self.generate_button.setEnabled(False)
            return

        available_columns = list(self.scheduler.data.columns)
        dialog = ColumnSelectionDialog(available_columns, [], self.current_language, self)
        if not dialog.exec_():
            self.csv_file_path = None
            self.csv_path_label.setText(get_translation("no_file_selected", self.current_language))
            self.csv_path_label.setStyleSheet("color: #64748b; font-size: 9pt; font-style: italic;")
            self.generate_button.setEnabled(False)
            return

        selected_columns = dialog.get_selected_columns()
        if self.scheduler.load_file(file_path, selected_columns=selected_columns):
            self.total_devices_label.setText(
                get_translation("total_devices", self.current_language).format(self.scheduler.total_devices)
            )
            self.generate_button.setEnabled(True)
            self.update_recommendations()
            self.statusBar().showMessage(
                f"✅ Loaded {self.scheduler.total_devices} devices · {len(selected_columns)} columns"
            )
            self.status_badge.setText("● File Loaded")
            self.status_badge.setStyleSheet("color: #4ade80; font-size: 9pt; font-weight: bold;")
        else:
            QMessageBox.critical(self, "Error", "Failed to load file with selected columns.")
            self.generate_button.setEnabled(False)

    # ── Recommendations ────────────────────────────────────────────────────────
    def update_recommendations(self):
        """Update the recommendation labels based on current values."""
        if self.scheduler.total_devices > 0:
            mb_per_device = self.mb_per_device_input.value()
            ideal_bandwidth = self.scheduler.total_devices * mb_per_device
            self.ideal_bandwidth_label.setText(f"{ideal_bandwidth:.1f} MB")

            current_bandwidth = self.bandwidth_input.value()
            if current_bandwidth >= ideal_bandwidth:
                self.ideal_bandwidth_label.setStyleSheet(
                    "color: #4ade80; font-weight: bold;"
                )
            else:
                self.ideal_bandwidth_label.setStyleSheet(
                    "color: #f87171; font-weight: bold;"
                )

            devices_per_wave = self.scheduler.calculate_devices_per_wave(current_bandwidth, mb_per_device)
            self.devices_per_wave_label.setText(f"{devices_per_wave:,}")

            ideal_waves = self.scheduler.calculate_ideal_waves(self.scheduler.total_devices, devices_per_wave)
            self.ideal_waves_label.setText(f"{ideal_waves}")

            current_waves = self.waves_input.value()
            if current_waves >= ideal_waves:
                self.ideal_waves_label.setStyleSheet("color: #4ade80; font-weight: bold;")
            else:
                self.ideal_waves_label.setStyleSheet("color: #f87171; font-weight: bold;")

            self.config["mb_per_device"] = mb_per_device
            save_config(self.config)
        else:
            for lbl in [self.ideal_bandwidth_label, self.ideal_waves_label, self.devices_per_wave_label]:
                lbl.setText("-")
                lbl.setStyleSheet("color: #64748b; font-weight: bold;")

    # ── Wave Generation ────────────────────────────────────────────────────────
    def generate_waves(self):
        """Generate waves using a QThread, shows preview and saves to history."""
        try:
            rfc = self.rfc_input.text().strip()
            bandwidth = self.bandwidth_input.value()
            num_waves = self.waves_input.value()
            mb_per_device = self.mb_per_device_input.value()
            start_date_str = self.start_date_input.date().toString("yyyy-MM-dd")
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            avoid_holidays = self.avoid_holidays_checkbox.isChecked()
            country_code = self._get_country_code()
            start_time = self.start_time_input.time().toPyTime()
            end_time = self.end_time_input.time().toPyTime()
            all_day = self.all_day_checkbox.isChecked()
            email_body = self.email_body_input.toPlainText().strip() or None
            devices_per_wave = self.scheduler.calculate_devices_per_wave(bandwidth, mb_per_device)

            # Validação Outlook
            required_emails = []
            optional_emails = []
            if self.outlook_checkbox.isChecked():
                required_participants = self.required_participants_input.text().strip()
                optional_participants = self.optional_participants_input.text().strip()
                location = self.location_input.text().strip() or "A definir"
                if not required_participants:
                    QMessageBox.warning(self, get_translation("warning", self.current_language),
                        get_translation("outlook_no_participants", self.current_language))
                    return
                email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
                required_emails = [e.strip() for e in required_participants.split(";") if e.strip()]
                optional_emails = [e.strip() for e in optional_participants.split(";") if e.strip()] if optional_participants else []
                invalid_emails = [e for e in required_emails + optional_emails if not re.match(email_pattern, e)]
                if invalid_emails:
                    QMessageBox.warning(self, get_translation("warning", self.current_language),
                        get_translation("invalid_emails", self.current_language) + "\n" + ", ".join(invalid_emails))
                    return
                if not is_outlook_available():
                    QMessageBox.warning(self, "Outlook",
                        get_translation("outlook_unavailable", self.current_language))

            # UI: mostrar progresso
            self.generate_button.setEnabled(False)
            self.progress_bar.setValue(0)
            self.progress_bar.show()
            self.progress_label.show()
            self.status_badge.setText("● " + get_translation("generating", self.current_language))
            self.status_badge.setStyleSheet("color: #fbbf24; font-size: 9pt; font-weight: bold; margin-left:6px;")
            self.statusBar().showMessage("⚙️ " + get_translation("generating", self.current_language))

            # Armazenar parâmetros Outlook para uso no callback
            self._pending_rfc = rfc
            self._pending_start_time = start_time
            self._pending_end_time = end_time
            self._pending_all_day = all_day
            self._pending_location = getattr(self, 'location_input', None) and self.location_input.text().strip() or "A definir"
            self._pending_required_emails = required_emails
            self._pending_optional_emails = optional_emails
            self._pending_email_body = email_body
            self._pending_country_code = country_code

            # Iniciar thread
            self._gen_thread = GenerateWavesThread(
                self.scheduler, num_waves, devices_per_wave,
                start_date, avoid_holidays, country_code
            )
            self._gen_thread.progress.connect(self._on_progress)
            self._gen_thread.finished.connect(self._on_generation_done)
            self._gen_thread.start()

        except Exception as e:
            logger.error(f"Error starting wave generation: {str(e)}", exc_info=True)
            self.generate_button.setEnabled(True)
            self.progress_bar.hide()
            self.progress_label.hide()
            self.status_badge.setText("● " + get_translation("error", self.current_language))
            self.status_badge.setStyleSheet("color: #f87171; font-size: 9pt; font-weight: bold; margin-left:6px;")
            QMessageBox.critical(self, get_translation("error", self.current_language),
                                 get_translation("error_occurred", self.current_language) + f" {str(e)}")

    def _on_generation_done(self, wave_distribution, wave_labels, error):
        """Callback do QThread: recebe resultado da geração."""
        self._on_progress(95, "step_excel")

        if error:
            self.generate_button.setEnabled(True)
            self.progress_bar.hide()
            self.progress_label.hide()
            self.status_badge.setText("● " + get_translation("error", self.current_language))
            self.status_badge.setStyleSheet("color: #f87171; font-size: 9pt; font-weight: bold; margin-left:6px;")
            QMessageBox.critical(self, get_translation("error", self.current_language),
                                 get_translation("error_occurred", self.current_language) + f" {error}")
            return

        # Preview antes de salvar
        preview = WavePreviewDialog(wave_labels, wave_distribution, self.current_language, self)
        if not preview.exec_() or not preview.confirmed:
            self.generate_button.setEnabled(True)
            self.progress_bar.hide()
            self.progress_label.hide()
            self.status_badge.setText("● " + get_translation("ready", self.current_language))
            self.status_badge.setStyleSheet("color: #4ade80; font-size: 9pt; font-weight: bold; margin-left:6px;")
            self.statusBar().showMessage("⚡ " + get_translation("ready", self.current_language))
            return

        # Escolher onde salvar
        save_dir = self.config.get("last_directory", "")
        default_filename = f"waves_schedule_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        excel_path, _ = QFileDialog.getSaveFileName(
            self, get_translation("save_excel_title", self.current_language),
            os.path.join(save_dir, default_filename),
            "Excel Files (*.xlsx)"
        )

        if not excel_path:
            self.generate_button.setEnabled(True)
            self.progress_bar.hide()
            self.progress_label.hide()
            self.status_badge.setText("● " + get_translation("ready", self.current_language))
            self.status_badge.setStyleSheet("color: #4ade80; font-size: 9pt; font-weight: bold; margin-left:6px;")
            self.statusBar().showMessage("⚡ " + get_translation("ready", self.current_language))
            return

        excel_saved = self.scheduler.generate_excel(
            excel_path, wave_distribution, wave_labels, self._pending_rfc
        )

        # Opcionalmente exportar CSV
        csv_saved = False
        if self.export_csv_checkbox.isChecked() and excel_saved:
            csv_path = excel_path.replace(".xlsx", ".csv")
            csv_saved = self.scheduler.generate_csv(csv_path, wave_distribution, wave_labels)

        # Integração Outlook
        outlook_events_created = False
        if self.outlook_checkbox.isChecked() and excel_saved:
            outlook_events_created = create_outlook_events(
                wave_labels,
                self._pending_start_time, self._pending_end_time,
                self._pending_rfc, self._pending_all_day,
                self._pending_location,
                self._pending_required_emails, self._pending_optional_emails,
                self._pending_email_body
            )

        # Adicionar ao histórico
        if excel_saved:
            add_entry(
                excel_path=excel_path,
                num_waves=len(wave_labels),
                num_devices=self.scheduler.total_devices,
                rfc=self._pending_rfc,
                country_code=self._pending_country_code,
                wave_labels=wave_labels,
            )

        self._on_progress(100, "step_done")
        self.generate_button.setEnabled(True)
        self.progress_bar.hide()
        self.progress_label.hide()

        if excel_saved:
            self.status_badge.setText("● Done")
            self.status_badge.setStyleSheet("color: #4ade80; font-size: 9pt; font-weight: bold; margin-left:6px;")
            msg_lines = [
                get_translation("waves_generated_msg", self.current_language, len(wave_labels)),
                get_translation("file_label", self.current_language) + " " + os.path.basename(excel_path),
            ]
            if csv_saved:
                msg_lines.append("📄 CSV exportado com sucesso.")
            if self.outlook_checkbox.isChecked():
                if outlook_events_created:
                    msg_lines.append("📅 " + get_translation("outlook_success", self.current_language))
                else:
                    msg_lines.append("⚠️ " + get_translation("outlook_failed", self.current_language))

            # Botão "Abrir Arquivo" no dialog de sucesso
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle(get_translation("success", self.current_language))
            msg_box.setText("\n".join(msg_lines))
            msg_box.setIcon(QMessageBox.Information)
            open_btn = msg_box.addButton(
                "📂 " + get_translation("open_file", self.current_language), QMessageBox.ActionRole
            )
            msg_box.addButton(QMessageBox.Ok)
            msg_box.exec_()

            if msg_box.clickedButton() == open_btn:
                try:
                    os.startfile(excel_path)
                except Exception as e:
                    logger.error(f"Error opening file: {e}")

            self.statusBar().showMessage(f"✅ " + get_translation("file_label", self.current_language) + f" {excel_path}")
        else:
            self.status_badge.setText("● " + get_translation("error", self.current_language))
            self.status_badge.setStyleSheet("color: #f87171; font-size: 9pt; font-weight: bold; margin-left:6px;")
            QMessageBox.critical(self, get_translation("error", self.current_language),
                                 get_translation("excel_error", self.current_language))
            self.statusBar().showMessage("❌ " + get_translation("excel_error", self.current_language))

        # ── Timezone Selector ──────────────────────────────────────────────────────
    def open_timezone_selector(self):
        current_timezone = self.timezone_combo.currentText()
        dialog = TimezoneSelectionDialog(current_timezone, self.current_language, self)
        if dialog.exec_() == QDialog.Accepted:
            selected_timezone = dialog.get_selected_timezone()
            idx = self.timezone_combo.findText(selected_timezone)
            if idx >= 0:
                self.timezone_combo.setCurrentIndex(idx)
            else:
                self.timezone_combo.addItem(selected_timezone)
                self.timezone_combo.setCurrentText(selected_timezone)
            self.config["timezone"] = selected_timezone
            save_config(self.config)
            logger.info(f"Timezone changed to: {selected_timezone}")

    # ── Holiday Check ──────────────────────────────────────────────────────────
    def check_holiday_date(self, date):
        """Check if the selected start date is a holiday and show warning."""
        try:
            python_date = date.toPyDate()
            country_code = self._get_country_code()
            is_holiday_date, holiday_name = is_holiday(python_date, country_code)

            if is_holiday_date:
                date_str = python_date.strftime("%d/%m/%Y")
                warning_msg = get_translation(
                    "holiday_warning", self.current_language,
                    date_str, holiday_name
                )
                QMessageBox.warning(self, "Aviso de Feriado", warning_msg)
                logger.info(f"Holiday warning shown for {date_str}: {holiday_name}")
        except Exception as e:
            logger.error(f"Error checking holiday date: {str(e)}")

    # ── Language Changed ───────────────────────────────────────────────────────
    def on_language_changed(self, language_text):
        try:
            language_code = self.language_combo.currentData()
            if not language_code:
                language_code = "pt-BR"
            self.current_language = language_code
            self.config["language"] = language_code
            save_config(self.config)
            logger.info(f"Language changed to: {language_code}")
            QMessageBox.information(
                self,
                get_translation("information", self.current_language),
                get_translation("restart_required", self.current_language)
            )
        except Exception as e:
            logger.error(f"Error changing language: {str(e)}")

    # ── Country Changed ────────────────────────────────────────────────────────
    def on_country_changed(self, country_text):
        try:
            country_code = self._get_country_code()
            self.config["country"] = country_code
            save_config(self.config)
            self.check_holiday_date(self.start_date_input.date())
            logger.info(f"Country changed to: {country_code}")
        except Exception as e:
            logger.error(f"Error changing country: {str(e)}")

    # ── Show Holidays ──────────────────────────────────────────────────────────
    def show_holidays(self):
        try:
            current_year = self.start_date_input.date().year()
            country_code = self._get_country_code()
            dialog = HolidaysViewDialog(current_year, country_code, None, self.current_language, self)
            dialog.exec_()
        except Exception as e:
            logger.error(f"Error showing holidays: {str(e)}")
            QMessageBox.critical(self, "Erro", f"Erro ao exibir feriados: {str(e)}")