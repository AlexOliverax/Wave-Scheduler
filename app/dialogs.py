"""
Dialogs avançados para o Waves Scheduler v2:
  - WavePreviewDialog: tabela de preview antes de salvar
  - HistoryDialog: histórico de schedules gerados
"""
import os
import subprocess
import logging
from datetime import datetime

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox,
    QAbstractItemView, QFrame, QSizePolicy
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QColor

from app.utils import get_translation
from app.history import get_history, clear_history, remove_entry

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# WavePreviewDialog
# ─────────────────────────────────────────────────────────────────────────────
class WavePreviewDialog(QDialog):
    """Mostra um preview das waves antes de salvar, pedindo confirmação."""

    WEEKDAY_KEYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]

    def __init__(self, wave_labels, wave_distribution, language="pt-BR", parent=None):
        super().__init__(parent)
        self.language = language
        self.wave_labels = wave_labels
        self.wave_distribution = wave_distribution
        self.confirmed = False

        self.setWindowTitle(get_translation("wave_preview_title", self.language))
        self.setMinimumSize(640, 480)
        self.resize(720, 520)
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(12)

        # Título
        title = QLabel("🌊  " + get_translation("wave_preview_title", self.language))
        title.setFont(QFont("Segoe UI", 13, QFont.Bold))
        title.setStyleSheet("color: #7bb3ff; padding: 4px;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        sep = QFrame(); sep.setFrameShape(QFrame.HLine)
        layout.addWidget(sep)

        # Tabela
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels([
            get_translation("wave_col_wave", self.language),
            get_translation("wave_col_date", self.language),
            get_translation("wave_col_weekday", self.language),
            get_translation("wave_col_devices", self.language),
        ])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.verticalHeader().setVisible(False)
        self.table.setShowGrid(False)

        total_devices = 0
        for i, label in enumerate(self.wave_labels):
            wave_key = f"Wave {i+1}"
            devices_in_wave = len(self.wave_distribution.get(wave_key, []))
            total_devices += devices_in_wave

            # Extrair data
            try:
                date_str = label.split(" - ")[1] if " - " in label else ""
                dt = datetime.strptime(date_str, "%d/%m/%Y") if date_str else None
                weekday_key = self.WEEKDAY_KEYS[dt.weekday()] if dt else ""
                weekday = get_translation(weekday_key, self.language) if weekday_key else ""
            except Exception:
                date_str = ""
                weekday = ""

            self.table.insertRow(i)
            self.table.setItem(i, 0, QTableWidgetItem(f"Wave {i+1}"))
            self.table.setItem(i, 1, QTableWidgetItem(date_str))
            self.table.setItem(i, 2, QTableWidgetItem(weekday))
            count_item = QTableWidgetItem(f"{devices_in_wave:,}")
            count_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(i, 3, count_item)

        layout.addWidget(self.table)

        # Resumo
        summary = QLabel(f"📊  Total: {len(self.wave_labels)} waves · {total_devices:,} dispositivos")
        summary.setStyleSheet("color: #94a3b8; font-size: 9pt; padding: 4px;")
        summary.setAlignment(Qt.AlignRight)
        layout.addWidget(summary)

        # Botões
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        cancel_btn = QPushButton(get_translation("cancel", self.language))
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        confirm_btn = QPushButton("✅  " + get_translation("confirm_save", self.language))
        confirm_btn.setObjectName("generateBtn")
        confirm_btn.clicked.connect(self._confirm)
        btn_layout.addWidget(confirm_btn)

        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def _confirm(self):
        self.confirmed = True
        self.accept()


# ─────────────────────────────────────────────────────────────────────────────
# HistoryDialog
# ─────────────────────────────────────────────────────────────────────────────
class HistoryDialog(QDialog):
    """Exibe o histórico de schedules gerados com opção de abrir arquivos."""

    def __init__(self, language="pt-BR", parent=None):
        super().__init__(parent)
        self.language = language
        self.setWindowTitle(get_translation("history_title", self.language))
        self.setMinimumSize(780, 460)
        self.resize(860, 520)
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(12)

        # Título
        title = QLabel("📋  " + get_translation("history_title", self.language))
        title.setFont(QFont("Segoe UI", 13, QFont.Bold))
        title.setStyleSheet("color: #7bb3ff; padding: 4px;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        sep = QFrame(); sep.setFrameShape(QFrame.HLine)
        layout.addWidget(sep)

        # Tabela
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels([
            get_translation("history_col_date", self.language),
            get_translation("history_col_rfc", self.language),
            get_translation("history_col_waves", self.language),
            get_translation("history_col_devices", self.language),
            get_translation("history_col_file", self.language),
        ])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.verticalHeader().setVisible(False)
        self.table.setShowGrid(False)
        self.table.doubleClicked.connect(self._open_selected)

        self._load_history()
        layout.addWidget(self.table)

        # Botões
        btn_layout = QHBoxLayout()
        clear_btn = QPushButton("🗑️  " + get_translation("history_clear", self.language))
        clear_btn.setStyleSheet("color: #f87171; border-color: rgba(248,113,113,0.3);")
        clear_btn.clicked.connect(self._clear_history)
        btn_layout.addWidget(clear_btn)

        btn_layout.addStretch()

        open_btn = QPushButton("📂  " + get_translation("history_open", self.language))
        open_btn.clicked.connect(self._open_selected)
        btn_layout.addWidget(open_btn)

        close_btn = QPushButton(get_translation("close", self.language))
        close_btn.setDefault(True)
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def _load_history(self):
        """Carrega e exibe as entradas do histórico."""
        self.table.setRowCount(0)
        entries = get_history()

        if not entries:
            self.table.setRowCount(1)
            empty_item = QTableWidgetItem(get_translation("history_empty", self.language))
            empty_item.setTextAlignment(Qt.AlignCenter)
            empty_item.setForeground(QColor("#64748b"))
            self.table.setItem(0, 0, empty_item)
            self.table.setSpan(0, 0, 1, 5)
            return

        self._entries = entries
        for i, entry in enumerate(entries):
            self.table.insertRow(i)
            ts = entry.get("timestamp", "")
            try:
                dt_str = datetime.fromisoformat(ts).strftime("%d/%m/%Y %H:%M")
            except Exception:
                dt_str = ts[:16] if ts else "—"

            self.table.setItem(i, 0, QTableWidgetItem(dt_str))
            self.table.setItem(i, 1, QTableWidgetItem(entry.get("rfc", "—")))

            waves_item = QTableWidgetItem(str(entry.get("num_waves", "—")))
            waves_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(i, 2, waves_item)

            dev_item = QTableWidgetItem(f"{entry.get('num_devices', 0):,}")
            dev_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(i, 3, dev_item)

            path = entry.get("excel_path", "")
            exists = os.path.exists(path)
            fname = entry.get("filename", os.path.basename(path))
            file_item = QTableWidgetItem(fname)
            if not exists:
                file_item.setForeground(QColor("#f87171"))
                file_item.setToolTip(f"Arquivo não encontrado: {path}")
            else:
                file_item.setForeground(QColor("#4ade80"))
                file_item.setToolTip(path)
            self.table.setItem(i, 4, file_item)

    def _open_selected(self):
        row = self.table.currentRow()
        if not hasattr(self, "_entries") or row < 0 or row >= len(self._entries):
            return
        path = self._entries[row].get("excel_path", "")
        if path and os.path.exists(path):
            try:
                os.startfile(path)
            except Exception as e:
                QMessageBox.warning(self, "Erro", f"Não foi possível abrir o arquivo:\n{e}")
        else:
            QMessageBox.warning(self, "Arquivo não encontrado",
                                f"O arquivo não foi encontrado:\n{path}")

    def _clear_history(self):
        reply = QMessageBox.question(
            self,
            get_translation("history_clear", self.language),
            get_translation("history_clear_confirm", self.language),
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            clear_history()
            self._load_history()
