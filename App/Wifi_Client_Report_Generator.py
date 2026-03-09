# gui_main.py

import sys
import os
import subprocess
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QListWidget, QListWidgetItem,
    QLabel, QMessageBox, QCheckBox, QDateEdit, QGroupBox,
    QStatusBar, QMenuBar, QSizePolicy
)
from PyQt6.QtCore import Qt, QDate, QUrl
from PyQt6.QtGui import QDesktopServices, QFont, QPalette

from report_generator import ingest_files, generate_excel_report, normalize_datetime
from datetime import datetime

APP_VERSION = "1.3"
REPO_URL    = "https://github.com/mfrisby-extreme/XIQ_Client_Report_CSV"

# ---------- Stylesheet ----------
# ---------- Stylesheet ----------
# Colors are intentionally omitted so Qt inherits the system palette,
# which automatically respects light/dark mode on all platforms.
# Only structural properties (radius, padding, spacing) are hardcoded.
STYLE = """
QGroupBox {
    font-weight: bold;
    font-size: 12px;
    border: 1px solid palette(mid);
    border-radius: 6px;
    margin-top: 10px;
    padding-top: 6px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 0 6px;
    left: 10px;
}

QListWidget {
    border: 1px solid palette(mid);
    border-radius: 4px;
    padding: 2px;
}
QListWidget::item {
    padding: 4px 6px;
    border-radius: 3px;
}
QListWidget::item:selected {
    background-color: palette(highlight);
    color: palette(highlighted-text);
}

QDateEdit {
    border: 1px solid palette(mid);
    border-radius: 4px;
    padding: 3px 6px;
}

QCheckBox {
    spacing: 6px;
}

QPushButton#secondary {
    border: 1px solid palette(mid);
    border-radius: 4px;
    padding: 5px 14px;
}
QPushButton#secondary:hover {
    background-color: palette(midlight);
}

QPushButton#primary {
    background-color: palette(highlight);
    color: palette(highlighted-text);
    border: none;
    border-radius: 4px;
    padding: 6px 20px;
    font-weight: bold;
}
QPushButton#primary:hover {
    background-color: palette(highlight);
    opacity: 0.85;
}
QPushButton#primary:disabled {
    background-color: palette(mid);
    color: palette(shadow);
}

QStatusBar {
    border-top: 1px solid palette(mid);
    font-size: 11px;
    padding: 2px 8px;
}

QMenuBar {
    border-bottom: 1px solid palette(mid);
}
"""


class ReportUI(QWidget):
    def __init__(self):
        super().__init__()
        self.data      = []
        self.file_path = ""

        self.setWindowTitle("WiFi Client Report Generator")
        self.setMinimumWidth(480)
        self.setStyleSheet(STYLE)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── Menu bar ──────────────────────────────────────────────
        menu_bar  = QMenuBar()
        help_menu = menu_bar.addMenu("Help")
        about_act = help_menu.addAction("About")
        about_act.triggered.connect(self.show_about)
        root.addWidget(menu_bar)

        # ── Header banner ─────────────────────────────────────────
        header = QLabel("WiFi Client Report Generator")
        header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header_font = QFont()
        header_font.setPointSize(13)
        header_font.setBold(True)
        header.setFont(header_font)
        header.setAutoFillBackground(True)
        palette = header.palette()
        palette.setColor(header.backgroundRole(), self.palette().color(QPalette.ColorRole.Highlight))
        palette.setColor(header.foregroundRole(), self.palette().color(QPalette.ColorRole.HighlightedText))
        header.setPalette(palette)
        header.setStyleSheet("padding: 14px;")
        root.addWidget(header)

        # ── Main content area ─────────────────────────────────────
        content = QVBoxLayout()
        content.setContentsMargins(16, 14, 16, 10)
        content.setSpacing(10)
        root.addLayout(content)

        # ── Data Source group ─────────────────────────────────────
        src_group  = QGroupBox("Data Source")
        src_layout = QVBoxLayout(src_group)
        src_layout.setContentsMargins(10, 14, 10, 10)
        src_layout.setSpacing(8)

        load_row = QHBoxLayout()
        self.load_btn = QPushButton("Load CSV / ZIP File(s)...")
        self.load_btn.setObjectName("secondary")
        self.load_btn.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.load_btn.clicked.connect(self.load_csv)
        self.file_label = QLabel("No file loaded")
        self.file_label.setStyleSheet("color: #6b7280; font-size: 11px;")
        self.file_label.setWordWrap(True)
        load_row.addWidget(self.load_btn)
        load_row.addWidget(self.file_label, 1)
        src_layout.addLayout(load_row)

        content.addWidget(src_group)

        # ── Sites group ───────────────────────────────────────────
        sites_group  = QGroupBox("Select Site(s)")
        sites_layout = QVBoxLayout(sites_group)
        sites_layout.setContentsMargins(10, 14, 10, 10)

        self.site_list = QListWidget()
        self.site_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.site_list.setMinimumHeight(120)
        placeholder = QListWidgetItem("— load a file to populate —")
        placeholder.setFlags(Qt.ItemFlag.NoItemFlags)
        placeholder.setForeground(Qt.GlobalColor.gray)
        self.site_list.addItem(placeholder)
        self.site_list.setEnabled(False)
        sites_layout.addWidget(self.site_list)

        content.addWidget(sites_group)

        # ── Options group ─────────────────────────────────────────
        opt_group  = QGroupBox("Options")
        opt_layout = QVBoxLayout(opt_group)
        opt_layout.setContentsMargins(10, 14, 10, 10)
        opt_layout.setSpacing(6)

        self.combine_floors_cb = QCheckBox("Aggregate per-floor data into buildings")
        self.combine_floors_cb.setChecked(True)
        self.combine_floors_cb.stateChanged.connect(self.toggle_combine_floors)

        self.tab_per_building = QCheckBox("Generate a separate tab per building")
        self.tab_per_building.setChecked(True)

        self.open_after_checkbox = QCheckBox("Open report file after generation")

        opt_layout.addWidget(self.combine_floors_cb)
        opt_layout.addWidget(self.tab_per_building)
        opt_layout.addWidget(self.open_after_checkbox)

        content.addWidget(opt_group)

        # ── Date Range group ──────────────────────────────────────
        date_group  = QGroupBox("Date Range")
        date_layout = QHBoxLayout(date_group)
        date_layout.setContentsMargins(10, 14, 10, 10)
        date_layout.setSpacing(8)

        self.date_from = QDateEdit()
        self.date_from.setDisplayFormat("yyyy-MM-dd")
        self.date_from.setCalendarPopup(True)
        self.date_from.setEnabled(False)

        self.date_to = QDateEdit()
        self.date_to.setDisplayFormat("yyyy-MM-dd")
        self.date_to.setCalendarPopup(True)
        self.date_to.setEnabled(False)

        date_layout.addWidget(QLabel("From:"))
        date_layout.addWidget(self.date_from, 1)
        date_layout.addSpacing(12)
        date_layout.addWidget(QLabel("To:"))
        date_layout.addWidget(self.date_to, 1)

        content.addWidget(date_group)

        # ── Action row ────────────────────────────────────────────
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.export_btn = QPushButton("Generate Report")
        self.export_btn.setObjectName("primary")
        self.export_btn.setFixedWidth(150)
        self.export_btn.clicked.connect(self.generate_report)
        btn_row.addWidget(self.export_btn)
        content.addLayout(btn_row)
        content.addStretch()

        # ── Status bar ────────────────────────────────────────────
        self.status_bar = QStatusBar()
        self.status_bar.showMessage("Ready  —  no file loaded")
        root.addWidget(self.status_bar)

    # ── Helpers ───────────────────────────────────────────────────

    @staticmethod
    def open_in_default_app(path: str) -> None:
        p = str(Path(path).resolve())
        if sys.platform.startswith("darwin"):
            subprocess.run(["open", p], check=False)
        elif os.name == "nt":
            os.startfile(p)          # type: ignore[attr-defined]
        else:
            subprocess.run(["xdg-open", p], check=False)

    def show_about(self):
        msg = QMessageBox(self)
        msg.setWindowTitle("About")
        msg.setText(
            f"<b>WiFi Client Report Generator</b><br>"
            f"Version {APP_VERSION}<br><br>"
            f"Source code:<br>"
            f'<a href="{REPO_URL}">{REPO_URL}</a>'
        )
        msg.setTextFormat(Qt.TextFormat.RichText)
        msg.setTextInteractionFlags(Qt.TextInteractionFlag.TextBrowserInteraction)
        msg.exec()

    # ── Slots ─────────────────────────────────────────────────────

    def load_csv(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Open CSV or ZIP Files", "",
            "CSV or ZIP Files (*.csv *.zip)"
        )
        if not paths:
            return

        try:
            self.data = ingest_files(paths)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load files:\n{e}")
            return

        if not self.data:
            QMessageBox.warning(self, "No Data", "No valid CSV data found.")
            return

        # Date bounds
        dates = []
        for row in self.data:
            end_time = row.get('end_time')
            if not end_time:
                continue
            try:
                dates.append(normalize_datetime(end_time))
            except ValueError:
                continue

        if not dates:
            QMessageBox.warning(self, "No Valid Dates", "No valid timestamps found in the loaded data.")
            return

        min_date, max_date = min(dates), max(dates)

        self.date_from.setMinimumDate(QDate(1900, 1, 1))
        self.date_to.setMaximumDate(QDate(3000, 1, 1))
        self.date_from.setDate(QDate(min_date.year, min_date.month, min_date.day))
        self.date_to.setDate(QDate(max_date.year, max_date.month, max_date.day))
        self.date_from.setMinimumDate(QDate(min_date.year, min_date.month, min_date.day))
        self.date_from.setMaximumDate(QDate(max_date.year, max_date.month, max_date.day))
        self.date_to.setMinimumDate(QDate(min_date.year, min_date.month, min_date.day))
        self.date_to.setMaximumDate(QDate(max_date.year, max_date.month, max_date.day))
        self.date_from.setEnabled(True)
        self.date_to.setEnabled(True)

        # Populate sites
        self.site_list.clear()
        self.site_list.setEnabled(True)
        sites = sorted(set(d['location'] for d in self.data))
        for site in sites:
            self.site_list.addItem(QListWidgetItem(site))

        # Update labels and status bar
        names = ", ".join(Path(p).name for p in paths)
        self.file_label.setText(names)
        self.status_bar.showMessage(
            f"Loaded {len(self.data):,} sessions across {len(sites)} site(s)"
            f"  |  {min_date.strftime('%Y-%m-%d')} – {max_date.strftime('%Y-%m-%d')}"
        )

    def toggle_combine_floors(self, state):
        if not state:
            self.tab_per_building.setChecked(False)
            self.tab_per_building.setEnabled(False)
        else:
            self.tab_per_building.setEnabled(True)

    def generate_report(self):
        if not self.data:
            QMessageBox.warning(self, "No Data", "Please load a CSV file first.")
            return

        selected_sites = [item.text() for item in self.site_list.selectedItems()]
        if not selected_sites:
            QMessageBox.warning(self, "No Sites", "Select at least one site.")
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self, "Save Excel Report", "", "Excel Files (*.xlsx)"
        )
        if not output_path:
            return
        if not output_path.endswith('.xlsx'):
            output_path += '.xlsx'

        date_from = datetime.combine(self.date_from.date().toPyDate(), datetime.min.time())
        date_to   = datetime.combine(self.date_to.date().toPyDate(),   datetime.max.time())

        self.status_bar.showMessage("Generating report…")
        QApplication.processEvents()

        try:
            generate_excel_report(
                data=self.data,
                selected_sites=selected_sites,
                output_path=output_path,
                date_from=date_from,
                date_to=date_to,
                aggregate_floors=self.combine_floors_cb.isChecked(),
                tab_per_building=self.tab_per_building.isChecked()
            )
            self.status_bar.showMessage(f"Report saved  —  {output_path}")
            if self.open_after_checkbox.isChecked():
                self.open_in_default_app(output_path)
            else:
                QMessageBox.information(self, "Done", f"Report saved to:\n{output_path}")
        except Exception as e:
            self.status_bar.showMessage("Error generating report")
            QMessageBox.critical(self, "Error", str(e))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ReportUI()
    window.show()
    sys.exit(app.exec())