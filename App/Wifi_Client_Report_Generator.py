# gui_main.py

import sys
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog,
    QListWidget, QLabel, QMessageBox, QListWidgetItem, QCheckBox, QDateEdit, QHBoxLayout
)
from PyQt6.QtCore import Qt, QDate
import webbrowser
from report_generator import ingest_files, generate_excel_report
from datetime import datetime
from report_generator import normalize_datetime

class ReportUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("📊 WiFi Client Report Generator")
        self.setGeometry(200, 200, 500, 600)

        layout = QVBoxLayout()

        self.load_btn = QPushButton("📂 Load CSV/ZIP File(s)")
        self.load_btn.clicked.connect(self.load_csv)
        layout.addWidget(self.load_btn)

        self.combine_floors_cb = QCheckBox("Aggregate per-floor data")
        self.combine_floors_cb.stateChanged.connect(self.toggle_combine_floors)

        self.tab_per_building = QCheckBox("Report Tab per building")
        self.tab_per_building.setEnabled(False)

        layout.addWidget(self.combine_floors_cb)

        layout.addWidget(self.tab_per_building)

        layout.addWidget(QLabel("Select Site(s):"))
        self.site_list = QListWidget()
        self.site_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(self.site_list)

        date_layout = QHBoxLayout()
        self.date_from = QDateEdit()
        self.date_from.setDisplayFormat("yyyy-MM-dd")
        self.date_from.setCalendarPopup(True)
        # self.date_from.setDate(QDate.currentDate().addMonths(-1))
        self.date_to = QDateEdit()
        self.date_to.setDisplayFormat("yyyy-MM-dd")
        self.date_to.setCalendarPopup(True)
        # self.date_to.setDate(QDate.currentDate())

        date_layout.addWidget(QLabel("From:"))
        date_layout.addWidget(self.date_from)
        date_layout.addWidget(QLabel("To:"))
        date_layout.addWidget(self.date_to)
        layout.addLayout(date_layout)
        self.date_from.setEnabled(False)
        self.date_to.setEnabled(False)
        self.open_after_checkbox = QCheckBox("📂 Open Excel file after creation")
        layout.addWidget(self.open_after_checkbox)

        self.export_btn = QPushButton("✅ Generate Excel Report")
        self.export_btn.clicked.connect(self.generate_report)
        layout.addWidget(self.export_btn)

        self.setLayout(layout)
        self.data = []
        self.file_path = ""

    def load_csv(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Open CSV or ZIP Files",
            "",
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
        sites = sorted(set(d['location'] for d in self.data))
        for site in sites:
            self.site_list.addItem(QListWidgetItem(site))

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

        output_path, _ = QFileDialog.getSaveFileName(self, "Save Excel Report", "", "Excel Files (*.xlsx)")
        if not output_path.endswith('.xlsx'):
            output_path += '.xlsx'
        date_from = datetime.combine(self.date_from.date().toPyDate(), datetime.min.time())
        date_to = datetime.combine(self.date_to.date().toPyDate(), datetime.max.time())
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

            QMessageBox.information(self, "Done", f"Report saved to:\n{output_path}")
            if self.open_after_checkbox.isChecked():
                webbrowser.open(output_path)
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ReportUI()
    window.show()
    sys.exit(app.exec())
