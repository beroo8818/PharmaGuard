import csv
import sys

from app.ui_helpers import create_scroll_layout

from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox,
    QFileDialog,
)
from PySide6.QtCore import Qt

from database.operational_forecasting import get_forecast_rows
from app.session import get_current_user_label


class OperationalForecast(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Operational Forecasting - PharmaGuard")
        self.resize(1400, 800)

        self.rows = []

        layout = create_scroll_layout(self)

        title = QLabel("Operational Forecasting / التنبؤ من حركات الصرف")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 22px; font-weight: bold;")
        layout.addWidget(title)

        self.user_label = QLabel(f"Current user: {get_current_user_label()}")
        self.user_label.setAlignment(Qt.AlignCenter)
        self.user_label.setStyleSheet("font-weight: bold; color: #2563eb;")
        layout.addWidget(self.user_label)

        note = QLabel(
            "هذه الشاشة تستخدم حركات الصرف الفعلية Issue من قاعدة البيانات "
            "لحساب الاستهلاك، أيام التغطية، وخطر النقص."
        )
        note.setWordWrap(True)
        note.setAlignment(Qt.AlignCenter)
        layout.addWidget(note)

        controls = QHBoxLayout()

        self.coverage_days = QSpinBox()
        self.coverage_days.setRange(7, 365)
        self.coverage_days.setValue(60)

        self.lead_time_days = QSpinBox()
        self.lead_time_days.setRange(1, 180)
        self.lead_time_days.setValue(14)

        self.refresh_btn = QPushButton("Refresh Forecast / تحديث")
        self.refresh_btn.clicked.connect(self.load_forecast)

        self.export_btn = QPushButton("Export CSV / تصدير CSV")
        self.export_btn.clicked.connect(self.export_csv)

        controls.addWidget(QLabel("Coverage days:"))
        controls.addWidget(self.coverage_days)
        controls.addWidget(QLabel("Lead time days:"))
        controls.addWidget(self.lead_time_days)
        controls.addWidget(self.refresh_btn)
        controls.addWidget(self.export_btn)
        controls.addStretch()

        layout.addLayout(controls)

        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet("font-weight: bold; color: #2563eb;")
        layout.addWidget(self.status_label)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        layout.addWidget(self.table)

        self.load_forecast()

    def load_forecast(self):
        try:
            self.rows = get_forecast_rows(
                coverage_days=int(self.coverage_days.value()),
                lead_time_days=int(self.lead_time_days.value()),
            )

            columns = [
                "item_code",
                "generic_name",
                "brand_name",
                "ven_class",
                "location_name",
                "current_stock",
                "issued_30d",
                "issued_90d",
                "avg_monthly_demand",
                "daily_demand",
                "days_left",
                "coverage_days",
                "target_stock",
                "suggested_reorder_qty",
                "shortage_risk",
                "forecast_basis",
                "recommended_action",
            ]

            self.table.setSortingEnabled(False)
            self.table.clear()
            self.table.setColumnCount(len(columns))
            self.table.setRowCount(len(self.rows))
            self.table.setHorizontalHeaderLabels(columns)

            for r, row in enumerate(self.rows):
                for c, col in enumerate(columns):
                    value = row.get(col, "")
                    item = QTableWidgetItem(str(value))
                    self.table.setItem(r, c, item)

            self.table.resizeColumnsToContents()
            self.table.setSortingEnabled(True)

            critical = sum(1 for r in self.rows if r.get("shortage_risk") == "Critical")
            high = sum(1 for r in self.rows if r.get("shortage_risk") == "High")
            medium = sum(1 for r in self.rows if r.get("shortage_risk") == "Medium")

            self.status_label.setText(
                f"Loaded {len(self.rows)} rows | "
                f"Critical: {critical} | High: {high} | Medium: {medium}"
            )

        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))

    def export_csv(self):
        if not self.rows:
            QMessageBox.warning(self, "No data", "No forecast rows to export.")
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Export operational forecast",
            "operational_forecast.csv",
            "CSV Files (*.csv)"
        )

        if not path:
            return

        if not path.lower().endswith(".csv"):
            path += ".csv"

        columns = list(self.rows[0].keys())

        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=columns)
            writer.writeheader()
            writer.writerows(self.rows)

        QMessageBox.information(
            self,
            "Export done",
            f"Forecast exported successfully:\n{path}"
        )


def main():
    app = QApplication(sys.argv)
    window = OperationalForecast()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
