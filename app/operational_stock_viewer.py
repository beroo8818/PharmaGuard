import csv
import sys
from app.ui_helpers import create_scroll_layout
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QLabel,
    QFileDialog,
    QMessageBox,
)
from PySide6.QtCore import Qt

from database.stock_views import get_current_stock_rows


class OperationalStockViewer(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Operational Stock Viewer - PharmaGuard")
        self.resize(1200, 700)

        self.rows = []

        layout = create_scroll_layout(self)

        title = QLabel("Operational Stock Viewer / عرض الرصيد المحسوب من الحركات")
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(title)

        note = QLabel(
            "This screen reads stock from SQLite stock movements. "
            "It does not change the original PharmaGuard program."
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        buttons = QHBoxLayout()

        self.refresh_btn = QPushButton("Refresh / تحديث")
        self.refresh_btn.clicked.connect(self.load_data)

        self.export_btn = QPushButton("Export CSV / تصدير CSV")
        self.export_btn.clicked.connect(self.export_csv)

        buttons.addWidget(self.refresh_btn)
        buttons.addWidget(self.export_btn)
        buttons.addStretch()

        layout.addLayout(buttons)

        self.status_label = QLabel("Ready")
        layout.addWidget(self.status_label)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        layout.addWidget(self.table)

        self.load_data()

    def load_data(self):
        try:
            self.rows = get_current_stock_rows()

            if not self.rows:
                self.table.setRowCount(0)
                self.table.setColumnCount(0)
                self.status_label.setText(
                    "No stock found. Run export_from_legacy first or insert stock movements."
                )
                return

            columns = list(self.rows[0].keys())

            self.table.setSortingEnabled(False)
            self.table.clear()
            self.table.setColumnCount(len(columns))
            self.table.setRowCount(len(self.rows))
            self.table.setHorizontalHeaderLabels(columns)

            for row_index, row in enumerate(self.rows):
                for col_index, col_name in enumerate(columns):
                    value = row.get(col_name, "")
                    item = QTableWidgetItem(str(value))
                    self.table.setItem(row_index, col_index, item)

            self.table.resizeColumnsToContents()
            self.table.setSortingEnabled(True)

            self.status_label.setText(f"Loaded {len(self.rows)} stock rows.")

        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))

    def export_csv(self):
        if not self.rows:
            QMessageBox.warning(self, "No data", "No rows to export.")
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Export current stock",
            "current_stock_from_ledger.csv",
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

        QMessageBox.information(self, "Export done", f"Saved:\n{path}")


def main():
    app = QApplication(sys.argv)
    window = OperationalStockViewer()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
