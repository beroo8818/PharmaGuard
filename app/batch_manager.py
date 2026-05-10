from datetime import datetime, date, timedelta
import sys
from app.ui_helpers import create_scroll_layout
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QFormLayout,
    QLabel,
    QComboBox,
    QLineEdit,
    QTextEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox,
)
from PySide6.QtCore import Qt

from database.schema import connect
from database.repositories import setup_database, add_supplier, add_batch


def load_items():
    conn = connect()
    cur = conn.cursor()
    cur.execute("""
        SELECT item_id, item_code, generic_name, brand_name
        FROM items
        WHERE is_active = 1
        ORDER BY generic_name;
    """)
    rows = cur.fetchall()
    conn.close()
    return rows


def load_suppliers():
    conn = connect()
    cur = conn.cursor()
    cur.execute("""
        SELECT supplier_id, supplier_name
        FROM suppliers
        WHERE is_active = 1
        ORDER BY supplier_name;
    """)
    rows = cur.fetchall()
    conn.close()
    return rows


def load_batches():
    conn = connect()
    cur = conn.cursor()
    cur.execute("""
        SELECT
            b.batch_id,
            i.item_code,
            i.generic_name,
            i.brand_name,
            b.batch_number,
            b.expiry_date,
            s.supplier_name,
            b.received_date,
            b.notes
        FROM batches b
        JOIN items i ON b.item_id = i.item_id
        LEFT JOIN suppliers s ON b.supplier_id = s.supplier_id
        ORDER BY
            i.generic_name,
            b.expiry_date,
            b.batch_number;
    """)
    rows = cur.fetchall()
    conn.close()
    return rows


def display_item(row):
    item_id, item_code, generic_name, brand_name = row
    brand = f" / {brand_name}" if brand_name else ""
    return f"{generic_name}{brand} ({item_code})"


class BatchManager(QWidget):
    def __init__(self):
        super().__init__()

        setup_database()

        self.setWindowTitle("Batch & Expiry Manager - PharmaGuard")
        self.resize(1100, 700)

        layout = create_scroll_layout(self)

        title = QLabel("Batch & Expiry Manager / إدارة الباتش والصلاحية")
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(title)

        warning = QLabel(
            "Important: this screen stores batch/expiry data in the new SQLite database. "
            "It does not edit the old PharmaGuard file directly."
        )
        warning.setWordWrap(True)
        layout.addWidget(warning)

        form_box = QFormLayout()

        self.item_combo = QComboBox()
        self.load_item_combo()

        self.supplier_combo = QComboBox()
        self.load_supplier_combo()

        self.new_supplier_name = QLineEdit()
        self.new_supplier_name.setPlaceholderText("Optional: write new supplier name then click Add Supplier")

        self.add_supplier_btn = QPushButton("Add supplier / إضافة مورد")
        self.add_supplier_btn.clicked.connect(self.add_supplier_clicked)

        supplier_row = QHBoxLayout()
        supplier_row.addWidget(self.supplier_combo)
        supplier_row.addWidget(self.new_supplier_name)
        supplier_row.addWidget(self.add_supplier_btn)

        self.batch_number = QLineEdit()
        self.batch_number.setPlaceholderText("Example: BATCH-2026-001")

        self.expiry_date = QLineEdit()
        self.expiry_date.setPlaceholderText("YYYY-MM-DD مثال: 2026-12-31")

        self.received_date = QLineEdit()
        self.received_date.setPlaceholderText("YYYY-MM-DD مثال: 2026-01-15")

        self.notes = QTextEdit()
        self.notes.setMaximumHeight(80)
        self.notes.setPlaceholderText("Optional notes")

        form_box.addRow("Medicine item:", self.item_combo)
        form_box.addRow("Supplier:", supplier_row)
        form_box.addRow("Batch number:", self.batch_number)
        form_box.addRow("Expiry date:", self.expiry_date)
        form_box.addRow("Received date:", self.received_date)
        form_box.addRow("Notes:", self.notes)

        layout.addLayout(form_box)

        buttons = QHBoxLayout()

        self.save_btn = QPushButton("Save batch / حفظ الباتش")
        self.save_btn.clicked.connect(self.save_batch)

        self.refresh_btn = QPushButton("Refresh table / تحديث الجدول")
        self.refresh_btn.clicked.connect(self.load_table)

        buttons.addWidget(self.save_btn)
        buttons.addWidget(self.refresh_btn)
        buttons.addStretch()

        layout.addLayout(buttons)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        layout.addWidget(self.table)

        self.load_table()

    def load_item_combo(self):
        self.item_combo.clear()
        self.items = load_items()

        for row in self.items:
            self.item_combo.addItem(display_item(row), row[0])

    def load_supplier_combo(self):
        self.supplier_combo.clear()
        self.suppliers = load_suppliers()

        for supplier_id, supplier_name in self.suppliers:
            self.supplier_combo.addItem(supplier_name, supplier_id)

    def add_supplier_clicked(self):
        name = self.new_supplier_name.text().strip()

        if not name:
            QMessageBox.warning(self, "Missing supplier", "Write supplier name first.")
            return

        supplier_id = add_supplier(name)

        QMessageBox.information(
            self,
            "Supplier saved",
            f"Supplier saved successfully.\nSupplier ID: {supplier_id}"
        )

        self.new_supplier_name.clear()
        self.load_supplier_combo()

    def save_batch(self):
        try:
            if self.item_combo.currentIndex() < 0:
                QMessageBox.warning(self, "Missing item", "Please choose a medicine item.")
                return

            batch_no = self.batch_number.text().strip()
            if not batch_no:
                QMessageBox.warning(self, "Missing batch", "Please write batch number.")
                return

            expiry = self.expiry_date.text().strip()
            received = self.received_date.text().strip()
            notes = self.notes.toPlainText().strip()

            item_id = self.item_combo.currentData()
            supplier_id = self.supplier_combo.currentData()

            batch_id = add_batch(
                item_id=item_id,
                batch_number=batch_no,
                expiry_date=expiry if expiry else None,
                supplier_id=supplier_id,
                received_date=received if received else None,
                notes=notes,
            )

            QMessageBox.information(
                self,
                "Batch saved",
                f"Batch saved successfully.\nBatch ID: {batch_id}"
            )

            self.batch_number.clear()
            self.expiry_date.clear()
            self.received_date.clear()
            self.notes.clear()
            self.load_table()

        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))

    def load_table(self):
        rows = load_batches()

        columns = [
            "batch_id",
            "item_code",
            "generic_name",
            "brand_name",
            "batch_number",
            "expiry_date",
            "supplier_name",
            "received_date",
            "notes",
        ]

        self.table.setSortingEnabled(False)
        self.table.clear()
        self.table.setColumnCount(len(columns))
        self.table.setRowCount(len(rows))
        self.table.setHorizontalHeaderLabels(columns)

        for row_index, row in enumerate(rows):
            for col_index, value in enumerate(row):
                self.table.setItem(row_index, col_index, QTableWidgetItem(str(value or "")))

        self.table.resizeColumnsToContents()
        self.table.setSortingEnabled(True)


def main():
    app = QApplication(sys.argv)
    window = BatchManager()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
