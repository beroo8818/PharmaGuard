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
    QDoubleSpinBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox,
)
from PySide6.QtCore import Qt

from database.schema import connect
from database.supplier_deliveries import (
    setup_delivery_tables,
    list_receivable_po_lines,
    receive_po_line_to_stock,
    list_delivery_receipts,
)
from app.session import get_current_user_id, get_current_user_label


def load_locations():
    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        SELECT location_id, location_name, location_type
        FROM locations
        WHERE is_active = 1
        ORDER BY location_name;
    """)

    rows = cur.fetchall()
    conn.close()
    return rows


def display_location(row):
    location_id, location_name, location_type = row
    return f"{location_name} [{location_type}]"


def display_po_line(row):
    (
        po_line_id,
        po_id,
        po_number,
        supplier_name,
        item_id,
        item_code,
        generic_name,
        brand_name,
        approved_qty,
        received_qty,
        remaining_qty,
        status,
    ) = row

    brand = f" / {brand_name}" if brand_name else ""

    return f"Line {po_line_id} | {po_number} | {generic_name}{brand} | Remaining: {remaining_qty}"


class SupplierDeliveryManager(QWidget):
    def __init__(self):
        super().__init__()

        setup_delivery_tables()

        self.setWindowTitle("Supplier Delivery Tracking - PharmaGuard")
        self.resize(1350, 760)

        layout = create_scroll_layout(self)

        title = QLabel("Supplier Delivery Tracking / استلام توريدات الموردين")
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(title)

        self.user_label = QLabel(f"Current user: {get_current_user_label()}")
        self.user_label.setStyleSheet("font-weight: bold; color: #2563eb;")
        layout.addWidget(self.user_label)

        note = QLabel(
            "هذه الشاشة تستلم كمية من طلب شراء حالته Sent أو Partially Received، "
            "وتنشئ Batch، وتعمل Receive stock movement، وبالتالي الرصيد يزيد."
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        form = QFormLayout()

        self.po_line_combo = QComboBox()
        self.location_combo = QComboBox()

        self.received_qty = QDoubleSpinBox()
        self.received_qty.setRange(0.01, 100000000)
        self.received_qty.setDecimals(2)
        self.received_qty.setValue(1)

        self.batch_number = QLineEdit()
        self.batch_number.setPlaceholderText("Example: BATCH-DEL-001")

        self.expiry_date = QLineEdit()
        self.expiry_date.setPlaceholderText("YYYY-MM-DD مثال: 2026-12-31")

        self.delivery_no = QLineEdit()
        self.delivery_no.setPlaceholderText("Example: INV-001")

        self.notes = QLineEdit()
        self.notes.setPlaceholderText("Optional notes")

        form.addRow("PO line to receive:", self.po_line_combo)
        form.addRow("Receive into location:", self.location_combo)
        form.addRow("Received quantity:", self.received_qty)
        form.addRow("Batch number:", self.batch_number)
        form.addRow("Expiry date:", self.expiry_date)
        form.addRow("Delivery / Invoice no:", self.delivery_no)
        form.addRow("Notes:", self.notes)

        layout.addLayout(form)

        buttons = QHBoxLayout()

        self.receive_btn = QPushButton("Receive delivery / استلام التوريد")
        self.receive_btn.clicked.connect(self.receive_clicked)

        self.refresh_btn = QPushButton("Refresh / تحديث")
        self.refresh_btn.clicked.connect(self.load_all)

        buttons.addWidget(self.receive_btn)
        buttons.addWidget(self.refresh_btn)
        buttons.addStretch()

        layout.addLayout(buttons)

        self.status_label = QLabel("Ready")
        layout.addWidget(self.status_label)

        tables = QHBoxLayout()

        left = QVBoxLayout()
        left.addWidget(QLabel("Receivable PO Lines / سطور قابلة للاستلام"))

        self.receivable_table = QTableWidget()
        self.receivable_table.setAlternatingRowColors(True)
        self.receivable_table.setSortingEnabled(True)
        self.receivable_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)

        left.addWidget(self.receivable_table)

        right = QVBoxLayout()
        right.addWidget(QLabel("Delivery Receipts / سجل الاستلام"))

        self.receipts_table = QTableWidget()
        self.receipts_table.setAlternatingRowColors(True)
        self.receipts_table.setSortingEnabled(True)
        self.receipts_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)

        right.addWidget(self.receipts_table)

        tables.addLayout(left, 2)
        tables.addLayout(right, 2)

        layout.addLayout(tables)

        self.load_all()

    def load_all(self):
        self.load_locations()
        self.load_receivable_lines()
        self.load_receipts()

    def load_locations(self):
        self.location_combo.clear()

        for row in load_locations():
            self.location_combo.addItem(display_location(row), row[0])

    def load_receivable_lines(self):
        rows = list_receivable_po_lines()

        self.po_line_combo.clear()

        for row in rows:
            self.po_line_combo.addItem(display_po_line(row), row[0])

        columns = [
            "po_line_id",
            "po_id",
            "po_number",
            "supplier_name",
            "item_id",
            "item_code",
            "generic_name",
            "brand_name",
            "approved_qty",
            "received_qty",
            "remaining_qty",
            "status",
        ]

        self.receivable_table.setSortingEnabled(False)
        self.receivable_table.clear()
        self.receivable_table.setColumnCount(len(columns))
        self.receivable_table.setRowCount(len(rows))
        self.receivable_table.setHorizontalHeaderLabels(columns)

        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                self.receivable_table.setItem(r, c, QTableWidgetItem(str(value or "")))

        self.receivable_table.resizeColumnsToContents()
        self.receivable_table.setSortingEnabled(True)

        self.status_label.setText(f"Receivable PO lines: {len(rows)}")

    def load_receipts(self):
        rows = list_delivery_receipts()

        columns = [
            "delivery_id",
            "received_at",
            "po_number",
            "po_line_id",
            "item_code",
            "generic_name",
            "batch_number",
            "expiry_date",
            "location_name",
            "received_qty",
            "delivery_no",
            "received_by",
            "notes",
        ]

        self.receipts_table.setSortingEnabled(False)
        self.receipts_table.clear()
        self.receipts_table.setColumnCount(len(columns))
        self.receipts_table.setRowCount(len(rows))
        self.receipts_table.setHorizontalHeaderLabels(columns)

        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                self.receipts_table.setItem(r, c, QTableWidgetItem(str(value or "")))

        self.receipts_table.resizeColumnsToContents()
        self.receipts_table.setSortingEnabled(True)

    def receive_clicked(self):
        try:
            po_line_id = self.po_line_combo.currentData()

            if po_line_id is None:
                QMessageBox.warning(
                    self,
                    "No receivable PO line",
                    "No PO line is available for receiving.\nMake sure PO status is Sent."
                )
                return

            location_id = self.location_combo.currentData()
            qty = float(self.received_qty.value())
            batch_no = self.batch_number.text().strip()
            expiry = self.expiry_date.text().strip()
            delivery_no = self.delivery_no.text().strip()
            notes = self.notes.text().strip()

            if not batch_no:
                QMessageBox.warning(self, "Missing batch", "Batch number is required.")
                return

            result = receive_po_line_to_stock(
                po_line_id=po_line_id,
                received_qty=qty,
                location_id=location_id,
                batch_number=batch_no,
                expiry_date=expiry,
                delivery_no=delivery_no,
                received_by=get_current_user_id(),
                notes=notes,
            )

            QMessageBox.information(
                self,
                "Delivery received",
                "Delivery received successfully.\n\n"
                f"Delivery ID: {result['delivery_id']}\n"
                f"Movement ID: {result['movement_id']}\n"
                f"Batch ID: {result['batch_id']}\n"
                f"PO status: {result['po_status']}"
            )

            self.batch_number.clear()
            self.expiry_date.clear()
            self.delivery_no.clear()
            self.notes.clear()
            self.received_qty.setValue(1)

            self.load_all()

        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))


def main():
    app = QApplication(sys.argv)
    window = SupplierDeliveryManager()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
