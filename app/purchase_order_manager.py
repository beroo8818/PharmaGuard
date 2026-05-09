import sys

from app.ui_helpers import create_scroll_layout
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QComboBox, QLineEdit, QDoubleSpinBox, QPushButton,
    QTableWidget, QTableWidgetItem, QMessageBox, QTextEdit
)
from PySide6.QtCore import Qt

from database.schema import connect
from database.purchase_orders import (
    PO_STATUSES,
    create_purchase_order,
    add_purchase_order_line,
    approve_all_requested_quantities,
    update_purchase_order_status,
    list_purchase_orders,
    list_purchase_order_lines,
)
from app.session import get_current_user_id, get_current_user_label


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


def load_items():
    conn = connect()
    cur = conn.cursor()
    cur.execute("""
        SELECT item_id, item_code, generic_name, brand_name, unit_cost
        FROM items
        WHERE is_active = 1
        ORDER BY generic_name;
    """)
    rows = cur.fetchall()
    conn.close()
    return rows


def display_item(row):
    item_id, item_code, generic_name, brand_name, unit_cost = row
    brand = f" / {brand_name}" if brand_name else ""
    return f"{generic_name}{brand} ({item_code})"


class PurchaseOrderManager(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Purchase Order Workflow - PharmaGuard")
        self.resize(1300, 760)

        layout = create_scroll_layout(self)

        title = QLabel("Purchase Order Workflow / دورة طلبات الشراء")
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(title)

        self.user_label = QLabel(f"Current user: {get_current_user_label()}")
        self.user_label.setStyleSheet("font-weight: bold; color: #2563eb;")
        layout.addWidget(self.user_label)

        note = QLabel(
            "هذه الشاشة تنشئ طلب شراء وتغير حالته بالترتيب الصحيح. "
            "الاستلام هنا لا يزيد المخزون بعد. ربط الاستلام بالمخزون سيكون في الخطوة القادمة."
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        form = QFormLayout()

        self.supplier_combo = QComboBox()
        self.notes = QTextEdit()
        self.notes.setMaximumHeight(60)
        self.notes.setPlaceholderText("Optional PO notes")

        form.addRow("Supplier:", self.supplier_combo)
        form.addRow("Notes:", self.notes)

        layout.addLayout(form)

        buttons1 = QHBoxLayout()

        self.create_btn = QPushButton("Create Draft PO / إنشاء طلب شراء")
        self.create_btn.clicked.connect(self.create_po_clicked)

        self.refresh_btn = QPushButton("Refresh / تحديث")
        self.refresh_btn.clicked.connect(self.load_all)

        buttons1.addWidget(self.create_btn)
        buttons1.addWidget(self.refresh_btn)
        buttons1.addStretch()

        layout.addLayout(buttons1)

        line_form = QFormLayout()

        self.po_combo = QComboBox()
        self.item_combo = QComboBox()

        self.requested_qty = QDoubleSpinBox()
        self.requested_qty.setRange(0.01, 100000000)
        self.requested_qty.setDecimals(2)
        self.requested_qty.setValue(1)

        self.unit_cost = QDoubleSpinBox()
        self.unit_cost.setRange(0, 100000000)
        self.unit_cost.setDecimals(2)

        self.line_notes = QLineEdit()
        self.line_notes.setPlaceholderText("Optional line notes")

        line_form.addRow("Purchase order:", self.po_combo)
        line_form.addRow("Medicine item:", self.item_combo)
        line_form.addRow("Requested qty:", self.requested_qty)
        line_form.addRow("Unit cost:", self.unit_cost)
        line_form.addRow("Line notes:", self.line_notes)

        layout.addLayout(line_form)

        buttons2 = QHBoxLayout()

        self.add_line_btn = QPushButton("Add line / إضافة صنف")
        self.add_line_btn.clicked.connect(self.add_line_clicked)

        self.approve_qty_btn = QPushButton("Approve requested qty / اعتماد الكميات")
        self.approve_qty_btn.clicked.connect(self.approve_qty_clicked)

        buttons2.addWidget(self.add_line_btn)
        buttons2.addWidget(self.approve_qty_btn)
        buttons2.addStretch()

        layout.addLayout(buttons2)

        buttons3 = QHBoxLayout()

        self.status_combo = QComboBox()
        self.status_combo.addItems(PO_STATUSES)

        self.change_status_btn = QPushButton("Change status / تغيير الحالة")
        self.change_status_btn.clicked.connect(self.change_status_clicked)

        buttons3.addWidget(QLabel("New status:"))
        buttons3.addWidget(self.status_combo)
        buttons3.addWidget(self.change_status_btn)
        buttons3.addStretch()

        layout.addLayout(buttons3)

        tables = QHBoxLayout()

        left = QVBoxLayout()
        left.addWidget(QLabel("Purchase Orders"))
        self.po_table = QTableWidget()
        self.po_table.setAlternatingRowColors(True)
        self.po_table.setSortingEnabled(True)
        self.po_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.po_table.cellClicked.connect(self.po_clicked)
        left.addWidget(self.po_table)

        right = QVBoxLayout()
        right.addWidget(QLabel("Purchase Order Lines"))
        self.lines_table = QTableWidget()
        self.lines_table.setAlternatingRowColors(True)
        self.lines_table.setSortingEnabled(True)
        self.lines_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        right.addWidget(self.lines_table)

        tables.addLayout(left, 2)
        tables.addLayout(right, 2)

        layout.addLayout(tables)

        self.status_label = QLabel("Ready")
        layout.addWidget(self.status_label)

        self.load_all()

    def load_all(self):
        self.load_suppliers()
        self.load_items()
        self.load_purchase_orders()

    def load_suppliers(self):
        self.supplier_combo.clear()
        self.supplier_combo.addItem("-- No supplier --", None)

        for supplier_id, supplier_name in load_suppliers():
            self.supplier_combo.addItem(supplier_name, supplier_id)

    def load_items(self):
        self.item_combo.clear()

        for row in load_items():
            self.item_combo.addItem(display_item(row), row[0])

    def load_purchase_orders(self):
        rows = list_purchase_orders()

        columns = [
            "po_id",
            "po_number",
            "supplier_name",
            "status",
            "requested_by",
            "approved_by",
            "created_at",
            "approved_at",
            "notes",
        ]

        self.po_table.setSortingEnabled(False)
        self.po_table.clear()
        self.po_table.setColumnCount(len(columns))
        self.po_table.setRowCount(len(rows))
        self.po_table.setHorizontalHeaderLabels(columns)

        self.po_combo.clear()

        for r, row in enumerate(rows):
            po_id = row[0]
            po_number = row[1]
            status = row[3]

            self.po_combo.addItem(f"{po_number} [{status}]", po_id)

            for c, value in enumerate(row):
                self.po_table.setItem(r, c, QTableWidgetItem(str(value or "")))

        self.po_table.resizeColumnsToContents()
        self.po_table.setSortingEnabled(True)

        self.status_label.setText(f"Loaded {len(rows)} purchase orders.")

        if rows:
            self.load_lines(rows[0][0])

    def selected_po_id(self):
        return self.po_combo.currentData()

    def create_po_clicked(self):
        try:
            supplier_id = self.supplier_combo.currentData()
            user_id = get_current_user_id()
            notes = self.notes.toPlainText().strip()

            po_id, po_number = create_purchase_order(
                supplier_id=supplier_id,
                requested_by=user_id,
                notes=notes,
            )

            QMessageBox.information(
                self,
                "Created",
                f"PO created successfully.\nPO ID: {po_id}\nPO Number: {po_number}"
            )

            self.notes.clear()
            self.load_purchase_orders()

        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))

    def add_line_clicked(self):
        try:
            po_id = self.selected_po_id()

            if po_id is None:
                QMessageBox.warning(self, "No PO", "Create or select a PO first.")
                return

            line_id = add_purchase_order_line(
                po_id=po_id,
                item_id=self.item_combo.currentData(),
                requested_qty=float(self.requested_qty.value()),
                unit_cost=float(self.unit_cost.value()),
                notes=self.line_notes.text().strip(),
            )

            QMessageBox.information(
                self,
                "Line added",
                f"Line added successfully.\nLine ID: {line_id}"
            )

            self.line_notes.clear()
            self.load_lines(po_id)

        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))

    def approve_qty_clicked(self):
        try:
            po_id = self.selected_po_id()

            if po_id is None:
                QMessageBox.warning(self, "No PO", "Select a PO first.")
                return

            approve_all_requested_quantities(po_id)
            self.load_lines(po_id)
            QMessageBox.information(self, "Done", "Requested quantities approved.")

        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))

    def change_status_clicked(self):
        try:
            po_id = self.selected_po_id()

            if po_id is None:
                QMessageBox.warning(self, "No PO", "Select a PO first.")
                return

            new_status = self.status_combo.currentText()

            updated = update_purchase_order_status(
                po_id=po_id,
                new_status=new_status,
                user_id=get_current_user_id(),
            )

            QMessageBox.information(
                self,
                "Status updated",
                f"PO status changed to: {updated}"
            )

            self.load_purchase_orders()

        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))

    def load_lines(self, po_id):
        rows = list_purchase_order_lines(po_id)

        columns = [
            "po_line_id",
            "item_code",
            "generic_name",
            "brand_name",
            "requested_qty",
            "approved_qty",
            "received_qty",
            "unit_cost",
            "value",
            "notes",
        ]

        self.lines_table.setSortingEnabled(False)
        self.lines_table.clear()
        self.lines_table.setColumnCount(len(columns))
        self.lines_table.setRowCount(len(rows))
        self.lines_table.setHorizontalHeaderLabels(columns)

        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                self.lines_table.setItem(r, c, QTableWidgetItem(str(value or "")))

        self.lines_table.resizeColumnsToContents()
        self.lines_table.setSortingEnabled(True)

    def po_clicked(self, row, col):
        item = self.po_table.item(row, 0)

        if item:
            po_id = int(item.text())
            self.load_lines(po_id)

            for i in range(self.po_combo.count()):
                if self.po_combo.itemData(i) == po_id:
                    self.po_combo.setCurrentIndex(i)
                    break


def main():
    app = QApplication(sys.argv)
    window = PurchaseOrderManager()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
