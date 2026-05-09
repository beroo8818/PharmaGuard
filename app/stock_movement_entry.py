import sys
from app.ui_helpers import create_scroll_layout
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QComboBox,
    QDoubleSpinBox,
    QLineEdit,
    QPushButton,
    QMessageBox,
)
from database.schema import connect
from app.session import get_current_user_id, get_current_user_label
from database.repositories import (
    setup_database,
    add_stock_movement,
    get_user_id,
)


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


def load_batches_for_item(item_id):
    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        SELECT
            batch_id,
            batch_number,
            expiry_date
        FROM batches
        WHERE item_id = ?
          AND expiry_date IS NOT NULL
          AND expiry_date != ''
          AND expiry_date >= date('now')
        ORDER BY
            expiry_date ASC,
            batch_number ASC;
    """, (item_id,))

    rows = cur.fetchall()
    conn.close()
    return rows


def calculate_available_stock(item_id, location_id, batch_id=None):
    conn = connect()
    cur = conn.cursor()

    batch_filter = ""
    params = [item_id, location_id, item_id, location_id, item_id, location_id, item_id, location_id]

    if batch_id is not None:
        batch_filter = " AND batch_id = ? "
        params = [
            item_id, location_id, batch_id,
            item_id, location_id, batch_id,
            item_id, location_id, batch_id,
            item_id, location_id, batch_id,
        ]

    query = f"""
    SELECT COALESCE(SUM(qty), 0) FROM (
        SELECT quantity AS qty
        FROM stock_movements
        WHERE item_id = ?
          AND to_location_id = ?
          {batch_filter}
          AND movement_type IN ('Receive', 'Return', 'Adjustment')

        UNION ALL

        SELECT quantity AS qty
        FROM stock_movements
        WHERE item_id = ?
          AND to_location_id = ?
          {batch_filter}
          AND movement_type = 'Transfer'

        UNION ALL

        SELECT -quantity AS qty
        FROM stock_movements
        WHERE item_id = ?
          AND from_location_id = ?
          {batch_filter}
          AND movement_type IN ('Issue', 'Waste')

        UNION ALL

        SELECT -quantity AS qty
        FROM stock_movements
        WHERE item_id = ?
          AND from_location_id = ?
          {batch_filter}
          AND movement_type = 'Transfer'
    );
    """

    cur.execute(query, params)
    value = cur.fetchone()[0]
    conn.close()

    try:
        return float(value or 0)
    except Exception:
        return 0.0


def display_item(row):
    item_id, item_code, generic_name, brand_name = row
    brand = f" / {brand_name}" if brand_name else ""
    return f"{generic_name}{brand} ({item_code})"


def display_location(row):
    location_id, location_name, location_type = row
    return f"{location_name} [{location_type}]"


def display_batch(row):
    batch_id, batch_number, expiry_date = row
    exp = expiry_date if expiry_date else "No expiry"
    return f"{batch_number} | Expiry: {exp}"


class StockMovementEntry(QWidget):
    def __init__(self):
        super().__init__()

        setup_database()

        self.setWindowTitle("Stock Movement Entry with Batch - PharmaGuard")
        self.resize(760, 520)

        self.items = load_items()
        self.locations = load_locations()

        layout = create_scroll_layout(self)

        title = QLabel("Stock Movement Entry / تسجيل حركة مخزون بالباتش")
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(title)

        note = QLabel(
            "This screen writes stock movements to the new SQLite database "
            "and links movements to Batch Number when available."
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        form = QFormLayout()

        self.movement_type = QComboBox()
        self.movement_type.addItems([
            "Receive",
            "Issue",
            "Transfer",
            "Return",
            "Waste",
            "Adjustment",
        ])
        self.movement_type.currentTextChanged.connect(self.update_help_text)

        self.item_combo = QComboBox()
        for row in self.items:
            self.item_combo.addItem(display_item(row), row[0])
        self.item_combo.currentIndexChanged.connect(self.load_batch_combo)

        self.batch_combo = QComboBox()
        self.batch_combo.currentIndexChanged.connect(self.update_available_stock_label)

        self.from_location_combo = QComboBox()
        self.from_location_combo.addItem("-- None --", None)
        for row in self.locations:
            self.from_location_combo.addItem(display_location(row), row[0])
        self.from_location_combo.currentIndexChanged.connect(self.update_available_stock_label)

        self.to_location_combo = QComboBox()
        self.to_location_combo.addItem("-- None --", None)
        for row in self.locations:
            self.to_location_combo.addItem(display_location(row), row[0])

        self.quantity = QDoubleSpinBox()
        self.quantity.setRange(0.01, 100000000)
        self.quantity.setDecimals(2)
        self.quantity.setValue(1)

        self.reference_no = QLineEdit()
        self.reference_no.setPlaceholderText("Example: GRN-001 / ISSUE-001 / TR-001")

        self.reason = QLineEdit()
        self.reason.setPlaceholderText("Example: Receive from supplier / Issue to ER / Transfer to ICU")

        self.available_stock_label = QLabel("Available stock: -")
        self.available_stock_label.setStyleSheet("font-weight: bold; color: #2563eb;")

        self.help_label = QLabel()
        self.help_label.setWordWrap(True)
        self.help_label.setStyleSheet("color: #555;")

        form.addRow("Movement type:", self.movement_type)
        form.addRow("Medicine item:", self.item_combo)
        form.addRow("Batch number:", self.batch_combo)
        form.addRow("From location:", self.from_location_combo)
        form.addRow("To location:", self.to_location_combo)
        form.addRow("Quantity:", self.quantity)
        form.addRow("Reference no:", self.reference_no)
        form.addRow("Reason:", self.reason)
        form.addRow("Stock check:", self.available_stock_label)

        layout.addLayout(form)
        layout.addWidget(self.help_label)

        buttons = QHBoxLayout()

        self.save_btn = QPushButton("Save movement / حفظ الحركة")
        self.save_btn.clicked.connect(self.save_movement)

        self.close_btn = QPushButton("Close")
        self.close_btn.clicked.connect(self.close)

        buttons.addWidget(self.save_btn)
        buttons.addWidget(self.close_btn)
        buttons.addStretch()

        layout.addLayout(buttons)

        self.load_batch_combo()
        self.update_help_text()

        if not self.items:
            QMessageBox.warning(
                self,
                "No items",
                "No medicines found in the database.\n\nRun export_from_legacy first."
            )

    def load_batch_combo(self):
        self.batch_combo.clear()
        self.batch_combo.addItem("-- No batch selected --", None)

        item_id = self.item_combo.currentData()
        if item_id is None:
            return

        batches = load_batches_for_item(item_id)

        for row in batches:
            self.batch_combo.addItem(display_batch(row), row[0])
        if self.batch_combo.count() > 1:
            self.batch_combo.setCurrentIndex(1)

        self.update_available_stock_label()

    def update_available_stock_label(self):
        item_id = self.item_combo.currentData()
        from_location_id = self.from_location_combo.currentData()
        batch_id = self.batch_combo.currentData()

        if item_id is None or from_location_id is None:
            self.available_stock_label.setText("Available stock: -")
            return

        available = calculate_available_stock(item_id, from_location_id, batch_id)
        self.available_stock_label.setText(f"Available stock in From location: {available}")

    def update_help_text(self):
        movement = self.movement_type.currentText()

        if movement == "Receive":
            text = (
            "Receive: استخدمها عند استلام دواء جديد. "
            "لازم يكون الباتش موجود في Batch Manager ومعاه Expiry Date. "
            "اختار To location والكمية."
            )

        elif movement == "Issue":
            text = (
            "Issue: استخدمها عند صرف دواء لقسم أو عيادة أو مريض. "
            "اختار From location. النظام سيختار أقرب Batch انتهاء تلقائيًا حسب FEFO. "
            "النظام سيمنع الصرف لو الكمية أكبر من الرصيد."
            )

        elif movement == "Transfer":
            text = (
            "Transfer: استخدمها عند نقل دواء من مخزن إلى صيدلية أو من موقع لموقع. "
            "لازم تختار From location و To location و Batch."
            )

        elif movement == "Return":
            text = (
            "Return: استخدمها عند رجوع دواء إلى المخزن أو الصيدلية. "
            "اختار To location ويفضل اختيار Batch."
            )

        elif movement == "Waste":
            text = (
            "Waste: استخدمها للهالك أو التالف. "
            "لازم تختار From location و Batch وتكتب سبب واضح."
            )

        else:
            text = (
            "Adjustment: استخدمها للتسوية بعد الجرد فقط. "
            "لازم تكتب سبب واضح. لا تستخدمها بدل الصرف أو الاستلام."
            )

        self.help_label.setText(text)

    def save_movement(self):
        try:
            if self.item_combo.currentIndex() < 0:
                QMessageBox.warning(self, "Missing item", "Please choose a medicine item.")
                return

            movement = self.movement_type.currentText()
            item_id = self.item_combo.currentData()
            batch_id = self.batch_combo.currentData()
            qty = float(self.quantity.value())
            from_location_id = self.from_location_combo.currentData()
            to_location_id = self.to_location_combo.currentData()
            reference_no = self.reference_no.text().strip()
            reason = self.reason.text().strip()
            admin_id = get_current_user_id()
            if movement in {"Waste", "Adjustment"} and not reason:
                QMessageBox.warning(
                    self,
                    "Reason required",
                    "You must write a reason for Waste or Adjustment."
                )
                return

            if movement in {"Issue", "Transfer", "Waste"}:
                if batch_id is None:
                    QMessageBox.warning(
                        self,
                        "Batch required",
                        "For Issue / Transfer / Waste, please choose a Batch Number."
                    )
                    return

                if from_location_id is None:
                    QMessageBox.warning(
                        self,
                        "From location required",
                        "Please choose From location."
                    )
                    return

                available = calculate_available_stock(item_id, from_location_id, batch_id)

                if qty > available:
                    QMessageBox.warning(
                        self,
                        "Not enough stock",
                        f"Quantity is greater than available stock.\n\nAvailable: {available}\nRequested: {qty}"
                    )
                    return

            movement_id = add_stock_movement(
                movement_type=movement,
                item_id=item_id,
                quantity=qty,
                from_location_id=from_location_id,
                to_location_id=to_location_id,
                batch_id=batch_id,
                reference_no=reference_no,
                reason=reason,
                user_id=admin_id,
            )

            QMessageBox.information(
                self,
                "Saved",
                f"Movement saved successfully.\nMovement ID: {movement_id}"
            )

            self.reference_no.clear()
            self.reason.clear()
            self.quantity.setValue(1)
            self.update_available_stock_label()

        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))


def main():
    app = QApplication(sys.argv)
    window = StockMovementEntry()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
