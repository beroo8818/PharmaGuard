import sys
from app.ui_helpers import create_scroll_layout
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QFormLayout, QHBoxLayout,
    QLabel, QComboBox, QDoubleSpinBox, QLineEdit, QPushButton,
    QTableWidget, QTableWidgetItem, QMessageBox
)
from PySide6.QtCore import Qt

from database.schema import connect
from app.session import get_current_user_id, get_current_user_label
from database.repositories import setup_database, add_stock_movement, get_user_id


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


def display_item(row):
    item_id, item_code, generic_name, brand_name = row
    brand = f" / {brand_name}" if brand_name else ""
    return f"{generic_name}{brand} ({item_code})"


def display_location(row):
    location_id, location_name, location_type = row
    return f"{location_name} [{location_type}]"


def get_available_batches_for_item_location(item_id, location_id):
    conn = connect()
    cur = conn.cursor()

    query = """
    WITH movement_lines AS (
        SELECT item_id, batch_id, to_location_id AS location_id, quantity AS signed_qty
        FROM stock_movements
        WHERE movement_type IN ('Receive', 'Return', 'Adjustment')
          AND to_location_id IS NOT NULL

        UNION ALL

        SELECT item_id, batch_id, to_location_id AS location_id, quantity AS signed_qty
        FROM stock_movements
        WHERE movement_type = 'Transfer'
          AND to_location_id IS NOT NULL

        UNION ALL

        SELECT item_id, batch_id, from_location_id AS location_id, -quantity AS signed_qty
        FROM stock_movements
        WHERE movement_type IN ('Issue', 'Waste')
          AND from_location_id IS NOT NULL

        UNION ALL

        SELECT item_id, batch_id, from_location_id AS location_id, -quantity AS signed_qty
        FROM stock_movements
        WHERE movement_type = 'Transfer'
          AND from_location_id IS NOT NULL
    )

    SELECT
        b.batch_id,
        b.batch_number,
        COALESCE(b.expiry_date, '') AS expiry_date,
        ROUND(SUM(ml.signed_qty), 2) AS available_stock
    FROM movement_lines ml
    JOIN batches b ON ml.batch_id = b.batch_id
    WHERE ml.item_id = ?
      AND ml.location_id = ?
      AND ml.batch_id IS NOT NULL
    GROUP BY b.batch_id, b.batch_number, b.expiry_date
    HAVING available_stock > 0
    ORDER BY
        CASE WHEN b.expiry_date IS NULL OR b.expiry_date = '' THEN 1 ELSE 0 END,
        b.expiry_date,
        b.batch_number;
    """

    cur.execute(query, (item_id, location_id))
    rows = cur.fetchall()
    conn.close()
    return rows


class FefoIssueHelper(QWidget):
    def __init__(self):
        super().__init__()

        setup_database()

        self.setWindowTitle("FEFO Issue Helper - PharmaGuard")
        self.resize(950, 650)

        self.items = load_items()
        self.locations = load_locations()
        self.suggestion_rows = []

        layout = create_scroll_layout(self)

        title = QLabel("FEFO Issue Helper / مساعد الصرف حسب أقرب صلاحية")
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(title)

        note = QLabel(
            "اختار الدواء والمكان والكمية، والبرنامج يقترح الصرف من أقرب Batch انتهاءً."
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        form = QFormLayout()

        self.item_combo = QComboBox()
        for row in self.items:
            self.item_combo.addItem(display_item(row), row[0])

        self.location_combo = QComboBox()
        for row in self.locations:
            self.location_combo.addItem(display_location(row), row[0])

        self.qty_spin = QDoubleSpinBox()
        self.qty_spin.setRange(0.01, 100000000)
        self.qty_spin.setDecimals(2)
        self.qty_spin.setValue(1)

        self.reference_no = QLineEdit()
        self.reference_no.setPlaceholderText("Example: FEFO-ISS-001")

        self.reason = QLineEdit()
        self.reason.setPlaceholderText("Example: Issue to ER")

        form.addRow("Medicine item:", self.item_combo)
        form.addRow("From location:", self.location_combo)
        form.addRow("Required issue quantity:", self.qty_spin)
        form.addRow("Reference no:", self.reference_no)
        form.addRow("Reason:", self.reason)

        layout.addLayout(form)

        buttons = QHBoxLayout()

        self.suggest_btn = QPushButton("Suggest FEFO issue / اقتراح الصرف")
        self.suggest_btn.clicked.connect(self.suggest_issue)

        self.save_btn = QPushButton("Save suggested issue / حفظ الصرف المقترح")
        self.save_btn.clicked.connect(self.save_issue)

        self.refresh_btn = QPushButton("Refresh batches / تحديث")
        self.refresh_btn.clicked.connect(self.show_available_batches)

        buttons.addWidget(self.suggest_btn)
        buttons.addWidget(self.save_btn)
        buttons.addWidget(self.refresh_btn)
        buttons.addStretch()

        layout.addLayout(buttons)

        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet("font-weight: bold; color: #2563eb;")
        layout.addWidget(self.status_label)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        layout.addWidget(self.table)

        self.show_available_batches()

    def show_available_batches(self):
        item_id = self.item_combo.currentData()
        location_id = self.location_combo.currentData()

        if item_id is None or location_id is None:
            self.status_label.setText("No item/location selected.")
            return

        rows = get_available_batches_for_item_location(item_id, location_id)

        columns = ["batch_id", "batch_number", "expiry_date", "available_stock"]

        self.table.clear()
        self.table.setColumnCount(len(columns))
        self.table.setRowCount(len(rows))
        self.table.setHorizontalHeaderLabels(columns)

        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                self.table.setItem(r, c, QTableWidgetItem(str(value or "")))

        self.table.resizeColumnsToContents()

        total = sum(float(row[3] or 0) for row in rows)
        self.status_label.setText(f"Available total stock in selected location: {total}")

    def suggest_issue(self):
        item_id = self.item_combo.currentData()
        location_id = self.location_combo.currentData()
        required_qty = float(self.qty_spin.value())

        rows = get_available_batches_for_item_location(item_id, location_id)

        if not rows:
            QMessageBox.warning(
                self,
                "No stock",
                "No available batch stock found for this item/location."
            )
            return

        total_available = sum(float(row[3] or 0) for row in rows)

        if required_qty > total_available:
            QMessageBox.warning(
                self,
                "Not enough stock",
                f"Available: {total_available}\nRequested: {required_qty}"
            )
            return

        remaining = required_qty
        suggestions = []

        for batch_id, batch_number, expiry_date, available_stock in rows:
            available_stock = float(available_stock or 0)

            if remaining <= 0:
                break

            issue_qty = min(available_stock, remaining)

            suggestions.append({
                "batch_id": batch_id,
                "batch_number": batch_number,
                "expiry_date": expiry_date,
                "available_stock": available_stock,
                "suggested_issue_qty": issue_qty,
            })

            remaining -= issue_qty

        self.suggestion_rows = suggestions

        columns = [
            "batch_id",
            "batch_number",
            "expiry_date",
            "available_stock",
            "suggested_issue_qty",
        ]

        self.table.clear()
        self.table.setColumnCount(len(columns))
        self.table.setRowCount(len(suggestions))
        self.table.setHorizontalHeaderLabels(columns)

        for r, row in enumerate(suggestions):
            for c, col in enumerate(columns):
                self.table.setItem(r, c, QTableWidgetItem(str(row.get(col, ""))))

        self.table.resizeColumnsToContents()
        self.status_label.setText(f"FEFO suggestion ready. Suggested rows: {len(suggestions)}")

    def save_issue(self):
        if not self.suggestion_rows:
            QMessageBox.warning(self, "No suggestion", "Click Suggest FEFO issue first.")
            return

        item_id = self.item_combo.currentData()
        location_id = self.location_combo.currentData()
        admin_id = get_current_user_id()
        reference = self.reference_no.text().strip() or "FEFO-ISSUE"
        reason = self.reason.text().strip() or "FEFO issue"

        try:
            count = 0

            for row in self.suggestion_rows:
                add_stock_movement(
                    movement_type="Issue",
                    item_id=item_id,
                    batch_id=row["batch_id"],
                    from_location_id=location_id,
                    to_location_id=None,
                    quantity=float(row["suggested_issue_qty"]),
                    reference_no=reference,
                    reason=reason,
                    user_id=admin_id,
                )
                count += 1

            QMessageBox.information(
                self,
                "Saved",
                f"FEFO issue saved successfully.\nMovements inserted: {count}"
            )

            self.suggestion_rows = []
            self.reference_no.clear()
            self.reason.clear()
            self.show_available_batches()

        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))


def main():
    app = QApplication(sys.argv)
    window = FefoIssueHelper()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
