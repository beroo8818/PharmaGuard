from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QLabel,
    QComboBox,
    QLineEdit,
    QPushButton,
    QMessageBox,
)

from database.schema import init_database
from database.repositories import list_items, update_item_minimum_stock


class MinimumStockManager(QWidget):
    def __init__(self):
        super().__init__()

        init_database()

        self.setWindowTitle("Minimum Stock Manager / إدارة الحد الأدنى")
        self.resize(500, 250)

        layout = QVBoxLayout(self)

        title = QLabel("Minimum Stock Manager / إدارة الحد الأدنى للمخزون")
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(title)

        layout.addWidget(QLabel("Select Item / اختر الصنف"))
        self.item_combo = QComboBox()
        layout.addWidget(self.item_combo)

        layout.addWidget(QLabel("Minimum Stock / الحد الأدنى"))
        self.minimum_stock_input = QLineEdit()
        self.minimum_stock_input.setPlaceholderText("Example: 100")
        layout.addWidget(self.minimum_stock_input)

        self.save_button = QPushButton("Save / حفظ")
        self.save_button.clicked.connect(self.save_minimum_stock)
        layout.addWidget(self.save_button)

        self.status_label = QLabel("")
        layout.addWidget(self.status_label)

        self.load_items()

    def load_items(self):
        self.item_combo.clear()

        rows = list_items()

        for row in rows:
            item_id = row[0]
            item_code = row[1]
            generic_name = row[2]
            brand_name = row[3] or ""

            label = f"{item_code} - {generic_name}"
            if brand_name:
                label += f" ({brand_name})"

            self.item_combo.addItem(label, item_id)

    def save_minimum_stock(self):
        item_id = self.item_combo.currentData()
        value_text = self.minimum_stock_input.text().strip()

        if not value_text:
            QMessageBox.warning(self, "Validation Error", "Minimum stock is required.")
            return

        try:
            minimum_stock = float(value_text)
        except ValueError:
            QMessageBox.warning(self, "Validation Error", "Minimum stock must be a number.")
            return

        try:
            update_item_minimum_stock(item_id, minimum_stock)
            self.status_label.setText("Minimum stock saved successfully.")
            QMessageBox.information(self, "Success", "Minimum stock saved successfully.")
        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))