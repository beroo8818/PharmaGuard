import sys
from app.ui_helpers import create_scroll_layout
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QFormLayout,
    QLabel,
    QLineEdit,
    QComboBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox,
)
from PySide6.QtCore import Qt

from database.security import (
    ROLES,
    ensure_security_columns,
    create_user,
    list_users,
    set_user_active,
)


class UserManager(QWidget):
    def __init__(self):
        super().__init__()

        ensure_security_columns()

        self.setWindowTitle("User Manager - PharmaGuard")
        self.resize(1000, 650)

        layout = create_scroll_layout(self)

        title = QLabel("User Manager / إدارة المستخدمين")
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(title)

        note = QLabel(
            "Use this screen to create users and assign roles. "
            "Later we will connect these users to movement entry and approvals."
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        form = QFormLayout()

        self.username = QLineEdit()
        self.username.setPlaceholderText("Example: ahmed")

        self.full_name = QLineEdit()
        self.full_name.setPlaceholderText("Example: Ahmed Mohamed")

        self.role_combo = QComboBox()
        self.role_combo.addItems(ROLES)

        self.password = QLineEdit()
        self.password.setEchoMode(QLineEdit.Password)
        self.password.setPlaceholderText("Write password")

        form.addRow("Username:", self.username)
        form.addRow("Full name:", self.full_name)
        form.addRow("Role:", self.role_combo)
        form.addRow("Password:", self.password)

        layout.addLayout(form)

        buttons = QHBoxLayout()

        self.add_btn = QPushButton("Create user / إضافة مستخدم")
        self.add_btn.clicked.connect(self.add_user_clicked)

        self.refresh_btn = QPushButton("Refresh / تحديث")
        self.refresh_btn.clicked.connect(self.load_table)

        self.disable_btn = QPushButton("Disable selected / تعطيل المحدد")
        self.disable_btn.clicked.connect(lambda: self.set_selected_active(False))

        self.enable_btn = QPushButton("Enable selected / تفعيل المحدد")
        self.enable_btn.clicked.connect(lambda: self.set_selected_active(True))

        buttons.addWidget(self.add_btn)
        buttons.addWidget(self.refresh_btn)
        buttons.addWidget(self.disable_btn)
        buttons.addWidget(self.enable_btn)
        buttons.addStretch()

        layout.addLayout(buttons)

        self.status_label = QLabel("Ready")
        layout.addWidget(self.status_label)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        layout.addWidget(self.table)

        self.load_table()

    def add_user_clicked(self):
        try:
            username = self.username.text().strip()
            full_name = self.full_name.text().strip()
            role = self.role_combo.currentText()
            password = self.password.text()

            if not username or not password:
                QMessageBox.warning(
                    self,
                    "Missing data",
                    "Username and password are required."
                )
                return

            user_id = create_user(
                username=username,
                full_name=full_name,
                role=role,
                password=password,
            )

            QMessageBox.information(
                self,
                "Saved",
                f"User created successfully.\nUser ID: {user_id}"
            )

            self.username.clear()
            self.full_name.clear()
            self.password.clear()
            self.load_table()

        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))

    def selected_user_id(self):
        row = self.table.currentRow()
        if row < 0:
            return None

        item = self.table.item(row, 0)
        if item is None:
            return None

        try:
            return int(item.text())
        except Exception:
            return None

    def set_selected_active(self, is_active):
        user_id = self.selected_user_id()

        if user_id is None:
            QMessageBox.warning(self, "No selection", "Select a user first.")
            return

        set_user_active(user_id, is_active)
        self.load_table()

    def load_table(self):
        rows = list_users()

        columns = [
            "user_id",
            "username",
            "full_name",
            "role",
            "is_active",
            "created_at",
            "last_login_at",
        ]

        self.table.setSortingEnabled(False)
        self.table.clear()
        self.table.setColumnCount(len(columns))
        self.table.setRowCount(len(rows))
        self.table.setHorizontalHeaderLabels(columns)

        for row_index, row in enumerate(rows):
            for col_index, value in enumerate(row):
                self.table.setItem(
                    row_index,
                    col_index,
                    QTableWidgetItem(str(value or ""))
                )

        self.table.resizeColumnsToContents()
        self.table.setSortingEnabled(True)

        self.status_label.setText(f"Loaded {len(rows)} users.")


def main():
    app = QApplication(sys.argv)
    window = UserManager()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
