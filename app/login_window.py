import sys
from app.ui_helpers import create_scroll_layout
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QFormLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QMessageBox,
)

from database.security import authenticate_user, ensure_security_columns
from app.session import save_current_user, clear_current_user


class LoginWindow(QWidget):
    def __init__(self):
        super().__init__()

        ensure_security_columns()

        self.setWindowTitle("Login - PharmaGuard")
        self.resize(450, 280)

        layout = create_scroll_layout(self)

        title = QLabel("Login / تسجيل الدخول")
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(title)

        note = QLabel(
            "بعد تسجيل الدخول، سيتم تسجيل حركات المخزون باسم هذا المستخدم."
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        form = QFormLayout()

        self.username = QLineEdit()
        self.username.setPlaceholderText("admin")

        self.password = QLineEdit()
        self.password.setEchoMode(QLineEdit.Password)
        self.password.setPlaceholderText("admin123")

        form.addRow("Username:", self.username)
        form.addRow("Password:", self.password)

        layout.addLayout(form)

        self.login_btn = QPushButton("Login / دخول")
        self.login_btn.clicked.connect(self.login_clicked)
        layout.addWidget(self.login_btn)

        self.logout_btn = QPushButton("Clear saved login / مسح تسجيل الدخول")
        self.logout_btn.clicked.connect(self.clear_login)
        layout.addWidget(self.logout_btn)

        self.result_label = QLabel("")
        self.result_label.setWordWrap(True)
        layout.addWidget(self.result_label)

    def login_clicked(self):
        user = authenticate_user(
            self.username.text().strip(),
            self.password.text()
        )

        if user is None:
            QMessageBox.warning(
                self,
                "Login failed",
                "Wrong username/password or inactive user."
            )
            self.result_label.setText("Login failed.")
            return

        save_current_user(user)

        self.result_label.setText(
            f"Login successful\n"
            f"Username: {user['username']}\n"
            f"Full name: {user['full_name']}\n"
            f"Role: {user['role']}"
        )

        QMessageBox.information(
            self,
            "Login successful",
            f"Logged in as:\n{user['username']} ({user['role']})"
        )

    def clear_login(self):
        clear_current_user()
        self.result_label.setText("Saved login cleared. Default will be admin.")
        QMessageBox.information(self, "Done", "Saved login cleared.")


def main():
    app = QApplication(sys.argv)
    window = LoginWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
