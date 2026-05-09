import sys
import subprocess
from pathlib import Path
from app.ui_helpers import create_scroll_layout
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QGridLayout,
    QLabel,
    QPushButton,
    QMessageBox,
)
from PySide6.QtCore import Qt

from app.session import get_current_user_label
from app.permissions import user_can, get_denied_message


PROJECT_ROOT = Path.cwd()
PYTHON_EXE = PROJECT_ROOT / "venv" / "Scripts" / "python.exe"


TOOLS = [
    {
        "title": "Login / تسجيل الدخول",
        "file": "run_login.py",
        "description": "اختيار المستخدم الحالي.",
        "permission": None,
    },
    {
        "title": "Operational Dashboard / لوحة التشغيل",
        "file": "run_operational_dashboard.py",
        "description": "ملخص الرصيد والحركات والطلبات.",
        "permission": "dashboard",
    },
    {
        "title": "Stock Viewer / عرض الرصيد",
        "file": "run_stock_viewer.py",
        "description": "عرض الرصيد المحسوب من الحركات.",
        "permission": "stock_view",
    },

    {
    "title": "Minimum Stock / الحد الأدنى",
    "file": "run_minimum_stock_manager.py",
    "description": "تعديل الحد الأدنى لكل صنف.",
    "permission": "stock_view",
    },
    {
        "title": "Movement Entry / تسجيل حركة",
        "file": "run_movement_entry.py",
        "description": "وارد، صرف، تحويل، مرتجع، هالك.",
        "permission": "movement_entry",
    },
    {
        "title": "Batch Manager / الباتش والصلاحية",
        "file": "run_batch_manager.py",
        "description": "إضافة Batch Number و Expiry Date.",
        "permission": "batch_manager",
    },
    {
        "title": "FEFO Helper / الصرف حسب الصلاحية",
        "file": "run_fefo_helper.py",
        "description": "اقتراح الصرف من أقرب صلاحية.",
        "permission": "fefo_helper",
    },
    {
        "title": "Operational Forecast / التنبؤ التشغيلي",
        "file": "run_operational_forecast.py",
        "description": "توقع النواقص من حركات الصرف.",
        "permission": "forecast",
    },
    {
        "title": "Reports Export / تصدير التقارير",
        "file": "run_reports.py",
        "description": "تقرير Excel للإدارة.",
        "permission": "reports",
    },
    {
        "title": "Live Sync Center / مزامنة البيانات",
        "description": "استيراد وتحديث Excel أو Google Sheet للبرنامجين.",
        "file": "run_pharmaguard.py",
        "permission": "admin",
    },
    {
        "title": "Audit Viewer / سجل المراجعة",
        "file": "run_audit_viewer.py",
        "description": "معرفة من عمل ماذا ومتى.",
        "permission": "audit_viewer",
    },
    {
        "title": "User Manager / إدارة المستخدمين",
        "file": "run_user_manager.py",
        "description": "إضافة مستخدمين وصلاحيات.",
        "permission": "user_manager",
    },
    {
        "title": "Purchase Orders / طلبات الشراء",
        "file": "run_purchase_orders.py",
        "description": "Draft → Reviewed → Approved → Sent.",
        "permission": "purchase_orders",
    },
    {
        "title": "Supplier Deliveries / استلام التوريد",
        "file": "run_supplier_deliveries.py",
        "description": "استلام طلب شراء وزيادة الرصيد.",
        "permission": "supplier_deliveries",
    },
    {
        "title": "Backup / Restore / النسخ الاحتياطي",
        "file": "run_backup_restore.py",
        "description": "حفظ واسترجاع قاعدة البيانات.",
        "permission": "backup_restore",
    },
]


class OperationalLauncher(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("PharmaGuard Operational Launcher")
        self.resize(1050, 760)

        self.layout = create_scroll_layout(self)

        title = QLabel("PharmaGuard Operational Launcher / لوحة تشغيل النظام")
        title.setStyleSheet("font-size: 22px; font-weight: bold;")
        title.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(title)

        self.user_label = QLabel("")
        self.user_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #2563eb;")
        self.user_label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.user_label)

        note = QLabel(
            "افتح Login أولًا، ثم افتح الأدوات حسب صلاحية المستخدم."
        )
        note.setWordWrap(True)
        note.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(note)

        self.grid = QGridLayout()
        self.layout.addLayout(self.grid)

        self.refresh_btn = QPushButton("Refresh current user / تحديث المستخدم الحالي")
        self.refresh_btn.clicked.connect(self.refresh_ui)
        self.layout.addWidget(self.refresh_btn)

        self.refresh_ui()

    def clear_grid(self):
        while self.grid.count():
            item = self.grid.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    def refresh_ui(self):
        self.user_label.setText(f"Current user: {get_current_user_label()}")
        self.clear_grid()

        row = 0
        col = 0

        for tool in TOOLS:
            permission = tool["permission"]
            allowed = True if permission is None else user_can(permission)

            title = tool["title"]
            description = tool["description"]

            if allowed:
                status = "Allowed"
            else:
                status = "Denied"

            btn = QPushButton(f"{title}\n\n{description}\n\n[{status}]")
            btn.setMinimumHeight(105)

            if allowed:
                bg = "#f8fafc"
                hover = "#e0f2fe"
                border = "#0284c7"
            else:
                bg = "#fee2e2"
                hover = "#fecaca"
                border = "#dc2626"

            btn.setStyleSheet(f"""
                QPushButton {{
                    font-size: 13px;
                    text-align: center;
                    padding: 10px;
                    border-radius: 10px;
                    border: 1px solid {border};
                    background-color: {bg};
                }}
                QPushButton:hover {{
                    background-color: {hover};
                }}
            """)

            btn.clicked.connect(
                lambda checked=False, file=tool["file"], perm=permission:
                    self.open_tool(file, perm)
            )

            self.grid.addWidget(btn, row, col)

            col += 1
            if col >= 2:
                col = 0
                row += 1

    def open_tool(self, file_name, permission):
        try:
            if permission is not None and not user_can(permission):
                QMessageBox.warning(
                    self,
                    "Access denied",
                    get_denied_message(permission)
                )
                return

            script_path = PROJECT_ROOT / file_name

            if not script_path.exists():
                QMessageBox.critical(
                    self,
                    "Missing file",
                    f"File not found:\n{script_path}"
                )
                return

            if not PYTHON_EXE.exists():
                QMessageBox.critical(
                    self,
                    "Python not found",
                    f"Python executable not found:\n{PYTHON_EXE}"
                )
                return

            subprocess.Popen(
                [str(PYTHON_EXE), str(script_path)],
                cwd=str(PROJECT_ROOT),
                shell=False
            )

        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))


def main():
    app = QApplication(sys.argv)
    window = OperationalLauncher()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
