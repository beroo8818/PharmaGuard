import sys
import subprocess
from pathlib import Path

from app.ui_helpers import create_scroll_layout

from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QLabel,
    QPushButton,
    QMessageBox,
)
from PySide6.QtCore import Qt


PROJECT_ROOT = Path.cwd()
PYTHON_EXE = PROJECT_ROOT / "venv" / "Scripts" / "python.exe"


class PharmaGuardGateway(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("PharmaGuard Main Gateway")
        self.resize(650, 420)

        layout = create_scroll_layout(self)

        title = QLabel("PharmaGuard Main Gateway / بوابة تشغيل PharmaGuard")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 22px; font-weight: bold;")
        layout.addWidget(title)

        note = QLabel(
            "هذه الشاشة تربط بين البرنامج القديم والنظام التشغيلي الجديد.\n\n"
            "البرنامج القديم يظل كما هو بدون كسر.\n"
            "النظام الجديد يحتوي على Database, Stock Movements, Batch, FEFO, PO, Reports."
        )
        note.setWordWrap(True)
        note.setAlignment(Qt.AlignCenter)
        layout.addWidget(note)

        self.legacy_btn = QPushButton("Open Legacy PharmaGuard / فتح البرنامج القديم")
        self.legacy_btn.setMinimumHeight(70)
        self.legacy_btn.clicked.connect(self.open_legacy)
        layout.addWidget(self.legacy_btn)

        self.operational_btn = QPushButton("Open Operational System / فتح النظام التشغيلي الجديد")
        self.operational_btn.setMinimumHeight(70)
        self.operational_btn.clicked.connect(self.open_operational)
        layout.addWidget(self.operational_btn)

        self.reports_btn = QPushButton("Open Reports / فتح التقارير")
        self.reports_btn.setMinimumHeight(60)
        self.reports_btn.clicked.connect(lambda: self.open_script("run_reports.py"))
        layout.addWidget(self.reports_btn)

        self.backup_btn = QPushButton("Open Backup / Restore / فتح النسخ الاحتياطي")
        self.backup_btn.setMinimumHeight(60)
        self.backup_btn.clicked.connect(lambda: self.open_script("run_backup_restore.py"))
        layout.addWidget(self.backup_btn)

    def open_script(self, script_name):
        try:
            script_path = PROJECT_ROOT / script_name

            if not PYTHON_EXE.exists():
                QMessageBox.critical(
                    self,
                    "Python not found",
                    f"Python executable not found:\n{PYTHON_EXE}"
                )
                return

            if not script_path.exists():
                QMessageBox.critical(
                    self,
                    "File not found",
                    f"Script not found:\n{script_path}"
                )
                return

            subprocess.Popen(
                [str(PYTHON_EXE), str(script_path)],
                cwd=str(PROJECT_ROOT),
                shell=False
            )

        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))

    def open_legacy(self):
        legacy_runner = PROJECT_ROOT / "run.py"

        if legacy_runner.exists():
            self.open_script("run.py")
            return

        possible_files = [
            "pharmaguard_v11_forecast_database.py",
            "pharmaguard_supplychain_v8_2_importfix.py",
        ]

        for file_name in possible_files:
            if (PROJECT_ROOT / file_name).exists():
                self.open_script(file_name)
                return

        QMessageBox.warning(
            self,
            "Legacy file not found",
            "Could not find run.py or the old PharmaGuard file."
        )

    def open_operational(self):
        self.open_script("run_operational_launcher.py")


def main():
    app = QApplication(sys.argv)
    window = PharmaGuardGateway()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
