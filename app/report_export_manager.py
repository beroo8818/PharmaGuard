
import sys
from pathlib import Path

from app.ui_helpers import create_scroll_layout

from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QMessageBox,
    QFileDialog,
    QListWidget,
)
from PySide6.QtCore import Qt

from database.operational_reports import export_operational_report, ensure_reports_dir
from app.session import get_current_user_label


class ReportExportManager(QWidget):
    def __init__(self):
        super().__init__()

        ensure_reports_dir()

        self.setWindowTitle("Report Export Manager - PharmaGuard")
        self.resize(900, 600)

        layout = create_scroll_layout(self)

        title = QLabel("Report Export Manager / تصدير تقارير الإدارة")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 22px; font-weight: bold;")
        layout.addWidget(title)

        self.user_label = QLabel(f"Current user: {get_current_user_label()}")
        self.user_label.setAlignment(Qt.AlignCenter)
        self.user_label.setStyleSheet("font-weight: bold; color: #2563eb;")
        layout.addWidget(self.user_label)

        note = QLabel(
            "هذه الشاشة تصدر تقرير Excel يحتوي على الرصيد، التوقعات، الصلاحيات، "
            "طلبات الشراء، الحركات، وسجل المراجعة."
        )
        note.setWordWrap(True)
        note.setAlignment(Qt.AlignCenter)
        layout.addWidget(note)

        buttons = QHBoxLayout()

        self.export_auto_btn = QPushButton("Export report automatically / تصدير تلقائي")
        self.export_auto_btn.clicked.connect(self.export_auto)

        self.export_as_btn = QPushButton("Export report as... / اختيار مكان الحفظ")
        self.export_as_btn.clicked.connect(self.export_as)

        self.refresh_btn = QPushButton("Refresh reports list / تحديث القائمة")
        self.refresh_btn.clicked.connect(self.load_reports)

        buttons.addWidget(self.export_auto_btn)
        buttons.addWidget(self.export_as_btn)
        buttons.addWidget(self.refresh_btn)
        buttons.addStretch()

        layout.addLayout(buttons)

        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet("font-weight: bold; color: #2563eb;")
        layout.addWidget(self.status_label)

        self.report_list = QListWidget()
        layout.addWidget(self.report_list)

        self.load_reports()

    def load_reports(self):
        self.report_list.clear()

        folder = ensure_reports_dir()
        files = sorted(folder.glob("*.xlsx"), reverse=True)

        for path in files:
            self.report_list.addItem(str(path))

        self.status_label.setText(f"Reports folder: {folder} | Files: {len(files)}")

    def export_auto(self):
        try:
            path = export_operational_report()
            QMessageBox.information(self, "Report exported", f"Saved:\n{path}")
            self.load_reports()
        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))

    def export_as(self):
        try:
            path, _ = QFileDialog.getSaveFileName(
                self,
                "Save operational report",
                str(Path.cwd() / "pharmaguard_operational_report.xlsx"),
                "Excel Files (*.xlsx)"
            )

            if not path:
                return

            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"

            saved = export_operational_report(path)
            QMessageBox.information(self, "Report exported", f"Saved:\n{saved}")
            self.load_reports()
        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))


def main():
    app = QApplication(sys.argv)
    window = ReportExportManager()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
