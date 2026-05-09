from __future__ import annotations

import json
import os
import subprocess
import sys

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QLabel, QListWidget, QMessageBox, QPushButton, QTextEdit, QVBoxLayout, QWidget, QHBoxLayout

from database.unified_report_builder import CONFIG_PATH, REPORTS_DIR, REPORT_CATALOG, build_report_package, ensure_sample_n8n_config, load_n8n_config, send_package_to_n8n, write_report_catalog, export_operational_report_safe, ensure_reports_dir


class UnifiedReportsHub(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PharmaGuard Reports Hub / مركز التقارير والإيميل")
        self.resize(1050, 720)
        ensure_reports_dir()
        self.last_package = None
        self._build_ui()
        self.refresh_report_list()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        title = QLabel("📦 مركز التقارير الموحد - PharmaGuard")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: #0f172a;")
        layout.addWidget(title)
        subtitle = QLabel("كل التقارير في مكان واحد: Operational report، ملخص/صورة مجمعة للإيميل، وتجميع آخر تقارير Analytics المصدّرة داخل فولدر reports.")
        subtitle.setWordWrap(True)
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("font-size: 14px; color: #475569;")
        layout.addWidget(subtitle)
        self.catalog = QListWidget()
        self.catalog.setMinimumHeight(210)
        layout.addWidget(self.catalog)
        buttons_1 = QHBoxLayout()
        self.export_op_btn = QPushButton("📊 Export Operational Excel")
        self.export_op_btn.clicked.connect(self.export_operational)
        buttons_1.addWidget(self.export_op_btn)
        self.package_btn = QPushButton("🖼️ Create Email Summary Package")
        self.package_btn.clicked.connect(self.create_package)
        buttons_1.addWidget(self.package_btn)
        self.email_btn = QPushButton("📧 Send Package to n8n")
        self.email_btn.clicked.connect(self.send_to_n8n)
        buttons_1.addWidget(self.email_btn)
        layout.addLayout(buttons_1)
        buttons_2 = QHBoxLayout()
        self.config_btn = QPushButton("⚙️ Create/Open n8n Config")
        self.config_btn.clicked.connect(self.create_or_open_config)
        buttons_2.addWidget(self.config_btn)
        self.open_folder_btn = QPushButton("📁 Open Reports Folder")
        self.open_folder_btn.clicked.connect(self.open_reports_folder)
        buttons_2.addWidget(self.open_folder_btn)
        self.refresh_btn = QPushButton("🔄 Refresh List")
        self.refresh_btn.clicked.connect(self.refresh_report_list)
        buttons_2.addWidget(self.refresh_btn)
        layout.addLayout(buttons_2)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(260)
        self.log.setStyleSheet("font-family: Consolas, monospace; font-size: 12px;")
        layout.addWidget(self.log)
        self._style_buttons()
        self._log("جاهز. الأفضل: Create Email Summary Package أولاً، وبعد مراجعة الملفات اضغط Send Package to n8n.")

    def _style_buttons(self):
        style = """QPushButton { font-size: 14px; font-weight: bold; color: white; background-color: #2563eb; border-radius: 10px; padding: 10px; min-height: 38px; } QPushButton:hover { background-color: #1d4ed8; } QPushButton:pressed { background-color: #1e40af; }"""
        for btn in [self.export_op_btn, self.package_btn, self.email_btn, self.config_btn, self.open_folder_btn, self.refresh_btn]:
            btn.setStyleSheet(style)

    def _log(self, text: str):
        self.log.append(text)

    def refresh_report_list(self):
        self.catalog.clear()
        for item in REPORT_CATALOG:
            self.catalog.addItem(f"{item['code']} | {item['name_ar']} | {item['recommended_frequency']} | {item['email_target']}")
        write_report_catalog()
        files = sorted(REPORTS_DIR.glob("*"), key=lambda p: p.stat().st_mtime, reverse=True)[:10]
        self._log("\nآخر الملفات في reports:")
        for path in files:
            self._log(f"- {path.name}")

    def export_operational(self):
        try:
            path = export_operational_report_safe()
            self._log(f"Operational report exported: {path}")
            QMessageBox.information(self, "Done", f"تم إنشاء التقرير التشغيلي:\n{path}")
            self.refresh_report_list()
        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))
            self._log(f"ERROR export operational: {exc}")

    def create_package(self):
        try:
            package = build_report_package(include_analytics=True)
            self.last_package = package
            self._log("\nCreated email summary package:")
            self._log(f"- Operational Excel: {package.operational_report}")
            self._log(f"- Summary HTML: {package.summary_html}")
            self._log(f"- Summary image: {package.summary_image}")
            self._log(f"- Catalog: {package.catalog_path}")
            if package.latest_analytics_files:
                self._log("- Latest Analytics exports attached if sending:")
                for p in package.latest_analytics_files: self._log(f"  * {p.name}")
            else:
                self._log("- No Analytics exported files found in reports folder yet.")
            self._log(f"- Metrics: {json.dumps(package.metrics, ensure_ascii=False)}")
            QMessageBox.information(self, "Done", "تم إنشاء ملخص الإيميل والصورة والتقرير التشغيلي داخل فولدر reports.")
            self.refresh_report_list()
        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))
            self._log(f"ERROR create package: {exc}")

    def create_or_open_config(self):
        path = ensure_sample_n8n_config()
        self._log(f"n8n config path: {path}")
        QMessageBox.information(self, "n8n Config", f"تم إنشاء/تأكيد ملف إعدادات n8n.\n\nافتح وعدّل:\n{path}\n\nغيّر webhook_url و token و to قبل الإرسال.")
        try:
            if sys.platform.startswith("win"): os.startfile(str(path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin": subprocess.Popen(["open", str(path)])
            else: subprocess.Popen(["xdg-open", str(path)])
        except Exception: pass

    def send_to_n8n(self):
        try:
            if self.last_package is None:
                self._log("No package in memory; creating a new package first...")
                self.last_package = build_report_package(include_analytics=True)
            config = load_n8n_config()
            response = send_package_to_n8n(self.last_package, config=config)
            self._log(f"n8n sent successfully. Status: {response.status_code}")
            self._log(response.text[:1000])
            QMessageBox.information(self, "Sent", "تم إرسال الملخص والتقرير إلى n8n بنجاح.")
        except Exception as exc:
            QMessageBox.critical(self, "n8n Send Failed", str(exc))
            self._log(f"ERROR send n8n: {exc}")

    def open_reports_folder(self):
        REPORTS_DIR.mkdir(exist_ok=True)
        try:
            if sys.platform.startswith("win"): os.startfile(str(REPORTS_DIR))  # type: ignore[attr-defined]
            elif sys.platform == "darwin": subprocess.Popen(["open", str(REPORTS_DIR)])
            else: subprocess.Popen(["xdg-open", str(REPORTS_DIR)])
        except Exception as exc:
            QMessageBox.warning(self, "Open folder", f"افتح الفولدر يدوياً:\n{REPORTS_DIR}\n\n{exc}")


def main():
    app = QApplication(sys.argv)
    window = UnifiedReportsHub()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
