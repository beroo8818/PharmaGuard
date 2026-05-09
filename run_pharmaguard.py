import sys
import subprocess
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QLabel,
    QPushButton,
    QMessageBox,
    QFrame,
    QFileDialog,
    QInputDialog,
)

from PySide6.QtCore import Qt


PROJECT_ROOT = Path(__file__).resolve().parent
LAST_GOOGLE_SHEET_FILE = PROJECT_ROOT / "last_google_sheet_url.txt"


class PharmaGuardUnifiedLauncher(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("PharmaGuard Unified Launcher")
        self.resize(850, 520)

        layout = QVBoxLayout(self)
        layout.setSpacing(18)

        title = QLabel("PharmaGuard")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("""
            QLabel {
                font-size: 32px;
                font-weight: bold;
                color: #0f172a;
            }
        """)
        layout.addWidget(title)

        subtitle = QLabel(
            "Government Hospital Pharmacy Inventory & Supply Chain Decision Support System"
        )
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("""
            QLabel {
                font-size: 16px;
                color: #475569;
            }
        """)
        layout.addWidget(subtitle)

        info = QLabel(
            "اختر الوحدة المطلوبة:\n\n"
            "1) Analytics & Forecasting: للاستيراد من Excel / Google Sheet، التحليل، النواقص، التوقعات، والتقارير.\n"
            "2) Operational Inventory: لتسجيل الوارد، الصرف، التحويل، الباتش، الصلاحية، والرصيد الفعلي."
        )
        info.setAlignment(Qt.AlignCenter)
        info.setWordWrap(True)
        info.setStyleSheet("""
            QLabel {
                font-size: 15px;
                color: #1e293b;
                padding: 12px;
            }
        """)
        layout.addWidget(info)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        layout.addWidget(line)

        import_opening_btn = QPushButton("📥 Import Opening Balance Excel / CSV")
        import_opening_btn.setMinimumHeight(70)
        import_opening_btn.clicked.connect(self.import_opening_balance)
        import_opening_btn.setStyleSheet(self.button_style())
        layout.addWidget(import_opening_btn)

        live_sync_excel_btn = QPushButton("🔄 Live Sync Excel / CSV")
        live_sync_excel_btn.setMinimumHeight(70)
        live_sync_excel_btn.clicked.connect(self.live_sync_excel)
        live_sync_excel_btn.setStyleSheet(self.button_style())
        layout.addWidget(live_sync_excel_btn)

        live_sync_google_btn = QPushButton("🌐 Live Sync Google Sheet")
        live_sync_google_btn.setMinimumHeight(70)
        live_sync_google_btn.clicked.connect(self.live_sync_google_sheet)
        live_sync_google_btn.setStyleSheet(self.button_style())
        layout.addWidget(live_sync_google_btn)

        analytics_btn = QPushButton("📊 Analytics & Forecasting Module")
        analytics_btn.setMinimumHeight(70)
        analytics_btn.clicked.connect(self.open_analytics)
        analytics_btn.setStyleSheet(self.button_style())
        layout.addWidget(analytics_btn)

        operational_btn = QPushButton("🏥 Operational Inventory Module")
        operational_btn.setMinimumHeight(70)
        operational_btn.clicked.connect(self.open_operational)
        operational_btn.setStyleSheet(self.button_style())
        layout.addWidget(operational_btn)

        reports_hub_btn = QPushButton("📦 Reports Hub / مركز التقارير والإيميل")
        reports_hub_btn.setMinimumHeight(70)
        reports_hub_btn.clicked.connect(self.open_reports_hub)
        reports_hub_btn.setStyleSheet(self.button_style())
        layout.addWidget(reports_hub_btn)

        executive_dashboard_btn = QPushButton("🏢 Executive + Storekeeper Dashboard")
        executive_dashboard_btn.setMinimumHeight(70)
        executive_dashboard_btn.clicked.connect(self.open_executive_storekeeper_dashboard)
        executive_dashboard_btn.setStyleSheet(self.button_style())
        layout.addWidget(executive_dashboard_btn)

        monthly_reports_btn = QPushButton("📅 Monthly Consumption Reports")
        monthly_reports_btn.setMinimumHeight(70)
        monthly_reports_btn.clicked.connect(self.open_monthly_consumption_reports)
        monthly_reports_btn.setStyleSheet(self.button_style())
        layout.addWidget(monthly_reports_btn)

        note = QLabel(
            "Recommended demo flow: افتح Analytics أولًا للعرض والتحليل، ثم افتح Operational لشرح التشغيل اليومي."
        )
        note.setAlignment(Qt.AlignCenter)
        note.setWordWrap(True)
        note.setStyleSheet("""
            QLabel {
                font-size: 13px;
                color: #64748b;
                padding-top: 8px;
            }
        """)
        layout.addWidget(note)

    def button_style(self):
        return """
            QPushButton {
                font-size: 18px;
                font-weight: bold;
                color: white;
                background-color: #2563eb;
                border-radius: 14px;
                padding: 14px;
            }
            QPushButton:hover {
                background-color: #1d4ed8;
            }
            QPushButton:pressed {
                background-color: #1e40af;
            }
        """

    def run_script(self, script_name):
        script_path = PROJECT_ROOT / script_name

        if not script_path.exists():
            QMessageBox.critical(
                self,
                "File not found",
                f"لم يتم العثور على الملف:\n{script_path}"
            )
            return

        try:
            subprocess.Popen(
                [sys.executable, str(script_path)],
                cwd=str(PROJECT_ROOT)
            )
        except Exception as e:
            QMessageBox.critical(
                self,
                "Error",
                f"حدث خطأ أثناء فتح الملف:\n{e}"
            )

    def open_analytics(self):
        self.run_script("pharmaguard_v11_forecast_database.py")

    def open_operational(self):
        # لو اسم ملف التشغيل عندك مختلف، غيّره هنا
        if (PROJECT_ROOT / "run_operational_launcher.py").exists():
            self.run_script("run_operational_launcher.py")
        elif (PROJECT_ROOT / "run.py").exists():
            self.run_script("run.py")
        else:
            QMessageBox.critical(
                self,
                "Operational module not found",
                "لم يتم العثور على run_operational_launcher.py أو run.py"
            )

    def open_reports_hub(self):
        self.run_script("run_unified_reports_hub.py")

    def open_executive_storekeeper_dashboard(self):
        self.run_script("run_executive_storekeeper_dashboard.py")

    def open_monthly_consumption_reports(self):
        self.run_script("run_monthly_consumption_reports.py")      

    def import_opening_balance(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Choose Excel or CSV File",
            "",
            "Excel / CSV Files (*.xlsx *.xlsm *.xls *.csv);;All Files (*.*)"
        )

        if not path:
            return
        



        confirm = QMessageBox.question(
            self,
            "Confirm Import",
            "هل أنت متأكد من الاستيراد؟\n\n"
            "لو ملف Excel لا يحتوي على Batch أو Expiry، سيتم استخدام:\n"
            "Batch = OPENING-BALANCE\n"
            "Expiry = 2099-12-31\n\n"
            "يفضل عمل Backup قبل الاستيراد."
        )

        if confirm != QMessageBox.Yes:
            return

        script_path = PROJECT_ROOT / "database" / "export_from_legacy.py"

        if not script_path.exists():
            QMessageBox.critical(
                self,
                "Import script not found",
                f"لم يتم العثور على ملف الاستيراد:\n{script_path}"
            )
            return

        try:
            result = subprocess.run(
            [sys.executable, str(script_path), path],
                cwd=str(PROJECT_ROOT),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
            )

            if result.returncode != 0:
                QMessageBox.critical(
                    self,
                    "Import Failed",
                    result.stderr or result.stdout or "Unknown import error"
                )
                return

            QMessageBox.information(
                self,
                "Import Done",
                "تم استيراد ملف Excel بنجاح.\n\n"
                "افتح Operational Inventory ثم Stock Viewer للتأكد من ظهور البيانات.\n"
                "وافتح Analytics & Forecasting لاستخدام الملف في التحليل.\n\n"
                f"{result.stdout}"
            )

        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def import_google_sheet_to_both(self):
        url, ok = QInputDialog.getText(
            self,
            "Google Sheet URL",
            "ضعي رابط Google Sheet هنا.\n\n"
            "مهم: لازم يكون Anyone with the link can view."
        )

        if not ok or not url.strip():
            return

        url = url.strip()

        try:
            LAST_GOOGLE_SHEET_FILE.write_text(url, encoding="utf-8")
        except Exception as e:
            QMessageBox.warning(
                self,
                "Warning",
                f"لم يتم حفظ رابط Google Sheet:\n{e}"
            )

        self.run_import_source(url)

    def refresh_last_google_sheet(self):
        if not LAST_GOOGLE_SHEET_FILE.exists():
            QMessageBox.warning(
                self,
                "No Google Sheet saved",
                "لم يتم حفظ رابط Google Sheet من قبل.\n\n"
                "استخدمي زر Import Google Sheet أولًا."
            )
            return

        url = LAST_GOOGLE_SHEET_FILE.read_text(encoding="utf-8").strip()

        if not url:
            QMessageBox.warning(
                self,
                "Empty Google Sheet URL",
                "رابط Google Sheet المحفوظ فارغ.\n\n"
                "استخدمي زر Import Google Sheet مرة أخرى."
            )
            return

        self.run_import_source(url)

    def run_import_source(self, path):
        confirm = QMessageBox.question(
            self,
            "Confirm Import / Refresh",
            "هل أنت متأكد من الاستيراد أو التحديث؟\n\n"
            "لو البيانات لا تحتوي على Batch أو Expiry، سيتم استخدام:\n"
            "Batch = OPENING-BALANCE\n"
            "Expiry = 2099-12-31\n\n"
            "ملاحظة: هذا Refresh يدوي وليس Live Sync لحظي."
        )

        if confirm != QMessageBox.Yes:
            return

        script_path = PROJECT_ROOT / "database" / "export_from_legacy.py"

        if not script_path.exists():
            QMessageBox.critical(
                self,
                "Import script not found",
                f"لم يتم العثور على ملف الاستيراد:\n{script_path}"
            )
            return

        try:
            result = subprocess.run(
                [sys.executable, str(script_path), path],
                cwd=str(PROJECT_ROOT),
                capture_output=True,
                text=True,
            )

            if result.returncode != 0:
                QMessageBox.critical(
                    self,
                    "Import Failed",
                    result.stderr or result.stdout or "Unknown import error"
                )
                return

            QMessageBox.information(
                self,
                "Import / Refresh Done",
                "تم الاستيراد أو التحديث بنجاح.\n\n"
                "افتحي Operational Inventory ثم Stock Viewer للتأكد من ظهور البيانات.\n"
                "وافتحي Analytics & Forecasting للتحليل.\n\n"
                f"{result.stdout}"
            )

        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

            
    def live_sync_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Choose Excel or CSV File",
            "",
            "Excel / CSV Files (*.xlsx *.xlsm *.xls *.csv);;All Files (*.*)"
        )

        if not path:
            return

        self.run_live_sync_source(path)

    def live_sync_google_sheet(self):
        url, ok = QInputDialog.getText(
            self,
            "Google Sheet URL",
            "ضعي رابط Google Sheet هنا.\n\n"
            "مهم: لازم يكون Anyone with the link can view."
        )

        if not ok or not url.strip():
            return

        url = url.strip()

        try:
            LAST_GOOGLE_SHEET_FILE.write_text(url, encoding="utf-8")
        except Exception as e:
            QMessageBox.warning(
                self,
                "Warning",
                f"لم يتم حفظ رابط Google Sheet:\n{e}"
            )

        self.run_live_sync_source(url)

    def refresh_last_google_sheet(self):
        if not LAST_GOOGLE_SHEET_FILE.exists():
            QMessageBox.warning(
                self,
                "No Google Sheet saved",
                "لم يتم حفظ رابط Google Sheet من قبل.\n\n"
                "استخدمي زر Live Sync Google Sheet أولًا."
            )
            return

        url = LAST_GOOGLE_SHEET_FILE.read_text(encoding="utf-8").strip()

        if not url:
            QMessageBox.warning(
                self,
                "Empty Google Sheet URL",
                "رابط Google Sheet المحفوظ فارغ.\n\n"
                "استخدمي زر Live Sync Google Sheet مرة أخرى."
            )
            return

        self.run_live_sync_source(url)

    def run_live_sync_source(self, source):
        confirm = QMessageBox.question(
            self,
            "Confirm Live Sync",
            "هل أنت متأكد من عمل Live Sync؟\n\n"
            "سيتم مقارنة الرصيد الموجود في الملف أو Google Sheet مع الرصيد داخل قاعدة البيانات.\n\n"
            "لو الرصيد زاد: سيتم تسجيل Receive بالفرق.\n"
            "لو الرصيد نقص: سيتم تسجيل Waste بالفرق.\n\n"
            "لو لا يوجد Batch أو Expiry، سيتم استخدام:\n"
            "OPENING-BALANCE / 2099-12-31"
        )

        if confirm != QMessageBox.Yes:
            return

        script_path = PROJECT_ROOT / "database" / "live_sync_from_source.py"

        if not script_path.exists():
            QMessageBox.critical(
                self,
                "Live Sync script not found",
                f"لم يتم العثور على ملف:\n{script_path}"
            )
            return

        try:
            result = subprocess.run(
                [sys.executable, str(script_path), source],
                cwd=str(PROJECT_ROOT),
                capture_output=True,
                text=True,
            )

            if result.returncode != 0:
                QMessageBox.critical(
                    self,
                    "Live Sync Failed",
                    result.stderr or result.stdout or "Unknown live sync error"
                )
                return

            QMessageBox.information(
                self,
                "Live Sync Done",
                "تم عمل Live Sync بنجاح.\n\n"
                "افتحي Operational Inventory ثم Stock Viewer للتأكد من الرصيد.\n\n"
                f"{result.stdout}"
            )

        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PharmaGuardUnifiedLauncher()
    window.show()
    sys.exit(app.exec())