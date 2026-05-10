import sys
import json
import traceback
from pathlib import Path
from datetime import datetime

import pandas as pd
import requests

from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QMessageBox, QComboBox, QLineEdit, QTextEdit, QFileDialog, QFrame
)
from PySide6.QtCore import Qt


PROJECT_ROOT = Path(__file__).resolve().parent
MONTHLY_DIR = PROJECT_ROOT / "monthly_outputs"
REPORTS_DIR = PROJECT_ROOT / "reports"
N8N_OUT_DIR = PROJECT_ROOT / "n8n_out"
CONFIG_PATH = PROJECT_ROOT / "n8n_email_config.json"

N8N_OUT_DIR.mkdir(exist_ok=True)

REPORT_CATALOG = [
    {
        "code": "critical_items",
        "name_ar": "الأصناف الحرجة",
        "type": "derived_sheet",
        "source_pattern": "monthly_reports_*.xlsx",
        "source_dir": "monthly_outputs",
        "sheet": "03_Critical_Items",
        "description": "أصناف Critical / High حسب الرصيد، الاستهلاك، وأيام التغطية."
    },
    {
        "code": "purchase_requests",
        "name_ar": "طلبات الشراء المقترحة",
        "type": "derived_purchase",
        "source_pattern": "monthly_reports_*.xlsx",
        "source_dir": "monthly_outputs",
        "sheet": "02_Item_Summary",
        "description": "الأصناف التي لها suggested_reorder_qty أكبر من صفر."
    },
    {
        "code": "monthly_full_workbook",
        "name_ar": "التقرير الشهري الكامل",
        "type": "latest_file",
        "source_pattern": "monthly_reports_*.xlsx",
        "source_dir": "monthly_outputs",
        "description": "كل شيتات التحليل الشهري: Dashboard، Normalized، Critical، فروع، رصيد افتتاحي، مقارنة شهرية."
    },
    {
        "code": "next_opening_balance",
        "name_ar": "رصيد افتتاحي للشهر التالي",
        "type": "latest_file",
        "source_pattern": "next_opening_balance_*.xlsx",
        "source_dir": "monthly_outputs",
        "description": "رصيد نهاية الشهر الحالي ليصبح افتتاحي الشهر التالي."
    },
    {
        "code": "normalized_monthly_data",
        "name_ar": "البيانات بعد Normalization",
        "type": "latest_file",
        "source_pattern": "normalized_*.xlsx",
        "source_dir": "monthly_outputs",
        "description": "تحويل ملف المنصرفات من Wide Format إلى Long Format للتحليل."
    },
    {
        "code": "branch_summary",
        "name_ar": "ملخص الفروع / الصيدليات",
        "type": "derived_sheet",
        "source_pattern": "monthly_reports_*.xlsx",
        "source_dir": "monthly_outputs",
        "sheet": "04_Branch_Summary",
        "description": "إجمالي المنصرف والرصيد وعدد الأصناف لكل فرع."
    },
    {
        "code": "new_medicines",
        "name_ar": "الأدوية الجديدة",
        "type": "derived_sheet",
        "source_pattern": "monthly_reports_*.xlsx",
        "source_dir": "monthly_outputs",
        "sheet": "05_New_Medicines",
        "description": "أصناف ظهرت في ملف الشهر ولم تكن موجودة في الماستر السابق."
    },
    {
        "code": "operational_excel",
        "name_ar": "التقرير التشغيلي Excel",
        "type": "latest_file",
        "source_pattern": "pharmaguard_operational_report_*.xlsx",
        "source_dir": "reports",
        "description": "تقرير التشغيل من قاعدة Operational: رصيد، توقع، صلاحيات، طلبيات، حركات، Audit."
    },
    {
        "code": "email_summary_image",
        "name_ar": "صورة الملخص للإيميل",
        "type": "latest_file",
        "source_pattern": "pharmaguard_email_summary_*.png",
        "source_dir": "reports",
        "description": "صورة مختصرة تصلح للإرسال للإدارة."
    },
    {
        "code": "email_summary_html",
        "name_ar": "ملخص الإيميل HTML",
        "type": "latest_file",
        "source_pattern": "pharmaguard_email_summary_*.html",
        "source_dir": "reports",
        "description": "ملخص HTML للتقرير التشغيلي."
    },
]


def load_config():
    if not CONFIG_PATH.exists():
        CONFIG_PATH.write_text(json.dumps({
            "webhook_url": "https://YOUR-N8N-DOMAIN/webhook/pharmaguard-report-send",
            "token": "CHANGE_ME_SECRET_TOKEN",
            "manager_email": "manager@hospital.gov",
            "default_subject_prefix": "PharmaGuard"
        }, ensure_ascii=False, indent=2), encoding="utf-8")

    return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))


def save_config(config):
    CONFIG_PATH.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")


def latest_file(folder: Path, pattern: str):
    if not folder.exists():
        return None
    files = sorted(folder.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0] if files else None


def source_folder(name: str) -> Path:
    if name == "monthly_outputs":
        return MONTHLY_DIR
    if name == "reports":
        return REPORTS_DIR
    return PROJECT_ROOT / name


def safe_html_table(df: pd.DataFrame, max_rows=20):
    if df.empty:
        return "<p>لا توجد بيانات في هذا التقرير.</p>"
    view = df.head(max_rows).copy()
    return view.to_html(index=False, escape=True, border=1)


def clean_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in out.columns:
        if out[col].dtype == object:
            out[col] = out[col].astype(str).replace({"nan": "", "None": ""})
    return out


def build_derived_report(report_def):
    folder = source_folder(report_def["source_dir"])
    source = latest_file(folder, report_def["source_pattern"])
    if not source:
        raise FileNotFoundError(f"لا يوجد ملف مصدر مطابق: {folder / report_def['source_pattern']}")

    sheet = report_def.get("sheet")
    df = pd.read_excel(source, sheet_name=sheet)

    if report_def["type"] == "derived_purchase":
        # Keep items where suggested reorder quantity is positive.
        if "suggested_reorder_qty" in df.columns:
            qty = pd.to_numeric(df["suggested_reorder_qty"], errors="coerce").fillna(0)
            df = df[qty > 0].copy()
        elif "shortage_risk" in df.columns:
            df = df[df["shortage_risk"].astype(str).isin(["Critical", "High"])].copy()

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = N8N_OUT_DIR / f"{report_def['code']}_{timestamp}.xlsx"

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        clean_for_excel(df).to_excel(writer, index=False, sheet_name=report_def["code"][:31])

    return out_path, df, source


def build_report_menu_html():
    rows = []
    for report in REPORT_CATALOG:
        rows.append(
            f"<tr><td><b>{report['name_ar']}</b></td>"
            f"<td>{report['code']}</td>"
            f"<td>{report['description']}</td></tr>"
        )
    return f"""
    <html>
    <body style="font-family:Arial, sans-serif; direction:rtl;">
      <h2>PharmaGuard - قائمة التقارير المتاحة</h2>
      <p>اختر التقرير المطلوب من شاشة البرنامج، ثم اضغط إرسال التقرير المختار.</p>
      <table border="1" cellpadding="6" cellspacing="0">
        <tr><th>اسم التقرير</th><th>الكود</th><th>الاستخدام</th></tr>
        {''.join(rows)}
      </table>
    </body>
    </html>
    """


def build_html_summary(report_code, report_name, df=None, note="", source_file=None):
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    count = 0 if df is None else len(df)
    table = "" if df is None else safe_html_table(df)

    source_line = ""
    if source_file:
        source_line = f"<p><b>مصدر البيانات:</b> {Path(source_file).name}</p>"

    return f"""
    <html>
    <body style="font-family:Arial, sans-serif; direction:rtl;">
      <h2>PharmaGuard - {report_name}</h2>
      <p><b>كود التقرير:</b> {report_code}</p>
      <p><b>وقت الإنشاء:</b> {generated_at}</p>
      <p><b>عدد السجلات:</b> {count}</p>
      {source_line}
      <p>{note}</p>
      {table}
    </body>
    </html>
    """


def send_to_n8n(webhook_url, token, to_email, report_code, report_name, html_body, attachment_path=None):
    if not webhook_url or "YOUR-N8N-DOMAIN" in webhook_url:
        raise RuntimeError("Webhook URL غير مضبوط في n8n_email_config.json")

    if not token or token == "CHANGE_ME_SECRET_TOKEN":
        raise RuntimeError("Token غير مضبوط في n8n_email_config.json")

    subject = f"PharmaGuard - {report_name}"

    data = {
        "report_code": report_code,
        "report_name_ar": report_name,
        "to_email": to_email,
        "subject": subject,
        "body_html": html_body,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    headers = {
        "x-pharmaguard-token": token
    }

    files = None
    file_handle = None

    try:
        if attachment_path:
            attachment_path = Path(attachment_path)
            file_handle = open(attachment_path, "rb")
            files = {
                "report_file": (
                    attachment_path.name,
                    file_handle,
                    "application/octet-stream"
                )
            }

        response = requests.post(
            webhook_url,
            data=data,
            files=files,
            headers=headers,
            timeout=120,
        )

        return response.status_code, response.text

    finally:
        if file_handle:
            file_handle.close()


class N8NEmailReportsCenter(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("PharmaGuard n8n Email Reports Center")
        self.resize(980, 720)

        self.config = load_config()

        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        title = QLabel("📧 PharmaGuard n8n Email Reports Center")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: #0f172a;")
        layout.addWidget(title)

        subtitle = QLabel("إرسال الأصناف الحرجة تلقائياً أو اختيار تقرير محدد وإرساله للمدير.")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("font-size: 14px; color: #475569;")
        layout.addWidget(subtitle)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        layout.addWidget(line)

        form = QVBoxLayout()

        self.webhook_input = QLineEdit(self.config.get("webhook_url", ""))
        self.token_input = QLineEdit(self.config.get("token", ""))
        self.manager_email_input = QLineEdit(self.config.get("manager_email", ""))

        form.addWidget(QLabel("n8n Webhook URL"))
        form.addWidget(self.webhook_input)

        form.addWidget(QLabel("Secret Token - يرسل في Header باسم x-pharmaguard-token"))
        form.addWidget(self.token_input)

        form.addWidget(QLabel("Manager Email"))
        form.addWidget(self.manager_email_input)

        layout.addLayout(form)

        save_btn = QPushButton("💾 Save n8n Config")
        save_btn.clicked.connect(self.save_config_clicked)
        layout.addWidget(save_btn)

        line2 = QFrame()
        line2.setFrameShape(QFrame.HLine)
        layout.addWidget(line2)

        self.report_combo = QComboBox()
        for report in REPORT_CATALOG:
            self.report_combo.addItem(f"{report['name_ar']}  |  {report['code']}", report)

        layout.addWidget(QLabel("اختيار تقرير لإرساله للمدير"))
        layout.addWidget(self.report_combo)

        row = QHBoxLayout()

        send_critical_btn = QPushButton("🚨 Send Critical Items Now")
        send_critical_btn.clicked.connect(self.send_critical_clicked)
        row.addWidget(send_critical_btn)

        send_selected_btn = QPushButton("📤 Send Selected Report")
        send_selected_btn.clicked.connect(self.send_selected_clicked)
        row.addWidget(send_selected_btn)

        send_menu_btn = QPushButton("📋 Send Report Menu")
        send_menu_btn.clicked.connect(self.send_menu_clicked)
        row.addWidget(send_menu_btn)

        layout.addLayout(row)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setStyleSheet("font-family: Consolas; font-size: 12px;")
        layout.addWidget(self.log)

        self.write_log("جاهز. اضبط Webhook URL وToken وإيميل المدير ثم اضغط Save.")

    def write_log(self, text):
        self.log.append(f"[{datetime.now().strftime('%H:%M:%S')}] {text}")

    def current_config(self):
        return {
            "webhook_url": self.webhook_input.text().strip(),
            "token": self.token_input.text().strip(),
            "manager_email": self.manager_email_input.text().strip(),
            "default_subject_prefix": "PharmaGuard"
        }

    def save_config_clicked(self):
        self.config = self.current_config()
        save_config(self.config)
        self.write_log("تم حفظ n8n_email_config.json")
        QMessageBox.information(self, "Saved", "تم حفظ إعدادات n8n.")

    def send_common(self, report_code, report_name, html_body, attachment_path=None):
        cfg = self.current_config()
        status, response_text = send_to_n8n(
            webhook_url=cfg["webhook_url"],
            token=cfg["token"],
            to_email=cfg["manager_email"],
            report_code=report_code,
            report_name=report_name,
            html_body=html_body,
            attachment_path=attachment_path,
        )
        self.write_log(f"n8n status: {status}")
        self.write_log(response_text[:1000])
        if 200 <= status < 300:
            QMessageBox.information(self, "Sent", "تم إرسال التقرير إلى n8n بنجاح.")
        else:
            QMessageBox.warning(self, "n8n Error", f"n8n رجع حالة: {status}\n\n{response_text[:1000]}")

    def send_critical_clicked(self):
        try:
            report_def = next(r for r in REPORT_CATALOG if r["code"] == "critical_items")
            out_path, df, source = build_derived_report(report_def)

            if df.empty:
                QMessageBox.information(
                    self,
                    "No Critical Items",
                    "لا توجد أصناف حرجة في آخر تقرير شهري. لن يتم إرسال إيميل."
                )
                self.write_log("لا توجد أصناف حرجة. لم يتم الإرسال.")
                return

            html = build_html_summary(
                "critical_items",
                "الأصناف الحرجة",
                df=df,
                note="تنبيه: هذه الأصناف تحتاج مراجعة مدير الصيدلية والمشتريات. لا يتم تنفيذ شراء تلقائي.",
                source_file=source
            )
            self.send_common("critical_items", "الأصناف الحرجة", html, out_path)

        except Exception as e:
            self.write_log(traceback.format_exc())
            QMessageBox.critical(self, "Error", str(e))

    def send_selected_clicked(self):
        try:
            report_def = self.report_combo.currentData()
            report_type = report_def["type"]

            if report_type in ["derived_sheet", "derived_purchase"]:
                out_path, df, source = build_derived_report(report_def)
                html = build_html_summary(
                    report_def["code"],
                    report_def["name_ar"],
                    df=df,
                    note=report_def["description"],
                    source_file=source
                )
                self.send_common(report_def["code"], report_def["name_ar"], html, out_path)
                return

            if report_type == "latest_file":
                folder = source_folder(report_def["source_dir"])
                file_path = latest_file(folder, report_def["source_pattern"])
                if not file_path:
                    raise FileNotFoundError(f"لا يوجد ملف مطابق: {folder / report_def['source_pattern']}")

                html = build_html_summary(
                    report_def["code"],
                    report_def["name_ar"],
                    df=None,
                    note=report_def["description"],
                    source_file=file_path
                )
                self.send_common(report_def["code"], report_def["name_ar"], html, file_path)
                return

            raise RuntimeError(f"Unknown report type: {report_type}")

        except Exception as e:
            self.write_log(traceback.format_exc())
            QMessageBox.critical(self, "Error", str(e))

    def send_menu_clicked(self):
        try:
            html = build_report_menu_html()
            self.send_common("report_menu", "قائمة التقارير المتاحة", html, None)
        except Exception as e:
            self.write_log(traceback.format_exc())
            QMessageBox.critical(self, "Error", str(e))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = N8NEmailReportsCenter()
    window.show()
    sys.exit(app.exec())
