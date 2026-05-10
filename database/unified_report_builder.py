from __future__ import annotations

import json
import os
import sqlite3
import sys
import zipfile
from dataclasses import dataclass
from datetime import datetime, date, timedelta
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import requests
from openpyxl import Workbook

try:
    from PIL import Image, ImageDraw, ImageFont
    HAS_PIL = True
except Exception:
    HAS_PIL = False

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

REPORTS_DIR = PROJECT_ROOT / "reports"
CONFIG_PATH = PROJECT_ROOT / "n8n_config.json"
OP_DB_NAME = "pharmaguard_operational_v12.db"
OP_DB_PATH = PROJECT_ROOT / OP_DB_NAME

REPORT_CATALOG = [
    {"code": "operational_full_report", "name_ar": "التقرير التشغيلي الكامل", "source": "Operational DB", "recommended_frequency": "يومي أو أسبوعي", "email_target": "مدير الصيدلية + المخزن + المشتريات", "description": "ملف Excel شامل: الرصيد، التوقعات، الصلاحيات، الطلبيات، الحركات، Audit Logs."},
    {"code": "executive_email_summary", "name_ar": "الصورة / الملخص المجمع للإيميل", "source": "Operational DB", "recommended_frequency": "يومي عند وجود نواقص أو أسبوعي كملخص", "email_target": "مدير الصيدلية والإدارة", "description": "صورة PNG وHTML مختصر بعدد النواقص الحرجة، تنبيهات الصلاحية، الطلبيات، والحركات الحديثة."},
    {"code": "operational_current_stock", "name_ar": "تقرير الرصيد الحالي", "source": "Operational DB", "recommended_frequency": "يومي/قبل الجرد", "email_target": "صيدلي المخزن + مدير الصيدلية", "description": "رصيد كل صنف حسب الموقع والباتش والصلاحية."},
    {"code": "operational_forecast_shortage", "name_ar": "تقرير النواقص والتوقعات", "source": "Operational DB", "recommended_frequency": "يومي عند وجود Critical/High", "email_target": "مدير الصيدلية + المشتريات", "description": "خطر النقص وكمية إعادة الطلب المقترحة."},
    {"code": "operational_expiry_alerts", "name_ar": "تقرير الصلاحية و FEFO", "source": "Operational DB", "recommended_frequency": "أسبوعي أو فوري عند Expired", "email_target": "صيدلي المخزن + مدير الصيدلية", "description": "Expired و Near Expiry من الباتشات التي لها رصيد."},
    {"code": "operational_purchase_orders", "name_ar": "تقرير الطلبيات", "source": "Operational DB", "recommended_frequency": "يومي للطلبات المفتوحة / أسبوعي للإدارة", "email_target": "المشتريات + مدير الصيدلية", "description": "حالة أوامر الشراء المفتوحة والمغلقة والملاحظات."},
    {"code": "operational_recent_movements", "name_ar": "تقرير الحركات الحديثة", "source": "Operational DB", "recommended_frequency": "داخلياً عند المراجعة أو فرق الجرد", "email_target": "صيدلي المخزن", "description": "آخر 500 حركة مخزون."},
    {"code": "operational_audit_logs", "name_ar": "تقرير سجل المراجعة Audit", "source": "Operational DB", "recommended_frequency": "أسبوعي أو عند التحقيق", "email_target": "IT + مدير الصيدلية عند الحاجة", "description": "العمليات والتعديلات المسجلة."},
    {"code": "analytics_exported_reports", "name_ar": "تقارير Analytics المصدّرة سابقاً", "source": "Analytics & Forecasting", "recommended_frequency": "بعد استيراد ملف المنصرفات الشهري", "email_target": "مدير الصيدلية + الإدارة", "description": "الملفات التي يصدّرها Analytics مثل final report أو workbook. تظهر هنا عند حفظها داخل فولدر reports."},
]

@dataclass
class ReportPackage:
    operational_report: Optional[Path]
    summary_html: Path
    summary_image: Optional[Path]
    catalog_path: Path
    latest_analytics_files: List[Path]
    metrics: Dict[str, object]


def ensure_reports_dir() -> Path:
    REPORTS_DIR.mkdir(exist_ok=True)
    return REPORTS_DIR


def connect_op() -> sqlite3.Connection:
    return sqlite3.connect(OP_DB_PATH)


def table_exists(table_name: str) -> bool:
    if not OP_DB_PATH.exists():
        return False
    try:
        with connect_op() as conn:
            return conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table_name,)).fetchone() is not None
    except Exception:
        return False


def has_operational_schema() -> bool:
    required = ["items", "locations", "batches", "stock_movements", "purchase_orders", "audit_logs"]
    return all(table_exists(t) for t in required)


def safe_query(query: str, params: Tuple = ()) -> Tuple[List[str], List[Tuple]]:
    if not OP_DB_PATH.exists():
        return [], []
    try:
        with connect_op() as conn:
            cur = conn.execute(query, params)
            cols = [d[0] for d in cur.description] if cur.description else []
            return cols, cur.fetchall()
    except Exception:
        return [], []


def rows_to_dicts(columns: Iterable[str], rows: Iterable[Tuple]) -> List[Dict[str, object]]:
    cols = list(columns)
    return [dict(zip(cols, row)) for row in rows]


def current_stock_query() -> str:
    return """
    WITH movement_lines AS (
        SELECT item_id, batch_id, to_location_id AS location_id, quantity AS signed_qty
        FROM stock_movements
        WHERE movement_type IN ('Receive', 'Return', 'Adjustment') AND to_location_id IS NOT NULL
        UNION ALL
        SELECT item_id, batch_id, to_location_id AS location_id, quantity AS signed_qty
        FROM stock_movements
        WHERE movement_type = 'Transfer' AND to_location_id IS NOT NULL
        UNION ALL
        SELECT item_id, batch_id, from_location_id AS location_id, -quantity AS signed_qty
        FROM stock_movements
        WHERE movement_type IN ('Issue', 'Waste') AND from_location_id IS NOT NULL
        UNION ALL
        SELECT item_id, batch_id, from_location_id AS location_id, -quantity AS signed_qty
        FROM stock_movements
        WHERE movement_type = 'Transfer' AND from_location_id IS NOT NULL
    )
    SELECT i.item_code, i.generic_name, COALESCE(i.brand_name, '') AS brand_name,
           COALESCE(i.dosage_form, '') AS dosage_form, l.location_name,
           COALESCE(b.batch_number, '') AS batch_number, COALESCE(b.expiry_date, '') AS expiry_date,
           ROUND(SUM(ml.signed_qty), 2) AS current_stock
    FROM movement_lines ml
    JOIN items i ON ml.item_id = i.item_id
    JOIN locations l ON ml.location_id = l.location_id
    LEFT JOIN batches b ON ml.batch_id = b.batch_id
    GROUP BY i.item_code, i.generic_name, i.brand_name, i.dosage_form, l.location_name, b.batch_number, b.expiry_date
    HAVING current_stock <> 0
    ORDER BY i.generic_name, l.location_name, b.expiry_date;
    """


def forecast_query() -> str:
    return """
    WITH stock_now AS (
        WITH movement_lines AS (
            SELECT item_id, to_location_id AS location_id, quantity AS signed_qty
            FROM stock_movements
            WHERE movement_type IN ('Receive', 'Return', 'Adjustment') AND to_location_id IS NOT NULL
            UNION ALL
            SELECT item_id, to_location_id AS location_id, quantity AS signed_qty
            FROM stock_movements
            WHERE movement_type = 'Transfer' AND to_location_id IS NOT NULL
            UNION ALL
            SELECT item_id, from_location_id AS location_id, -quantity AS signed_qty
            FROM stock_movements
            WHERE movement_type IN ('Issue', 'Waste') AND from_location_id IS NOT NULL
            UNION ALL
            SELECT item_id, from_location_id AS location_id, -quantity AS signed_qty
            FROM stock_movements
            WHERE movement_type = 'Transfer' AND from_location_id IS NOT NULL
        )
        SELECT item_id, location_id, ROUND(SUM(signed_qty), 2) AS current_stock
        FROM movement_lines
        GROUP BY item_id, location_id
    ),
    issue_30 AS (
        SELECT item_id, from_location_id AS location_id, SUM(quantity) AS issued_30d
        FROM stock_movements
        WHERE movement_type = 'Issue' AND DATE(movement_datetime) >= DATE('now', '-30 day')
        GROUP BY item_id, from_location_id
    )
    SELECT i.item_code, i.generic_name, l.location_name,
           COALESCE(sn.current_stock, 0) AS current_stock,
           COALESCE(i30.issued_30d, 0) AS issued_30d,
           CASE
             WHEN COALESCE(i30.issued_30d, 0) = 0 THEN 'No consumption data'
             WHEN sn.current_stock <= 0 THEN 'Critical'
             WHEN sn.current_stock / (i30.issued_30d / 30.0) <= 14 THEN 'Critical'
             WHEN sn.current_stock / (i30.issued_30d / 30.0) <= 28 THEN 'High'
             WHEN sn.current_stock / (i30.issued_30d / 30.0) <= 60 THEN 'Medium'
             ELSE 'Low'
           END AS shortage_risk,
           ROUND(MAX(0, (COALESCE(i30.issued_30d, 0) / 30.0 * 60) - sn.current_stock), 2) AS suggested_reorder_qty
    FROM stock_now sn
    JOIN items i ON sn.item_id = i.item_id
    JOIN locations l ON sn.location_id = l.location_id
    LEFT JOIN issue_30 i30 ON sn.item_id = i30.item_id AND sn.location_id = i30.location_id
    ORDER BY shortage_risk, i.generic_name;
    """


def expiry_alerts(columns: List[str], rows: List[Tuple]) -> Tuple[List[str], List[List[object]]]:
    idx = {name: i for i, name in enumerate(columns)}
    today = date.today()
    near_limit = today + timedelta(days=90)
    result: List[List[object]] = []
    for row in rows:
        expiry_text = str(row[idx.get("expiry_date", -1)] or "").strip() if "expiry_date" in idx else ""
        stock = float(row[idx.get("current_stock", -1)] or 0) if "current_stock" in idx else 0
        if not expiry_text or stock <= 0:
            continue
        try:
            exp = datetime.strptime(expiry_text[:10], "%Y-%m-%d").date()
        except Exception:
            continue
        if exp < today:
            risk = "Expired"
        elif exp <= near_limit:
            risk = "Near expiry"
        else:
            continue
        result.append([risk, row[idx.get("item_code")], row[idx.get("generic_name")], row[idx.get("location_name")], row[idx.get("batch_number")], expiry_text, stock])
    return ["risk", "item_code", "generic_name", "location_name", "batch_number", "expiry_date", "current_stock"], result


def add_sheet(wb: Workbook, title: str, columns: List[str], rows: Iterable[Iterable[object]]) -> None:
    ws = wb.create_sheet(title[:31])
    ws.append(columns or ["Message"])
    for row in rows:
        ws.append(list(row))
    for col in ws.columns:
        width = 10
        letter = col[0].column_letter
        for cell in col:
            width = max(width, len(str(cell.value or "")))
        ws.column_dimensions[letter].width = min(width + 2, 45)


def export_operational_report_safe(output_path: Optional[Path] = None) -> Path:
    ensure_reports_dir()
    if output_path is None:
        output_path = REPORTS_DIR / f"pharmaguard_operational_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append(["Metric", "Value"])
    ws.append(["Generated at", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    ws.append(["Database path", str(OP_DB_PATH)])
    if not has_operational_schema():
        ws.append(["Status", "Operational schema/data not ready. Run database/schema.py and import operational inventory first."])
        ws.append(["Note", "Analytics exported reports can still be attached separately from reports folder."])
        wb.save(output_path)
        return output_path

    stock_cols, stock_rows = safe_query(current_stock_query())
    forecast_cols, forecast_rows = safe_query(forecast_query())
    expiry_cols, expiry_rows = expiry_alerts(stock_cols, stock_rows)
    ws.append(["Current stock rows", len(stock_rows)])
    ws.append(["Forecast rows", len(forecast_rows)])
    ws.append(["Expiry alerts", len(expiry_rows)])
    add_sheet(wb, "Current Stock", stock_cols, stock_rows)
    add_sheet(wb, "Forecast", forecast_cols, forecast_rows)
    add_sheet(wb, "Expiry Alerts", expiry_cols, expiry_rows)
    po_cols, po_rows = safe_query("SELECT po_id, po_number, status, created_at, notes FROM purchase_orders ORDER BY po_id DESC;")
    add_sheet(wb, "Purchase Orders", po_cols, po_rows)
    mov_cols, mov_rows = safe_query("""
        SELECT sm.movement_id, sm.movement_datetime, sm.movement_type, i.generic_name, sm.quantity,
               COALESCE(sm.reference_no, '') AS reference_no, COALESCE(sm.reason, '') AS reason
        FROM stock_movements sm JOIN items i ON sm.item_id = i.item_id
        ORDER BY sm.movement_id DESC LIMIT 500;
    """)
    add_sheet(wb, "Recent Movements", mov_cols, mov_rows)
    audit_cols, audit_rows = safe_query("SELECT audit_id, created_at, action, table_name, record_id, old_value, new_value FROM audit_logs ORDER BY audit_id DESC LIMIT 500;")
    add_sheet(wb, "Audit Logs", audit_cols, audit_rows)
    wb.save(output_path)
    return output_path


def collect_operational_metrics() -> Dict[str, object]:
    metrics: Dict[str, object] = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "db_ready": has_operational_schema(),
        "stock_rows": 0,
        "forecast_rows": 0,
        "critical_shortage": 0,
        "high_shortage": 0,
        "medium_shortage": 0,
        "low_shortage": 0,
        "no_consumption_data": 0,
        "suggested_reorder_total": 0,
        "expiry_alerts": 0,
        "expired_count": 0,
        "near_expiry_count": 0,
        "purchase_orders_total": 0,
        "open_purchase_orders": 0,
        "recent_movements": 0,
        "audit_rows": 0,
    }
    if not metrics["db_ready"]:
        return metrics
    stock_cols, stock_rows = safe_query(current_stock_query())
    metrics["stock_rows"] = len(stock_rows)
    f_cols, f_rows = safe_query(forecast_query())
    forecast = rows_to_dicts(f_cols, f_rows)
    metrics["forecast_rows"] = len(forecast)
    for row in forecast:
        risk = str(row.get("shortage_risk") or "").strip()
        if risk == "Critical": metrics["critical_shortage"] += 1
        elif risk == "High": metrics["high_shortage"] += 1
        elif risk == "Medium": metrics["medium_shortage"] += 1
        elif risk == "Low": metrics["low_shortage"] += 1
        elif risk: metrics["no_consumption_data"] += 1
        try: metrics["suggested_reorder_total"] += float(row.get("suggested_reorder_qty") or 0)
        except Exception: pass
    e_cols, e_rows = expiry_alerts(stock_cols, stock_rows)
    metrics["expiry_alerts"] = len(e_rows)
    for row in e_rows:
        if row[0] == "Expired": metrics["expired_count"] += 1
        elif row[0] == "Near expiry": metrics["near_expiry_count"] += 1
    p_cols, p_rows = safe_query("SELECT status, COUNT(*) AS count_rows FROM purchase_orders GROUP BY status;")
    metrics["purchase_orders_total"] = sum(int(row[1] or 0) for row in p_rows) if p_rows else 0
    open_statuses = {"Draft", "Reviewed", "Approved", "Sent", "Partially Received", "Open", "Pending"}
    metrics["open_purchase_orders"] = sum(int(row[1] or 0) for row in p_rows if str(row[0]) in open_statuses) if p_rows else 0
    _, m_rows = safe_query("SELECT COUNT(*) FROM stock_movements;")
    if m_rows: metrics["recent_movements"] = int(m_rows[0][0] or 0)
    _, a_rows = safe_query("SELECT COUNT(*) FROM audit_logs;")
    if a_rows: metrics["audit_rows"] = int(a_rows[0][0] or 0)
    metrics["suggested_reorder_total"] = round(float(metrics["suggested_reorder_total"]), 2)
    return metrics


def latest_analytics_reports(limit: int = 5) -> List[Path]:
    ensure_reports_dir()
    candidates: List[Path] = []
    for pattern in ("*.xlsx", "*.html"):
        candidates.extend(REPORTS_DIR.glob(pattern))
    candidates = [p for p in candidates if not p.name.startswith("pharmaguard_operational_report_")]
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[:limit]


def write_report_catalog(path: Optional[Path] = None) -> Path:
    ensure_reports_dir()
    path = path or REPORTS_DIR / "pharmaguard_report_catalog.json"
    path.write_text(json.dumps(REPORT_CATALOG, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def create_summary_html(metrics: Dict[str, object], output_path: Optional[Path] = None) -> Path:
    ensure_reports_dir()
    output_path = output_path or REPORTS_DIR / f"pharmaguard_email_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
    if not metrics.get("db_ready"):
        risk_note = "<p style='color:#b45309;font-weight:bold'>قاعدة التشغيل غير جاهزة أو لم يتم تشغيل database/schema.py بعد. سيتم إرسال هيكل التقرير، لكن الأرقام ستظل صفرية.</p>"
    elif int(metrics.get("critical_shortage", 0)) > 0:
        risk_note = "<p style='color:#b91c1c;font-weight:bold'>يوجد نواقص حرجة تحتاج مراجعة بشرية قبل أي إجراء شراء أو تحويل.</p>"
    elif int(metrics.get("high_shortage", 0)) > 0:
        risk_note = "<p style='color:#b45309;font-weight:bold'>يوجد أصناف عالية الخطورة، يفضل مراجعتها اليوم.</p>"
    else:
        risk_note = "<p style='color:#166534;font-weight:bold'>لا توجد نواقص حرجة حسب بيانات قاعدة التشغيل الحالية.</p>"
    rows = [
        ("وقت الإنشاء", metrics.get("generated_at")), ("قاعدة التشغيل جاهزة", metrics.get("db_ready")),
        ("صفوف الرصيد الحالي", metrics.get("stock_rows")), ("صفوف التوقع", metrics.get("forecast_rows")),
        ("نواقص حرجة Critical", metrics.get("critical_shortage")), ("نواقص عالية High", metrics.get("high_shortage")),
        ("تنبيهات صلاحية", metrics.get("expiry_alerts")), ("Expired", metrics.get("expired_count")),
        ("Near Expiry", metrics.get("near_expiry_count")), ("إجمالي كمية إعادة الطلب المقترحة", metrics.get("suggested_reorder_total")),
        ("إجمالي الطلبيات", metrics.get("purchase_orders_total")), ("الطلبيات المفتوحة", metrics.get("open_purchase_orders")),
        ("عدد الحركات المسجلة", metrics.get("recent_movements")), ("Audit rows", metrics.get("audit_rows")),
    ]
    table = "".join(f"<tr><td>{k}</td><td>{v}</td></tr>" for k, v in rows)
    html = f"""<!doctype html><html lang='ar' dir='rtl'><head><meta charset='utf-8'><style>
body {{ font-family: Arial, Tahoma, sans-serif; color:#0f172a; }} .wrapper {{ max-width: 900px; margin:0 auto; }}
h1 {{ font-size:24px; }} .note {{ background:#f8fafc; border:1px solid #e2e8f0; border-radius:12px; padding:12px; }}
table {{ border-collapse:collapse; width:100%; margin-top:12px; }} td {{ border:1px solid #cbd5e1; padding:9px; }} td:first-child {{ font-weight:bold; background:#f8fafc; width:55%; }}
.footer {{ color:#64748b; margin-top:14px; font-size:12px; }}
</style></head><body><div class='wrapper'><h1>PharmaGuard - ملخص تشغيلي مجمع</h1><div class='note'>{risk_note}<p>هذا الملخص للدعم والمتابعة فقط، ولا يعتمد كقرار شراء أو تعديل مخزون بدون موافقة بشرية.</p></div><table>{table}</table><p class='footer'>مرفق التقرير التشغيلي الكامل وصورة الملخص وكتالوج التقارير.</p></div></body></html>"""
    output_path.write_text(html, encoding="utf-8")
    return output_path


def _font(size: int, bold: bool = False):
    if not HAS_PIL:
        return None
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    ]
    for path in candidates:
        if Path(path).exists():
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def create_summary_image(metrics: Dict[str, object], output_path: Optional[Path] = None) -> Optional[Path]:
    ensure_reports_dir()
    if not HAS_PIL:
        return None
    output_path = output_path or REPORTS_DIR / f"pharmaguard_email_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
    width, height = 1300, 760
    img = Image.new("RGB", (width, height), "#f8fafc")
    draw = ImageDraw.Draw(img)
    title_font, subtitle_font = _font(44, True), _font(22, False)
    card_title_font, card_value_font, small_font = _font(20, True), _font(38, True), _font(18, False)
    draw.rounded_rectangle((30, 30, width - 30, 140), radius=24, fill="#0f172a")
    draw.text((60, 52), "PharmaGuard Executive Summary", fill="white", font=title_font)
    draw.text((62, 104), f"Generated: {metrics.get('generated_at')}", fill="#cbd5e1", font=subtitle_font)
    cards = [
        ("DB ready", metrics.get("db_ready", False), "Operational database", "#e0e7ff"),
        ("Critical shortages", metrics.get("critical_shortage", 0), "Need same-day review", "#fee2e2"),
        ("High shortages", metrics.get("high_shortage", 0), "Review today", "#ffedd5"),
        ("Expiry alerts", metrics.get("expiry_alerts", 0), f"Expired: {metrics.get('expired_count', 0)}", "#fef9c3"),
        ("Open POs", metrics.get("open_purchase_orders", 0), f"Total POs: {metrics.get('purchase_orders_total', 0)}", "#dbeafe"),
        ("Reorder total", metrics.get("suggested_reorder_total", 0), "Suggested quantity", "#dcfce7"),
    ]
    x0, y0, card_w, card_h, gap_x, gap_y = 50, 180, 380, 145, 30, 35
    for i, (title, value, sub, fill) in enumerate(cards):
        row, col = divmod(i, 3)
        x, y = x0 + col * (card_w + gap_x), y0 + row * (card_h + gap_y)
        draw.rounded_rectangle((x, y, x + card_w, y + card_h), radius=22, fill=fill, outline="#cbd5e1", width=2)
        draw.text((x + 24, y + 20), title, fill="#0f172a", font=card_title_font)
        draw.text((x + 24, y + 55), str(value), fill="#0f172a", font=card_value_font)
        draw.text((x + 24, y + 112), sub, fill="#334155", font=small_font)
    draw.rounded_rectangle((50, 560, width - 50, 700), radius=22, fill="white", outline="#cbd5e1", width=2)
    draw.text((80, 590), "Decision support only: human approval is required before purchase or stock adjustment.", fill="#b91c1c", font=subtitle_font)
    draw.text((80, 635), "Attachments: operational Excel report + this summary image + latest Analytics exports if available.", fill="#475569", font=small_font)
    img.save(output_path)
    return output_path


def build_report_package(include_analytics: bool = True) -> ReportPackage:
    metrics = collect_operational_metrics()
    operational_path = export_operational_report_safe()
    html_path = create_summary_html(metrics)
    image_path = create_summary_image(metrics)
    catalog_path = write_report_catalog()
    analytics_files = latest_analytics_reports() if include_analytics else []
    return ReportPackage(operational_report=operational_path, summary_html=html_path, summary_image=image_path, catalog_path=catalog_path, latest_analytics_files=analytics_files, metrics=metrics)


def ensure_sample_n8n_config() -> Path:
    """Create a simple n8n config if it does not exist.

    The user workflow expects:
    - POST /webhook/pharmaguard-report-send
    - Header auth: x-pharmaguard-token
    - Binary attachment field: report_file
    """
    if not CONFIG_PATH.exists():
        email_cfg = PROJECT_ROOT / "n8n_email_config.json"
        if email_cfg.exists():
            try:
                data = json.loads(email_cfg.read_text(encoding="utf-8"))
                sample = {
                    "webhook_url": data.get("webhook_url", ""),
                    "token": data.get("token", ""),
                    "to": data.get("manager_email", data.get("to", "")),
                    "subject_prefix": data.get("default_subject_prefix", "PharmaGuard"),
                    "include_latest_analytics_exports": True,
                }
            except Exception:
                sample = {"webhook_url": "https://YOUR-N8N-DOMAIN/webhook/pharmaguard-report-send", "token": "CHANGE_ME_SECRET_TOKEN", "to": "pharmacy.manager@hospital.gov", "subject_prefix": "PharmaGuard", "include_latest_analytics_exports": True}
        else:
            sample = {"webhook_url": "https://YOUR-N8N-DOMAIN/webhook/pharmaguard-report-send", "token": "CHANGE_ME_SECRET_TOKEN", "to": "pharmacy.manager@hospital.gov", "subject_prefix": "PharmaGuard", "include_latest_analytics_exports": True}
        CONFIG_PATH.write_text(json.dumps(sample, ensure_ascii=False, indent=2), encoding="utf-8")
    return CONFIG_PATH


def load_n8n_config() -> Dict[str, object]:
    ensure_sample_n8n_config()
    config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    # Also accept the older n8n_email_config.json names used by the Email Reports Center.
    if not config.get("webhook_url") or "YOUR-N8N-DOMAIN" in str(config.get("webhook_url")):
        email_cfg = PROJECT_ROOT / "n8n_email_config.json"
        if email_cfg.exists():
            try:
                other = json.loads(email_cfg.read_text(encoding="utf-8"))
                config["webhook_url"] = other.get("webhook_url", config.get("webhook_url", ""))
                config["token"] = other.get("token", config.get("token", ""))
                config["to"] = other.get("manager_email", other.get("to", config.get("to", "")))
                config["subject_prefix"] = other.get("default_subject_prefix", config.get("subject_prefix", "PharmaGuard"))
            except Exception:
                pass
    return config


def _build_n8n_report_zip(package: ReportPackage) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_path = REPORTS_DIR / f"pharmaguard_n8n_report_package_{stamp}.zip"
    files_to_add = [
        package.operational_report,
        package.summary_html,
        package.summary_image,
        package.catalog_path,
    ] + list(package.latest_analytics_files[:5])
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in files_to_add:
            if p and Path(p).exists():
                zf.write(str(p), arcname=Path(p).name)
    return zip_path


def send_package_to_n8n(package: ReportPackage, config: Optional[Dict[str, object]] = None) -> requests.Response:
    config = config or load_n8n_config()
    url = str(config.get("webhook_url") or "").strip()
    token = str(config.get("token") or "").strip()
    recipient = str(config.get("to") or config.get("manager_email") or "").strip()
    subject_prefix = str(config.get("subject_prefix") or config.get("default_subject_prefix") or "PharmaGuard").strip()
    if not url or "YOUR-N8N-DOMAIN" in url:
        raise RuntimeError(f"n8n webhook URL is not configured. Edit: {CONFIG_PATH}")
    if not token or token == "CHANGE_ME_SECRET_TOKEN":
        raise RuntimeError(f"n8n token is not configured. Edit: {CONFIG_PATH}")

    html_body = package.summary_html.read_text(encoding="utf-8") if package.summary_html.exists() else "<h2>PharmaGuard</h2><p>Report package attached.</p>"
    payload = {
        "report_code": "combined_operational_package",
        "report_name_ar": "ملخص مجمع + تقرير تشغيلي",
        "generated_at": package.metrics.get("generated_at"),
        "to": recipient,
        "subject": f"{subject_prefix} - ملخص تشغيلي وتقرير Operational",
        "body_html": html_body,
        "metrics_json": json.dumps(package.metrics, ensure_ascii=False),
    }
    headers = {"x-pharmaguard-token": token}

    report_zip = _build_n8n_report_zip(package)
    with open(report_zip, "rb") as f:
        files = {
            # IMPORTANT: the uploaded n8n workflow checks $binary.report_file
            # and the Gmail node attaches binary property report_file.
            "report_file": (report_zip.name, f, "application/zip")
        }
        response = requests.post(url, data=payload, files=files, headers=headers, timeout=120)
        response.raise_for_status()
        return response

if __name__ == "__main__":
    pkg = build_report_package(include_analytics=True)
    print("Package created:")
    print("Operational:", pkg.operational_report)
    print("Summary HTML:", pkg.summary_html)
    print("Summary image:", pkg.summary_image)
    print("Catalog:", pkg.catalog_path)
    print("Metrics:", pkg.metrics)
