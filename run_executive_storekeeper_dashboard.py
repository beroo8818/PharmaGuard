import sys
import sqlite3
import subprocess
from datetime import date, datetime, timedelta
from pathlib import Path

try:
    import pandas as pd
except Exception:
    pd = None

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QLabel,
    QPushButton,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox,
)

PROJECT_ROOT = Path(__file__).resolve().parent
MONTHLY_OUTPUTS = PROJECT_ROOT / "monthly_outputs"
DB_CANDIDATES = [
    PROJECT_ROOT / "pharmaguard_operational_v12.db",
    PROJECT_ROOT / "pharmaguard_operational.db",
]
LEAD_TIME_DAYS = 14
NEAR_EXPIRY_DAYS = 90
OVERDUE_PO_DAYS = 14
ADJUSTMENT_ALERT_DAYS = 30


def safe_float(value):
    try:
        if value is None:
            return 0.0
        if str(value).strip().lower() in ("", "nan", "none", "n/a"):
            return 0.0
        return float(value)
    except Exception:
        return 0.0


def latest_file(folder: Path, pattern: str):
    if not folder.exists():
        return None
    files = sorted(folder.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0] if files else None


def latest_monthly_report():
    return latest_file(MONTHLY_OUTPUTS, "monthly_reports_*.xlsx")


def latest_next_opening():
    return latest_file(MONTHLY_OUTPUTS, "next_opening_balance_*.xlsx")


def monthly_available():
    return pd is not None and latest_monthly_report() is not None


def read_monthly_sheet(sheet_name):
    if pd is None:
        return None
    path = latest_monthly_report()
    if not path:
        return None
    try:
        return pd.read_excel(path, sheet_name=sheet_name)
    except Exception:
        return None


def df_rows(df, limit=500):
    if df is None:
        return []
    try:
        clean = df.head(limit).fillna("")
        return clean.to_dict(orient="records")
    except Exception:
        return []


def metric_dict():
    df = read_monthly_sheet("00_Dashboard")
    if df is None or df.empty or "metric" not in df.columns or "value" not in df.columns:
        return {}
    return {str(r["metric"]): r["value"] for _, r in df.iterrows()}


def get_db_path():
    for path in DB_CANDIDATES:
        if path.exists():
            return path
    return DB_CANDIDATES[0]


def connect():
    conn = sqlite3.connect(get_db_path())
    conn.row_factory = sqlite3.Row
    return conn


def run_query(sql, params=()):
    try:
        conn = connect()
        rows = conn.execute(sql, params).fetchall()
        conn.close()
        return rows
    except Exception:
        return []


def scalar(sql, params=(), default=0):
    try:
        conn = connect()
        row = conn.execute(sql, params).fetchone()
        conn.close()
        if not row:
            return default
        return row[0] if row[0] is not None else default
    except Exception:
        return default


def movement_stock_cte():
    return """
    WITH movement_lines AS (
        SELECT item_id, batch_id, to_location_id AS location_id, quantity AS signed_qty
        FROM stock_movements
        WHERE movement_type IN ('Receive', 'Return', 'Adjustment')
          AND to_location_id IS NOT NULL
        UNION ALL
        SELECT item_id, batch_id, to_location_id AS location_id, quantity AS signed_qty
        FROM stock_movements
        WHERE movement_type = 'Transfer'
          AND to_location_id IS NOT NULL
        UNION ALL
        SELECT item_id, batch_id, from_location_id AS location_id, -quantity AS signed_qty
        FROM stock_movements
        WHERE movement_type IN ('Issue', 'Waste')
          AND from_location_id IS NOT NULL
        UNION ALL
        SELECT item_id, batch_id, from_location_id AS location_id, -quantity AS signed_qty
        FROM stock_movements
        WHERE movement_type = 'Transfer'
          AND from_location_id IS NOT NULL
    ),
    stock_now AS (
        SELECT item_id, batch_id, location_id, ROUND(SUM(signed_qty), 2) AS current_stock
        FROM movement_lines
        GROUP BY item_id, batch_id, location_id
        HAVING ROUND(SUM(signed_qty), 2) <> 0
    )
    """


def get_current_stock_rows(limit=500):
    sql = movement_stock_cte() + """
    SELECT
        COALESCE(i.item_code, '') AS item_code,
        i.generic_name,
        COALESCE(i.brand_name, '') AS brand_name,
        COALESCE(i.dosage_form, '') AS dosage_form,
        COALESCE(i.strength, '') AS strength,
        COALESCE(i.ven_class, '') AS ven_class,
        COALESCE(l.location_name, '') AS location_name,
        COALESCE(b.batch_number, '') AS batch_number,
        COALESCE(b.expiry_date, '') AS expiry_date,
        sn.current_stock,
        COALESCE(i.minimum_stock, 0) AS minimum_stock,
        COALESCE(i.unit_cost, 0) AS unit_cost,
        ROUND(sn.current_stock * COALESCE(i.unit_cost, 0), 2) AS stock_value
    FROM stock_now sn
    JOIN items i ON sn.item_id = i.item_id
    LEFT JOIN batches b ON sn.batch_id = b.batch_id
    LEFT JOIN locations l ON sn.location_id = l.location_id
    WHERE i.is_active = 1
    ORDER BY i.generic_name, l.location_name
    LIMIT ?;
    """
    return run_query(sql, (limit,))


def get_forecast_rows():
    sql = """
    WITH movement_lines AS (
        SELECT item_id, to_location_id AS location_id, quantity AS signed_qty
        FROM stock_movements
        WHERE movement_type IN ('Receive', 'Return', 'Adjustment')
          AND to_location_id IS NOT NULL
        UNION ALL
        SELECT item_id, to_location_id AS location_id, quantity AS signed_qty
        FROM stock_movements
        WHERE movement_type = 'Transfer'
          AND to_location_id IS NOT NULL
        UNION ALL
        SELECT item_id, from_location_id AS location_id, -quantity AS signed_qty
        FROM stock_movements
        WHERE movement_type IN ('Issue', 'Waste')
          AND from_location_id IS NOT NULL
        UNION ALL
        SELECT item_id, from_location_id AS location_id, -quantity AS signed_qty
        FROM stock_movements
        WHERE movement_type = 'Transfer'
          AND from_location_id IS NOT NULL
    ),
    stock_now AS (
        SELECT item_id, location_id, ROUND(SUM(signed_qty), 2) AS current_stock
        FROM movement_lines
        GROUP BY item_id, location_id
    ),
    issue_30 AS (
        SELECT item_id, from_location_id AS location_id, SUM(quantity) AS issued_30d
        FROM stock_movements
        WHERE movement_type = 'Issue'
          AND from_location_id IS NOT NULL
          AND DATE(movement_datetime) >= DATE('now', '-30 day')
        GROUP BY item_id, from_location_id
    ),
    issue_90 AS (
        SELECT item_id, from_location_id AS location_id, SUM(quantity) AS issued_90d
        FROM stock_movements
        WHERE movement_type = 'Issue'
          AND from_location_id IS NOT NULL
          AND DATE(movement_datetime) >= DATE('now', '-90 day')
        GROUP BY item_id, from_location_id
    )
    SELECT
        COALESCE(i.item_code, '') AS item_code,
        i.generic_name,
        COALESCE(i.brand_name, '') AS brand_name,
        COALESCE(i.ven_class, '') AS ven_class,
        COALESCE(l.location_name, '') AS location_name,
        COALESCE(sn.current_stock, 0) AS current_stock,
        COALESCE(i.minimum_stock, 0) AS minimum_stock,
        COALESCE(i30.issued_30d, 0) AS issued_30d,
        COALESCE(i90.issued_90d, 0) AS issued_90d,
        COALESCE(i.unit_cost, 0) AS unit_cost
    FROM stock_now sn
    JOIN items i ON sn.item_id = i.item_id
    LEFT JOIN locations l ON sn.location_id = l.location_id
    LEFT JOIN issue_30 i30 ON sn.item_id = i30.item_id AND sn.location_id = i30.location_id
    LEFT JOIN issue_90 i90 ON sn.item_id = i90.item_id AND sn.location_id = i90.location_id
    WHERE i.is_active = 1
    ORDER BY i.generic_name, l.location_name;
    """
    raw_rows = run_query(sql)
    result = []
    for row in raw_rows:
        current_stock = safe_float(row["current_stock"])
        issued_30d = safe_float(row["issued_30d"])
        issued_90d = safe_float(row["issued_90d"])
        minimum_stock = safe_float(row["minimum_stock"])
        ven_class = str(row["ven_class"] or "").strip().upper()

        if issued_90d > 0:
            avg_monthly = round((issued_90d / 90) * 30, 2)
            basis = "90 days"
        elif issued_30d > 0:
            avg_monthly = round(issued_30d, 2)
            basis = "30 days"
        else:
            avg_monthly = 0
            basis = "No issue history"

        daily = avg_monthly / 30 if avg_monthly > 0 else 0
        days_left = round(current_stock / daily, 1) if daily > 0 else None

        is_vital = ven_class in ("V", "VITAL", "حيوي", "حيوية")
        if current_stock <= 0 and is_vital:
            risk = "Critical"
            reason = "Vital item out of stock"
        elif days_left is not None and days_left <= LEAD_TIME_DAYS and is_vital:
            risk = "Critical"
            reason = "Vital item will run out before lead time"
        elif minimum_stock > 0 and current_stock < minimum_stock:
            risk = "High"
            reason = "Below minimum stock"
        elif days_left is not None and days_left <= LEAD_TIME_DAYS:
            risk = "High"
            reason = "Will run out before lead time"
        elif days_left is not None and days_left <= LEAD_TIME_DAYS + 14:
            risk = "Medium"
            reason = "Needs monitoring"
        else:
            risk = "Low"
            reason = "No urgent action"

        reorder_qty = max(0, round((daily * 60) - current_stock, 2)) if daily > 0 else 0

        result.append({
            "risk": risk,
            "reason": reason,
            "item_code": row["item_code"],
            "generic_name": row["generic_name"],
            "brand_name": row["brand_name"],
            "ven_class": row["ven_class"],
            "location_name": row["location_name"],
            "current_stock": current_stock,
            "minimum_stock": minimum_stock,
            "issued_30d": issued_30d,
            "issued_90d": issued_90d,
            "avg_monthly": avg_monthly,
            "days_left": "" if days_left is None else days_left,
            "reorder_qty": reorder_qty,
            "basis": basis,
        })
    return result


def get_expiry_alert_rows():
    today = date.today()
    limit = today + timedelta(days=NEAR_EXPIRY_DAYS)
    rows = get_current_stock_rows(limit=5000)
    result = []
    for row in rows:
        expiry_text = str(row["expiry_date"] or "")[:10]
        if not expiry_text:
            continue
        try:
            expiry = datetime.strptime(expiry_text, "%Y-%m-%d").date()
        except Exception:
            continue
        stock = safe_float(row["current_stock"])
        if stock <= 0:
            continue
        if expiry < today:
            risk = "Expired"
        elif expiry <= limit:
            risk = "Near expiry"
        else:
            continue
        days_to_expiry = (expiry - today).days
        result.append([
            risk,
            row["generic_name"],
            row["location_name"],
            row["batch_number"],
            row["expiry_date"],
            days_to_expiry,
            stock,
            row["stock_value"],
        ])
    result.sort(key=lambda x: x[5])
    return result


def get_po_overdue_rows():
    sql = """
    SELECT
        COALESCE(po.po_number, po.po_id) AS po_number,
        COALESCE(s.supplier_name, '') AS supplier_name,
        po.status,
        po.created_at,
        CAST(julianday('now') - julianday(po.created_at) AS INTEGER) AS age_days,
        COALESCE(po.notes, '') AS notes
    FROM purchase_orders po
    LEFT JOIN suppliers s ON po.supplier_id = s.supplier_id
    WHERE po.status NOT IN ('Closed', 'Cancelled')
      AND DATE(po.created_at) <= DATE('now', ?)
    ORDER BY age_days DESC;
    """
    return run_query(sql, (f"-{OVERDUE_PO_DAYS} day",))


def get_month_comparison():
    current_qty = scalar("""
        SELECT COALESCE(SUM(quantity), 0)
        FROM stock_movements
        WHERE movement_type = 'Issue'
          AND strftime('%Y-%m', movement_datetime) = strftime('%Y-%m', 'now');
    """)
    previous_qty = scalar("""
        SELECT COALESCE(SUM(quantity), 0)
        FROM stock_movements
        WHERE movement_type = 'Issue'
          AND strftime('%Y-%m', movement_datetime) = strftime('%Y-%m', date('now', 'start of month', '-1 month'));
    """)
    if previous_qty:
        change = round(((current_qty - previous_qty) / previous_qty) * 100, 1)
    else:
        change = "N/A"
    return current_qty, previous_qty, change


def get_adjustment_alert_rows():
    sql = """
    SELECT
        sm.movement_datetime,
        i.generic_name,
        COALESCE(lf.location_name, lt.location_name, '') AS location_name,
        sm.quantity,
        COALESCE(sm.reason, '') AS reason,
        COALESCE(sm.reference_no, '') AS reference_no
    FROM stock_movements sm
    JOIN items i ON sm.item_id = i.item_id
    LEFT JOIN locations lf ON sm.from_location_id = lf.location_id
    LEFT JOIN locations lt ON sm.to_location_id = lt.location_id
    WHERE sm.movement_type = 'Adjustment'
      AND DATE(sm.movement_datetime) >= DATE('now', ?)
    ORDER BY ABS(sm.quantity) DESC
    LIMIT 100;
    """
    return run_query(sql, (f"-{ADJUSTMENT_ALERT_DAYS} day",))


def get_supplier_delay_rows():
    sql = """
    SELECT
        COALESCE(s.supplier_name, 'Unknown Supplier') AS supplier_name,
        COUNT(*) AS delayed_po_count,
        MAX(CAST(julianday('now') - julianday(po.created_at) AS INTEGER)) AS max_delay_days
    FROM purchase_orders po
    LEFT JOIN suppliers s ON po.supplier_id = s.supplier_id
    WHERE po.status NOT IN ('Closed', 'Cancelled')
      AND DATE(po.created_at) <= DATE('now', ?)
    GROUP BY COALESCE(s.supplier_name, 'Unknown Supplier')
    ORDER BY delayed_po_count DESC, max_delay_days DESC;
    """
    return run_query(sql, (f"-{OVERDUE_PO_DAYS} day",))


def get_seasonal_forecast_rows():
    sql = """
    SELECT
        i.generic_name,
        COALESCE(i.ven_class, '') AS ven_class,
        strftime('%m', sm.movement_datetime) AS month_no,
        SUM(sm.quantity) AS issued_qty
    FROM stock_movements sm
    JOIN items i ON sm.item_id = i.item_id
    WHERE sm.movement_type = 'Issue'
    GROUP BY i.item_id, strftime('%m', sm.movement_datetime)
    ORDER BY i.generic_name, month_no;
    """
    rows = run_query(sql)
    grouped = {}
    for row in rows:
        name = row["generic_name"]
        grouped.setdefault(name, {"ven_class": row["ven_class"], "months": {}})
        grouped[name]["months"][row["month_no"]] = safe_float(row["issued_qty"])

    result = []
    for name, data in grouped.items():
        values = list(data["months"].values())
        if not values:
            continue
        avg = sum(values) / len(values)
        peak_month = max(data["months"], key=lambda m: data["months"][m])
        peak_qty = data["months"][peak_month]
        seasonal_factor = round(peak_qty / avg, 2) if avg else "N/A"
        result.append([
            name,
            data["ven_class"],
            peak_month,
            round(peak_qty, 2),
            round(avg, 2),
            seasonal_factor,
            "Use after 12 months data" if len(values) < 12 else "Usable",
        ])
    result.sort(key=lambda x: x[0])
    return result


class ExecutiveStorekeeperDashboard(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Executive + Storekeeper Dashboard - PharmaGuard")
        self.resize(1450, 880)

        main_layout = QVBoxLayout(self)

        title = QLabel("Executive Dashboard + Storekeeper Smart Alerts / لوحة الإدارة والصيدلي")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size:22px;font-weight:bold;color:#0f172a;")
        main_layout.addWidget(title)

        self.data_label = QLabel("")
        self.data_label.setAlignment(Qt.AlignCenter)
        self.data_label.setStyleSheet("color:#475569;font-weight:bold;")
        main_layout.addWidget(self.data_label)

        top_buttons = QHBoxLayout()
        self.refresh_btn = QPushButton("Refresh / تحديث")
        self.refresh_btn.clicked.connect(self.load_all)
        top_buttons.addWidget(self.refresh_btn)

        self.open_monthly_btn = QPushButton("Open Monthly Reports / التقارير الشهرية")
        self.open_monthly_btn.clicked.connect(lambda: self.open_script("run_monthly_consumption_reports.py"))
        top_buttons.addWidget(self.open_monthly_btn)

        self.open_n8n_btn = QPushButton("Open n8n Email Reports / إرسال n8n")
        self.open_n8n_btn.clicked.connect(lambda: self.open_script("run_n8n_email_reports_center.py"))
        top_buttons.addWidget(self.open_n8n_btn)

        self.open_stock_btn = QPushButton("Open Stock Viewer / عرض الرصيد")
        self.open_stock_btn.clicked.connect(lambda: self.open_script("run_stock_viewer.py"))
        top_buttons.addWidget(self.open_stock_btn)

        self.open_movement_btn = QPushButton("Open Movement Entry / تسجيل حركة")
        self.open_movement_btn.clicked.connect(lambda: self.open_script("run_movement_entry.py"))
        top_buttons.addWidget(self.open_movement_btn)

        self.open_fefo_btn = QPushButton("Open Batch / FEFO / الصلاحية")
        self.open_fefo_btn.clicked.connect(lambda: self.open_script("run_fefo_helper.py"))
        top_buttons.addWidget(self.open_fefo_btn)

        main_layout.addLayout(top_buttons)

        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        self.executive_tab = QWidget()
        self.storekeeper_tab = QWidget()
        self.alerts_tab = QWidget()
        self.forecast_tab = QWidget()

        self.tabs.addTab(self.executive_tab, "Executive Dashboard")
        self.tabs.addTab(self.storekeeper_tab, "Storekeeper / Monthly Stock")
        self.tabs.addTab(self.alerts_tab, "Smart Alerts")
        self.tabs.addTab(self.forecast_tab, "Forecast / Month Comparison")

        self.build_executive_tab()
        self.build_storekeeper_tab()
        self.build_alerts_tab()
        self.build_forecast_tab()

        self.load_all()

    def card(self, title, value, note=""):
        label = QLabel(
            f"<div style='font-size:13px;color:#475569'>{title}</div>"
            f"<div style='font-size:30px;font-weight:bold;color:#0f172a'>{value}</div>"
            f"<div style='font-size:11px;color:#64748b'>{note}</div>"
        )
        label.setAlignment(Qt.AlignCenter)
        label.setMinimumHeight(100)
        label.setStyleSheet("""
            QLabel {
                background:#f8fafc;
                border:1px solid #cbd5e1;
                border-radius:12px;
                padding:10px;
            }
        """)
        return label

    def table(self):
        tbl = QTableWidget()
        tbl.setSortingEnabled(True)
        return tbl

    def build_executive_tab(self):
        layout = QVBoxLayout(self.executive_tab)
        self.cards_grid = QGridLayout()
        layout.addLayout(self.cards_grid)
        layout.addWidget(QLabel("Critical / Vital / Priority items / الأصناف الحرجة أو ذات الأولوية"))
        self.vital_missing_table = self.table()
        layout.addWidget(self.vital_missing_table)
        layout.addWidget(QLabel("Branch summary or overdue purchase orders / ملخص الفروع أو أوامر الشراء المتأخرة"))
        self.overdue_po_table = self.table()
        layout.addWidget(self.overdue_po_table)
        layout.addWidget(QLabel("New medicines or delayed suppliers / الأدوية الجديدة أو الموردون المتأخرون"))
        self.delayed_suppliers_table = self.table()
        layout.addWidget(self.delayed_suppliers_table)

    def build_storekeeper_tab(self):
        layout = QVBoxLayout(self.storekeeper_tab)
        note = QLabel(
            "لو يوجد monthly_outputs، تعرض هذه الشاشة رصيد افتتاحي الشهر التالي من ملف التقارير الشهرية. "
            "ولو لا يوجد، ترجع لبيانات Operational DB."
        )
        note.setWordWrap(True)
        layout.addWidget(note)
        layout.addWidget(QLabel("Current / Opening stock / كشف الرصيد أو الرصيد الافتتاحي للشهر التالي"))
        self.stock_table = self.table()
        layout.addWidget(self.stock_table)
        layout.addWidget(QLabel("Branch stock or daily movements / رصيد الفروع أو حركات اليوم"))
        self.today_movements_table = self.table()
        layout.addWidget(self.today_movements_table)

    def build_alerts_tab(self):
        layout = QVBoxLayout(self.alerts_tab)
        layout.addWidget(QLabel("Critical + High alerts only / تنبيهات مهمة فقط"))
        self.smart_alerts_table = self.table()
        layout.addWidget(self.smart_alerts_table)
        layout.addWidget(QLabel("Purchase suggestions / طلبات شراء مقترحة"))
        self.expiry_table = self.table()
        layout.addWidget(self.expiry_table)
        layout.addWidget(QLabel("Stock count gaps or notes / فروق جرد أو ملاحظات"))
        self.adjustment_table = self.table()
        layout.addWidget(self.adjustment_table)

    def build_forecast_tab(self):
        layout = QVBoxLayout(self.forecast_tab)
        note = QLabel(
            "لو تقارير شهرية موجودة، ستعرض الشاشة ملخص الأصناف والتوقعات من الملف الشهري. "
            "التنبؤ الموسمي الحقيقي يحتاج عدة أشهر، ويفضل 12 شهر."
        )
        note.setWordWrap(True)
        note.setStyleSheet("color:#92400e;font-weight:bold;")
        layout.addWidget(note)
        self.seasonal_table = self.table()
        layout.addWidget(self.seasonal_table)

    def clear_cards(self):
        while self.cards_grid.count():
            item = self.cards_grid.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    def set_table(self, table, columns, rows):
        table.setSortingEnabled(False)
        table.clear()
        table.setColumnCount(len(columns))
        table.setRowCount(len(rows))
        table.setHorizontalHeaderLabels(columns)
        for r, row in enumerate(rows):
            if isinstance(row, sqlite3.Row):
                values = [row[col] for col in columns if col in row.keys()]
                if len(values) < len(columns):
                    values = [row[col] if col in row.keys() else "" for col in columns]
            elif isinstance(row, dict):
                values = [row.get(col, "") for col in columns]
            else:
                values = list(row)
            for c, value in enumerate(values[:len(columns)]):
                table.setItem(r, c, QTableWidgetItem(str(value if value is not None else "")))
        table.resizeColumnsToContents()
        table.setSortingEnabled(True)

    def monthly_mode(self):
        return monthly_available()

    def load_all(self):
        try:
            if self.monthly_mode():
                self.data_label.setText(f"Data source: Monthly report → {latest_monthly_report().name}")
                self.load_monthly_executive()
                self.load_monthly_storekeeper()
                self.load_monthly_alerts()
                self.load_monthly_forecast()
            else:
                self.data_label.setText(f"Data source: Operational DB → {get_db_path()} | No monthly_outputs/monthly_reports_*.xlsx found")
                self.load_operational_executive()
                self.load_operational_storekeeper()
                self.load_operational_alerts()
                self.load_operational_forecast()
        except Exception as exc:
            QMessageBox.critical(self, "Dashboard Error", str(exc))

    def load_monthly_executive(self):
        self.clear_cards()
        md = metric_dict()
        item_summary = read_monthly_sheet("02_Item_Summary")
        critical_df = read_monthly_sheet("03_Critical_Items")
        branch_df = read_monthly_sheet("04_Branch_Summary")
        new_df = read_monthly_sheet("05_New_Medicines")
        comparison_df = read_monthly_sheet("08_Month_Comparison")

        if item_summary is None:
            item_summary = pd.DataFrame() if pd is not None else None
        if critical_df is None:
            critical_df = pd.DataFrame() if pd is not None else None

        critical_count = md.get("critical_items", 0)
        high_count = md.get("high_risk_items", 0)
        items_count = md.get("items_count", len(item_summary) if item_summary is not None else 0)
        total_stock = md.get("total_current_stock", "")
        if total_stock == "" and item_summary is not None and "current_stock" in item_summary.columns:
            total_stock = round(item_summary["current_stock"].apply(safe_float).sum(), 2)
        total_issued = md.get("total_issued_month", "")
        if total_issued == "" and item_summary is not None and "total_issued_month" in item_summary.columns:
            total_issued = round(item_summary["total_issued_month"].apply(safe_float).sum(), 2)

        reorder_count = 0
        if item_summary is not None and "suggested_reorder_qty" in item_summary.columns:
            reorder_count = int((item_summary["suggested_reorder_qty"].apply(safe_float) > 0).sum())
        new_count = len(new_df) if new_df is not None else 0
        month = md.get("month", "Unknown")
        next_month = md.get("next_month", "Unknown")

        cards = [
            ("Month / الشهر", month, f"Next: {next_month}"),
            ("Items / عدد الأصناف", items_count, "من الملف الشهري"),
            ("Critical items / أصناف حرجة", critical_count, "حسب الرصيد وأيام التغطية"),
            ("High risk / خطر عالي", high_count, "تحتاج متابعة"),
            ("Stock quantity / إجمالي الرصيد", total_stock, "كمية وليس قيمة مالية"),
            ("Issued this month / منصرف الشهر", total_issued, "من ملف المنصرفات"),
            ("Suggested purchases / شراء مقترح", reorder_count, "suggested_reorder_qty > 0"),
            ("New medicines / أدوية جديدة", new_count, "مقارنة بالماستر"),
        ]
        for i, (title, value, note) in enumerate(cards):
            self.cards_grid.addWidget(self.card(title, value, note), i // 4, i % 4)

        crit_cols = [c for c in ["shortage_risk", "scientific_name", "trade_names", "unit", "current_stock", "avg_monthly_consumption", "days_left", "dynamic_min", "suggested_reorder_qty", "source_sheet"] if critical_df is not None and c in critical_df.columns]
        if not crit_cols and critical_df is not None:
            crit_cols = list(critical_df.columns[:10])
        self.set_table(self.vital_missing_table, crit_cols, df_rows(critical_df[crit_cols] if critical_df is not None and crit_cols else critical_df, 100))

        branch_cols = ["branch", "total_issued", "total_stock", "item_count"] if branch_df is not None and "branch" in branch_df.columns else (list(branch_df.columns) if branch_df is not None else [])
        self.set_table(self.overdue_po_table, branch_cols, df_rows(branch_df[branch_cols] if branch_df is not None and branch_cols else branch_df, 100))

        new_cols = list(new_df.columns[:8]) if new_df is not None else []
        self.set_table(self.delayed_suppliers_table, new_cols, df_rows(new_df[new_cols] if new_df is not None and new_cols else new_df, 100))

    def load_monthly_storekeeper(self):
        opening_df = read_monthly_sheet("06_Next_Opening_Total")
        by_branch_df = read_monthly_sheet("07_Next_Opening_By_Branch")
        if opening_df is None:
            opening_df = pd.DataFrame() if pd is not None else None
        if by_branch_df is None:
            by_branch_df = pd.DataFrame() if pd is not None else None

        cols1 = [c for c in ["next_month", "source_sheet", "scientific_name", "trade_names", "unit", "opening_stock_next_month"] if opening_df is not None and c in opening_df.columns]
        if not cols1 and opening_df is not None:
            cols1 = list(opening_df.columns[:10])
        self.set_table(self.stock_table, cols1, df_rows(opening_df[cols1] if opening_df is not None and cols1 else opening_df, 500))

        cols2 = [c for c in ["next_month", "source_sheet", "scientific_name", "trade_name", "unit", "branch", "opening_stock"] if by_branch_df is not None and c in by_branch_df.columns]
        if not cols2 and by_branch_df is not None:
            cols2 = list(by_branch_df.columns[:10])
        self.set_table(self.today_movements_table, cols2, df_rows(by_branch_df[cols2] if by_branch_df is not None and cols2 else by_branch_df, 500))

    def load_monthly_alerts(self):
        critical_df = read_monthly_sheet("03_Critical_Items")
        item_summary = read_monthly_sheet("02_Item_Summary")
        comparison_df = read_monthly_sheet("08_Month_Comparison")

        if critical_df is None:
            critical_df = pd.DataFrame() if pd is not None else None
        if item_summary is None:
            item_summary = pd.DataFrame() if pd is not None else None

        alert_cols = [c for c in ["shortage_risk", "scientific_name", "trade_names", "unit", "current_stock", "days_left", "dynamic_min", "suggested_reorder_qty", "note"] if critical_df is not None and c in critical_df.columns]
        if not alert_cols and critical_df is not None:
            alert_cols = list(critical_df.columns[:10])
        self.set_table(self.smart_alerts_table, alert_cols, df_rows(critical_df[alert_cols] if critical_df is not None and alert_cols else critical_df, 200))

        purchase_df = item_summary
        if purchase_df is not None and "suggested_reorder_qty" in purchase_df.columns:
            purchase_df = purchase_df[purchase_df["suggested_reorder_qty"].apply(safe_float) > 0].copy()
        purchase_cols = [c for c in ["scientific_name", "trade_names", "unit", "current_stock", "avg_monthly_consumption", "days_left", "suggested_reorder_qty", "shortage_risk"] if purchase_df is not None and c in purchase_df.columns]
        if not purchase_cols and purchase_df is not None:
            purchase_cols = list(purchase_df.columns[:10])
        self.set_table(self.expiry_table, purchase_cols, df_rows(purchase_df[purchase_cols] if purchase_df is not None and purchase_cols else purchase_df, 200))

        comp_cols = list(comparison_df.columns) if comparison_df is not None else []
        self.set_table(self.adjustment_table, comp_cols, df_rows(comparison_df, 50))

    def load_monthly_forecast(self):
        item_summary = read_monthly_sheet("02_Item_Summary")
        if item_summary is None:
            item_summary = pd.DataFrame() if pd is not None else None
        cols = [c for c in ["scientific_name", "trade_names", "unit", "avg_monthly_year", "avg_monthly_recent", "avg_monthly_consumption", "current_stock", "days_left", "dynamic_min", "suggested_reorder_qty", "shortage_risk"] if item_summary is not None and c in item_summary.columns]
        if not cols and item_summary is not None:
            cols = list(item_summary.columns[:12])
        self.set_table(self.seasonal_table, cols, df_rows(item_summary[cols] if item_summary is not None and cols else item_summary, 500))

    def load_operational_executive(self):
        self.clear_cards()
        forecast = get_forecast_rows()
        critical = [r for r in forecast if r["risk"] == "Critical"]
        vital_missing = [r for r in forecast if str(r["ven_class"]).strip().upper() in ("V", "VITAL") and safe_float(r["current_stock"]) <= 0]
        stock_value = scalar(movement_stock_cte() + """
            SELECT ROUND(SUM(sn.current_stock * COALESCE(i.unit_cost, 0)), 2)
            FROM stock_now sn JOIN items i ON sn.item_id = i.item_id;
        """)
        expiry_rows = get_expiry_alert_rows()
        expected_waste = sum(safe_float(r[7]) for r in expiry_rows)
        overdue_pos = get_po_overdue_rows()
        supplier_delays = get_supplier_delay_rows()
        current_qty, previous_qty, change = get_month_comparison()

        cards = [
            ("Critical items / أصناف حرجة", len(critical), "Critical فقط"),
            ("Inventory value / قيمة المخزون", round(safe_float(stock_value), 2), "حسب unit_cost"),
            ("Expected waste / هالك متوقع", round(expected_waste, 2), f"Expiry <= {NEAR_EXPIRY_DAYS} days"),
            ("Overdue POs / طلبيات متأخرة", len(overdue_pos), f"> {OVERDUE_PO_DAYS} days"),
            ("Vital missing / حيوي ناقص", len(vital_missing), "VEN = V and stock <= 0"),
            ("Delayed suppliers / موردون متأخرون", len(supplier_delays), "حسب PO مفتوح"),
            ("Current month issue / صرف هذا الشهر", round(safe_float(current_qty), 2), "Issue movements"),
            ("Previous month issue / صرف الشهر السابق", round(safe_float(previous_qty), 2), f"Change: {change}%"),
        ]
        for i, (title, value, note) in enumerate(cards):
            self.cards_grid.addWidget(self.card(title, value, note), i // 4, i % 4)

        self.set_table(self.vital_missing_table, ["risk", "reason", "item_code", "generic_name", "ven_class", "location_name", "current_stock", "days_left", "reorder_qty"], vital_missing[:100])
        self.set_table(self.overdue_po_table, ["po_number", "supplier_name", "status", "created_at", "age_days", "notes"], overdue_pos)
        self.set_table(self.delayed_suppliers_table, ["supplier_name", "delayed_po_count", "max_delay_days"], supplier_delays)

    def load_operational_storekeeper(self):
        self.set_table(self.stock_table, ["item_code", "generic_name", "brand_name", "dosage_form", "strength", "ven_class", "location_name", "batch_number", "expiry_date", "current_stock", "minimum_stock", "unit_cost", "stock_value"], get_current_stock_rows())
        today_rows = run_query("""
            SELECT
                sm.movement_datetime,
                sm.movement_type,
                i.generic_name,
                sm.quantity,
                COALESCE(lf.location_name, '') AS from_location,
                COALESCE(lt.location_name, '') AS to_location,
                COALESCE(sm.reason, '') AS reason,
                COALESCE(sm.reference_no, '') AS reference_no
            FROM stock_movements sm
            JOIN items i ON sm.item_id = i.item_id
            LEFT JOIN locations lf ON sm.from_location_id = lf.location_id
            LEFT JOIN locations lt ON sm.to_location_id = lt.location_id
            WHERE DATE(sm.movement_datetime) = DATE('now')
            ORDER BY sm.movement_id DESC;
        """)
        self.set_table(self.today_movements_table, ["movement_datetime", "movement_type", "generic_name", "quantity", "from_location", "to_location", "reason", "reference_no"], today_rows)

    def load_operational_alerts(self):
        forecast = get_forecast_rows()
        alerts = [r for r in forecast if r["risk"] in ("Critical", "High")]
        self.set_table(self.smart_alerts_table, ["risk", "reason", "item_code", "generic_name", "ven_class", "location_name", "current_stock", "minimum_stock", "issued_30d", "days_left", "reorder_qty"], alerts[:200])
        self.set_table(self.expiry_table, ["risk", "generic_name", "location_name", "batch_number", "expiry_date", "days_to_expiry", "current_stock", "stock_value"], get_expiry_alert_rows())
        self.set_table(self.adjustment_table, ["movement_datetime", "generic_name", "location_name", "quantity", "reason", "reference_no"], get_adjustment_alert_rows())

    def load_operational_forecast(self):
        self.set_table(self.seasonal_table, ["generic_name", "ven_class", "peak_month", "peak_qty", "average_monthly_qty", "seasonal_factor", "confidence_note"], get_seasonal_forecast_rows())

    def open_script(self, script_name):
        script_path = PROJECT_ROOT / script_name
        if not script_path.exists():
            QMessageBox.warning(self, "File not found", f"لم يتم العثور على:\n{script_path}")
            return
        try:
            subprocess.Popen([sys.executable, str(script_path)], cwd=str(PROJECT_ROOT))
        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))


def main():
    app = QApplication(sys.argv)
    window = ExecutiveStorekeeperDashboard()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
