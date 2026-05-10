
from datetime import datetime, date, timedelta
from pathlib import Path
import sqlite3
from openpyxl import Workbook

DB_NAME = "pharmaguard_operational_v12.db"
REPORTS_DIR = Path.cwd() / "reports"


def db_path():
    return Path.cwd() / DB_NAME


def connect():
    return sqlite3.connect(db_path())


def ensure_reports_dir():
    REPORTS_DIR.mkdir(exist_ok=True)
    return REPORTS_DIR


def fetch_rows(query, params=()):
    conn = connect()
    cur = conn.cursor()
    cur.execute(query, params)
    columns = [d[0] for d in cur.description]
    rows = cur.fetchall()
    conn.close()
    return columns, rows


def add_sheet(wb, title, columns, rows):
    ws = wb.create_sheet(title[:31])
    ws.append(columns)

    for row in rows:
        ws.append(list(row))

    for col in ws.columns:
        max_len = 10
        letter = col[0].column_letter
        for cell in col:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(value))
        ws.column_dimensions[letter].width = min(max_len + 2, 45)


def current_stock_query():
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
    )
    SELECT
        i.item_code,
        i.generic_name,
        COALESCE(i.brand_name, '') AS brand_name,
        COALESCE(i.dosage_form, '') AS dosage_form,
        l.location_name,
        COALESCE(b.batch_number, '') AS batch_number,
        COALESCE(b.expiry_date, '') AS expiry_date,
        ROUND(SUM(ml.signed_qty), 2) AS current_stock
    FROM movement_lines ml
    JOIN items i ON ml.item_id = i.item_id
    JOIN locations l ON ml.location_id = l.location_id
    LEFT JOIN batches b ON ml.batch_id = b.batch_id
    GROUP BY
        i.item_code,
        i.generic_name,
        i.brand_name,
        i.dosage_form,
        l.location_name,
        b.batch_number,
        b.expiry_date
    HAVING current_stock <> 0
    ORDER BY i.generic_name, l.location_name, b.expiry_date;
    """


def forecast_query():
    return """
    WITH stock_now AS (
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
        )
        SELECT item_id, location_id, ROUND(SUM(signed_qty), 2) AS current_stock
        FROM movement_lines
        GROUP BY item_id, location_id
    ),
    issue_30 AS (
        SELECT item_id, from_location_id AS location_id, SUM(quantity) AS issued_30d
        FROM stock_movements
        WHERE movement_type = 'Issue'
          AND DATE(movement_datetime) >= DATE('now', '-30 day')
        GROUP BY item_id, from_location_id
    )
    SELECT
        i.item_code,
        i.generic_name,
        l.location_name,
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
    LEFT JOIN issue_30 i30
        ON sn.item_id = i30.item_id
       AND sn.location_id = i30.location_id
    ORDER BY shortage_risk, i.generic_name;
    """


def expiry_alerts_rows():
    columns, rows = fetch_rows(current_stock_query())
    idx = {name: i for i, name in enumerate(columns)}

    today = date.today()
    near_limit = today + timedelta(days=90)

    result = []

    for row in rows:
        expiry_text = str(row[idx["expiry_date"]] or "").strip()
        stock = float(row[idx["current_stock"]] or 0)

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

        result.append([
            risk,
            row[idx["item_code"]],
            row[idx["generic_name"]],
            row[idx["location_name"]],
            row[idx["batch_number"]],
            expiry_text,
            stock,
        ])

    return [
        "risk",
        "item_code",
        "generic_name",
        "location_name",
        "batch_number",
        "expiry_date",
        "current_stock",
    ], result


def export_operational_report(output_path=None):
    ensure_reports_dir()

    if output_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = REPORTS_DIR / f"pharmaguard_operational_report_{timestamp}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"

    stock_cols, stock_rows = fetch_rows(current_stock_query())
    forecast_cols, forecast_rows = fetch_rows(forecast_query())
    expiry_cols, expiry_rows = expiry_alerts_rows()

    ws.append(["Metric", "Value"])
    ws.append(["Generated at", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    ws.append(["Current stock rows", len(stock_rows)])
    ws.append(["Forecast rows", len(forecast_rows)])
    ws.append(["Expiry alerts", len(expiry_rows)])

    add_sheet(wb, "Current Stock", stock_cols, stock_rows)
    add_sheet(wb, "Forecast", forecast_cols, forecast_rows)
    add_sheet(wb, "Expiry Alerts", expiry_cols, expiry_rows)

    po_cols, po_rows = fetch_rows("""
        SELECT po_id, po_number, status, created_at, notes
        FROM purchase_orders
        ORDER BY po_id DESC;
    """)
    add_sheet(wb, "Purchase Orders", po_cols, po_rows)

    mov_cols, mov_rows = fetch_rows("""
        SELECT
            sm.movement_id,
            sm.movement_datetime,
            sm.movement_type,
            i.generic_name,
            sm.quantity,
            COALESCE(sm.reference_no, '') AS reference_no,
            COALESCE(sm.reason, '') AS reason
        FROM stock_movements sm
        JOIN items i ON sm.item_id = i.item_id
        ORDER BY sm.movement_id DESC
        LIMIT 500;
    """)
    add_sheet(wb, "Recent Movements", mov_cols, mov_rows)

    audit_cols, audit_rows = fetch_rows("""
        SELECT audit_id, created_at, action, table_name, record_id, old_value, new_value
        FROM audit_logs
        ORDER BY audit_id DESC
        LIMIT 500;
    """)
    add_sheet(wb, "Audit Logs", audit_cols, audit_rows)

    wb.save(output_path)
    return str(output_path)


if __name__ == "__main__":
    path = export_operational_report()
    print("Report exported:")
    print(path)
