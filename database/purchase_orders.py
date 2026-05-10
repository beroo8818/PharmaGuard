from datetime import datetime
from database.schema import connect, init_database, seed_basic_data


PO_STATUSES = [
    "Draft",
    "Reviewed",
    "Approved",
    "Sent",
    "Partially Received",
    "Closed",
    "Cancelled",
]


ALLOWED_TRANSITIONS = {
    "Draft": ["Reviewed", "Cancelled"],
    "Reviewed": ["Approved", "Draft", "Cancelled"],
    "Approved": ["Sent", "Cancelled"],
    "Sent": ["Partially Received", "Closed", "Cancelled"],
    "Partially Received": ["Closed", "Sent", "Cancelled"],
    "Closed": [],
    "Cancelled": [],
}


def setup_purchase_orders():
    init_database()
    seed_basic_data()


def generate_po_number():
    return "PO-" + datetime.now().strftime("%Y%m%d-%H%M%S")


def create_purchase_order(supplier_id=None, requested_by=None, notes=""):
    setup_purchase_orders()

    po_number = generate_po_number()

    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO purchase_orders (
            po_number,
            supplier_id,
            status,
            requested_by,
            notes
        )
        VALUES (?, ?, 'Draft', ?, ?);
    """, (po_number, supplier_id, requested_by, notes))

    conn.commit()
    po_id = cur.lastrowid
    conn.close()

    return po_id, po_number


def add_purchase_order_line(po_id, item_id, requested_qty, unit_cost=0, notes=""):
    if requested_qty <= 0:
        raise ValueError("Requested quantity must be greater than zero.")

    conn = connect()
    cur = conn.cursor()

    cur.execute("SELECT status FROM purchase_orders WHERE po_id = ?;", (po_id,))
    row = cur.fetchone()

    if row is None:
        conn.close()
        raise ValueError("Purchase order not found.")

    status = row[0]

    if status not in ["Draft", "Reviewed"]:
        conn.close()
        raise ValueError("You can add lines only while PO is Draft or Reviewed.")

    cur.execute("""
        INSERT INTO purchase_order_lines (
            po_id,
            item_id,
            requested_qty,
            approved_qty,
            received_qty,
            unit_cost,
            notes
        )
        VALUES (?, ?, ?, 0, 0, ?, ?);
    """, (po_id, item_id, requested_qty, unit_cost, notes))

    conn.commit()
    line_id = cur.lastrowid
    conn.close()

    return line_id


def approve_all_requested_quantities(po_id):
    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        UPDATE purchase_order_lines
        SET approved_qty = requested_qty
        WHERE po_id = ?;
    """, (po_id,))

    conn.commit()
    conn.close()


def update_purchase_order_status(po_id, new_status, user_id=None):
    if new_status not in PO_STATUSES:
        raise ValueError("Invalid status.")

    conn = connect()
    cur = conn.cursor()

    cur.execute("SELECT status FROM purchase_orders WHERE po_id = ?;", (po_id,))
    row = cur.fetchone()

    if row is None:
        conn.close()
        raise ValueError("Purchase order not found.")

    old_status = row[0]

    if new_status == old_status:
        conn.close()
        return old_status

    allowed = ALLOWED_TRANSITIONS.get(old_status, [])

    if new_status not in allowed:
        conn.close()
        raise ValueError(f"Invalid status transition: {old_status} -> {new_status}")

    if new_status == "Approved":
        cur.execute("""
            UPDATE purchase_orders
            SET status = ?, approved_by = ?, approved_at = CURRENT_TIMESTAMP
            WHERE po_id = ?;
        """, (new_status, user_id, po_id))
    else:
        cur.execute("""
            UPDATE purchase_orders
            SET status = ?
            WHERE po_id = ?;
        """, (new_status, po_id))

    conn.commit()
    conn.close()

    return new_status


def list_purchase_orders():
    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        SELECT
            po.po_id,
            po.po_number,
            COALESCE(s.supplier_name, '') AS supplier_name,
            po.status,
            COALESCE(ru.username, '') AS requested_by,
            COALESCE(au.username, '') AS approved_by,
            po.created_at,
            COALESCE(po.approved_at, '') AS approved_at,
            COALESCE(po.notes, '') AS notes
        FROM purchase_orders po
        LEFT JOIN suppliers s ON po.supplier_id = s.supplier_id
        LEFT JOIN users ru ON po.requested_by = ru.user_id
        LEFT JOIN users au ON po.approved_by = au.user_id
        ORDER BY po.po_id DESC;
    """)

    rows = cur.fetchall()
    conn.close()
    return rows


def list_purchase_order_lines(po_id):
    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        SELECT
            pol.po_line_id,
            i.item_code,
            i.generic_name,
            i.brand_name,
            pol.requested_qty,
            pol.approved_qty,
            pol.received_qty,
            pol.unit_cost,
            ROUND(pol.requested_qty * pol.unit_cost, 2) AS value,
            COALESCE(pol.notes, '') AS notes
        FROM purchase_order_lines pol
        JOIN items i ON pol.item_id = i.item_id
        WHERE pol.po_id = ?
        ORDER BY pol.po_line_id;
    """, (po_id,))

    rows = cur.fetchall()
    conn.close()
    return rows
