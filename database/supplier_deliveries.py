from database.schema import connect, init_database, seed_basic_data
from database.repositories import add_batch, add_stock_movement


def setup_delivery_tables():
    init_database()
    seed_basic_data()

    conn = connect()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS delivery_receipts (
        delivery_id INTEGER PRIMARY KEY AUTOINCREMENT,
        po_id INTEGER NOT NULL,
        po_line_id INTEGER NOT NULL,
        item_id INTEGER NOT NULL,
        supplier_id INTEGER,
        batch_id INTEGER,
        location_id INTEGER NOT NULL,
        received_qty REAL NOT NULL,
        delivery_no TEXT,
        received_by INTEGER,
        received_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        notes TEXT,
        FOREIGN KEY (po_id) REFERENCES purchase_orders(po_id),
        FOREIGN KEY (po_line_id) REFERENCES purchase_order_lines(po_line_id),
        FOREIGN KEY (item_id) REFERENCES items(item_id),
        FOREIGN KEY (supplier_id) REFERENCES suppliers(supplier_id),
        FOREIGN KEY (batch_id) REFERENCES batches(batch_id),
        FOREIGN KEY (location_id) REFERENCES locations(location_id),
        FOREIGN KEY (received_by) REFERENCES users(user_id)
    );
    """)

    conn.commit()
    conn.close()


def list_receivable_po_lines():
    setup_delivery_tables()

    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        SELECT
            pol.po_line_id,
            po.po_id,
            po.po_number,
            COALESCE(s.supplier_name, '') AS supplier_name,
            i.item_id,
            i.item_code,
            i.generic_name,
            i.brand_name,
            pol.approved_qty,
            pol.received_qty,
            ROUND(pol.approved_qty - pol.received_qty, 2) AS remaining_qty,
            po.status
        FROM purchase_order_lines pol
        JOIN purchase_orders po ON pol.po_id = po.po_id
        JOIN items i ON pol.item_id = i.item_id
        LEFT JOIN suppliers s ON po.supplier_id = s.supplier_id
        WHERE po.status IN ('Sent', 'Partially Received')
          AND pol.approved_qty > pol.received_qty
        ORDER BY po.po_id DESC, pol.po_line_id;
    """)

    rows = cur.fetchall()
    conn.close()
    return rows


def get_or_create_batch(item_id, batch_number, expiry_date=None, supplier_id=None, notes=""):
    batch_number = str(batch_number).strip()

    if not batch_number:
        raise ValueError("Batch number is required.")

    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        SELECT batch_id
        FROM batches
        WHERE item_id = ?
          AND batch_number = ?
        LIMIT 1;
    """, (item_id, batch_number))

    row = cur.fetchone()
    conn.close()

    if row:
        return row[0]

    return add_batch(
        item_id=item_id,
        batch_number=batch_number,
        expiry_date=expiry_date if expiry_date else None,
        supplier_id=supplier_id,
        received_date=None,
        notes=notes,
    )


def receive_po_line_to_stock(
    po_line_id,
    received_qty,
    location_id,
    batch_number,
    expiry_date=None,
    delivery_no="",
    received_by=None,
    notes="",
):
    setup_delivery_tables()

    if received_qty <= 0:
        raise ValueError("Received quantity must be greater than zero.")

    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        SELECT
            pol.po_id,
            pol.item_id,
            pol.approved_qty,
            pol.received_qty,
            po.status,
            po.supplier_id,
            po.po_number
        FROM purchase_order_lines pol
        JOIN purchase_orders po ON pol.po_id = po.po_id
        WHERE pol.po_line_id = ?;
    """, (po_line_id,))

    row = cur.fetchone()

    if row is None:
        conn.close()
        raise ValueError("PO line not found.")

    po_id, item_id, approved_qty, old_received_qty, po_status, supplier_id, po_number = row

    if po_status not in ("Sent", "Partially Received"):
        conn.close()
        raise ValueError("PO must be Sent or Partially Received before receiving.")

    approved_qty = float(approved_qty or 0)
    old_received_qty = float(old_received_qty or 0)
    received_qty = float(received_qty)

    new_received_qty = old_received_qty + received_qty

    if new_received_qty > approved_qty:
        conn.close()
        raise ValueError(
            f"Received quantity cannot exceed approved quantity. "
            f"Approved={approved_qty}, Already received={old_received_qty}, Trying={received_qty}"
        )

    conn.close()

    batch_id = get_or_create_batch(
        item_id=item_id,
        batch_number=batch_number,
        expiry_date=expiry_date,
        supplier_id=supplier_id,
        notes=f"Created from PO {po_number}",
    )

    reference_no = delivery_no.strip() if delivery_no.strip() else f"DELIVERY-{po_number}-LINE-{po_line_id}"

    movement_id = add_stock_movement(
        movement_type="Receive",
        item_id=item_id,
        batch_id=batch_id,
        from_location_id=None,
        to_location_id=location_id,
        quantity=received_qty,
        reference_no=reference_no,
        reason=f"Supplier delivery from PO {po_number}. {notes}".strip(),
        user_id=received_by,
    )

    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        UPDATE purchase_order_lines
        SET received_qty = ?
        WHERE po_line_id = ?;
    """, (new_received_qty, po_line_id))

    cur.execute("""
        SELECT
            SUM(approved_qty),
            SUM(received_qty)
        FROM purchase_order_lines
        WHERE po_id = ?;
    """, (po_id,))

    total_approved, total_received = cur.fetchone()
    total_approved = float(total_approved or 0)
    total_received = float(total_received or 0)

    if total_approved > 0 and total_received >= total_approved:
        new_status = "Closed"
    else:
        new_status = "Partially Received"

    cur.execute("""
        UPDATE purchase_orders
        SET status = ?
        WHERE po_id = ?;
    """, (new_status, po_id))

    cur.execute("""
        INSERT INTO delivery_receipts (
            po_id,
            po_line_id,
            item_id,
            supplier_id,
            batch_id,
            location_id,
            received_qty,
            delivery_no,
            received_by,
            notes
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
    """, (
        po_id,
        po_line_id,
        item_id,
        supplier_id,
        batch_id,
        location_id,
        received_qty,
        reference_no,
        received_by,
        notes,
    ))

    conn.commit()
    delivery_id = cur.lastrowid
    conn.close()

    return {
        "delivery_id": delivery_id,
        "movement_id": movement_id,
        "batch_id": batch_id,
        "po_status": new_status,
    }


def list_delivery_receipts(limit=200):
    setup_delivery_tables()

    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        SELECT
            dr.delivery_id,
            dr.received_at,
            po.po_number,
            dr.po_line_id,
            i.item_code,
            i.generic_name,
            b.batch_number,
            b.expiry_date,
            l.location_name,
            dr.received_qty,
            dr.delivery_no,
            COALESCE(u.username, '') AS received_by,
            COALESCE(dr.notes, '') AS notes
        FROM delivery_receipts dr
        JOIN purchase_orders po ON dr.po_id = po.po_id
        JOIN items i ON dr.item_id = i.item_id
        LEFT JOIN batches b ON dr.batch_id = b.batch_id
        JOIN locations l ON dr.location_id = l.location_id
        LEFT JOIN users u ON dr.received_by = u.user_id
        ORDER BY dr.delivery_id DESC
        LIMIT ?;
    """, (limit,))

    rows = cur.fetchall()
    conn.close()
    return rows
