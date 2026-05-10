from datetime import date, datetime, timedelta
from database.schema import connect, init_database, seed_basic_data

VALID_MOVEMENT_TYPES = {
    "Receive",
    "Issue",
    "Transfer",
    "Return",
    "Waste",
    "Adjustment",
}


def setup_database():
    init_database()
    seed_basic_data()


def add_location(location_name, location_type="pharmacy"):
    conn = connect()
    cur = conn.cursor()

    cur.execute("""
    INSERT OR IGNORE INTO locations (location_name, location_type)
    VALUES (?, ?);
    """, (location_name, location_type))

    conn.commit()

    cur.execute("SELECT location_id FROM locations WHERE location_name = ?;", (location_name,))
    location_id = cur.fetchone()[0]

    conn.close()
    return location_id


def add_supplier(supplier_name="Unknown Supplier"):
    conn = connect()
    cur = conn.cursor()

    cur.execute("""
    INSERT OR IGNORE INTO suppliers (supplier_name)
    VALUES (?);
    """, (supplier_name,))

    conn.commit()

    cur.execute("SELECT supplier_id FROM suppliers WHERE supplier_name = ?;", (supplier_name,))
    supplier_id = cur.fetchone()[0]

    conn.close()
    return supplier_id


def add_item(
    item_code,
    generic_name,
    brand_name="",
    dosage_form="",
    strength="",
    ven_class="E",
    unit_cost=0,
):
    conn = connect()
    cur = conn.cursor()

    cur.execute("""
    INSERT OR IGNORE INTO items (
        item_code, generic_name, brand_name, dosage_form, strength, ven_class, unit_cost
    )
    VALUES (?, ?, ?, ?, ?, ?, ?);
    """, (
        item_code,
        generic_name,
        brand_name,
        dosage_form,
        strength,
        ven_class,
        unit_cost,
    ))

    conn.commit()

    cur.execute("SELECT item_id FROM items WHERE item_code = ?;", (item_code,))
    item_id = cur.fetchone()[0]

    conn.close()
    return item_id


def add_batch(
    item_id,
    batch_number,
    expiry_date=None,
    supplier_id=None,
    received_date=None,
    notes="",
):
    if not batch_number or not str(batch_number).strip():
        raise ValueError("Batch number is required.")

    if not expiry_date or not str(expiry_date).strip():
        raise ValueError("Expiry date is required.")

    expiry_date = str(expiry_date).strip()

    try:
        expiry_obj = datetime.strptime(expiry_date, "%Y-%m-%d").date()
    except ValueError:
        raise ValueError("Expiry date must be in YYYY-MM-DD format. Example: 2026-12-31")

    if expiry_obj < date.today():
        raise ValueError("Expiry date cannot be in the past.")

    if received_date:
        received_date = str(received_date).strip()
        try:
            datetime.strptime(received_date, "%Y-%m-%d")
        except ValueError:
            raise ValueError("Received date must be in YYYY-MM-DD format. Example: 2026-05-08")

    conn = connect()
    cur = conn.cursor()

    try:
        cur.execute("""
        INSERT INTO batches (
            item_id, batch_number, expiry_date, supplier_id, received_date, notes
        )
        VALUES (?, ?, ?, ?, ?, ?);
        """, (
            item_id,
            batch_number.strip(),
            expiry_date,
            supplier_id,
            received_date,
            notes,
        ))

        conn.commit()
        batch_id = cur.lastrowid
        return batch_id

    except Exception:
        conn.rollback()
        raise

    finally:
        conn.close()
def get_user_id(username="admin"):
    conn = connect()
    cur = conn.cursor()

    cur.execute("SELECT user_id FROM users WHERE username = ?;", (username,))
    row = cur.fetchone()

    conn.close()

    if row:
        return row[0]

    return None


def get_current_stock(item_id, location_id, batch_id=None):
    conn = connect()
    cur = conn.cursor()

    cur.execute("""
    SELECT
        COALESCE(SUM(
            CASE
                WHEN to_location_id = ? THEN quantity
                WHEN from_location_id = ? THEN -quantity
                ELSE 0
            END
        ), 0)
    FROM stock_movements
    WHERE item_id = ?
      AND (? IS NULL OR batch_id = ?);
    """, (
        location_id,
        location_id,
        item_id,
        batch_id,
        batch_id,
    ))

    row = cur.fetchone()
    conn.close()

    return row[0] if row else 0


def get_batch_expiry_date(batch_id):
    conn = connect()
    cur = conn.cursor()

    cur.execute("""
    SELECT expiry_date
    FROM batches
    WHERE batch_id = ?;
    """, (batch_id,))

    row = cur.fetchone()
    conn.close()

    return row[0] if row else None

def item_has_batches(item_id):
    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        SELECT COUNT(*)
        FROM batches
        WHERE item_id = ?
          AND expiry_date IS NOT NULL
          AND expiry_date != ''
          AND expiry_date >= date('now');
    """, (item_id,))

    count = cur.fetchone()[0]
    conn.close()

    return count > 0


def add_stock_movement(
    movement_type,
    item_id,
    quantity,
    from_location_id=None,
    to_location_id=None,
    batch_id=None,
    reference_no="",
    reason="",
    user_id=None,
):
    if movement_type not in VALID_MOVEMENT_TYPES:
        raise ValueError(f"Invalid movement type: {movement_type}")

    if quantity is None or quantity <= 0:
        raise ValueError("Quantity must be greater than zero.")

    if movement_type == "Receive" and to_location_id is None:
        raise ValueError("Receive movement needs to_location_id.")

    if movement_type in {"Issue", "Waste"} and from_location_id is None:
        raise ValueError(f"{movement_type} movement needs from_location_id.")

    if movement_type == "Transfer" and (from_location_id is None or to_location_id is None):
        raise ValueError("Transfer movement needs both from_location_id and to_location_id.")

    if movement_type in {"Waste", "Adjustment"} and not str(reason).strip():
        raise ValueError("Reason is required for Waste/Adjustment.")

    if movement_type in {"Issue", "Waste", "Transfer"}:
        available_stock = get_current_stock(
            item_id=item_id,
            location_id=from_location_id,
            batch_id=batch_id,
        )

        if available_stock < quantity:
            raise ValueError(
                f"Insufficient stock. Available: {available_stock}, requested: {quantity}"
            )

    conn = connect()
    cur = conn.cursor()

    try:
        cur.execute("""
        INSERT INTO stock_movements (
            movement_type,
            item_id,
            batch_id,
            from_location_id,
            to_location_id,
            quantity,
            reference_no,
            reason,
            user_id
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);
        """, (
            movement_type,
            item_id,
            batch_id,
            from_location_id,
            to_location_id,
            quantity,
            reference_no,
            reason,
            user_id,
        ))

        conn.commit()
        movement_id = cur.lastrowid
        return movement_id

    except Exception:
        conn.rollback()
        raise

    finally:
        conn.close()

def list_near_expiry_batches(days=90):
    today = date.today()
    end_date = today + timedelta(days=days)

    conn = connect()
    cur = conn.cursor()

    cur.execute("""
    SELECT
        b.batch_id,
        i.item_code,
        i.generic_name,
        i.brand_name,
        b.batch_number,
        b.expiry_date
    FROM batches b
    JOIN items i ON b.item_id = i.item_id
    WHERE b.expiry_date IS NOT NULL
      AND b.expiry_date >= ?
      AND b.expiry_date <= ?
    ORDER BY b.expiry_date ASC;
    """, (
        today.isoformat(),
        end_date.isoformat(),
    ))

    rows = cur.fetchall()
    conn.close()
    return rows

def list_low_stock_items():
    conn = connect()
    cur = conn.cursor()

    cur.execute("""
    SELECT
        i.item_id,
        i.item_code,
        i.generic_name,
        i.brand_name,
        COALESCE(SUM(
            CASE
                WHEN sm.to_location_id IS NOT NULL THEN sm.quantity
                WHEN sm.from_location_id IS NOT NULL THEN -sm.quantity
                ELSE 0
            END
        ), 0) AS current_stock,
        COALESCE(i.minimum_stock, 0) AS minimum_stock
    FROM items i
    LEFT JOIN stock_movements sm ON i.item_id = sm.item_id
    WHERE i.is_active = 1
    GROUP BY
        i.item_id,
        i.item_code,
        i.generic_name,
        i.brand_name,
        i.minimum_stock
    HAVING minimum_stock > 0
       AND current_stock < minimum_stock
    ORDER BY generic_name;
    """)

    rows = cur.fetchall()
    conn.close()
    return rows

def update_item_minimum_stock(item_id, minimum_stock):
    if minimum_stock < 0:
        raise ValueError("Minimum stock cannot be negative.")

    conn = connect()
    cur = conn.cursor()

    cur.execute("""
    UPDATE items
    SET minimum_stock = ?
    WHERE item_id = ?;
    """, (
        minimum_stock,
        item_id,
    ))

    if cur.rowcount == 0:
        conn.close()
        raise ValueError("Item not found.")

    conn.commit()
    conn.close()

def list_items():
    conn = connect()
    cur = conn.cursor()

    cur.execute("""
    SELECT item_id, item_code, generic_name, brand_name, dosage_form, strength
    FROM items
    ORDER BY generic_name;
    """)

    rows = cur.fetchall()
    conn.close()
    return rows


def list_recent_movements(limit=20):
    conn = connect()
    cur = conn.cursor()

    cur.execute("""
    SELECT
        sm.movement_id,
        sm.movement_datetime,
        sm.movement_type,
        i.generic_name,
        sm.quantity,
        lf.location_name AS from_location,
        lt.location_name AS to_location,
        sm.reason
    FROM stock_movements sm
    JOIN items i ON sm.item_id = i.item_id
    LEFT JOIN locations lf ON sm.from_location_id = lf.location_id
    LEFT JOIN locations lt ON sm.to_location_id = lt.location_id
    ORDER BY sm.movement_id DESC
    LIMIT ?;
    """, (limit,))

    rows = cur.fetchall()
    conn.close()
    return rows
