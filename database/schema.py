import sqlite3
from pathlib import Path


DB_NAME = "pharmaguard_operational_v12.db"


def get_db_path():
    return Path.cwd() / DB_NAME


def connect(db_path=None):
    path = Path(db_path) if db_path else get_db_path()
    conn = sqlite3.connect(path)
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def init_database(db_path=None):
    conn = connect(db_path)
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS users (
        user_id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT NOT NULL UNIQUE,
        full_name TEXT,
        role TEXT NOT NULL DEFAULT 'viewer',
        is_active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS locations (
        location_id INTEGER PRIMARY KEY AUTOINCREMENT,
        location_name TEXT NOT NULL UNIQUE,
        location_type TEXT NOT NULL DEFAULT 'pharmacy',
        is_active INTEGER NOT NULL DEFAULT 1
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS suppliers (
        supplier_id INTEGER PRIMARY KEY AUTOINCREMENT,
        supplier_name TEXT NOT NULL UNIQUE,
        phone TEXT,
        notes TEXT,
        is_active INTEGER NOT NULL DEFAULT 1
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS items (
        item_id INTEGER PRIMARY KEY AUTOINCREMENT,
        item_code TEXT UNIQUE,
        generic_name TEXT NOT NULL,
        brand_name TEXT,
        dosage_form TEXT,
        strength TEXT,
        ven_class TEXT DEFAULT 'E',
        abc_class TEXT,
        unit_cost REAL DEFAULT 0,
        minimum_stock REAL DEFAULT 0,
        is_active INTEGER NOT NULL DEFAULT 1
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS batches (
        batch_id INTEGER PRIMARY KEY AUTOINCREMENT,
        item_id INTEGER NOT NULL,
        batch_number TEXT NOT NULL,
        expiry_date TEXT,
        supplier_id INTEGER,
        received_date TEXT,
        notes TEXT,
        FOREIGN KEY (item_id) REFERENCES items(item_id),
        FOREIGN KEY (supplier_id) REFERENCES suppliers(supplier_id)
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS stock_movements (
        movement_id INTEGER PRIMARY KEY AUTOINCREMENT,
        movement_datetime TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        movement_type TEXT NOT NULL,
        item_id INTEGER NOT NULL,
        batch_id INTEGER,
        from_location_id INTEGER,
        to_location_id INTEGER,
        quantity REAL NOT NULL,
        reference_no TEXT,
        reason TEXT,
        user_id INTEGER,
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (item_id) REFERENCES items(item_id),
        FOREIGN KEY (batch_id) REFERENCES batches(batch_id),
        FOREIGN KEY (from_location_id) REFERENCES locations(location_id),
        FOREIGN KEY (to_location_id) REFERENCES locations(location_id),
        FOREIGN KEY (user_id) REFERENCES users(user_id)
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS purchase_orders (
        po_id INTEGER PRIMARY KEY AUTOINCREMENT,
        po_number TEXT UNIQUE,
        supplier_id INTEGER,
        status TEXT NOT NULL DEFAULT 'Draft',
        requested_by INTEGER,
        approved_by INTEGER,
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        approved_at TEXT,
        notes TEXT,
        FOREIGN KEY (supplier_id) REFERENCES suppliers(supplier_id),
        FOREIGN KEY (requested_by) REFERENCES users(user_id),
        FOREIGN KEY (approved_by) REFERENCES users(user_id)
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS purchase_order_lines (
        po_line_id INTEGER PRIMARY KEY AUTOINCREMENT,
        po_id INTEGER NOT NULL,
        item_id INTEGER NOT NULL,
        requested_qty REAL NOT NULL,
        approved_qty REAL DEFAULT 0,
        received_qty REAL DEFAULT 0,
        unit_cost REAL DEFAULT 0,
        notes TEXT,
        FOREIGN KEY (po_id) REFERENCES purchase_orders(po_id),
        FOREIGN KEY (item_id) REFERENCES items(item_id)
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS audit_logs (
        audit_id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        action TEXT NOT NULL,
        table_name TEXT,
        record_id TEXT,
        old_value TEXT,
        new_value TEXT,
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (user_id) REFERENCES users(user_id)
    );
    """)

    conn.commit()
    conn.close()

    ensure_items_minimum_stock_column()

    return str(get_db_path())

def ensure_items_minimum_stock_column(db_path=None):
    conn = connect(db_path)
    cur = conn.cursor()

    cur.execute("PRAGMA table_info(items);")
    columns = [row[1] for row in cur.fetchall()]

    if "minimum_stock" not in columns:
        cur.execute("ALTER TABLE items ADD COLUMN minimum_stock REAL DEFAULT 0;")

    conn.commit()
    conn.close()


def seed_basic_data(db_path=None):
    conn = connect(db_path)
    cur = conn.cursor()

    cur.execute("""
    INSERT OR IGNORE INTO users (username, full_name, role)
    VALUES ('admin', 'System Admin', 'admin');
    """)

    default_locations = [
        ('Main Store', 'store'),
        ('Main Pharmacy', 'pharmacy'),
        ('ER Pharmacy', 'pharmacy'),
        ('ICU Pharmacy', 'pharmacy')
    ]

    for name, loc_type in default_locations:
        cur.execute("""
        INSERT OR IGNORE INTO locations (location_name, location_type)
        VALUES (?, ?);
        """, (name, loc_type))

    cur.execute("""
    INSERT OR IGNORE INTO suppliers (supplier_name, notes)
    VALUES ('Unknown Supplier', 'Default supplier placeholder');
    """)

    conn.commit()
    conn.close()


def list_tables(db_path=None):
    conn = connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name;")
    rows = [r[0] for r in cur.fetchall()]
    conn.close()
    return rows


if __name__ == "__main__":
    db_path = init_database()
    seed_basic_data()
    print("Database created:")
    print(db_path)
    print("Tables:")
    for table in list_tables():
        print("-", table)
