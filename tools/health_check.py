from pathlib import Path
import sqlite3
import importlib.util
import sys

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))

DB_FILE = PROJECT / "pharmaguard_operational_v12.db"

IMPORTANT_FILES = [
    "run_pharmaguard_gateway.py",
    "run_operational_launcher.py",
    "run_stock_viewer.py",
    "run_movement_entry.py",
    "run_batch_manager.py",
    "run_fefo_helper.py",
    "run_audit_viewer.py",
    "run_user_manager.py",
    "run_purchase_orders.py",
    "run_supplier_deliveries.py",
    "run_backup_restore.py",
    "run_operational_dashboard.py",
    "run_operational_forecast.py",
    "run_reports.py",
    "START_PHARMAGUARD.bat",
]

IMPORTANT_MODULES = [
    "app.session",
    "app.permissions",
    "app.operational_launcher",
    "app.main_gateway",
    "database.schema",
    "database.repositories",
    "database.stock_views",
    "database.purchase_orders",
    "database.supplier_deliveries",
    "database.operational_forecasting",
    "database.operational_reports",
]

IMPORTANT_TABLES = [
    "items",
    "locations",
    "batches",
    "stock_movements",
    "suppliers",
    "purchase_orders",
    "purchase_order_lines",
    "users",
    "audit_logs",
    "delivery_receipts",
]


def ok(message):
    print("[OK] " + message)


def warn(message):
    print("[WARN] " + message)


def fail(message):
    print("[FAIL] " + message)


def check_files():
    print("\n=== FILE CHECK ===")
    for file_name in IMPORTANT_FILES:
        path = PROJECT / file_name
        if path.exists():
            ok(file_name)
        else:
            warn(file_name + " is missing")


def check_modules():
    print("\n=== MODULE CHECK ===")
    for module_name in IMPORTANT_MODULES:
        try:
            spec = importlib.util.find_spec(module_name)
            if spec is None:
                warn(module_name + " not found")
            else:
                ok(module_name)
        except Exception as exc:
            fail(module_name + " -> " + str(exc))


def check_packages():
    print("\n=== PACKAGE CHECK ===")
    packages = ["PySide6", "openpyxl"]

    for package in packages:
        try:
            __import__(package)
            ok(package)
        except Exception:
            warn(package + " not installed")


def check_database():
    print("\n=== DATABASE CHECK ===")

    if not DB_FILE.exists():
        fail("Database file not found: " + str(DB_FILE))
        return

    ok("Database found: " + str(DB_FILE))

    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    cur.execute("SELECT name FROM sqlite_master WHERE type='table';")
    existing_tables = {row[0] for row in cur.fetchall()}

    for table in IMPORTANT_TABLES:
        if table not in existing_tables:
            warn(table + " table is missing")
            continue

        try:
            cur.execute(f"SELECT COUNT(*) FROM {table};")
            count = cur.fetchone()[0]
            ok(f"{table}: {count} rows")
        except Exception as exc:
            fail(f"{table}: {exc}")

    conn.close()


def main():
    print("PharmaGuard System Health Check")
    print("Project:", PROJECT)

    check_files()
    check_modules()
    check_packages()
    check_database()

    print("\nHealth check finished.")


if __name__ == "__main__":
    main()