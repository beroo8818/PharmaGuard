import sys
from pathlib import Path
from datetime import datetime

# Make project root visible so Python can import legacy and database modules
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from legacy import pharmaguard_legacy as legacy
from database.schema import connect
from database.repositories import (
    setup_database,
    get_user_id,
    add_supplier,
    add_location,
    add_item,
    add_batch,
    add_stock_movement,
    get_current_stock,
)


def safe_text(value, default=""):
    if value is None:
        return default

    text = str(value).strip()

    if text.lower() in {"", "nan", "nat", "none", "null"}:
        return default

    return text


def safe_number(value, default=0):
    try:
        if value is None:
            return default

        text = str(value).strip()

        if text.lower() in {"", "nan", "nat", "none", "null"}:
            return default

        return float(text)

    except Exception:
        return default


def ensure_sync_table():
    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS import_sync_state (
            row_key TEXT PRIMARY KEY,
            source_name TEXT,
            item_code TEXT,
            generic_name TEXT,
            location_name TEXT,
            batch_number TEXT,
            last_sheet_stock REAL DEFAULT 0,
            last_operational_stock REAL DEFAULT 0,
            last_sync_at TEXT
        );
    """)

    conn.commit()
    conn.close()


def get_or_create_opening_batch(item_id, supplier_id):
    """
    Creates or reuses the placeholder batch for legacy data without batch/expiry.
    """
    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        SELECT batch_id
        FROM batches
        WHERE item_id = ?
          AND batch_number = 'OPENING-BALANCE'
        LIMIT 1;
    """, (item_id,))

    row = cur.fetchone()
    conn.close()

    if row:
        return row[0]

    return add_batch(
        item_id=item_id,
        batch_number="OPENING-BALANCE",
        expiry_date="2099-12-31",
        supplier_id=supplier_id,
        received_date=None,
        notes=(
            "Live sync placeholder batch. "
            "OPENING-BALANCE / 2099-12-31 means the source file has no real batch or expiry."
        ),
    )


def make_row_key(item_code, generic_name, location_name):
    """
    This key tells the sync system that the same medicine in the same location
    is the same row every time we refresh.
    """
    item_part = item_code or generic_name or "UNKNOWN_ITEM"
    location_part = location_name or "Main Store"

    return f"{item_part}__{location_part}".replace(" ", "_").upper()


def upsert_sync_state(
    row_key,
    source_name,
    item_code,
    generic_name,
    location_name,
    batch_number,
    last_sheet_stock,
    last_operational_stock,
):
    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO import_sync_state (
            row_key,
            source_name,
            item_code,
            generic_name,
            location_name,
            batch_number,
            last_sheet_stock,
            last_operational_stock,
            last_sync_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(row_key)
        DO UPDATE SET
            source_name = excluded.source_name,
            item_code = excluded.item_code,
            generic_name = excluded.generic_name,
            location_name = excluded.location_name,
            batch_number = excluded.batch_number,
            last_sheet_stock = excluded.last_sheet_stock,
            last_operational_stock = excluded.last_operational_stock,
            last_sync_at = excluded.last_sync_at;
    """, (
        row_key,
        source_name,
        item_code,
        generic_name,
        location_name,
        batch_number,
        last_sheet_stock,
        last_operational_stock,
        datetime.now().isoformat(timespec="seconds"),
    ))

    conn.commit()
    conn.close()


def load_dataframe_from_source(source):
    """
    Source can be:
    - Excel file path
    - CSV file path
    - Google Sheet URL, if the legacy importer supports it
    """
    df_raw, sheet_used, report, preview = legacy.import_any_file(source)
    df = legacy.compute_metrics(df_raw)
    return df, sheet_used


def live_sync_dataframe_to_operational_db(df, source_name="LiveSync"):
    setup_database()
    ensure_sync_table()

    admin_id = get_user_id("admin")
    if admin_id is None:
        raise RuntimeError("Admin user was not found. Run setup_database first.")

    synced_rows = 0
    no_change_rows = 0
    increased_rows = 0
    decreased_rows = 0
    skipped_rows = 0
    errors = []

    for index, row in df.iterrows():
        try:
            generic_name = (
                safe_text(row.get("generic_name"))
                or safe_text(row.get("medicine_name"))
                or "Unknown Item"
            )

            brand_name = (
                safe_text(row.get("brand_names_display"))
                or safe_text(row.get("brand_name"))
            )

            dosage_form = safe_text(row.get("dosage_form_group"))
            item_code = safe_text(row.get("item_code")) or f"AUTO-{index + 1:05d}"
            location_name = (
                safe_text(row.get("branch_name"))
                or safe_text(row.get("storage_location"))
                or "Main Store"
            )
            supplier_name = safe_text(row.get("supplier_name")) or "Unknown Supplier"
            ven_class = safe_text(row.get("ven_class"), "E") or "E"
            unit_cost = safe_number(row.get("unit_cost"), 0)
            sheet_stock = safe_number(row.get("current_stock"), 0)

            if sheet_stock < 0:
                skipped_rows += 1
                errors.append(f"Skipped {item_code}: negative stock is not allowed.")
                continue

            supplier_id = add_supplier(supplier_name)
            location_id = add_location(location_name, "pharmacy")

            item_id = add_item(
                item_code=item_code,
                generic_name=generic_name,
                brand_name=brand_name,
                dosage_form=dosage_form,
                strength="",
                ven_class=ven_class,
                unit_cost=unit_cost,
            )

            batch_id = get_or_create_opening_batch(item_id, supplier_id)

            row_key = make_row_key(
                item_code=item_code,
                generic_name=generic_name,
                location_name=location_name,
            )

            operational_stock = get_current_stock(
                item_id=item_id,
                location_id=location_id,
                batch_id=batch_id,
            )

            difference = sheet_stock - operational_stock

            # Avoid tiny floating point noise
            if abs(difference) < 0.000001:
                no_change_rows += 1
                upsert_sync_state(
                    row_key=row_key,
                    source_name=source_name,
                    item_code=item_code,
                    generic_name=generic_name,
                    location_name=location_name,
                    batch_number="OPENING-BALANCE",
                    last_sheet_stock=sheet_stock,
                    last_operational_stock=operational_stock,
                )
                continue

            reference_no = (
                f"LSYNC-{datetime.now().strftime('%Y%m%d%H%M%S')}-"
                f"{index + 1}-{item_code}-{location_id}"
            )

            if difference > 0:
                add_stock_movement(
                    movement_type="Receive",
                    item_id=item_id,
                    quantity=abs(difference),
                    from_location_id=None,
                    to_location_id=location_id,
                    batch_id=batch_id,
                    reference_no=reference_no,
                    reason=f"Live sync increase from {operational_stock} to {sheet_stock}",
                    user_id=admin_id,
                )
                increased_rows += 1

            else:
                # Temporary practical solution:
                # current system supports Waste as a decrease movement.
                # In a professional version, this should become AdjustmentOut.
                add_stock_movement(
                    movement_type="Waste",
                    item_id=item_id,
                    quantity=abs(difference),
                    from_location_id=location_id,
                    to_location_id=None,
                    batch_id=batch_id,
                    reference_no=reference_no,
                    reason=f"Live sync decrease from {operational_stock} to {sheet_stock}",
                    user_id=admin_id,
                )
                decreased_rows += 1

            synced_rows += 1

            upsert_sync_state(
                row_key=row_key,
                source_name=source_name,
                item_code=item_code,
                generic_name=generic_name,
                location_name=location_name,
                batch_number="OPENING-BALANCE",
                last_sheet_stock=sheet_stock,
                last_operational_stock=sheet_stock,
            )

        except Exception as exc:
            skipped_rows += 1
            errors.append(f"Row {index + 1} failed: {exc}")

    return {
        "synced_rows": synced_rows,
        "no_change_rows": no_change_rows,
        "increased_rows": increased_rows,
        "decreased_rows": decreased_rows,
        "skipped_rows": skipped_rows,
        "errors": errors,
    }


def main():
    if len(sys.argv) < 2:
        print("Usage:")
        print("python database/live_sync_from_source.py <excel_csv_path_or_google_sheet_url>")
        sys.exit(1)

    source = sys.argv[1]

    df, sheet_used = load_dataframe_from_source(source)

    result = live_sync_dataframe_to_operational_db(
        df=df,
        source_name=sheet_used,
    )

    print("Live sync completed.")
    print(f"Source: {source}")
    print(f"Sheet used: {sheet_used}")
    print(f"Synced rows: {result['synced_rows']}")
    print(f"No change rows: {result['no_change_rows']}")
    print(f"Increased rows: {result['increased_rows']}")
    print(f"Decreased rows: {result['decreased_rows']}")
    print(f"Skipped rows: {result['skipped_rows']}")

    if result["errors"]:
        print("")
        print("Errors:")
        for error in result["errors"][:20]:
            print("-", error)


if __name__ == "__main__":
    main()