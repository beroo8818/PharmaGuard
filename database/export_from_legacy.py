from pathlib import Path
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from legacy import pharmaguard_legacy as legacy
from database.schema import connect
from database.repositories import (
    setup_database,
    add_location,
    add_supplier,
    add_item,
    add_batch,
    get_user_id,
    add_stock_movement,
)


def safe_text(value, default=""):
    if value is None:
        return default
    text = str(value).strip()
    if text.lower() in {"nan", "none", "nat"}:
        return default
    return text


def safe_number(value, default=0):
    try:
        if value is None:
            return default
        text = str(value).strip()
        if text.lower() in {"nan", "none", "nat", ""}:
            return default
        return float(text)
    except Exception:
        return default


def movement_reference_exists(reference_no):
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        "SELECT movement_id FROM stock_movements WHERE reference_no = ? LIMIT 1;",
        (reference_no,),
    )
    row = cur.fetchone()
    conn.close()
    return row is not None


def load_legacy_dataframe():
    # لو كتبت مسار ملف Excel بعد اسم السكريبت، هيستورده.
    # لو ما كتبتش مسار، هيستخدم Demo Data.
    if len(sys.argv) >= 2:
        path = sys.argv[1]
        print("Importing file:")
        # print(path)

        df_raw, sheet_used, report, preview = legacy.import_any_file(path)
        df = legacy.compute_metrics(df_raw)

        print("Imported sheet/source:", sheet_used)
        print("Rows:", len(df))
        return df, sheet_used

    print("No Excel path provided. Using demo data.")
    df_raw = legacy.create_demo_df()
    df = legacy.compute_metrics(df_raw)
    return df, "Demo"


def export_dataframe_to_operational_db(df, source_name="Legacy"):
    setup_database()

    admin_id = get_user_id("admin")
    if admin_id is None:
        raise RuntimeError("Admin user was not found. Run setup_database first.")

    inserted_items = 0
    inserted_movements = 0
    skipped_movements = 0

    for index, row in df.iterrows():
        generic_name = safe_text(row.get("generic_name")) or safe_text(row.get("medicine_name")) or "Unknown Item"
        brand_name = safe_text(row.get("brand_names_display")) or safe_text(row.get("brand_name"))
        dosage_form = safe_text(row.get("dosage_form_group"))
        item_code = safe_text(row.get("item_code")) or f"AUTO-{index + 1:05d}"
        location_name = safe_text(row.get("branch_name")) or safe_text(row.get("storage_location")) or "Main Store"
        supplier_name = safe_text(row.get("supplier_name")) or "Unknown Supplier"
        ven_class = safe_text(row.get("ven_class"), "E") or "E"
        unit_cost = safe_number(row.get("unit_cost"), 0)
        current_stock = safe_number(row.get("current_stock"), 0)

        if current_stock <= 0:
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
        inserted_items += 1

        expiry_value = row.get("expiry_date")
        expiry_date = ""
        try:
            if expiry_value is not None and str(expiry_value).lower() not in {"nan", "nat", "none"}:
                expiry_date = str(expiry_value)[:10]
        except Exception:
            expiry_date = ""

        batch_number = f"OPENING-{item_code}-{location_id}"

                # ==============================
        # Batch / Expiry fallback
        # ==============================
        # If Excel has no batch_number or expiry_date,
        # create a safe legacy placeholder so both modules can work.

        batch_value = row.get("batch_number")
        expiry_value = row.get("expiry_date")

        batch_number = safe_text(batch_value)
        expiry_date = safe_text(expiry_value)

        # If expiry comes as datetime text, keep only YYYY-MM-DD
        if expiry_date:
            expiry_date = expiry_date[:10]

        # If no batch exists in Excel, use legacy placeholder
        if not batch_number:
            batch_number = "OPENING-BALANCE"

        # If no expiry exists in Excel, use legacy placeholder date
        # This is NOT a real expiry date.
        if not expiry_date:
            expiry_date = "2099-12-31"

        batch_id = add_batch(
            item_id=item_id,
            batch_number=batch_number,
            expiry_date=expiry_date,
            supplier_id=supplier_id,
            received_date=None,
            notes=(
                f"Imported from {source_name}. "
                f"OPENING-BALANCE / 2099-12-31 means legacy placeholder "
                f"because the imported file has no batch or expiry."
            ),
        )

        reference_no = f"OPENING-{source_name}-{index + 1}-{item_code}-{location_id}"

        if movement_reference_exists(reference_no):
            skipped_movements += 1
            continue

        add_stock_movement(
            movement_type="Receive",
            item_id=item_id,
            batch_id=batch_id,
            to_location_id=location_id,
            quantity=current_stock,
            reference_no=reference_no,
            reason=f"Opening balance imported from {source_name}",
            user_id=admin_id,
        )

        inserted_movements += 1

    print("")
    print("Export finished.")
    print("Items processed:", inserted_items)
    print("Opening balance movements inserted:", inserted_movements)
    print("Skipped existing movements:", skipped_movements)


if __name__ == "__main__":
    dataframe, source = load_legacy_dataframe()
    export_dataframe_to_operational_db(dataframe, source)
