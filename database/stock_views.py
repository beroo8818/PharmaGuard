import csv
from database.schema import connect


def get_current_stock_rows(include_zero=False):
    conn = connect()
    cur = conn.cursor()

    having_clause = "" if include_zero else "HAVING current_stock <> 0"

    query = f"""
    WITH movement_lines AS (
        SELECT
            item_id,
            batch_id,
            to_location_id AS location_id,
            quantity AS signed_qty
        FROM stock_movements
        WHERE movement_type IN ('Receive', 'Return', 'Adjustment')
          AND to_location_id IS NOT NULL

        UNION ALL

        SELECT
            item_id,
            batch_id,
            to_location_id AS location_id,
            quantity AS signed_qty
        FROM stock_movements
        WHERE movement_type = 'Transfer'
          AND to_location_id IS NOT NULL

        UNION ALL

        SELECT
            item_id,
            batch_id,
            from_location_id AS location_id,
            -quantity AS signed_qty
        FROM stock_movements
        WHERE movement_type IN ('Issue', 'Waste')
          AND from_location_id IS NOT NULL

        UNION ALL

        SELECT
            item_id,
            batch_id,
            from_location_id AS location_id,
            -quantity AS signed_qty
        FROM stock_movements
        WHERE movement_type = 'Transfer'
          AND from_location_id IS NOT NULL
    )

    SELECT
        i.item_code,
        i.generic_name,
        i.brand_name,
        i.dosage_form,
        i.strength,
        l.location_name,
        l.location_type,
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
        i.strength,
        l.location_name,
        l.location_type,
        b.batch_number,
        b.expiry_date
    {having_clause}
    ORDER BY
        i.generic_name,
        l.location_name,
        b.expiry_date;
    """

    cur.execute(query)
    columns = [desc[0] for desc in cur.description]
    rows = [dict(zip(columns, row)) for row in cur.fetchall()]

    conn.close()
    return rows


def print_current_stock(limit=100):
    rows = get_current_stock_rows()

    print("Current stock from movement ledger")
    print("-" * 60)

    if not rows:
        print("No stock rows found.")
        return

    for row in rows[:limit]:
        print(
            f"{row['generic_name']} | "
            f"{row['location_name']} | "
            f"Batch: {row['batch_number']} | "
            f"Expiry: {row['expiry_date']} | "
            f"Stock: {row['current_stock']}"
        )

    if len(rows) > limit:
        print(f"... showing first {limit} of {len(rows)} rows")


def export_current_stock_csv(path="current_stock_from_ledger.csv"):
    rows = get_current_stock_rows()

    if not rows:
        print("No stock rows to export.")
        return None

    columns = list(rows[0].keys())

    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=columns)
        writer.writeheader()
        writer.writerows(rows)

    print("CSV exported:")
    print(path)
    return path


if __name__ == "__main__":
    print_current_stock()
    export_current_stock_csv()
