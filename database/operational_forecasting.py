from database.schema import connect, init_database


def setup_forecasting():
    init_database()


def get_forecast_rows(coverage_days=60, lead_time_days=14):
    setup_forecasting()

    conn = connect()
    cur = conn.cursor()

    query = """
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
        i.item_id,
        i.item_code,
        i.generic_name,
        COALESCE(i.brand_name, '') AS brand_name,
        COALESCE(i.ven_class, '') AS ven_class,
        l.location_id,
        l.location_name,
        COALESCE(sn.current_stock, 0) AS current_stock,
        COALESCE(i30.issued_30d, 0) AS issued_30d,
        COALESCE(i90.issued_90d, 0) AS issued_90d
    FROM stock_now sn
    JOIN items i ON sn.item_id = i.item_id
    JOIN locations l ON sn.location_id = l.location_id
    LEFT JOIN issue_30 i30
        ON sn.item_id = i30.item_id
       AND sn.location_id = i30.location_id
    LEFT JOIN issue_90 i90
        ON sn.item_id = i90.item_id
       AND sn.location_id = i90.location_id
    WHERE i.is_active = 1
    ORDER BY i.generic_name, l.location_name;
    """

    cur.execute(query)
    rows = cur.fetchall()
    conn.close()

    result = []

    for row in rows:
        (
            item_id,
            item_code,
            generic_name,
            brand_name,
            ven_class,
            location_id,
            location_name,
            current_stock,
            issued_30d,
            issued_90d,
        ) = row

        current_stock = float(current_stock or 0)
        issued_30d = float(issued_30d or 0)
        issued_90d = float(issued_90d or 0)

        if issued_90d > 0:
            avg_monthly_demand = round((issued_90d / 90) * 30, 2)
            forecast_basis = "90d movement history"
        elif issued_30d > 0:
            avg_monthly_demand = round(issued_30d, 2)
            forecast_basis = "30d movement history"
        else:
            avg_monthly_demand = 0
            forecast_basis = "No issue history"

        daily_demand = round(avg_monthly_demand / 30, 4) if avg_monthly_demand > 0 else 0

        if daily_demand > 0:
            days_left = round(current_stock / daily_demand, 1)
        else:
            days_left = None

        target_stock = round(daily_demand * coverage_days, 2)
        suggested_reorder_qty = round(max(0, target_stock - current_stock), 2)

        if avg_monthly_demand <= 0:
            shortage_risk = "No consumption data"
        elif current_stock <= 0:
            shortage_risk = "Critical"
        elif days_left is not None and days_left <= lead_time_days:
            shortage_risk = "Critical"
        elif days_left is not None and days_left <= lead_time_days + 14:
            shortage_risk = "High"
        elif days_left is not None and days_left <= coverage_days:
            shortage_risk = "Medium"
        else:
            shortage_risk = "Low"

        if shortage_risk == "Critical":
            action = "Urgent replenishment / طلب عاجل"
        elif shortage_risk == "High":
            action = "Prepare purchase order / جهز طلب شراء"
        elif shortage_risk == "Medium":
            action = "Monitor weekly / متابعة أسبوعية"
        elif shortage_risk == "Low":
            action = "No action now / لا إجراء الآن"
        else:
            action = "Need issue history / يحتاج بيانات صرف"

        result.append({
            "item_id": item_id,
            "item_code": item_code,
            "generic_name": generic_name,
            "brand_name": brand_name,
            "ven_class": ven_class,
            "location_id": location_id,
            "location_name": location_name,
            "current_stock": current_stock,
            "issued_30d": issued_30d,
            "issued_90d": issued_90d,
            "avg_monthly_demand": avg_monthly_demand,
            "daily_demand": daily_demand,
            "days_left": "" if days_left is None else days_left,
            "coverage_days": coverage_days,
            "target_stock": target_stock,
            "suggested_reorder_qty": suggested_reorder_qty,
            "shortage_risk": shortage_risk,
            "forecast_basis": forecast_basis,
            "recommended_action": action,
        })

    return result


if __name__ == "__main__":
    rows = get_forecast_rows()
    print("Forecast rows:", len(rows))
    for row in rows[:20]:
        print(row)
