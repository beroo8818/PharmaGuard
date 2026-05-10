APP_VERSION = "V11 Forecast Database"
APP_TITLE_EN = f"PharmaGuard AI Pro {APP_VERSION}"
APP_TITLE_AR = f"فارماجارد برو {APP_VERSION}"

MAX_TABLE_ROWS = 800
CHART_TOP_N = 8

CORE_COLUMNS = [
    "medicine_name",
    "item_code",
    "current_stock",
    "min_stock_level",
    "avg_daily_consumption",
    "lead_time_days",
    "pending_po_qty",
    "expiry_date",
    "clinical_priority",
    "storage_location",
]

OPTIONAL_COLUMNS = [
    "branch_name",
    "avg_monthly_consumption",
    "unit_cost",
    "supplier_name",
    "supplier_on_time_rate",
    "ven_class",
    "abc_class",
    "generic_name",
    "brand_name",
    "brand_names_display",
    "dosage_form_group",
    "analysis_name",
    "source_medicine_names",
    "stock_status",
    "overstock_risk",
    "predicted_stock_30d",
    "predicted_stock_60d",
    "predicted_stock_90d",
    "selected_horizon_months",
    "recommended_reorder_qty_horizon",
]

ALL_COLUMNS = CORE_COLUMNS + OPTIONAL_COLUMNS
