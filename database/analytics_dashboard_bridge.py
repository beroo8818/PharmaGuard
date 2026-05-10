from __future__ import annotations

import json
import shutil
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
MONTHLY_OUTPUTS = PROJECT_ROOT / "monthly_outputs"
REPORTS_DIR = PROJECT_ROOT / "reports"
CONFIG_PATHS = [
    PROJECT_ROOT / "n8n_config.json",
    PROJECT_ROOT / "n8n_email_config.json",
]


def _append_sync_log(message: str):
    try:
        REPORTS_DIR.mkdir(exist_ok=True)
        log_path = REPORTS_DIR / "import_sync_log.txt"
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with log_path.open("a", encoding="utf-8") as f:
            f.write(f"[{stamp}] {message}\n")
    except Exception:
        pass


def _safe_num(series, default=0):
    try:
        return pd.to_numeric(series, errors="coerce").fillna(default)
    except Exception:
        return pd.Series(dtype=float)


def _pick(df: pd.DataFrame, names, default=""):
    for name in names:
        if name in df.columns:
            return df[name]
    return pd.Series([default] * len(df), index=df.index)


def make_monthly_report_from_analytics(df: pd.DataFrame, source_name: str = "Analytics import") -> Optional[Path]:
    """
    Bridge Analytics import into Executive + Storekeeper Dashboard.

    The dashboard reads monthly_outputs/monthly_reports_*.xlsx and expects sheets:
    00_Dashboard, 02_Item_Summary, 03_Critical_Items, 04_Branch_Summary,
    05_New_Medicines, 06_Next_Opening_Total, 07_Next_Opening_By_Branch,
    08_Month_Comparison.
    """
    if df is None or df.empty:
        return None

    MONTHLY_OUTPUTS.mkdir(exist_ok=True)
    REPORTS_DIR.mkdir(exist_ok=True)

    work = df.copy()
    n = len(work)

    medicine = _pick(work, ["generic_name", "medicine_name", "analysis_name"], "Unknown Item").astype(str)
    trade = _pick(work, ["brand_names_display", "brand_name", "source_medicine_names"], "").astype(str)
    unit = _pick(work, ["dosage_form_group"], "").astype(str)
    branch = _pick(work, ["branch_name", "storage_location"], "Main Store").astype(str)
    source_sheet = pd.Series([source_name] * n, index=work.index)

    current_stock = _safe_num(_pick(work, ["current_stock"], 0))
    avg_monthly = _safe_num(_pick(work, ["avg_monthly_consumption"], 0))
    avg_daily = _safe_num(_pick(work, ["avg_daily_consumption"], 0))
    if avg_monthly.sum() == 0 and avg_daily.sum() > 0:
        avg_monthly = avg_daily * 30
    if avg_daily.sum() == 0 and avg_monthly.sum() > 0:
        avg_daily = avg_monthly / 30

    dyn_min = _safe_num(_pick(work, ["dynamic_min_stock", "min_stock_level"], 0))
    if dyn_min.sum() == 0:
        dyn_min = (avg_monthly * 0.5).round(2)

    reorder = _safe_num(_pick(work, ["recommended_reorder_qty", "recommended_reorder_qty_horizon"], 0))
    if reorder.sum() == 0:
        reorder = (dyn_min - current_stock).clip(lower=0).round(2)

    shortage = _pick(work, ["shortage_risk"], "Low").astype(str)
    # If source did not calculate risks yet, infer a simple one.
    if set(shortage.str.lower().unique()) <= {"", "nan", "none", "low"}:
        days_left_tmp = current_stock / avg_daily.replace(0, pd.NA)
        shortage = pd.Series("Low", index=work.index)
        shortage.loc[(current_stock <= 0) & (avg_monthly > 0)] = "Critical"
        shortage.loc[(days_left_tmp.notna()) & (days_left_tmp <= 14)] = "Critical"
        shortage.loc[(days_left_tmp.notna()) & (days_left_tmp > 14) & (days_left_tmp <= 30)] = "High"
        shortage.loc[(current_stock < dyn_min) & (shortage == "Low")] = "High"

    days_left = current_stock / avg_daily.replace(0, pd.NA)
    days_left = days_left.fillna(9999).round(1)

    item_summary = pd.DataFrame({
        "source_sheet": source_sheet,
        "scientific_name": medicine,
        "trade_names": trade,
        "unit": unit,
        "branch": branch,
        "current_stock": current_stock.round(2),
        "avg_monthly_consumption": avg_monthly.round(2),
        "total_issued_month": avg_monthly.round(2),
        "days_left": days_left,
        "dynamic_min": dyn_min.round(2),
        "suggested_reorder_qty": reorder.round(2),
        "shortage_risk": shortage,
        "note": _pick(work, ["reason_explanation", "recommended_action"], "").astype(str),
    })

    critical = item_summary[item_summary["shortage_risk"].isin(["Critical", "High"])].copy()
    branch_summary = item_summary.groupby("branch", dropna=False).agg(
        total_issued=("total_issued_month", "sum"),
        total_stock=("current_stock", "sum"),
        item_count=("scientific_name", "count"),
    ).reset_index()

    next_opening_total = item_summary.groupby(["source_sheet", "scientific_name", "trade_names", "unit"], dropna=False).agg(
        opening_stock_next_month=("current_stock", "sum")
    ).reset_index()
    next_opening_total.insert(0, "next_month", "Next")

    next_opening_branch = item_summary[["source_sheet", "scientific_name", "trade_names", "unit", "branch", "current_stock"]].copy()
    next_opening_branch.insert(0, "next_month", "Next")
    next_opening_branch = next_opening_branch.rename(columns={"current_stock": "opening_stock", "trade_names": "trade_name"})

    new_medicines = pd.DataFrame(columns=["scientific_name", "trade_names", "unit", "branch"])
    comparison = pd.DataFrame([{
        "metric": "Generated from Analytics import",
        "current": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "previous": "",
        "change": "",
    }])

    dashboard = pd.DataFrame([
        {"metric": "month", "value": datetime.now().strftime("%Y-%m")},
        {"metric": "next_month", "value": "Next"},
        {"metric": "items_count", "value": int(len(item_summary))},
        {"metric": "critical_items", "value": int((item_summary["shortage_risk"] == "Critical").sum())},
        {"metric": "high_risk_items", "value": int((item_summary["shortage_risk"] == "High").sum())},
        {"metric": "total_current_stock", "value": float(current_stock.sum())},
        {"metric": "total_issued_month", "value": float(avg_monthly.sum())},
        {"metric": "source", "value": source_name},
        {"metric": "generated_at", "value": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
    ])

    safe_source = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in str(source_name))[:40] or "analytics"
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    monthly_path = MONTHLY_OUTPUTS / f"monthly_reports_{safe_source}_{stamp}.xlsx"

    with pd.ExcelWriter(monthly_path, engine="openpyxl") as writer:
        dashboard.to_excel(writer, index=False, sheet_name="00_Dashboard")
        item_summary.to_excel(writer, index=False, sheet_name="02_Item_Summary")
        critical.to_excel(writer, index=False, sheet_name="03_Critical_Items")
        branch_summary.to_excel(writer, index=False, sheet_name="04_Branch_Summary")
        new_medicines.to_excel(writer, index=False, sheet_name="05_New_Medicines")
        next_opening_total.to_excel(writer, index=False, sheet_name="06_Next_Opening_Total")
        next_opening_branch.to_excel(writer, index=False, sheet_name="07_Next_Opening_By_Branch")
        comparison.to_excel(writer, index=False, sheet_name="08_Month_Comparison")

    # Keep a stable latest copy so Executive Dashboard always reads the newest Analytics import.
    try:
        latest_monthly = MONTHLY_OUTPUTS / "monthly_reports_latest.xlsx"
        shutil.copy2(monthly_path, latest_monthly)
    except Exception as copy_exc:
        _append_sync_log(f"WARNING: could not create latest monthly copy: {copy_exc}")

    _append_sync_log(
        f"Analytics import synced to Executive Dashboard | source={source_name} | rows={len(item_summary)} | file={monthly_path.name}"
    )

    # Also copy a stable analytics report into reports folder for n8n package attachment.
    latest_report = REPORTS_DIR / "analytics_latest_for_email.xlsx"
    with pd.ExcelWriter(latest_report, engine="openpyxl") as writer:
        dashboard.to_excel(writer, index=False, sheet_name="Dashboard")
        item_summary.to_excel(writer, index=False, sheet_name="Item Summary")
        critical.to_excel(writer, index=False, sheet_name="Critical High Alerts")
        branch_summary.to_excel(writer, index=False, sheet_name="Branch Summary")

    _append_sync_log(f"Analytics email report updated | file={latest_report.name}")

    return monthly_path


def after_analytics_import(window) -> Optional[Path]:
    """Call this after Analytics recompute. It updates the Executive Dashboard source."""
    df = getattr(window, "df", None)
    source = getattr(window, "current_sheet", "Analytics import") or "Analytics import"
    path = make_monthly_report_from_analytics(df, source)
    try:
        if path and hasattr(window, "status"):
            window.status.showMessage(f"Executive Dashboard monthly report updated: {path.name}", 8000)
    except Exception:
        pass
    return path
