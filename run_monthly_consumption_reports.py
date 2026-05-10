# -*- coding: utf-8 -*-
"""
PharmaGuard Monthly Consumption Processor
Reads the Ministry-style monthly workbook (multiple sheets with branch columns: منصرف / رصيد),
normalizes it, creates reports, detects new medicines, and creates next-month opening balance.

Run:
    python run_monthly_consumption_reports.py
or:
    python run_monthly_consumption_reports.py "file.xlsx"
"""
from __future__ import annotations

import sys
import re
import json
import shutil
from pathlib import Path
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import numpy as np

PROJECT_ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = PROJECT_ROOT / "monthly_outputs"
ARCHIVE_DIR = PROJECT_ROOT / "imported_files"
MASTER_FILE = OUTPUT_DIR / "master_medicines.xlsx"
LAST_ITEM_SUMMARY_FILE = OUTPUT_DIR / "last_item_summary.xlsx"
LAST_NORMALIZED_FILE = OUTPUT_DIR / "last_normalized.xlsx"

LEAD_TIME_DAYS = 14
CRITICAL_DAYS = 14
HIGH_DAYS = 30
MODERATE_DAYS = 60

ARABIC_MONTHS = {
    "يناير": 1, "فبراير": 2, "مارس": 3, "ابريل": 4, "إبريل": 4, "أبريل": 4,
    "مايو": 5, "يونيو": 6, "يوليو": 7, "اغسطس": 8, "أغسطس": 8, "سبتمبر": 9,
    "اكتوبر": 10, "أكتوبر": 10, "نوفمبر": 11, "ديسمبر": 12,
}


def clean_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    text = str(value).replace("\n", " ").replace("\r", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return text


def clean_number(value: Any) -> float:
    if value is None or pd.isna(value):
        return 0.0
    if isinstance(value, str):
        value = value.strip()
        if value in {"", "-", "--", "—"}:
            return 0.0
        value = value.replace(",", "")
    try:
        return float(value)
    except Exception:
        return 0.0


def norm_key(text: str) -> str:
    text = clean_text(text).lower()
    replacements = {
        "أ": "ا", "إ": "ا", "آ": "ا", "ى": "ي", "ة": "ه",
        "ـ": "", "mg": "mg", "mcg": "mcg", "ml": "ml",
    }
    for a, b in replacements.items():
        text = text.replace(a, b)
    text = re.sub(r"[^a-z0-9\u0600-\u06FF]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def detect_month_year(path: Path) -> Tuple[str, Optional[int], Optional[int], str]:
    name = path.stem
    month_name = "غير محدد"
    month_num = None
    year = None
    for m, n in ARABIC_MONTHS.items():
        if m in name:
            month_name = m
            month_num = n
            break
    y = re.search(r"(20\d{2})", name)
    if y:
        year = int(y.group(1))
    label = f"{month_name} {year or ''}".strip()
    return label, month_num, year, name


def next_month_label(month_num: Optional[int], year: Optional[int]) -> str:
    if not month_num or not year:
        return "الشهر التالي"
    inv = {v: k for k, v in ARABIC_MONTHS.items() if not k.startswith("إ") and not k.startswith("أ")}
    nm = month_num + 1
    ny = year
    if nm == 13:
        nm = 1
        ny += 1
    return f"{inv.get(nm, str(nm))} {ny}"


def find_header_row(df: pd.DataFrame) -> Optional[int]:
    for idx in range(min(len(df), 20)):
        row_text = " ".join(clean_text(x) for x in df.iloc[idx].tolist())
        if "الاسم" in row_text and ("العلم" in row_text or "التجاري" in row_text) and "رصيد" in row_text:
            return idx
    return None


def find_col(columns: List[str], patterns: List[str]) -> Optional[int]:
    for i, col in enumerate(columns):
        c = clean_text(col)
        for p in patterns:
            if p in c:
                return i
    return None


def detect_branch_columns(header_values: List[Any], subheader_values: List[Any], start_col: int) -> List[Dict[str, Any]]:
    branches: Dict[str, Dict[str, Any]] = {}
    current_branch = ""

    for c in range(start_col, len(header_values)):
        header = clean_text(header_values[c])
        sub = clean_text(subheader_values[c]) if c < len(subheader_values) else ""

        if header:
            current_branch = header

        if not current_branch:
            continue

        # Ignore aggregate text columns, keep only branch columns with منصرف/رصيد signals.
        if "منصرف" in sub:
            branches.setdefault(current_branch, {"branch": current_branch, "issued_col": None, "stock_col": None})["issued_col"] = c
        elif "رصيد" in sub:
            branches.setdefault(current_branch, {"branch": current_branch, "issued_col": None, "stock_col": None})["stock_col"] = c

    return list(branches.values())


def normalize_workbook(path: Path) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, Any]]:
    xls = pd.ExcelFile(path)
    normalized_rows: List[Dict[str, Any]] = []
    item_rows: List[Dict[str, Any]] = []
    import_log: Dict[str, Any] = {
        "source_file": str(path),
        "sheets_seen": xls.sheet_names,
        "sheets_imported": [],
        "sheets_skipped": [],
        "generated_at": datetime.now().isoformat(timespec="seconds"),
    }

    month_label, month_num, year, source_name = detect_month_year(path)
    next_label = next_month_label(month_num, year)
    import_log.update({"month_label": month_label, "next_month_label": next_label})

    for sheet in xls.sheet_names:
        # Avoid accidental duplicate copy sheets.
        if "copy" in sheet.lower() or "نسخة" in sheet.lower():
            import_log["sheets_skipped"].append({"sheet": sheet, "reason": "duplicate/copy sheet"})
            continue

        df = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object)
        header_idx = find_header_row(df)
        if header_idx is None or header_idx + 1 >= len(df):
            import_log["sheets_skipped"].append({"sheet": sheet, "reason": "header row not found"})
            continue

        header = [clean_text(x) for x in df.iloc[header_idx].tolist()]
        subheader = [clean_text(x) for x in df.iloc[header_idx + 1].tolist()]

        serial_col = find_col(header, ["م"])
        system_col = find_col(header, ["الاسم كما هو مدون", "مدون علي المنظومة", "مدون على المنظومة"])
        scientific_col = find_col(header, ["الاسم العلم"])
        trade_col = find_col(header, ["الاسم التجاري", "الاسم التجارى"])
        unit_col = find_col(header, ["الوحدة"])
        pack_col = find_col(header, ["عدد الأقراص"])
        avg_year_col = find_col(header, ["سنويا", "سنوياً"])
        avg_recent_col = find_col(header, ["نصف سنوي", "نصف سنوى"])
        end_stock_col = find_col(header, ["رصيد شهر"])

        if end_stock_col is None:
            import_log["sheets_skipped"].append({"sheet": sheet, "reason": "end stock column not found"})
            continue

        branch_defs = detect_branch_columns(header, subheader, end_stock_col + 1)
        if not branch_defs:
            import_log["sheets_skipped"].append({"sheet": sheet, "reason": "branch columns not found"})
            continue

        import_log["sheets_imported"].append({"sheet": sheet, "header_row_excel": header_idx + 1, "branches": [b["branch"] for b in branch_defs]})

        last_system = ""
        last_scientific = ""
        last_unit = ""

        for r in range(header_idx + 2, len(df)):
            row = df.iloc[r].tolist()
            serial = row[serial_col] if serial_col is not None and serial_col < len(row) else None
            system_name = clean_text(row[system_col]) if system_col is not None and system_col < len(row) else ""
            scientific_name = clean_text(row[scientific_col]) if scientific_col is not None and scientific_col < len(row) else ""
            trade_name = clean_text(row[trade_col]) if trade_col is not None and trade_col < len(row) else ""
            unit = clean_text(row[unit_col]) if unit_col is not None and unit_col < len(row) else ""

            if system_name:
                last_system = system_name
            if scientific_name:
                last_scientific = scientific_name
            if unit:
                last_unit = unit

            system_name = system_name or last_system
            scientific_name = scientific_name or last_scientific
            unit = unit or last_unit

            # Skip truly empty rows.
            if not any([system_name, scientific_name, trade_name]):
                continue

            avg_year = clean_number(row[avg_year_col]) if avg_year_col is not None and avg_year_col < len(row) else 0.0
            avg_recent = clean_number(row[avg_recent_col]) if avg_recent_col is not None and avg_recent_col < len(row) else 0.0
            end_stock = clean_number(row[end_stock_col]) if end_stock_col is not None and end_stock_col < len(row) else 0.0
            pack_count = clean_number(row[pack_col]) if pack_col is not None and pack_col < len(row) else 0.0

            item_key = norm_key(scientific_name or system_name or trade_name) + " | " + norm_key(unit)
            item_base = {
                "month": month_label,
                "next_month": next_label,
                "source_sheet": sheet,
                "excel_row": r + 1,
                "serial": serial,
                "item_key": item_key,
                "system_name": system_name,
                "scientific_name": scientific_name,
                "trade_name": trade_name,
                "unit": unit,
                "pack_count": pack_count,
                "avg_monthly_year": avg_year,
                "avg_monthly_recent": avg_recent,
                "end_stock_total": end_stock,
            }
            item_rows.append(item_base.copy())

            for b in branch_defs:
                issued = clean_number(row[b["issued_col"]]) if b["issued_col"] is not None and b["issued_col"] < len(row) else 0.0
                stock = clean_number(row[b["stock_col"]]) if b["stock_col"] is not None and b["stock_col"] < len(row) else 0.0
                if issued == 0 and stock == 0:
                    continue
                normalized_rows.append({
                    **item_base,
                    "branch": b["branch"],
                    "issued_qty": issued,
                    "branch_stock": stock,
                })

    normalized_df = pd.DataFrame(normalized_rows)
    items_df = pd.DataFrame(item_rows)
    return normalized_df, items_df, import_log


def create_reports(normalized_df: pd.DataFrame, items_df: pd.DataFrame, import_log: Dict[str, Any]) -> Dict[str, pd.DataFrame]:
    reports: Dict[str, pd.DataFrame] = {}
    reports["01_Normalized"] = normalized_df
    reports["01_Items_Raw"] = items_df

    if items_df.empty:
        reports["02_Item_Summary"] = pd.DataFrame()
        return reports

    if normalized_df.empty:
        issued_by_item = pd.DataFrame(columns=["item_key", "total_issued_month", "branch_stock_sum"])
    else:
        issued_by_item = normalized_df.groupby("item_key", as_index=False).agg(
            total_issued_month=("issued_qty", "sum"),
            branch_stock_sum=("branch_stock", "sum"),
            active_branches=("branch", "nunique"),
        )

    item_summary = items_df.groupby("item_key", as_index=False).agg(
        month=("month", "first"),
        next_month=("next_month", "first"),
        source_sheet=("source_sheet", "first"),
        system_name=("system_name", "first"),
        scientific_name=("scientific_name", "first"),
        trade_names=("trade_name", lambda s: " | ".join(sorted({clean_text(x) for x in s if clean_text(x)})[:8])),
        unit=("unit", "first"),
        avg_monthly_year=("avg_monthly_year", "max"),
        avg_monthly_recent=("avg_monthly_recent", "max"),
        closing_stock_file=("end_stock_total", "sum"),
    )
    item_summary = item_summary.merge(issued_by_item, on="item_key", how="left")
    for col in ["total_issued_month", "branch_stock_sum", "active_branches"]:
        if col in item_summary.columns:
            item_summary[col] = item_summary[col].fillna(0)

    # Choose the safer current stock: if file total stock exists use it, otherwise use branch sum.
    item_summary["current_stock"] = np.where(item_summary["closing_stock_file"] > 0, item_summary["closing_stock_file"], item_summary["branch_stock_sum"])
    item_summary["avg_monthly_consumption"] = np.where(item_summary["avg_monthly_recent"] > 0, item_summary["avg_monthly_recent"], item_summary["avg_monthly_year"])
    item_summary["avg_monthly_consumption"] = np.where(item_summary["avg_monthly_consumption"] > 0, item_summary["avg_monthly_consumption"], item_summary["total_issued_month"])

    daily = item_summary["avg_monthly_consumption"] / 30.0
    item_summary["days_left"] = np.where(daily > 0, item_summary["current_stock"] / daily, np.nan)
    item_summary["dynamic_min"] = np.ceil(item_summary["avg_monthly_consumption"] * LEAD_TIME_DAYS / 30.0)
    item_summary["suggested_reorder_qty"] = np.ceil(np.maximum(0, item_summary["dynamic_min"] - item_summary["current_stock"]))

    def risk(row):
        stock = row.get("current_stock", 0)
        days = row.get("days_left", np.nan)
        dyn = row.get("dynamic_min", 0)
        if stock <= 0:
            return "Critical"
        if pd.notna(days) and days <= CRITICAL_DAYS:
            return "Critical"
        if stock < dyn:
            return "High"
        if pd.notna(days) and days <= HIGH_DAYS:
            return "High"
        if pd.notna(days) and days <= MODERATE_DAYS:
            return "Moderate"
        return "Low"

    item_summary["shortage_risk"] = item_summary.apply(risk, axis=1)
    item_summary["note"] = np.where(
        item_summary["avg_monthly_consumption"] <= 0,
        "لا يوجد متوسط استهلاك كاف؛ يحتاج مراجعة بشرية",
        ""
    )
    item_summary = item_summary.sort_values(["shortage_risk", "suggested_reorder_qty"], ascending=[True, False])
    reports["02_Item_Summary"] = item_summary

    reports["03_Critical_Items"] = item_summary[item_summary["shortage_risk"].isin(["Critical", "High"])].copy()

    if normalized_df.empty:
        reports["04_Branch_Summary"] = pd.DataFrame()
        reports["07_Next_Opening_By_Branch"] = pd.DataFrame()
    else:
        branch_summary = normalized_df.groupby("branch", as_index=False).agg(
            total_issued=("issued_qty", "sum"),
            total_stock=("branch_stock", "sum"),
            item_count=("item_key", "nunique"),
        )
        reports["04_Branch_Summary"] = branch_summary.sort_values("total_issued", ascending=False)

        next_branch = normalized_df.groupby(["next_month", "source_sheet", "item_key", "scientific_name", "trade_name", "unit", "branch"], as_index=False).agg(
            opening_stock=("branch_stock", "sum")
        )
        reports["07_Next_Opening_By_Branch"] = next_branch.sort_values(["source_sheet", "scientific_name", "branch"])

    # New medicines detection.
    current_master = item_summary[["item_key", "source_sheet", "scientific_name", "trade_names", "unit"]].drop_duplicates("item_key").copy()
    if MASTER_FILE.exists():
        try:
            master = pd.read_excel(MASTER_FILE)
            old_keys = set(master["item_key"].astype(str)) if "item_key" in master.columns else set()
        except Exception:
            master = pd.DataFrame()
            old_keys = set()
        new_meds = current_master[~current_master["item_key"].astype(str).isin(old_keys)].copy()
        updated_master = pd.concat([master, new_meds.assign(first_seen_month=import_log.get("month_label", ""), first_seen_at=datetime.now().isoformat(timespec="seconds"))], ignore_index=True)
    else:
        new_meds = current_master.copy()
        new_meds["first_run_note"] = "أول تشغيل؛ كل الأصناف تعتبر جديدة بالنسبة للماستر"
        updated_master = current_master.assign(first_seen_month=import_log.get("month_label", ""), first_seen_at=datetime.now().isoformat(timespec="seconds"))
    reports["05_New_Medicines"] = new_meds
    reports["_updated_master"] = updated_master.drop_duplicates("item_key") if not updated_master.empty else updated_master

    next_total = item_summary[["next_month", "source_sheet", "item_key", "system_name", "scientific_name", "trade_names", "unit", "current_stock"]].copy()
    next_total = next_total.rename(columns={"current_stock": "opening_stock_next_month"})
    reports["06_Next_Opening_Total"] = next_total

    # Month comparison against previous run.
    if LAST_ITEM_SUMMARY_FILE.exists():
        try:
            prev = pd.read_excel(LAST_ITEM_SUMMARY_FILE)
            cols = ["item_key", "current_stock", "total_issued_month", "shortage_risk"]
            prev_small = prev[[c for c in cols if c in prev.columns]].copy()
            prev_small = prev_small.rename(columns={
                "current_stock": "previous_stock",
                "total_issued_month": "previous_issued",
                "shortage_risk": "previous_risk",
            })
            comp = item_summary.merge(prev_small, on="item_key", how="left")
            comp["stock_change"] = comp["current_stock"] - comp["previous_stock"].fillna(0)
            comp["issued_change"] = comp["total_issued_month"] - comp["previous_issued"].fillna(0)
            reports["08_Month_Comparison"] = comp[[
                "item_key", "scientific_name", "trade_names", "unit", "current_stock", "previous_stock",
                "stock_change", "total_issued_month", "previous_issued", "issued_change", "shortage_risk", "previous_risk"
            ]]
        except Exception as e:
            reports["08_Month_Comparison"] = pd.DataFrame([{"error": str(e)}])
    else:
        reports["08_Month_Comparison"] = pd.DataFrame([{"note": "لا يوجد شهر سابق محفوظ للمقارنة. سيتم إنشاء المقارنة من التشغيل القادم."}])

    dashboard = pd.DataFrame([
        {"metric": "month", "value": import_log.get("month_label", "")},
        {"metric": "next_month", "value": import_log.get("next_month_label", "")},
        {"metric": "items_count", "value": int(item_summary["item_key"].nunique())},
        {"metric": "critical_items", "value": int((item_summary["shortage_risk"] == "Critical").sum())},
        {"metric": "high_risk_items", "value": int((item_summary["shortage_risk"] == "High").sum())},
        {"metric": "total_current_stock", "value": float(item_summary["current_stock"].sum())},
        {"metric": "total_issued_month", "value": float(item_summary["total_issued_month"].sum())},
        {"metric": "new_medicines", "value": int(len(reports["05_New_Medicines"]))},
    ])
    reports["00_Dashboard"] = dashboard
    reports["09_Import_Log"] = pd.DataFrame([import_log])
    return reports


def save_outputs(reports: Dict[str, pd.DataFrame], source_path: Path, import_log: Dict[str, Any]) -> Dict[str, Path]:
    OUTPUT_DIR.mkdir(exist_ok=True)
    ARCHIVE_DIR.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_stem = re.sub(r"[^a-zA-Z0-9\u0600-\u06FF_-]+", "_", source_path.stem)[:80]
    report_path = OUTPUT_DIR / f"monthly_reports_{safe_stem}_{timestamp}.xlsx"
    opening_path = OUTPUT_DIR / f"next_opening_balance_{safe_stem}_{timestamp}.xlsx"
    normalized_path = OUTPUT_DIR / f"normalized_{safe_stem}_{timestamp}.xlsx"
    log_path = OUTPUT_DIR / f"import_log_{safe_stem}_{timestamp}.json"

    # Keep original untouched: archive a copy only.
    archived_source = ARCHIVE_DIR / f"{timestamp}_{source_path.name}"
    try:
        shutil.copy2(source_path, archived_source)
    except Exception:
        archived_source = source_path

    # Full report workbook.
    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        for name, df in reports.items():
            if name.startswith("_"):
                continue
            sheet_name = name[:31]
            if isinstance(df, pd.DataFrame):
                df.to_excel(writer, index=False, sheet_name=sheet_name)

    # Separate files for frequent use.
    if "01_Normalized" in reports:
        reports["01_Normalized"].to_excel(normalized_path, index=False)
    with pd.ExcelWriter(opening_path, engine="openpyxl") as writer:
        if "06_Next_Opening_Total" in reports:
            reports["06_Next_Opening_Total"].to_excel(writer, index=False, sheet_name="Opening_Total")
        if "07_Next_Opening_By_Branch" in reports:
            reports["07_Next_Opening_By_Branch"].to_excel(writer, index=False, sheet_name="Opening_By_Branch")

    if "_updated_master" in reports:
        reports["_updated_master"].to_excel(MASTER_FILE, index=False)
    if "02_Item_Summary" in reports:
        reports["02_Item_Summary"].to_excel(LAST_ITEM_SUMMARY_FILE, index=False)
    if "01_Normalized" in reports:
        reports["01_Normalized"].to_excel(LAST_NORMALIZED_FILE, index=False)

    log_data = dict(import_log)
    log_data["outputs"] = {"report": str(report_path), "opening": str(opening_path), "normalized": str(normalized_path), "archived_source": str(archived_source)}
    log_path.write_text(json.dumps(log_data, ensure_ascii=False, indent=2), encoding="utf-8")

    return {"report": report_path, "opening": opening_path, "normalized": normalized_path, "log": log_path, "archived_source": archived_source}


def process_file(path: str | Path) -> Dict[str, Path]:
    source_path = Path(path).resolve()
    if not source_path.exists():
        raise FileNotFoundError(f"File not found: {source_path}")
    normalized_df, items_df, import_log = normalize_workbook(source_path)
    reports = create_reports(normalized_df, items_df, import_log)
    outputs = save_outputs(reports, source_path, import_log)
    return outputs


def choose_file_with_dialog() -> Optional[str]:
    # Standard library dialog: avoids requiring PySide6 for this tool.
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        path = filedialog.askopenfilename(
            title="Choose monthly consumption Excel file",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        root.destroy()
        return path or None
    except Exception:
        return None


def main() -> int:
    if len(sys.argv) >= 2:
        path = sys.argv[1]
    else:
        path = choose_file_with_dialog()
        if not path:
            print("No file selected.")
            return 1

    try:
        outputs = process_file(path)
    except Exception as e:
        print("ERROR:", e)
        return 2

    print("Monthly reports created successfully.")
    for k, v in outputs.items():
        print(f"{k}: {v}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
