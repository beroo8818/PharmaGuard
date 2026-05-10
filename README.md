# PharmaGuard

PharmaGuard is a pharmacy inventory and supply-chain decision-support system designed for government hospital pharmacies.

## Quick Start

### Option 1: Windows one-click run

Double-click:

```bat
START_PHARMAGUARD.bat
```

The batch file will create a local virtual environment, install requirements, and run the main launcher.

### Option 2: Manual run

```bash
python -m venv venv
venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python run_pharmaguard.py
```

## Main Launcher

Run:

```bash
python run_pharmaguard.py
```

The launcher opens:

- Import Opening Balance Excel / CSV
- Live Sync Excel / CSV
- Live Sync Google Sheet
- Analytics & Forecasting Module
- Operational Inventory Module
- Reports Hub

## Main Modules

### Analytics & Forecasting

File:

```bash
python pharmaguard_v11_forecast_database.py
```

Includes import, forecasting, shortage/overstock analysis, purchase suggestions, redistribution suggestions, FEFO/expiry analysis, and executive dashboards.

### Operational Inventory

File:

```bash
python run_operational_launcher.py
```

Includes stock viewer, movement entry, batch and expiry management, minimum stock, FEFO helper, operational forecast, purchase orders, supplier deliveries, audit viewer, users, backup/restore, and reports.

## Data Notes

The system supports raw government hospital pharmacy Excel files with Arabic sheets and branch-level `منصرف / رصيد` columns.

If imported data does not contain batch number or expiry date, the system treats it as legacy opening balance data and uses:

- Batch Number: `OPENING-BALANCE`
- Expiry Date: `2099-12-31`

This is a placeholder for old imported data, not a real expiry date.

## Requirements

The main GUI uses PySide6. Install dependencies with:

```bash
python -m pip install -r requirements.txt
```

Recommended Python version: 3.11 or 3.12.
