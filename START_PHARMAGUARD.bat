@echo off
cd /d "%~dp0"

if exist "venv\Scripts\python.exe" (
    "venv\Scripts\python.exe" run_pharmaguard.py
    pause
    exit /b
)

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" run_pharmaguard.py
    pause
    exit /b
)

echo Python environment not found.
echo Please run SETUP_ONCE.bat first.
pause