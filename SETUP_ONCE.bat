@echo off
cd /d "%~dp0"

echo =====================================
echo Creating venv environment...
echo =====================================

py -3.12 -m venv venv

echo =====================================
echo Installing required packages...
echo =====================================

venv\Scripts\python.exe -m pip install --upgrade pip setuptools wheel
venv\Scripts\python.exe -m pip install PySide6 pandas openpyxl numpy requests matplotlib Pillow plotly streamlit

echo =====================================
echo Setup finished.
echo You can now open START_PHARMAGUARD.bat
echo =====================================

pause