@echo off
cd /d "%~dp0"
echo Current folder:
echo %CD%
echo.
echo Files:
dir /b
echo.
echo Checking Python:
python --version
echo.
echo Checking run_pharmaguard.py:
if exist run_pharmaguard.py (
    echo FOUND run_pharmaguard.py
) else (
    echo NOT FOUND run_pharmaguard.py
)
echo.
pause
