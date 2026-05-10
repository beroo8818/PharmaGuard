@echo off
setlocal
cd /d "%~dp0"

echo ==========================================
echo PharmaGuard startup diagnostic
echo Project folder: %CD%
echo ==========================================
echo.

if not exist "run_pharmaguard.py" (
    echo ERROR: run_pharmaguard.py was not found in this folder.
    echo You are probably running this BAT from the wrong folder,
    echo or the ZIP was not fully extracted.
    echo.
    echo Current folder:
    echo %CD%
    echo.
    echo Files here:
    dir /b
    echo.
    pause
    exit /b 1
)

where python >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python was not found.
    echo Install Python 3.11 or 3.12 and tick "Add Python to PATH".
    pause
    exit /b 1
)

if not exist "venv\Scripts\python.exe" (
    echo Creating virtual environment...
    python -m venv venv
    if errorlevel 1 (
        echo ERROR: Failed to create venv.
        pause
        exit /b 1
    )
)

call "venv\Scripts\activate"

if not exist ".installed" (
    echo Installing requirements for the first time...
    python -m pip install --upgrade pip
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo ERROR: Failed to install requirements.
        pause
        exit /b 1
    )
    echo installed > .installed
) else (
    echo Requirements already installed. Skipping install.
)

echo.
echo Starting PharmaGuard...
python run_pharmaguard.py

if errorlevel 1 (
    echo.
    echo ERROR: PharmaGuard crashed. See the message above.
    pause
    exit /b 1
)

pause
