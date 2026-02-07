@echo off
REM Activate venv and install requirements (run from project root)
cd /d "%~dp0"

if not exist "venv\Scripts\activate.bat" (
    echo Creating virtual environment...
    python -m venv venv
    if errorlevel 1 (
        echo Failed to create venv. Ensure Python is installed and on PATH.
        pause
        exit /b 1
    )
)

echo Activating venv...
call venv\Scripts\activate.bat

echo Installing requirements...
pip install -r requirements.txt
if errorlevel 1 (
    echo pip install failed.
    pause
    exit /b 1
)

echo Done. Venv is active in this window. Run: python fetch_market_history.py --test
pause
