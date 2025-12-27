@echo off
REM Quick script to update prices in Excel file
cd /d "%~dp0"
call venv\Scripts\activate.bat
python update_prices.py
pause

