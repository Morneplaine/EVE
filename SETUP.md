# EVE Manufacturing Database - Setup Guide

## Virtual Environment Setup

A virtual environment has been created for this project. Follow these steps to activate it and install dependencies:

### Activate the Virtual Environment

**PowerShell:**
```powershell
.\venv\Scripts\Activate.ps1
```

**Command Prompt:**
```cmd
venv\Scripts\activate.bat
```

### Install Dependencies

Once the virtual environment is activated, install the required packages:

```bash
pip install -r requirements.txt
```

### Deactivate the Virtual Environment

When you're done working, you can deactivate the virtual environment:

```bash
deactivate
```

## Running the Script

After activating the virtual environment and installing dependencies, you can run the script:

```bash
python eve_manufacturing_database.py
```

## Market History (EVE Tycoon)

Historical trade data for Jita (The Forge, region id 44992) is fetched from [EVE Tycoon](https://evetycoon.com/docs):

- **Endpoint:** `GET /api/v1/market/history/{regionId}/{typeId}`
- **Response:** Array of daily records: `date` (Unix ms), `average`, `highest`, `lowest`, `orderCount`, `volume`.
- **History:** The API returns full history (many days) in one response. No need to accumulate over time.
- **Batching:** There is no batch endpoint; one request per `typeId`. The script fetches only reprocessable modules (same set as Top 30 analysis) to limit requests. Use `--types 34,35` for a custom list or `--limit 100` for testing.

Run:

```bash
python fetch_market_history.py [--limit N] [--delay 1] [--types 34,35,...]
```

Data is stored in `market_history_daily` with `transaction_skew = 1 - ((average - lowest) / (highest - lowest))` per day. Use SQL to get most recent, last 7 days, or last 30 days (e.g. `WHERE date_utc >= date('now', '-7 days')`).

## Dependencies

- pandas: Data manipulation and analysis
- openpyxl: Excel file reading/writing
- requests: HTTP library for API calls
- xlsxwriter: Excel file writing with formatting

