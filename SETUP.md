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
- **Batching:** There is no batch endpoint; one request per `typeId`. With `--all-items`, types are limited to the **same set as Update All Prices** (types in `prices` table; ~50k). Use `--scope blueprint_consensus_mineral` for a smaller set (blueprint + group_consensus + mineral only; ~12k). Use `--types 34,35` for a custom list or `--limit 100` for testing.

**Test on one item first (recommended):**

```bash
python fetch_market_history.py --types 34 --delay 0.5
```

Then check the DB: `SELECT * FROM market_history_daily WHERE type_id = 34 LIMIT 10;` — you should see `type_name`, `date_utc`, `average`, `highest`, `lowest`, `order_count`, `volume`. Type 34 is Tritanium; one request populates many days.

**Then run for more items:**

```bash
python fetch_market_history.py [--limit N] [--delay 1] [--types 34,35,...]
# or for all types that have prices (same as Update All Prices; ~50k types):
python fetch_market_history.py --all-items --delay 1
# or only blueprint + group_consensus + mineral (~12k types):
python fetch_market_history.py --all-items --scope blueprint_consensus_mineral --delay 1
```

Data is stored in `market_history_daily` with item name in `type_name`. To get transaction skew when needed: `1 - ((average - lowest) / (highest - lowest))` (when highest ≠ lowest). Use SQL for most recent, last 7 days, or last 30 days (e.g. `WHERE date_utc >= date('now', '-7 days')`).

## Syncing the database to git

The database `eve_manufacturing.db` is tracked in git. Only the **last 4 versions** are kept (current + 3 backups). When you push:

1. Run: `python push_db_to_git.py`
2. The script rotates backups (current → v1, v1 → v2, v2 → v3), commits the 4 files, and pushes. The oldest backup (v3) is overwritten so you never have more than 4 versions.
3. On another computer, `git pull` to get the latest DB; the app uses `eve_manufacturing.db` (current).

## Dependencies

- pandas: Data manipulation and analysis
- openpyxl: Excel file reading/writing
- requests: HTTP library for API calls
- xlsxwriter: Excel file writing with formatting

