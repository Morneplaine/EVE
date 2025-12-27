# EVE Online Manufacturing Database

A comprehensive database tool for EVE Online manufacturing and reprocessing analysis with live market prices from Jita.

## Features

- **Manufacturing Database**: All blueprints with materials, skills, and outputs in pivot format
- **Reprocessing Database**: All items with reprocessing outputs in pivot format
- **Live Market Prices**: Integrated with Fuzzwork Market API for Jita prices
- **Volume Data**: Includes packaged volumes for transportation planning
- **Smart Column Ordering**: Materials sorted by usage frequency (most used first)

## Quick Start

### First Time Setup

1. **Activate virtual environment:**
   ```powershell
   .\venv\Scripts\Activate.ps1
   ```

2. **Create database (with prices):**
   ```powershell
   python eve_manufacturing_database.py
   ```
   This will take 5-10 minutes as it downloads SDE data and fetches prices.

3. **Or create database without prices (faster):**
   ```powershell
   python eve_manufacturing_database.py --skip-prices
   ```
   Then add prices later using the update script.

### Updating Prices (Fast - Recommended)

Since prices change frequently but other data rarely changes, use the dedicated price updater:

1. **Close the Excel file** (if open)

2. **Update prices:**
   ```powershell
   python update_prices.py
   ```
   Or simply double-click: `update_prices.bat`

3. **Re-open the Excel file** - prices will automatically update via VLOOKUP formulas

This takes only 2-3 minutes vs 5-10 minutes for a full rebuild.

## Excel File Structure

### Manufacturing Sheet
- **Row 0**: Headers (material names, product info)
- **Row 1**: Buy prices (VLOOKUP formulas that pull from Prices sheet)
- **Row 2+**: Data rows (blueprints with material quantities)

### Reprocessing Sheet
- **Row 0**: Headers (material names, item info)
- **Row 1**: Buy prices (VLOOKUP formulas)
- **Row 2+**: Data rows (items with reprocessing outputs)

### Prices Sheet
- Contains all price data (Buy Max, Buy Volume, Sell Min, etc.)
- Used by VLOOKUP formulas in Manufacturing and Reprocessing sheets
- Update this sheet to refresh all prices

## Files

- `eve_manufacturing_database.py` - Main script (creates full database)
- `update_prices.py` - Price updater (updates only prices, fast)
- `update_prices.bat` - Windows batch file for easy price updates
- `requirements.txt` - Python dependencies

## Data Sources

- **Static Data**: Fuzzwork SDE (www.fuzzwork.co.uk)
- **Market Prices**: Fuzzwork Market API (market.fuzzwork.co.uk)
- **System**: Jita 4-4 (system ID: 30000142)

## Notes

- Material quantities are for ME 0 (unresearched blueprints)
- Reprocessing values assume 100% efficiency (actual: 50-60%)
- Prices are fetched from Jita market data
- VLOOKUP formulas in row 1 automatically update when Prices sheet is updated

