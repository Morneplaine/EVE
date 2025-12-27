# Migration Guide: Excel to SQL Database

## Overview

The project now supports both approaches:

1. **Original Excel-based approach** (`eve_manufacturing_database.py`)
   - Simple, single Excel file
   - Good for basic analysis
   - Limited querying capabilities

2. **New SQL-based approach** (recommended for advanced users)
   - SQLite database for fast queries
   - Advanced profitability analysis
   - Skill and resource filtering
   - Still generates Excel when needed

## Why Switch to SQL?

✅ **Fast Queries**: Find profitable items in seconds
✅ **Advanced Filtering**: Filter by skills, resources, ME level
✅ **Easy Updates**: Update only prices (2-3 min vs 5-10 min)
✅ **Custom Analysis**: Write your own SQL queries
✅ **Still Get Excel**: Generate static Excel files when needed

## Migration Steps

### Step 1: Build the Database

```powershell
python build_database.py
```

This is a one-time setup that takes ~5-10 minutes. It:
- Downloads SDE data
- Populates SQLite database
- Creates all tables and relationships

### Step 2: Update Prices

```powershell
python update_prices_db.py
```

Updates only prices in ~2-3 minutes (vs 5-10 min for full rebuild).

### Step 3: Import Your Data

**Import Skills:**
```powershell
# Option 1: CSV (easiest)
python import_character_skills.py csv example_skills.csv

# Option 2: Manual entry
python import_character_skills.py manual

# Option 3: ESI API (if you have access token)
python import_character_skills.py esi <token> <character_id>
```

**Import Inventory:**
```powershell
python import_inventory.py example_inventory.csv
```

### Step 4: Analyze Profitability

```powershell
# Basic analysis
python analyze_profitability.py

# With filters: ME level 5, filter by skills, min profit 1000 ISK
python analyze_profitability.py 5 true false 1000

# Only show items you can actually make (skills + resources)
python analyze_profitability.py 0 true true 0
```

Results are saved to `profitability_analysis.csv`.

### Step 5: Generate Excel (Optional)

If you still want an Excel file:

```powershell
python generate_excel.py
```

This creates the same Excel format as before, but from the database.

## Daily Workflow

1. **Update prices** (2-3 min):
   ```powershell
   python update_prices_db.py
   ```

2. **Analyze profitability** (instant):
   ```powershell
   python analyze_profitability.py 5 true true 0
   ```

3. **Review results** in `profitability_analysis.csv`

## Weekly Workflow

1. **Update inventory** when you get new materials:
   ```powershell
   python import_inventory.py inventory.csv
   ```

2. **Re-analyze** to see new opportunities

## Monthly Workflow

1. **Rebuild database** if SDE updated (rare):
   ```powershell
   python build_database.py
   ```

2. **Update prices**:
   ```powershell
   python update_prices_db.py
   ```

3. **Generate fresh Excel** if needed:
   ```powershell
   python generate_excel.py
   ```

## Comparison

| Feature | Excel Approach | SQL Approach |
|---------|---------------|--------------|
| Setup Time | 5-10 min | 5-10 min (one-time) |
| Price Update | 5-10 min | 2-3 min |
| Query Speed | Slow (Excel formulas) | Fast (SQL) |
| Skill Filtering | Manual | Automatic |
| Resource Filtering | Manual | Automatic |
| ME Level Calculation | Manual | Automatic |
| Custom Queries | Limited | Full SQL |
| Excel Output | Always | On-demand |

## Example: Finding Best Items to Make

**With Excel:**
1. Open Excel file
2. Manually check each row
3. Calculate profit manually
4. Filter by skills manually
5. Check resources manually
6. Sort by profit

**With SQL:**
```powershell
python analyze_profitability.py 5 true true 0
```

Done! Results in `profitability_analysis.csv` sorted by profit.

## Getting Your Character Data

### Skills

**Option 1: ESI API** (most accurate)
1. Create app at https://developers.eveonline.com/
2. Get OAuth token
3. Use `import_character_skills.py esi <token> <char_id>`

**Option 2: CSV Export**
- Use third-party tools to export skills
- Format: `typeID,skillName,level`
- See `example_skills.csv` for format

**Option 3: Manual Entry**
- Use `import_character_skills.py manual`
- Enter skills interactively

### Inventory

**Option 1: CSV Export**
- Export from EVE client tools
- Format: `typeID,typeName,quantity`
- See `example_inventory.csv` for format

**Option 2: Manual Entry**
- Create CSV with your main materials
- Focus on materials you use most

## Tips

1. **Start Simple**: Build database, update prices, analyze
2. **Add Skills Later**: You can analyze without skills first
3. **Update Inventory Regularly**: More accurate = better decisions
4. **Use ME Level**: Higher ME = lower costs = more profit
5. **Filter Resources**: Only see items you can actually make

## Troubleshooting

**Database not found:**
- Run `python build_database.py` first

**Prices are zero:**
- Run `python update_prices_db.py`

**No profitable items:**
- Check if prices are updated
- Lower `min_profit` threshold
- Disable skill/resource filters temporarily

**Skills not importing:**
- Check CSV format matches `example_skills.csv`
- Ensure typeIDs are correct
- Try manual entry mode

## Next Steps

1. Build database: `python build_database.py`
2. Update prices: `python update_prices_db.py`
3. Try analysis: `python analyze_profitability.py`
4. Import your data when ready
5. Generate Excel if needed: `python generate_excel.py`

Enjoy your new powerful manufacturing analysis tool! 🚀

