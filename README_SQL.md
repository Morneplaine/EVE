# EVE Manufacturing Database - SQL Edition

A powerful SQLite-based database system for EVE Online manufacturing analysis with advanced querying capabilities.

## Why SQL?

- **Fast Queries**: Find profitable items instantly with SQL
- **Advanced Calculations**: Filter by skills, resources, ME level, etc.
- **Easy Updates**: Update only prices without regenerating everything
- **Flexible Analysis**: Custom queries for any scenario
- **Static Excel Output**: Generate Excel files when needed

## Quick Start

### 1. Build the Database (One-time setup)

```powershell
python build_database.py
```

This downloads SDE data and populates the SQLite database. Takes ~5-10 minutes.

### 2. Update Prices (Fast - run daily)

```powershell
python update_prices_db.py
```

Updates only prices in ~2-3 minutes.

### 3. Import Your Character Skills

**Option A: CSV Import**
```powershell
python import_character_skills.py csv skills.csv
```

CSV format:
```csv
typeID,skillName,level
3300,Science,5
3301,Research,4
```

**Option B: ESI API** (requires OAuth token)
```powershell
python import_character_skills.py esi <access_token> <character_id>
```

**Option C: Manual Entry**
```powershell
python import_character_skills.py manual
```

### 4. Import Your Inventory/Resources

```powershell
python import_inventory.py inventory.csv
```

CSV format:
```csv
typeID,typeName,quantity
34,Tritanium,1000000
35,Pyerite,500000
```

### 5. Analyze Profitability

```powershell
# Basic analysis (all items, no filters)
python analyze_profitability.py

# With ME level 5, filter by skills, min profit 1000 ISK
python analyze_profitability.py 5 true false 1000

# With resources filter (only show items you can make)
python analyze_profitability.py 0 true true 0
```

Arguments:
- `me_level`: Material Efficiency level (0-10)
- `filter_skills`: true/false - only show items where you have skills
- `filter_resources`: true/false - only show items you can make with current inventory
- `min_profit`: Minimum profit per unit (ISK)

### 6. Generate Excel (Optional)

```powershell
python generate_excel.py
```

Creates a static Excel file from the database (same format as before).

## Workflow

### Daily Workflow
1. Update prices: `python update_prices_db.py`
2. Analyze profitability: `python analyze_profitability.py 5 true true 0`
3. Review results in `profitability_analysis.csv`

### Weekly Workflow
1. Update inventory: `python import_inventory.py inventory.csv`
2. Update skills if needed: `python import_character_skills.py csv skills.csv`
3. Re-analyze profitability

### Monthly Workflow
1. Rebuild database if SDE updated: `python build_database.py`
2. Update prices: `python update_prices_db.py`
3. Generate fresh Excel: `python generate_excel.py`

## Advanced Usage

### Custom SQL Queries

You can query the database directly:

```python
import sqlite3
import pandas as pd

conn = sqlite3.connect("eve_manufacturing.db")

# Find all blueprints that use Tritanium
query = """
    SELECT DISTINCT b.productName, mm.quantity
    FROM blueprints b
    JOIN manufacturing_materials mm ON b.blueprintTypeID = mm.blueprintTypeID
    WHERE mm.materialTypeID = 34
    ORDER BY mm.quantity DESC
"""

df = pd.read_sql_query(query, conn)
print(df)
```

### Example Queries

**Find items you can make with current inventory:**
```sql
SELECT 
    b.productName,
    b.outputQuantity,
    MIN(inv.quantity / mm.quantity) as max_units
FROM blueprints b
JOIN manufacturing_materials mm ON b.blueprintTypeID = mm.blueprintTypeID
LEFT JOIN inventory inv ON mm.materialTypeID = inv.typeID
GROUP BY b.blueprintTypeID, b.productName, b.outputQuantity
HAVING MIN(inv.quantity / mm.quantity) > 0
ORDER BY max_units DESC
```

**Find most profitable items (ignoring skills/resources):**
```sql
SELECT 
    b.productName,
    p.sell_min as product_price,
    SUM(mm.quantity * pm.sell_min) as material_cost,
    (p.sell_min - SUM(mm.quantity * pm.sell_min) / b.outputQuantity) as profit_per_unit
FROM blueprints b
JOIN prices p ON b.productTypeID = p.typeID
JOIN manufacturing_materials mm ON b.blueprintTypeID = mm.blueprintTypeID
JOIN prices pm ON mm.materialTypeID = pm.typeID
GROUP BY b.blueprintTypeID, b.productName, p.sell_min, b.outputQuantity
HAVING profit_per_unit > 0
ORDER BY profit_per_unit DESC
LIMIT 20
```

## Database Schema

- `items`: All EVE items with volumes
- `blueprints`: Manufacturing blueprints
- `manufacturing_materials`: Materials required per blueprint
- `manufacturing_skills`: Skills required per blueprint
- `reprocessing_outputs`: Reprocessing data
- `prices`: Market prices (updated frequently)
- `character_skills`: Your character's skills
- `inventory`: Your available resources

See `database_schema.sql` for full schema.

## Files

- `build_database.py` - Build database from SDE (one-time)
- `update_prices_db.py` - Update prices only (fast)
- `generate_excel.py` - Generate Excel from database
- `analyze_profitability.py` - Find profitable items
- `import_character_skills.py` - Import skills (CSV/ESI/manual)
- `import_inventory.py` - Import inventory/resources
- `database_schema.sql` - Database schema

## Tips

1. **Update prices daily** - Prices change frequently
2. **Keep inventory updated** - More accurate profitability analysis
3. **Use ME level in analysis** - Higher ME = lower material costs
4. **Filter by skills** - Only see items you can actually make
5. **Filter by resources** - Only see items you can afford to make

## Getting Character Skills from EVE

### Option 1: ESI API (Recommended)
1. Create EVE SSO application at https://developers.eveonline.com/
2. Get OAuth access token
3. Use `import_character_skills.py esi <token> <character_id>`

### Option 2: Export from EVE Client
Some third-party tools can export skills to CSV format.

### Option 3: Manual Entry
Use `import_character_skills.py manual` for interactive entry.

## Getting Inventory Data

### Option 1: Export from EVE Client
Use tools like EVE Inventory Manager or similar to export inventory to CSV.

### Option 2: Manual Entry
Create a CSV with your most important materials:
```csv
typeID,typeName,quantity
34,Tritanium,1000000
35,Pyerite,500000
```

### Option 3: ESI API
You can extend `import_inventory.py` to fetch from ESI API if needed.

