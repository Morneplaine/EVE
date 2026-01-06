"""
Check the structure of reprocessing data to find batch size information
"""

import pandas as pd
import sqlite3
from pathlib import Path

# Check database
print("=" * 60)
print("Checking database for reprocessing data")
print("=" * 60)

conn = sqlite3.connect('eve_manufacturing.db')

# Check Thorium Charge S
query = """
SELECT * FROM reprocessing_outputs 
WHERE itemName = 'Thorium Charge S'
"""
df = pd.read_sql_query(query, conn)
print("\nThorium Charge S reprocessing outputs:")
print(df)
print(f"\nTotal materials: {len(df)}")

# Check Iron Charge S
query2 = """
SELECT * FROM reprocessing_outputs 
WHERE itemName = 'Iron Charge S'
"""
df2 = pd.read_sql_query(query2, conn)
print("\n\nIron Charge S reprocessing outputs:")
print(df2)

# Check a few more items to see patterns
query3 = """
SELECT itemName, COUNT(*) as material_count, SUM(quantity) as total_quantity
FROM reprocessing_outputs
WHERE itemName LIKE '%Charge%'
GROUP BY itemName
ORDER BY itemName
LIMIT 10
"""
df3 = pd.read_sql_query(query3, conn)
print("\n\nCharge items summary:")
print(df3)

conn.close()

# Now check the raw SDE file structure
print("\n\n" + "=" * 60)
print("Checking raw SDE file structure")
print("=" * 60)

data_dir = Path('eve_data')
if data_dir.exists():
    inv_type_materials_file = data_dir / 'invTypeMaterials.csv'
    if inv_type_materials_file.exists():
        print(f"\nReading {inv_type_materials_file}")
        df_raw = pd.read_csv(inv_type_materials_file)
        print(f"\nColumns in invTypeMaterials.csv:")
        print(df_raw.columns.tolist())
        print(f"\nFirst few rows:")
        print(df_raw.head(10))
        
        # Check for Thorium Charge S (need to find typeID first)
        print("\n\nChecking for Thorium Charge S in raw data...")
        inv_types_file = data_dir / 'invTypes.csv'
        if inv_types_file.exists():
            inv_types = pd.read_csv(inv_types_file)
            thorium_charge = inv_types[inv_types['typeName'] == 'Thorium Charge S']
            if len(thorium_charge) > 0:
                type_id = thorium_charge.iloc[0]['typeID']
                print(f"Thorium Charge S typeID: {type_id}")
                thorium_materials = df_raw[df_raw['typeID'] == type_id]
                print(f"\nRaw reprocessing data for Thorium Charge S:")
                print(thorium_materials)
    else:
        print(f"File not found: {inv_type_materials_file}")
else:
    print(f"Data directory not found: {data_dir}")

