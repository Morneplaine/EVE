"""Check missile batch sizes"""

import sqlite3
import pandas as pd

conn = sqlite3.connect('eve_manufacturing.db')

# Check Inferno Precision Light Missile
query = """
SELECT 
    ro.itemName,
    ro.itemTypeID,
    g.groupName,
    ro.batch_size,
    ro.materialName,
    ro.quantity
FROM reprocessing_outputs ro
JOIN items i ON ro.itemTypeID = i.typeID
JOIN groups g ON i.groupID = g.groupID
WHERE ro.itemName LIKE '%Inferno Precision Light Missile%'
"""
df = pd.read_sql_query(query, conn)
print("Inferno Precision Light Missile:")
print(df)
print()

# Check what group missiles belong to
query2 = """
SELECT DISTINCT g.groupName, COUNT(*) as item_count
FROM reprocessing_outputs ro
JOIN items i ON ro.itemTypeID = i.typeID
JOIN groups g ON i.groupID = g.groupID
WHERE ro.itemName LIKE '%Missile%'
GROUP BY g.groupName
ORDER BY item_count DESC
LIMIT 10
"""
df2 = pd.read_sql_query(query2, conn)
print("Missile groups:")
print(df2)
print()

# Check a few missile items and their batch sizes
query3 = """
SELECT DISTINCT
    ro.itemName,
    g.groupName,
    ro.batch_size,
    SUM(ro.quantity) as total_materials
FROM reprocessing_outputs ro
JOIN items i ON ro.itemTypeID = i.typeID
JOIN groups g ON i.groupID = g.groupID
WHERE ro.itemName LIKE '%Missile%'
GROUP BY ro.itemName, g.groupName, ro.batch_size
ORDER BY ro.itemName
LIMIT 10
"""
df3 = pd.read_sql_query(query3, conn)
print("Sample missiles with batch sizes:")
print(df3)

conn.close()

