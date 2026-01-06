"""Check missile and plasma batch sizes"""

import sqlite3
import pandas as pd

conn = sqlite3.connect('eve_manufacturing.db')

# Check different missile types
print("=" * 60)
print("Checking Missile Groups and Batch Sizes")
print("=" * 60)

query = """
SELECT DISTINCT
    g.groupName,
    COUNT(DISTINCT ro.itemTypeID) as item_count,
    MIN(ro.batch_size) as min_batch_size,
    MAX(ro.batch_size) as max_batch_size
FROM reprocessing_outputs ro
JOIN items i ON ro.itemTypeID = i.typeID
JOIN groups g ON i.groupID = g.groupID
WHERE g.groupName LIKE '%Missile%'
GROUP BY g.groupName
ORDER BY item_count DESC
"""
df = pd.read_sql_query(query, conn)
print("\nMissile groups:")
print(df)

# Check specific missile items
print("\n" + "=" * 60)
print("Sample Missile Items:")
print("=" * 60)
query2 = """
SELECT DISTINCT
    ro.itemName,
    g.groupName,
    ro.batch_size,
    SUM(ro.quantity) as total_materials
FROM reprocessing_outputs ro
JOIN items i ON ro.itemTypeID = i.typeID
JOIN groups g ON i.groupID = g.groupID
WHERE g.groupName LIKE '%Missile%' 
  AND g.groupName NOT LIKE '%Launcher%'
GROUP BY ro.itemName, g.groupName, ro.batch_size
ORDER BY ro.itemName
LIMIT 15
"""
df2 = pd.read_sql_query(query2, conn)
print(df2)

# Check plasma items
print("\n" + "=" * 60)
print("Checking Plasma Items:")
print("=" * 60)
query3 = """
SELECT DISTINCT
    ro.itemName,
    g.groupName,
    ro.batch_size,
    SUM(ro.quantity) as total_materials
FROM reprocessing_outputs ro
JOIN items i ON ro.itemTypeID = i.typeID
JOIN groups g ON i.groupID = g.groupID
WHERE ro.itemName LIKE '%Plasma%'
GROUP BY ro.itemName, g.groupName, ro.batch_size
ORDER BY ro.itemName
LIMIT 15
"""
df3 = pd.read_sql_query(query3, conn)
print(df3)

# Check plasma groups
print("\n" + "=" * 60)
print("Plasma Groups:")
print("=" * 60)
query4 = """
SELECT DISTINCT
    g.groupName,
    COUNT(DISTINCT ro.itemTypeID) as item_count,
    MIN(ro.batch_size) as min_batch_size,
    MAX(ro.batch_size) as max_batch_size
FROM reprocessing_outputs ro
JOIN items i ON ro.itemTypeID = i.typeID
JOIN groups g ON i.groupID = g.groupID
WHERE g.groupName LIKE '%Plasma%'
GROUP BY g.groupName
ORDER BY item_count DESC
"""
df4 = pd.read_sql_query(query4, conn)
print(df4)

# Check Inferno Precision Light Missile specifically
print("\n" + "=" * 60)
print("Inferno Precision Light Missile Details:")
print("=" * 60)
query5 = """
SELECT 
    ro.itemName,
    g.groupName,
    ro.batch_size,
    ro.materialName,
    ro.quantity
FROM reprocessing_outputs ro
JOIN items i ON ro.itemTypeID = i.typeID
JOIN groups g ON i.groupID = g.groupID
WHERE ro.itemName = 'Inferno Precision Light Missile'
"""
df5 = pd.read_sql_query(query5, conn)
print(df5)

conn.close()


