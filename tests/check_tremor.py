import sqlite3

conn = sqlite3.connect('eve_manufacturing.db')
cursor = conn.cursor()

cursor.execute("""
    SELECT itemTypeID, itemName, materialTypeID, materialName, quantity, batch_size 
    FROM reprocessing_outputs 
    WHERE itemName = 'Tremor L'
""")

rows = cursor.fetchall()
print("Tremor L reprocessing data:")
for row in rows:
    print(f"  Item: {row[1]}, Material: {row[3]}, Quantity: {row[4]}, Batch Size: {row[5]}")

# Check prices - need to join with items table
cursor.execute("""
    SELECT p.typeID, i.typeName, p.buy_max, p.sell_min 
    FROM prices p
    JOIN items i ON p.typeID = i.typeID
    WHERE i.typeName IN ('Tremor L', 'Morphite', 'Fernite Carbide', 'Fullerides')
""")
print("\nPrices:")
for row in cursor.fetchall():
    print(f"  {row[1]}: buy_max={row[2]}, sell_min={row[3]}")

conn.close()

