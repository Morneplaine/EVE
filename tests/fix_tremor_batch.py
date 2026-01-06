import sqlite3

conn = sqlite3.connect('eve_manufacturing.db')
cursor = conn.cursor()

# Update Tremor L batch_size to 100
cursor.execute("UPDATE reprocessing_outputs SET batch_size = 100 WHERE itemName = 'Tremor L'")
conn.commit()

print("Updated Tremor L batch_size to 100")
cursor.execute("SELECT itemName, materialName, quantity, batch_size FROM reprocessing_outputs WHERE itemName = 'Tremor L'")
for row in cursor.fetchall():
    print(f"  {row[1]}: quantity={row[2]}, batch_size={row[3]}")

conn.close()


