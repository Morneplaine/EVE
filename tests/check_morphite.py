import sqlite3

conn = sqlite3.connect('eve_manufacturing.db')
cursor = conn.cursor()

cursor.execute("""
    SELECT itemName, materialName, quantity, batch_size 
    FROM reprocessing_outputs 
    WHERE itemName = 'Tremor L' AND materialName = 'Morphite'
""")

row = cursor.fetchone()
if row:
    print(f"Tremor L - Morphite:")
    print(f"  Quantity: {row[2]}")
    print(f"  Batch Size: {row[3]}")
    print(f"\nCalculation:")
    print(f"  Per item: {row[2]} / {row[3]} = {row[2] / row[3]}")
    print(f"  Per module (55% yield): {row[2] / row[3] * 0.55}")
    print(f"  For 5000 modules: {row[2] / row[3] * 0.55 * 5000}")
    print(f"  Rounded: {int(row[2] / row[3] * 0.55 * 5000)}")
    print(f"\nUser expects: 30")
    print(f"  Which would be: batch_quantity * yield = {row[2]} * 0.55 = {row[2] * 0.55}")

conn.close()

