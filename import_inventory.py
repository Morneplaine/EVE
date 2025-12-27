"""
Import inventory/resources into the database
Supports CSV import with columns: typeID, typeName, quantity
"""

import sqlite3
import pandas as pd
import logging
import sys

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

DB_FILE = "eve_manufacturing.db"

def import_from_csv(csv_file, replace=True):
    """Import inventory from CSV file
    
    CSV format should be:
    typeID,typeName,quantity
    34,Tritanium,1000000
    35,Pyerite,500000
    """
    logger.info(f"Importing inventory from CSV: {csv_file}")
    
    conn = sqlite3.connect(DB_FILE)
    
    try:
        # Read CSV
        inventory_df = pd.read_csv(csv_file)
        
        # Validate columns
        required_cols = ['typeID', 'quantity']
        if not all(col in inventory_df.columns for col in required_cols):
            raise ValueError(f"CSV must contain columns: {required_cols}")
        
        # If typeName not provided, try to get from items table
        if 'typeName' not in inventory_df.columns:
            items_query = "SELECT typeID, typeName FROM items WHERE typeID IN ({})".format(
                ','.join(['?'] * len(inventory_df))
            )
            items_df = pd.read_sql_query(items_query, conn, params=inventory_df['typeID'].tolist())
            inventory_df = inventory_df.merge(items_df, on='typeID', how='left')
            inventory_df['typeName'] = inventory_df['typeName'].fillna('Unknown')
        
        # Clear existing inventory if replace=True
        if replace:
            conn.execute("DELETE FROM inventory")
            logger.info("Cleared existing inventory")
        
        # Insert/update inventory
        for _, row in inventory_df.iterrows():
            conn.execute("""
                INSERT OR REPLACE INTO inventory (typeID, typeName, quantity)
                VALUES (?, ?, ?)
            """, (
                int(row['typeID']),
                str(row.get('typeName', 'Unknown')),
                int(row['quantity'])
            ))
        
        conn.commit()
        logger.info(f"Imported {len(inventory_df)} inventory items")
        
        # Show summary
        total_query = "SELECT SUM(quantity) as total FROM inventory"
        result = conn.execute(total_query).fetchone()
        logger.info(f"Total items in inventory: {result[0] if result[0] else 0}")
        
    except Exception as e:
        logger.error(f"Error importing from CSV: {e}")
        conn.rollback()
        raise
    finally:
        conn.close()

def main():
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python import_inventory.py <file.csv> [replace]")
        print("  replace: true/false (default: true) - whether to replace existing inventory")
        return
    
    csv_file = sys.argv[1]
    replace = sys.argv[2].lower() == 'true' if len(sys.argv) > 2 else True
    
    import_from_csv(csv_file, replace=replace)

if __name__ == "__main__":
    main()

