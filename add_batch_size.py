"""
Add batch_size column to reprocessing_outputs table and populate it based on item groups.
For charges and similar items, batch size is typically 100. For modules, it's 1.
"""

import sqlite3
import pandas as pd
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

DB_FILE = "eve_manufacturing.db"

# Item groups that typically have batch size of 100
BATCH_SIZE_100_GROUPS = [
    'Charge',  # All charge types
    'Projectile Ammo',
    'Hybrid Charge',
    'Frequency Crystal',
    'Capacitor Charge',
]

# Item groups that typically have batch size of 5000
BATCH_SIZE_5000_GROUPS = [
    'Missile',  # All missile types (Light Missile, Heavy Missile, Cruise Missile, etc.)
    'Rocket',
    'Torpedo',
]

def determine_batch_size(group_name):
    """Determine batch size based on group name"""
    # Check for 5000 batch size first (missiles, rockets, torpedoes)
    if any(keyword in group_name for keyword in BATCH_SIZE_5000_GROUPS):
        return 5000
    # Check for 100 batch size (charges, ammo)
    if any(keyword in group_name for keyword in BATCH_SIZE_100_GROUPS):
        return 100
    # Default to 1 for modules and other items
    return 1

def add_batch_size_column():
    """Add batch_size column and populate it"""
    conn = sqlite3.connect(DB_FILE)
    
    try:
        # Check if column already exists
        cursor = conn.execute("PRAGMA table_info(reprocessing_outputs)")
        columns = [row[1] for row in cursor.fetchall()]
        
        if 'batch_size' not in columns:
            logger.info("Adding batch_size column to reprocessing_outputs table...")
            conn.execute("ALTER TABLE reprocessing_outputs ADD COLUMN batch_size INTEGER DEFAULT 1")
            conn.commit()
            logger.info("Column added successfully")
        else:
            logger.info("batch_size column already exists")
        
        # Get item groups for all items in reprocessing_outputs
        logger.info("Determining batch sizes based on item groups...")
        
        query = """
        SELECT DISTINCT ro.itemTypeID, ro.itemName, i.groupID, g.groupName
        FROM reprocessing_outputs ro
        JOIN items i ON ro.itemTypeID = i.typeID
        JOIN groups g ON i.groupID = g.groupID
        """
        df = pd.read_sql_query(query, conn)
        
        # Update batch_size for each item
        updated = 0
        for _, row in df.iterrows():
            item_type_id = row['itemTypeID']
            group_name = row['groupName']
            batch_size = determine_batch_size(group_name)
            
            conn.execute(
                "UPDATE reprocessing_outputs SET batch_size = ? WHERE itemTypeID = ?",
                (batch_size, item_type_id)
            )
            updated += 1
        
        conn.commit()
        logger.info(f"Updated batch_size for {updated} unique items")
        
        # Show some examples
        logger.info("\nSample batch sizes:")
        sample_query = """
        SELECT DISTINCT ro.itemName, g.groupName, ro.batch_size
        FROM reprocessing_outputs ro
        JOIN items i ON ro.itemTypeID = i.typeID
        JOIN groups g ON i.groupID = g.groupID
        WHERE ro.itemName LIKE '%Charge%'
        LIMIT 5
        """
        sample_df = pd.read_sql_query(sample_query, conn)
        print(sample_df)
        
        logger.info("\n" + "=" * 60)
        logger.info("SUCCESS! Batch sizes have been added and populated")
        logger.info("=" * 60)
        
    except Exception as e:
        logger.error(f"Error: {e}")
        import traceback
        logger.error(traceback.format_exc())
        conn.rollback()
    finally:
        conn.close()

if __name__ == "__main__":
    add_batch_size_column()

