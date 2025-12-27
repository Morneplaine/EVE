"""
Update only mineral prices in SQLite database from Fuzzwork Market API
This script identifies minerals and updates only their prices, leaving all other prices unchanged.
"""

import sqlite3
import logging
from eve_manufacturing_database import get_fuzzwork_market_prices, JITA_SYSTEM_ID

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

DB_FILE = "eve_manufacturing.db"

# Basic minerals only (refined minerals, not ores)
BASIC_MINERALS = [
    'Tritanium',
    'Pyerite',
    'Mexallon',
    'Isogen',
    'Nocxium',
    'Zydrine',
    'Megacyte',
    'Morphite',
]

# Additional items to update (mutaplasmid residues and other materials)
ADDITIONAL_ITEMS = [
    'Armor Mutaplasmid Residue',
    'Astronautic Mutaplasmid Residue',
    'Crystalline Isogen-10',
    'Damage Control Mutaplasmid Residue',
    'Drone Mutaplasmid Residue',
    'Engineering Mutaplasmid Residue',
    'Large Mutaplasmid Residue',
    'Medium Mutaplasmid Residue',
    'Mutaplasmid Residue',
    'Shield Mutaplasmid Residue',
    'Small Mutaplasmid Residue',
    'Stasis Webifier Mutaplasmid Residue',
    'Warp Disruption Mutaplasmid Residue',
    'Weapon Upgrade Mutaplasmid Residue',
    'X-Large Mutaplasmid Residue',
    'Zero-Point Condensate',
]

def get_mineral_type_ids(conn):
    """
    Get all typeIDs for basic minerals and additional items.
    Both minerals and additional items are identified by exact item name.
    """
    # Combine basic minerals and additional items - both searched by item name
    all_items = BASIC_MINERALS + ADDITIONAL_ITEMS
    placeholders = ','.join(['?'] * len(all_items))
    
    query = f"""
        SELECT DISTINCT i.typeID, i.typeName
        FROM items i
        WHERE i.typeName IN ({placeholders})
        ORDER BY i.typeName
    """
    
    cursor = conn.execute(query, all_items)
    items = cursor.fetchall()
    return items

def update_mineral_prices():
    """Update only mineral prices in the database"""
    logger.info("=" * 60)
    logger.info("Updating MINERAL and MATERIAL prices in database")
    logger.info("=" * 60)
    
    conn = sqlite3.connect(DB_FILE)
    
    try:
        # Get all mineral typeIDs
        items = get_mineral_type_ids(conn)
        type_ids = [row[0] for row in items]
        item_names = {row[0]: row[1] for row in items}
        
        if len(type_ids) == 0:
            logger.warning("No minerals or items found in database!")
            return
        
        logger.info(f"Found {len(type_ids)} items to update:")
        for type_id, name in list(item_names.items())[:20]:  # Show first 20
            logger.info(f"  - {name} (TypeID: {type_id})")
        if len(item_names) > 20:
            logger.info(f"  ... and {len(item_names) - 20} more")
        
        # Check which minerals already have price entries
        placeholders = ','.join(['?'] * len(type_ids))
        existing_query = f"SELECT typeID FROM prices WHERE typeID IN ({placeholders})"
        existing_cursor = conn.execute(existing_query, type_ids)
        existing_type_ids = set(row[0] for row in existing_cursor.fetchall())
        
        # Create price entries for items that don't have them yet
        missing_type_ids = set(type_ids) - existing_type_ids
        if missing_type_ids:
            logger.info(f"Creating price entries for {len(missing_type_ids)} items without price data...")
            for type_id in missing_type_ids:
                conn.execute("""
                    INSERT INTO prices (typeID, buy_max, buy_volume, sell_min, sell_avg, sell_median, sell_volume)
                    VALUES (?, 0, 0, 0, 0, 0, 0)
                """, (type_id,))
            conn.commit()
        
        # Fetch prices from Fuzzwork Market API
        logger.info("Fetching prices from Fuzzwork Market API for Jita...")
        logger.info("This may take a minute...")
        fuzzwork_prices = get_fuzzwork_market_prices(type_ids, station_id=JITA_SYSTEM_ID)
        
        # Update database
        logger.info("Updating database...")
        updated_count = 0
        no_price_count = 0
        
        for type_id in type_ids:
            price_data = fuzzwork_prices.get(type_id, {})
            
            # Update price entry
            conn.execute("""
                UPDATE prices SET
                    buy_max = ?,
                    buy_volume = ?,
                    sell_min = ?,
                    sell_avg = ?,
                    sell_median = ?,
                    sell_volume = ?,
                    updated_at = CURRENT_TIMESTAMP
                WHERE typeID = ?
            """, (
                price_data.get('buy_max', 0),
                price_data.get('buy_volume', 0),
                price_data.get('sell_min', 0),
                price_data.get('sell_min', 0),  # Use sell_min as avg
                price_data.get('sell_min', 0),  # Use sell_min as median
                price_data.get('sell_volume', 0),
                type_id
            ))
            
            if price_data.get('buy_max', 0) > 0 or price_data.get('sell_min', 0) > 0:
                updated_count += 1
            else:
                no_price_count += 1
                item_name = item_names.get(type_id, f"TypeID {type_id}")
                logger.warning(f"No price data found for: {item_name}")
        
        conn.commit()
        
        logger.info("=" * 60)
        logger.info(f"SUCCESS! Updated prices for {updated_count} items")
        if no_price_count > 0:
            logger.info(f"Warning: {no_price_count} items had no market data")
        logger.info("=" * 60)
        
    except Exception as e:
        logger.error(f"Error updating mineral prices: {e}")
        import traceback
        logger.error(traceback.format_exc())
        conn.rollback()
        raise
    finally:
        conn.close()

if __name__ == "__main__":
    update_mineral_prices()

