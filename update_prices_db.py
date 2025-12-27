"""
Update prices in SQLite database from Fuzzwork Market API
This is fast and only updates prices, leaving all other data unchanged.
"""

import sqlite3
import logging
from eve_manufacturing_database import get_fuzzwork_market_prices, JITA_SYSTEM_ID

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

DB_FILE = "eve_manufacturing.db"

def update_prices():
    """Update all prices in the database"""
    logger.info("=" * 60)
    logger.info("Updating prices in database")
    logger.info("=" * 60)
    
    conn = sqlite3.connect(DB_FILE)
    
    try:
        # Get all typeIDs that need prices
        cursor = conn.execute("SELECT typeID FROM prices")
        type_ids = [row[0] for row in cursor.fetchall()]
        
        logger.info(f"Found {len(type_ids)} items to update")
        
        # Fetch prices from Fuzzwork Market API
        logger.info("Fetching prices from Fuzzwork Market API...")
        fuzzwork_prices = get_fuzzwork_market_prices(type_ids, station_id=JITA_SYSTEM_ID)
        
        # Update database
        logger.info("Updating database...")
        updated_count = 0
        
        for type_id, price_data in fuzzwork_prices.items():
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
        
        conn.commit()
        
        logger.info("=" * 60)
        logger.info(f"SUCCESS! Updated prices for {updated_count} items")
        logger.info("=" * 60)
        
    except Exception as e:
        logger.error(f"Error updating prices: {e}")
        import traceback
        logger.error(traceback.format_exc())
        conn.rollback()
        raise
    finally:
        conn.close()

if __name__ == "__main__":
    update_prices()

