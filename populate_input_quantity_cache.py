"""
Populate input_quantity_cache table for all items in the database.

This script will:
1. Create the input_quantity_cache table if it doesn't exist
2. For each item in the database, determine its input_quantity using:
   - Blueprints table (if available)
   - Group-based lookup (if no blueprint)
   - Default to 1 (if no group data)
3. Cache all results for fast future lookups
4. Mark items that need manual review
"""

import sqlite3
import pandas as pd
import logging
from pathlib import Path
from calculate_reprocessing_value import (
    ensure_input_quantity_cache_table,
    get_input_quantity
)

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

DB_FILE = "eve_manufacturing.db"


def populate_input_quantity_cache():
    """Populate input_quantity_cache for all items in the database"""
    logger.info("=" * 60)
    logger.info("Populating Input Quantity Cache")
    logger.info("=" * 60)
    
    if not Path(DB_FILE).exists():
        logger.error(f"Database file not found: {DB_FILE}")
        logger.error("Please run build_database.py first to create the database.")
        return False
    
    conn = sqlite3.connect(DB_FILE)
    
    try:
        # Ensure cache table exists
        ensure_input_quantity_cache_table(conn)
        
        # Get all items from the database
        logger.info("Fetching all items from database...")
        items_query = "SELECT typeID, typeName FROM items ORDER BY typeID"
        items_df = pd.read_sql_query(items_query, conn)
        
        total_items = len(items_df)
        logger.info(f"Found {total_items} items to process")
        logger.info("This may take several minutes...")
        logger.info("")
        
        # Statistics
        stats = {
            'blueprint': 0,
            'group_consensus': 0,
            'group_most_frequent': 0,
            'default': 0,
            'needs_review': 0,
            'cached': 0,
            'processed': 0
        }
        
        # Process each item from the items table (one time only)
        processed_type_ids = set()  # Track processed items to avoid duplicates
        
        for idx, row in items_df.iterrows():
            type_id = int(row['typeID'])
            type_name = row['typeName']
            
            # Skip if already processed (shouldn't happen, but safety check)
            if type_id in processed_type_ids:
                logger.warning(f"Skipping duplicate typeID: {type_id} ({type_name})")
                continue
            
            processed_type_ids.add(type_id)
            
            # Get input quantity (this will cache it automatically)
            # This function should only process items that exist in the items table
            input_quantity, source, needs_review = get_input_quantity(conn, type_id)
            
            # Update statistics
            stats[source] = stats.get(source, 0) + 1
            if needs_review:
                stats['needs_review'] += 1
            
            stats['processed'] += 1
            
            # Progress update every 1000 items
            if (idx + 1) % 1000 == 0:
                logger.info(f"Processed {idx + 1}/{total_items} items... "
                          f"(Blueprint: {stats['blueprint']}, "
                          f"Consensus: {stats['group_consensus']}, "
                          f"Most Frequent: {stats['group_most_frequent']}, "
                          f"Default: {stats['default']}, "
                          f"Needs Review: {stats['needs_review']})")
        
        logger.info("")
        logger.info("=" * 60)
        logger.info("Cache Population Complete!")
        logger.info("=" * 60)
        logger.info(f"Total items processed: {stats['processed']}")
        logger.info(f"  - From blueprints: {stats['blueprint']}")
        logger.info(f"  - From group consensus: {stats['group_consensus']}")
        logger.info(f"  - From group most frequent: {stats['group_most_frequent']}")
        logger.info(f"  - Default (1): {stats['default']}")
        logger.info(f"  - Items needing review: {stats['needs_review']}")
        logger.info("")
        
        # Verify we processed the correct number of items
        cache_count_query = "SELECT COUNT(*) as count FROM input_quantity_cache"
        cache_count_df = pd.read_sql_query(cache_count_query, conn)
        cache_count = cache_count_df.iloc[0]['count'] if len(cache_count_df) > 0 else 0
        
        logger.info(f"Cache table now contains {cache_count} entries")
        if cache_count != total_items:
            logger.warning(f"WARNING: Processed {total_items} items but cache has {cache_count} entries!")
            logger.warning("This might indicate duplicates or items not in the items table")
        
        # Show items that need review
        review_query = """
            SELECT typeID, typeName, input_quantity, source
            FROM input_quantity_cache
            WHERE needs_review = 1
            ORDER BY source, typeName
            LIMIT 50
        """
        review_df = pd.read_sql_query(review_query, conn)
        
        if len(review_df) > 0:
            total_needing_review = pd.read_sql_query(
                "SELECT COUNT(*) as count FROM input_quantity_cache WHERE needs_review = 1", 
                conn
            ).iloc[0]['count']
            logger.info(f"Sample items needing review (showing first 50 of {total_needing_review}):")
            logger.info("-" * 60)
            for _, row in review_df.iterrows():
                logger.info(f"  {row['typeName']:<40} (ID: {row['typeID']:>6}) - "
                          f"Qty: {row['input_quantity']:>4} - Source: {row['source']}")
            if len(review_df) == 50:
                logger.info(f"  ... and more (query the database for full list)")
        
        logger.info("=" * 60)
        return True
        
    except Exception as e:
        logger.error(f"Error populating cache: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return False
    
    finally:
        conn.close()


if __name__ == "__main__":
    success = populate_input_quantity_cache()
    exit(0 if success else 1)
