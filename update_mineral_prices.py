"""
Update only mineral prices in SQLite database from Fuzzwork Market API
This script identifies minerals and updates only their prices, leaving all other prices unchanged.
"""

import sqlite3
import logging
from eve_manufacturing_database import get_fuzzwork_market_prices, JITA_SYSTEM_ID
from decryptors_data import get_decryptor_type_ids

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


def _read_prices(conn, type_ids):
    """Return {typeID: {'buy': buy_max, 'sell': sell_min}} for the given type IDs."""
    if not type_ids:
        return {}
    placeholders = ','.join(['?'] * len(type_ids))
    cur = conn.execute(
        f"SELECT typeID, buy_max, sell_min FROM prices WHERE typeID IN ({placeholders})",
        list(type_ids),
    )
    return {row[0]: {'buy': float(row[1] or 0), 'sell': float(row[2] or 0)} for row in cur.fetchall()}


def update_mineral_prices(extra_type_ids=None):
    """Update mineral/material prices and return a before/after comparison.

    Parameters
    ----------
    extra_type_ids : iterable of int, optional
        Additional typeIDs to update (e.g. shopping-list products and materials).

    Returns
    -------
    dict
        {typeID: {'name': str, 'is_mineral': bool, 'old_buy': float, 'new_buy': float,
                  'old_sell': float, 'new_sell': float}}
        Ordered so that the 8 basic minerals come first, then extra items, then the rest.
    """
    logger.info("=" * 60)
    logger.info("Updating MINERAL and MATERIAL prices in database")
    logger.info("=" * 60)
    
    conn = sqlite3.connect(DB_FILE)
    
    try:
        # ---------- core mineral + additional items ----------
        items = get_mineral_type_ids(conn)
        type_ids = [row[0] for row in items]
        item_names = {row[0]: row[1] for row in items}
        mineral_set = {row[0] for row in items if row[1] in BASIC_MINERALS}

        # ---------- extra type IDs (shopping list) ----------
        extra_set = set()
        if extra_type_ids:
            extra_set = set(int(t) for t in extra_type_ids if t) - set(type_ids)
            if extra_set:
                placeholders_ex = ','.join(['?'] * len(extra_set))
                cur_ex = conn.execute(
                    f"SELECT typeID, typeName FROM items WHERE typeID IN ({placeholders_ex})",
                    list(extra_set),
                )
                for row in cur_ex.fetchall():
                    item_names[row[0]] = row[1]
                type_ids = type_ids + list(extra_set)

        # ---------- decryptors ----------
        decryptor_type_ids = get_decryptor_type_ids()
        if decryptor_type_ids:
            placeholders_dec = ",".join(["?"] * len(decryptor_type_ids))
            cur_dec = conn.execute(
                f"SELECT typeID FROM prices WHERE typeID IN ({placeholders_dec})",
                decryptor_type_ids,
            )
            existing_dec = {row[0] for row in cur_dec.fetchall()}
            missing_dec = set(decryptor_type_ids) - existing_dec
            for tid in missing_dec:
                conn.execute(
                    """
                    INSERT INTO prices (typeID, buy_max, buy_volume, sell_min, sell_avg, sell_median, sell_volume)
                    VALUES (?, 0, 0, 0, 0, 0, 0)
                    """,
                    (tid,),
                )
            conn.commit()
            type_ids = type_ids + decryptor_type_ids
        
        if len(type_ids) == 0:
            logger.warning("No minerals or items found in database!")
            return {}
        
        logger.info(f"Found {len(type_ids)} items to update:")
        for type_id, name in list(item_names.items())[:20]:
            logger.info(f"  - {name} (TypeID: {type_id})")
        if len(item_names) > 20:
            logger.info(f"  ... and {len(item_names) - 20} more")
        
        # ---------- ensure price rows exist ----------
        placeholders = ','.join(['?'] * len(type_ids))
        existing_query = f"SELECT typeID FROM prices WHERE typeID IN ({placeholders})"
        existing_cursor = conn.execute(existing_query, type_ids)
        existing_type_ids = set(row[0] for row in existing_cursor.fetchall())
        
        missing_type_ids = set(type_ids) - existing_type_ids
        if missing_type_ids:
            logger.info(f"Creating price entries for {len(missing_type_ids)} items without price data...")
            for type_id in missing_type_ids:
                conn.execute("""
                    INSERT INTO prices (typeID, buy_max, buy_volume, sell_min, sell_avg, sell_median, sell_volume)
                    VALUES (?, 0, 0, 0, 0, 0, 0)
                """, (type_id,))
            conn.commit()
        
        # ---------- read BEFORE prices ----------
        old_prices = _read_prices(conn, type_ids)

        # ---------- fetch from Fuzzwork ----------
        logger.info("Fetching prices from Fuzzwork Market API for Jita...")
        logger.info("This may take a minute...")
        fuzzwork_prices = get_fuzzwork_market_prices(type_ids, station_id=JITA_SYSTEM_ID)
        
        # ---------- update database ----------
        logger.info("Updating database...")
        updated_count = 0
        no_price_count = 0
        
        for type_id in type_ids:
            price_data = fuzzwork_prices.get(type_id, {})
            
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
                price_data.get('sell_min', 0),
                price_data.get('sell_min', 0),
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

        # ---------- read AFTER prices ----------
        new_prices = _read_prices(conn, type_ids)
        
        logger.info("=" * 60)
        logger.info(f"SUCCESS! Updated prices for {updated_count} items")
        if no_price_count > 0:
            logger.info(f"Warning: {no_price_count} items had no market data")
        logger.info("=" * 60)

        # ---------- build comparison result ----------
        # Order: basic minerals first (in BASIC_MINERALS order), then extra shopping-list items, then rest
        mineral_order = {name: i for i, name in enumerate(BASIC_MINERALS)}
        def _sort_key(tid):
            name = item_names.get(tid, "")
            if name in mineral_order:
                return (0, mineral_order[name], name)
            if tid in extra_set:
                return (1, 0, name)
            return (2, 0, name)

        comparison = {}
        for tid in sorted(type_ids, key=_sort_key):
            name = item_names.get(tid, f"TypeID {tid}")
            comparison[tid] = {
                'name': name,
                'is_mineral': name in BASIC_MINERALS,
                'is_extra': tid in extra_set,
                'old_buy':  old_prices.get(tid, {}).get('buy',  0.0),
                'new_buy':  new_prices.get(tid, {}).get('buy',  0.0),
                'old_sell': old_prices.get(tid, {}).get('sell', 0.0),
                'new_sell': new_prices.get(tid, {}).get('sell', 0.0),
            }
        return comparison
        
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
