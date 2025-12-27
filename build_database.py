"""
Build SQLite database from EVE SDE data
This script populates the database with manufacturing, reprocessing, and item data.
Prices are fetched separately via update_prices_db.py
"""

import sqlite3
import pandas as pd
import requests
from pathlib import Path
import logging
from eve_manufacturing_database import (
    download_sde_file,
    load_sde_data,
    DATA_DIR
)

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

DB_FILE = "eve_manufacturing.db"

def create_database():
    """Create database schema"""
    logger.info(f"Creating database: {DB_FILE}")
    conn = sqlite3.connect(DB_FILE)
    
    with open('database_schema.sql', 'r') as f:
        schema = f.read()
        conn.executescript(schema)
    
    conn.commit()
    logger.info("Database schema created")
    return conn

def populate_items_and_groups(conn, sde_data):
    """Populate items and groups tables"""
    logger.info("Populating items and groups...")
    
    inv_types = sde_data['invTypes']
    inv_groups = sde_data['invGroups']
    inv_volumes = sde_data['invVolumes']
    
    # Create volume lookup from invVolumes (preferred - has packaged volumes)
    volume_lookup = {}
    packaged_volume_lookup = {}
    
    for _, vol in inv_volumes.iterrows():
        type_id = vol['typeID']
        volume_lookup[type_id] = vol.get('volume', 0.0)
        packaged_volume_lookup[type_id] = vol.get('packagedVolume', 0.0)
    
    # Fallback: Use volume from invTypes if not in invVolumes
    # invTypes has a 'volume' column that we can use as fallback
    for _, item in inv_types.iterrows():
        type_id = item['typeID']
        if type_id not in volume_lookup or volume_lookup[type_id] == 0:
            inv_types_volume = item.get('volume', 0.0)
            if inv_types_volume and inv_types_volume > 0:
                volume_lookup[type_id] = inv_types_volume
                # Use same volume for packaged if not specified
                if type_id not in packaged_volume_lookup or packaged_volume_lookup[type_id] == 0:
                    packaged_volume_lookup[type_id] = inv_types_volume
    
    # Insert groups
    groups_data = inv_groups[['groupID', 'groupName', 'categoryID']].copy()
    groups_data.to_sql('groups', conn, if_exists='replace', index=False)
    logger.info(f"Inserted {len(groups_data)} groups")
    
    # Insert items with volumes
    # Note: categoryID is not in invTypes, it's in invGroups, so we merge to get it
    items_data = inv_types[['typeID', 'typeName', 'groupID']].copy()
    
    # Merge with groups to get categoryID
    items_data = items_data.merge(
        inv_groups[['groupID', 'categoryID']],
        on='groupID',
        how='left'
    )
    
    items_data['volume'] = items_data['typeID'].map(volume_lookup).fillna(0.0)
    items_data['packaged_volume'] = items_data['typeID'].map(packaged_volume_lookup).fillna(0.0)
    items_data.to_sql('items', conn, if_exists='replace', index=False)
    logger.info(f"Inserted {len(items_data)} items")

def populate_blueprints(conn, sde_data):
    """Populate blueprints table"""
    logger.info("Populating blueprints...")
    
    inv_types = sde_data['invTypes']
    inv_groups = sde_data['invGroups']
    products = sde_data['industryActivityProducts']
    
    MANUFACTURING_ACTIVITY = 1
    manufacturing_products = products[products['activityID'] == MANUFACTURING_ACTIVITY].copy()
    
    # Merge with item names
    blueprints = manufacturing_products.merge(
        inv_types[['typeID', 'typeName', 'groupID']],
        left_on='productTypeID',
        right_on='typeID',
        how='left',
        suffixes=('', '_product')
    )
    
    # Merge with groups
    blueprints = blueprints.merge(
        inv_groups[['groupID', 'groupName']],
        on='groupID',
        how='left'
    )
    
    # Prepare for database
    blueprint_data = blueprints[[
        'typeID',  # blueprintTypeID
        'productTypeID',
        'typeName',  # productName
        'quantity',  # outputQuantity
        'groupID',
        'groupName'
    ]].copy()
    blueprint_data.columns = [
        'blueprintTypeID',
        'productTypeID',
        'productName',
        'outputQuantity',
        'groupID',
        'groupName'
    ]
    
    blueprint_data.to_sql('blueprints', conn, if_exists='replace', index=False)
    logger.info(f"Inserted {len(blueprint_data)} blueprints")

def populate_manufacturing_materials(conn, sde_data):
    """Populate manufacturing materials"""
    logger.info("Populating manufacturing materials...")
    
    inv_types = sde_data['invTypes']
    materials = sde_data['industryActivityMaterials']
    
    MANUFACTURING_ACTIVITY = 1
    manufacturing_materials = materials[materials['activityID'] == MANUFACTURING_ACTIVITY].copy()
    
    # Merge with material names
    materials_with_names = manufacturing_materials.merge(
        inv_types[['typeID', 'typeName']],
        left_on='materialTypeID',
        right_on='typeID',
        how='left',
        suffixes=('', '_material')
    )
    
    # Prepare for database
    materials_data = materials_with_names[[
        'typeID',  # blueprintTypeID
        'materialTypeID',
        'typeName',  # materialName
        'quantity'
    ]].copy()
    materials_data.columns = [
        'blueprintTypeID',
        'materialTypeID',
        'materialName',
        'quantity'
    ]
    
    materials_data.to_sql('manufacturing_materials', conn, if_exists='replace', index=False)
    logger.info(f"Inserted {len(materials_data)} material requirements")

def populate_manufacturing_skills(conn, sde_data):
    """Populate manufacturing skills"""
    logger.info("Populating manufacturing skills...")
    
    inv_types = sde_data['invTypes']
    skills = sde_data['industryActivitySkills']
    
    MANUFACTURING_ACTIVITY = 1
    manufacturing_skills = skills[skills['activityID'] == MANUFACTURING_ACTIVITY].copy()
    
    # Merge with skill names
    skills_with_names = manufacturing_skills.merge(
        inv_types[['typeID', 'typeName']],
        left_on='skillID',
        right_on='typeID',
        how='left',
        suffixes=('', '_skill')
    )
    
    # Prepare for database
    skills_data = skills_with_names[[
        'typeID',  # blueprintTypeID
        'skillID',
        'typeName',  # skillName
        'level'
    ]].copy()
    skills_data.columns = [
        'blueprintTypeID',
        'skillID',
        'skillName',
        'level'
    ]
    
    skills_data.to_sql('manufacturing_skills', conn, if_exists='replace', index=False)
    logger.info(f"Inserted {len(skills_data)} skill requirements")

def populate_reprocessing(conn, sde_data):
    """Populate reprocessing outputs"""
    logger.info("Populating reprocessing outputs...")
    
    inv_types = sde_data['invTypes']
    reprocess = sde_data['invTypeMaterials']
    
    # Merge with item names
    reprocess_data = reprocess.merge(
        inv_types[['typeID', 'typeName']],
        left_on='typeID',
        right_on='typeID',
        how='left',
        suffixes=('', '_item')
    )
    
    # Merge with material names
    reprocess_data = reprocess_data.merge(
        inv_types[['typeID', 'typeName']],
        left_on='materialTypeID',
        right_on='typeID',
        how='left',
        suffixes=('', '_material')
    )
    
    # Prepare for database
    reprocess_final = reprocess_data[[
        'typeID',  # itemTypeID
        'typeName',  # itemName
        'materialTypeID',
        'typeName_material',  # materialName
        'quantity'
    ]].copy()
    reprocess_final.columns = [
        'itemTypeID',
        'itemName',
        'materialTypeID',
        'materialName',
        'quantity'
    ]
    
    reprocess_final.to_sql('reprocessing_outputs', conn, if_exists='replace', index=False)
    logger.info(f"Inserted {len(reprocess_final)} reprocessing outputs")

def main():
    logger.info("=" * 60)
    logger.info("Building EVE Manufacturing Database")
    logger.info("=" * 60)
    
    try:
        # Create database
        conn = create_database()
        
        # Load SDE data
        logger.info("Loading SDE data...")
        sde_data = load_sde_data()
        
        # Populate tables
        populate_items_and_groups(conn, sde_data)
        populate_blueprints(conn, sde_data)
        populate_manufacturing_materials(conn, sde_data)
        populate_manufacturing_skills(conn, sde_data)
        populate_reprocessing(conn, sde_data)
        
        # Initialize prices table with all items (prices will be 0 until updated)
        logger.info("Initializing prices table...")
        conn.execute("""
            INSERT OR IGNORE INTO prices (typeID)
            SELECT typeID FROM items
        """)
        conn.commit()
        
        conn.close()
        
        logger.info("=" * 60)
        logger.info("SUCCESS! Database built successfully")
        logger.info(f"Database file: {DB_FILE}")
        logger.info("=" * 60)
        logger.info("Next steps:")
        logger.info("  1. Update prices: python update_prices_db.py")
        logger.info("  2. Generate Excel: python generate_excel.py")
        logger.info("  3. Analyze profitability: python analyze_profitability.py")
        logger.info("=" * 60)
        
    except Exception as e:
        logger.error(f"Error building database: {e}")
        import traceback
        logger.error(traceback.format_exc())
        raise

if __name__ == "__main__":
    main()

