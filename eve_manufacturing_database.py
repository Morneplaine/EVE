"""
EVE Online Manufacturing Database Creator
==========================================
This script downloads EVE Online Static Data Export (SDE) from Fuzzwork,
processes manufacturing blueprints, reprocessing data, and skill requirements,
then creates an Excel spreadsheet with live market prices from Jita 4-4.

Requirements: pandas, openpyxl, requests, xlsxwriter
"""

import pandas as pd
import requests
import time
from pathlib import Path
import logging
import sys

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Configuration
FUZZWORK_BASE = "https://www.fuzzwork.co.uk/dump/latest/"
JITA_SYSTEM_ID = 30000142  # Jita 4-4
DATA_DIR = Path("eve_data")
DATA_DIR.mkdir(exist_ok=True)

# Required CSV files from Fuzzwork SDE
REQUIRED_FILES = [
    "invTypes.csv.bz2",              # All items, modules, ships, etc.
    "invGroups.csv.bz2",             # Item group information
    "industryActivityMaterials.csv.bz2",  # Manufacturing materials
    "industryActivityProducts.csv.bz2",   # Manufacturing outputs
    "industryActivitySkills.csv.bz2",     # Required skills
    "invTypeMaterials.csv.bz2",      # Reprocessing outputs
    "industryActivity.csv.bz2",      # Activity types (manufacturing, etc.)
    "invVolumes.csv.bz2",            # Packaged volumes for items
]

def download_sde_file(filename, force_update=False):
    """
    Download a single SDE file from Fuzzwork if it doesn't exist locally or if update is needed.
    
    Args:
        filename (str): Name of the file to download
        force_update (bool): If True, always re-download even if file exists
        
    Returns:
        Path: Path to the downloaded file
    """
    local_path = DATA_DIR / filename
    url = FUZZWORK_BASE + filename
    
    # Check if update is needed
    if local_path.exists() and not force_update:
        try:
            # Get remote file size
            head_response = requests.head(url, timeout=10, allow_redirects=True)
            if head_response.status_code == 200:
                remote_size_str = head_response.headers.get('Content-Length')
                if remote_size_str:
                    try:
                        remote_size = int(remote_size_str)
                        local_size = local_path.stat().st_size
                        if local_size == remote_size:
                            logger.info(f"File {filename} already exists and is up to date, skipping download")
                            return local_path
                        else:
                            logger.info(f"File {filename} exists but size differs (local: {local_size:,}, remote: {remote_size:,}), re-downloading...")
                    except (ValueError, TypeError):
                        # Can't compare sizes, assume up to date
                        logger.info(f"File {filename} already exists, skipping download")
                        return local_path
                else:
                    # No size info, assume up to date
                    logger.info(f"File {filename} already exists, skipping download")
                    return local_path
        except requests.exceptions.RequestException:
            # Can't check remote, assume local is fine
            logger.info(f"File {filename} already exists, skipping download (could not check for updates)")
            return local_path
    
    logger.info(f"Downloading {filename} from {url}")
    
    try:
        response = requests.get(url, stream=True, timeout=30)
        response.raise_for_status()
        
        with open(local_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        
        logger.info(f"Successfully downloaded {filename}")
        return local_path
    
    except requests.exceptions.RequestException as e:
        logger.error(f"Failed to download {filename}: {e}")
        raise

def load_sde_data():
    """
    Download and load all required SDE CSV files.
    
    Returns:
        dict: Dictionary of DataFrames with SDE data
    """
    logger.info("Starting SDE data download and loading")
    
    sde_data = {}
    
    for filename in REQUIRED_FILES:
        try:
            file_path = download_sde_file(filename)
            table_name = filename.replace('.csv.bz2', '')
            
            # Load the compressed CSV
            logger.info(f"Loading {table_name} into memory")
            df = pd.read_csv(file_path, compression='bz2')
            sde_data[table_name] = df
            logger.info(f"Loaded {table_name}: {len(df)} rows")
            
        except Exception as e:
            logger.error(f"Error processing {filename}: {e}")
            raise
    
    return sde_data

def get_price_from_esi(type_id, region_id=10000002):
    """
    Fetch market price from ESI API (official EVE Online API) for a single type.
    Region 10000002 = The Forge (where Jita is located).
    
    Args:
        type_id (int): Type ID to fetch price for
        region_id (int): Region ID (default: The Forge)
        
    Returns:
        dict: Price info or None if failed
    """
    try:
        # ESI API endpoint for market orders
        url = f"https://esi.evetech.net/latest/markets/{region_id}/orders/"
        params = {
            'type_id': type_id,
            'order_type': 'sell',  # Sell orders (what you'd buy at)
            'page': 1
        }
        
        response = requests.get(url, params=params, timeout=5)
        if response.status_code == 200:
            orders = response.json()
            if orders:
                # Get minimum sell price (best price to buy at)
                sell_prices = [o['price'] for o in orders if o.get('location_id') == 60008494]  # Jita 4-4 station
                if not sell_prices:
                    # If no station-specific orders, use region minimum
                    sell_prices = [o['price'] for o in orders]
                
                if sell_prices:
                    min_price = min(sell_prices)
                    avg_price = sum(sell_prices) / len(sell_prices)
                    sorted_prices = sorted(sell_prices)
                    median_price = sorted_prices[len(sorted_prices) // 2]
                    
                    return {
                        'sell_min': min_price,
                        'sell_avg': avg_price,
                        'sell_median': median_price
                    }
    except Exception as e:
        logger.debug(f"ESI API failed for type {type_id}: {e}")
    return None

def get_price_from_eve_central(type_id, system_id):
    """
    Fetch market price from EVE-Central API.
    
    Args:
        type_id (int): Type ID to fetch price for
        system_id (int): System ID
        
    Returns:
        dict: Price info or None if failed
    """
    try:
        url = "https://api.eve-central.com/api/marketstat"
        params = {
            'typeid': type_id,
            'usesystem': system_id
        }
        
        response = requests.get(url, params=params, timeout=5)
        if response.status_code == 200:
            import xml.etree.ElementTree as ET
            root = ET.fromstring(response.content)
            type_elem = root.find('.//type')
            
            if type_elem is not None:
                sell_elem = type_elem.find('.//sell')
                if sell_elem is not None:
                    min_sell = sell_elem.find('min')
                    avg_sell = sell_elem.find('avg')
                    median_sell = sell_elem.find('median')
                    
                    if min_sell is not None and min_sell.text:
                        return {
                            'sell_min': float(min_sell.text),
                            'sell_avg': float(avg_sell.text) if avg_sell is not None and avg_sell.text else 0,
                            'sell_median': float(median_sell.text) if median_sell is not None and median_sell.text else 0
                        }
    except Exception as e:
        logger.debug(f"EVE-Central API failed for type {type_id}: {e}")
    return None

def get_price_from_evemarketer(type_id, system_id):
    """
    Fetch market price from EVEMarketer API (original method).
    
    Args:
        type_id (int): Type ID to fetch price for
        system_id (int): System ID
        
    Returns:
        dict: Price info or None if failed
    """
    try:
        url = f"https://api.evemarketer.com/ec/marketstat"
        params = {
            'typeid': type_id,
            'usesystem': system_id
        }
        
        response = requests.get(url, params=params, timeout=5)
        if response.status_code == 200:
            import xml.etree.ElementTree as ET
            root = ET.fromstring(response.content)
            type_elem = root.find('.//type')
            
            if type_elem is not None:
                sell_elem = type_elem.find('.//sell')
                if sell_elem is not None:
                    min_sell = sell_elem.find('min')
                    avg_sell = sell_elem.find('avg')
                    median_sell = sell_elem.find('median')
                    
                    if min_sell is not None and min_sell.text:
                        return {
                            'sell_min': float(min_sell.text),
                            'sell_avg': float(avg_sell.text) if avg_sell is not None and avg_sell.text else 0,
                            'sell_median': float(median_sell.text) if median_sell is not None and median_sell.text else 0
                        }
    except Exception as e:
        logger.debug(f"EVEMarketer API failed for type {type_id}: {e}")
    return None

def get_fuzzwork_market_prices(type_ids, station_id=30000142, batch_size=100):
    """
    Fetch market prices from Fuzzwork Market API for Jita.
    
    Based on: https://market.fuzzwork.co.uk/api/
    The API accepts multiple type IDs in one call, so we batch them efficiently.
    
    Args:
        type_ids (list): List of type IDs to fetch prices for
        station_id (int): Station/System ID (default: 30000142 for Jita)
        batch_size (int): Number of type IDs per API call (default: 100)
        
    Returns:
        dict: Dictionary mapping typeID to price info with keys:
            - buy_max: Maximum buy price
            - buy_volume: Total buy volume
            - sell_min: Minimum sell price
            - sell_volume: Total sell volume (bonus field)
    """
    logger.info(f"Fetching Fuzzwork Market prices for {len(type_ids)} items at station {station_id}")
    
    prices = {}
    failed_count = 0
    
    # Process in batches
    for i in range(0, len(type_ids), batch_size):
        batch = type_ids[i:i+batch_size]
        type_ids_str = ','.join(map(str, batch))
        
        url = f"https://market.fuzzwork.co.uk/aggregates/"
        params = {
            'station': station_id,
            'types': type_ids_str
        }
        
        try:
            logger.info(f"Fetching batch {i//batch_size + 1}/{(len(type_ids)-1)//batch_size + 1} ({len(batch)} items)")
            response = requests.get(url, params=params, timeout=30)
            response.raise_for_status()
            
            data = response.json()
            
            # Process each type ID in the response
            for type_id_str, type_data in data.items():
                type_id = int(type_id_str)
                
                buy_data = type_data.get('buy', {})
                sell_data = type_data.get('sell', {})
                
                prices[type_id] = {
                    'buy_max': float(buy_data.get('max', 0)) if buy_data.get('max') else 0,
                    'buy_volume': float(buy_data.get('volume', 0)) if buy_data.get('volume') else 0,
                    'sell_min': float(sell_data.get('min', 0)) if sell_data.get('min') else 0,
                    'sell_volume': float(sell_data.get('volume', 0)) if sell_data.get('volume') else 0,
                }
            
            # Rate limiting - be nice to the API
            time.sleep(0.5)
            
        except Exception as e:
            logger.warning(f"Failed to fetch batch {i//batch_size + 1}: {e}")
            failed_count += len(batch)
            # Mark all items in this batch as failed
            for type_id in batch:
                prices[type_id] = {
                    'buy_max': 0,
                    'buy_volume': 0,
                    'sell_min': 0,
                    'sell_volume': 0,
                }
    
    success_count = len([p for p in prices.values() if p['sell_min'] > 0 or p['buy_max'] > 0])
    logger.info(f"Fuzzwork Market price fetching complete: {success_count} successful, {failed_count} failed")
    logger.info(f"Success rate: {success_count/len(type_ids)*100:.1f}%")
    
    return prices

def get_market_price_batch(type_ids, max_retries=2):
    """
    Fetch market prices for multiple items from Jita using multiple API sources.
    
    Tries sources in order: ESI (official), EVE-Central, EVEMarketer.
    Continues with partial data if some requests fail.
    
    Args:
        type_ids (list): List of typeIDs to fetch prices for
        max_retries (int): Maximum number of retry attempts per source
        
    Returns:
        dict: Dictionary mapping typeID to price info (may be incomplete)
    """
    logger.info(f"Fetching market prices for {len(type_ids)} items using multiple sources")
    
    prices = {}
    failed_count = 0
    
    # Process items one at a time to avoid overwhelming APIs
    for idx, type_id in enumerate(type_ids):
        if (idx + 1) % 100 == 0:
            logger.info(f"Progress: {idx + 1}/{len(type_ids)} items processed ({len(prices)} prices found, {failed_count} failed)")
        
        price_data = None
        
        # Try ESI API first (official, most reliable)
        for attempt in range(max_retries):
            price_data = get_price_from_esi(type_id)
            if price_data:
                break
            if attempt < max_retries - 1:
                time.sleep(0.2)
        
        # Fallback to EVE-Central
        if not price_data:
            for attempt in range(max_retries):
                price_data = get_price_from_eve_central(type_id, JITA_SYSTEM_ID)
                if price_data:
                    break
                if attempt < max_retries - 1:
                    time.sleep(0.2)
        
        # Last resort: EVEMarketer
        if not price_data:
            for attempt in range(max_retries):
                price_data = get_price_from_evemarketer(type_id, JITA_SYSTEM_ID)
                if price_data:
                    break
                if attempt < max_retries - 1:
                    time.sleep(0.2)
        
        if price_data:
            prices[type_id] = price_data
        else:
            failed_count += 1
            # Set default values so Excel can still be created
            prices[type_id] = {
                'sell_min': 0,
                'sell_avg': 0,
                'sell_median': 0
            }
        
        # Rate limiting - be nice to the APIs
        time.sleep(0.1)
    
    success_count = len([p for p in prices.values() if p['sell_min'] > 0])
    logger.info(f"Price fetching complete: {success_count} successful, {failed_count} failed (set to 0)")
    logger.info(f"Success rate: {success_count/len(type_ids)*100:.1f}%")
    
    return prices

def process_manufacturing_data(sde_data):
    """
    Process SDE data to create manufacturing database with materials, skills, and products.
    
    This function:
    1. Filters for modules (you can adjust the category/group filter)
    2. Joins manufacturing materials, products, and skills
    3. Creates a comprehensive dataset
    
    Args:
        sde_data (dict): Dictionary of SDE DataFrames
        
    Returns:
        pd.DataFrame: Processed manufacturing data
    """
    logger.info("Processing manufacturing data")
    
    # Get the data tables
    inv_types = sde_data['invTypes']
    inv_groups = sde_data['invGroups']
    materials = sde_data['industryActivityMaterials']
    products = sde_data['industryActivityProducts']
    skills = sde_data['industryActivitySkills']
    
    # Activity ID 1 = Manufacturing
    MANUFACTURING_ACTIVITY = 1
    
    # Filter for manufacturing activity
    manufacturing_materials = materials[materials['activityID'] == MANUFACTURING_ACTIVITY].copy()
    manufacturing_products = products[products['activityID'] == MANUFACTURING_ACTIVITY].copy()
    manufacturing_skills = skills[skills['activityID'] == MANUFACTURING_ACTIVITY].copy()
    
    # Get blueprints that produce items
    blueprints = manufacturing_products[['typeID', 'productTypeID', 'quantity']].copy()
    blueprints.columns = ['blueprintTypeID', 'productTypeID', 'outputQuantity']
    
    # Merge with item names to get product names
    blueprints = blueprints.merge(
        inv_types[['typeID', 'typeName', 'groupID']],
        left_on='productTypeID',
        right_on='typeID',
        how='left'
    )
    blueprints.rename(columns={'typeName': 'productName'}, inplace=True)
    
    # Merge with groups to filter (optional - you can filter by specific groups like "Modules")
    blueprints = blueprints.merge(
        inv_groups[['groupID', 'groupName', 'categoryID']],
        on='groupID',
        how='left'
    )
    
    # Filter for specific categories if desired
    # Category 7 = Module, Category 6 = Ship, etc.
    # Uncomment to filter only modules:
    # blueprints = blueprints[blueprints['categoryID'] == 7]
    
    # For each blueprint, get materials
    blueprint_materials = []
    
    for idx, bp in blueprints.iterrows():
        bp_type_id = bp['blueprintTypeID']
        
        # Get materials for this blueprint
        bp_mats = manufacturing_materials[manufacturing_materials['typeID'] == bp_type_id].copy()
        
        if len(bp_mats) == 0:
            continue
        
        # Add material names
        bp_mats = bp_mats.merge(
            inv_types[['typeID', 'typeName']],
            left_on='materialTypeID',
            right_on='typeID',
            suffixes=('', '_material'),
            how='left'
        )
        
        # Aggregate materials into a string
        materials_list = []
        for _, mat in bp_mats.iterrows():
            materials_list.append(f"{mat['typeName']} x{int(mat['quantity'])}")
        
        materials_str = " | ".join(materials_list)
        
        # Get skills for this blueprint
        bp_skills = manufacturing_skills[manufacturing_skills['typeID'] == bp_type_id].copy()
        
        if len(bp_skills) > 0:
            bp_skills = bp_skills.merge(
                inv_types[['typeID', 'typeName']],
                left_on='skillID',
                right_on='typeID',
                suffixes=('', '_skill'),
                how='left'
            )
            skills_list = []
            for _, skill in bp_skills.iterrows():
                skills_list.append(f"{skill['typeName']} {int(skill['level'])}")
            skills_str = " | ".join(skills_list)
        else:
            skills_str = "None"
        
        blueprint_materials.append({
            'productTypeID': bp['productTypeID'],
            'productName': bp['productName'],
            'blueprintTypeID': bp['blueprintTypeID'],
            'groupName': bp['groupName'],
            'outputQuantity': bp['outputQuantity'],
            'materials': materials_str,
            'requiredSkills': skills_str,
            'materialsList': bp_mats[['materialTypeID', 'typeName', 'quantity']].to_dict('records') if len(bp_mats) > 0 else []
        })
    
    result_df = pd.DataFrame(blueprint_materials)
    logger.info(f"Processed {len(result_df)} blueprints with manufacturing data")
    
    return result_df

def process_reprocessing_data(sde_data):
    """
    Process reprocessing data to show what materials items break down into.
    
    Args:
        sde_data (dict): Dictionary of SDE DataFrames
        
    Returns:
        pd.DataFrame: Reprocessing data
    """
    logger.info("Processing reprocessing data")
    
    inv_types = sde_data['invTypes']
    reprocess = sde_data['invTypeMaterials']
    
    # Merge with item names
    reprocess_data = reprocess.merge(
        inv_types[['typeID', 'typeName']],
        left_on='typeID',
        right_on='typeID',
        how='left'
    ).rename(columns={'typeName': 'itemName'})
    
    reprocess_data = reprocess_data.merge(
        inv_types[['typeID', 'typeName']],
        left_on='materialTypeID',
        right_on='typeID',
        suffixes=('', '_material'),
        how='left'
    ).rename(columns={'typeName': 'materialName'})
    
    # Aggregate reprocessing outputs
    reprocess_summary = []
    
    for item_id in reprocess_data['typeID'].unique():
        item_data = reprocess_data[reprocess_data['typeID'] == item_id]
        item_name = item_data.iloc[0]['itemName']
        
        materials_list = []
        for _, mat in item_data.iterrows():
            materials_list.append(f"{mat['materialName']} x{int(mat['quantity'])}")
        
        materials_str = " | ".join(materials_list)
        
        reprocess_summary.append({
            'itemTypeID': item_id,
            'itemName': item_name,
            'reprocessingOutputs': materials_str,
            'outputsList': item_data[['materialTypeID', 'materialName', 'quantity']].to_dict('records')
        })
    
    result_df = pd.DataFrame(reprocess_summary)
    logger.info(f"Processed reprocessing data for {len(result_df)} items")
    
    return result_df

def create_excel_with_prices(manufacturing_df, reprocessing_df, output_filename, sde_data=None, prices=None):
    """
    Create an Excel file with manufacturing and reprocessing data, plus market prices.
    
    Manufacturing sheet uses separate columns for each material type (pivot format).
    
    Args:
        manufacturing_df (pd.DataFrame): Manufacturing data
        reprocessing_df (pd.DataFrame): Reprocessing data
        output_filename (str): Output Excel filename
        sde_data (dict, optional): SDE data dictionary for volume information
        prices (dict, optional): Pre-fetched prices dict. If None, will fetch prices.
    """
    logger.info(f"Creating Excel file: {output_filename}")
    
    # Collect all unique type IDs we need prices for
    all_type_ids = set()
    
    # Add product type IDs
    all_type_ids.update(manufacturing_df['productTypeID'].unique())
    
    # Add material type IDs from manufacturing
    for materials_list in manufacturing_df['materialsList']:
        for mat in materials_list:
            all_type_ids.add(mat['materialTypeID'])
    
    # Add item type IDs from reprocessing
    all_type_ids.update(reprocessing_df['itemTypeID'].unique())
    
    # Fetch prices if not provided - use Fuzzwork Market API
    if prices is None:
        logger.info("Fetching prices from Fuzzwork Market API...")
        fuzzwork_prices = get_fuzzwork_market_prices(list(all_type_ids), station_id=JITA_SYSTEM_ID)
        # Convert Fuzzwork format to expected format
        prices = {}
        for tid in all_type_ids:
            if tid in fuzzwork_prices:
                fw_price = fuzzwork_prices[tid]
                prices[tid] = {
                    'buy_max': fw_price.get('buy_max', 0),
                    'buy_volume': fw_price.get('buy_volume', 0),
                    'sell_min': fw_price.get('sell_min', 0),
                    'sell_avg': fw_price.get('sell_min', 0),  # Use sell_min as avg for compatibility
                    'sell_median': fw_price.get('sell_min', 0),
                    'sell_volume': fw_price.get('sell_volume', 0)
                }
            else:
                prices[tid] = {
                    'buy_max': 0,
                    'buy_volume': 0,
                    'sell_min': 0,
                    'sell_avg': 0,
                    'sell_median': 0,
                    'sell_volume': 0
                }
    else:
        # Ensure all type IDs have entries (fill missing ones with zeros)
        for tid in all_type_ids:
            if tid not in prices:
                prices[tid] = {
                    'buy_max': 0,
                    'buy_volume': 0,
                    'sell_min': 0,
                    'sell_avg': 0,
                    'sell_median': 0,
                    'sell_volume': 0
                }
    
    # ===== PREPARE MANUFACTURING DATA IN PIVOT FORMAT =====
    logger.info("Preparing manufacturing data in pivot format")
    
    # Collect all unique materials across all blueprints and count usage frequency
    all_materials = {}  # materialTypeID -> name
    material_usage_count = {}  # materialTypeID -> count of how many products use it
    
    for materials_list in manufacturing_df['materialsList']:
        used_materials = set()  # Track unique materials per blueprint
        for mat in materials_list:
            mat_id = mat['materialTypeID']
            mat_name = mat.get('typeName', 'Unknown')
            if mat_name is None or (isinstance(mat_name, float) and pd.isna(mat_name)):
                mat_name = 'Unknown'
            all_materials[mat_id] = str(mat_name)
            used_materials.add(mat_id)
        
        # Count how many blueprints use each material
        for mat_id in used_materials:
            material_usage_count[mat_id] = material_usage_count.get(mat_id, 0) + 1
    
    # Sort materials by usage frequency (most used first), then by name
    sorted_materials = sorted(
        all_materials.items(), 
        key=lambda x: (-material_usage_count.get(x[0], 0), str(x[1]))  # Negative for descending order
    )
    
    logger.info(f"Material columns sorted by usage frequency (most used first)")
    logger.info(f"Top 10 most used materials:")
    for i, (mat_id, mat_name) in enumerate(sorted_materials[:10], 1):
        count = material_usage_count.get(mat_id, 0)
        logger.info(f"  {i}. {mat_name}: used in {count} products")
    
    # Prepare volume lookup if SDE data is available
    volume_lookup = {}
    packaged_volume_lookup = {}
    if sde_data:
        logger.info("Loading volume data from SDE")
        inv_types = sde_data.get('invTypes')
        inv_volumes = sde_data.get('invVolumes')
        
        if inv_types is not None:
            # Create lookup for regular volume
            for _, row in inv_types.iterrows():
                volume_lookup[row['typeID']] = row.get('volume', 0.0)
        
        if inv_volumes is not None:
            # Create lookup for packaged volume
            for _, row in inv_volumes.iterrows():
                packaged_volume_lookup[row['typeID']] = row.get('volume', 0.0)
        
        logger.info(f"Loaded volumes for {len(volume_lookup)} items")
        if packaged_volume_lookup:
            logger.info(f"Loaded packaged volumes for {len(packaged_volume_lookup)} items")
    
    # Build pivot table
    mfg_pivot_data = []
    for idx, row in manufacturing_df.iterrows():
        pivot_row = {
            'Product Name': row['productName'],
            'Product TypeID': row['productTypeID'],
            'Group': row['groupName'],
            'Output Qty': row['outputQuantity']
        }
        
        # Initialize all material columns to 0
        for mat_id, mat_name in sorted_materials:
            pivot_row[mat_name] = 0
        
        # Fill in actual quantities from materialsList
        for mat in row['materialsList']:
            mat_name = mat.get('typeName', 'Unknown')
            if mat_name is None or (isinstance(mat_name, float) and pd.isna(mat_name)):
                mat_name = 'Unknown'
            mat_name = str(mat_name)
            if mat_name in pivot_row:  # Only if column exists
                pivot_row[mat_name] = int(mat['quantity'])
        
        # Add skills, price, and volume columns at the end
        pivot_row['Required Skills'] = row['requiredSkills']
        # Use sell_min for product price
        pivot_row['Product Price (ISK)'] = prices.get(row['productTypeID'], {}).get('sell_min', 0)
        
        # Add volume (prefer packaged volume if available)
        product_type_id = row['productTypeID']
        if product_type_id in packaged_volume_lookup:
            pivot_row['Product Volume (m³)'] = packaged_volume_lookup[product_type_id]
        else:
            pivot_row['Product Volume (m³)'] = volume_lookup.get(product_type_id, 0.0)
        
        mfg_pivot_data.append(pivot_row)
    
    mfg_pivot_df = pd.DataFrame(mfg_pivot_data)
    
    # Create Excel writer
    with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1
        })
        
        price_header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#70AD47',
            'font_color': 'white',
            'border': 1,
            'num_format': '#,##0.00'
        })
        
        money_format = workbook.add_format({'num_format': '#,##0.00'})
        number_format = workbook.add_format({'num_format': '#,##0'})
        formula_format = workbook.add_format({'num_format': '#,##0.00'})
        
        # ===== PRICES SHEET (for VLOOKUP reference) =====
        logger.info("Creating Prices sheet for VLOOKUP reference")
        
        price_data = []
        for type_id in sorted(all_type_ids):
            price_info = prices.get(type_id, {})
            price_data.append({
                'typeID': type_id,
                'Buy Max': price_info.get('buy_max', 0),
                'Buy Volume': price_info.get('buy_volume', 0),
                'Sell Min': price_info.get('sell_min', 0),
                'Sell Avg': price_info.get('sell_avg', 0),
                'Sell Median': price_info.get('sell_median', 0),
                'Sell Volume': price_info.get('sell_volume', 0)
            })
        
        prices_df = pd.DataFrame(price_data)
        prices_df.to_excel(writer, sheet_name='Prices', index=False)
        
        prices_worksheet = writer.sheets['Prices']
        prices_worksheet.set_column('A:A', 12)  # typeID
        prices_worksheet.set_column('B:G', 18)  # Price columns
        
        # ===== MANUFACTURING SHEET =====
        logger.info("Creating Manufacturing sheet with pivot format and buy price row")
        
        # Write data without headers, starting from row 2 (row 0 = headers, row 1 = buy prices, row 2+ = data)
        mfg_pivot_df.to_excel(writer, sheet_name='Manufacturing', index=False, startrow=2, header=False)
        
        worksheet = writer.sheets['Manufacturing']
        
        # Write header row (row 0)
        col = 0
        worksheet.write(0, col, 'Product Name', header_format)
        col += 1
        worksheet.write(0, col, 'Product TypeID', header_format)
        col += 1
        worksheet.write(0, col, 'Group', header_format)
        col += 1
        worksheet.write(0, col, 'Output Qty', header_format)
        col += 1
        
        # Material column headers and buy price row
        num_materials = len(sorted_materials)
        start_col = 4  # Column E (0-indexed: A=0, B=1, C=2, D=3, E=4)
        
        for i, (mat_id, mat_name) in enumerate(sorted_materials):
            col_idx = start_col + i
            # Write header in row 0
            worksheet.write(0, col_idx, mat_name, header_format)
            # Write buy price formula in row 1 (VLOOKUP from Prices sheet)
            # Formula: =VLOOKUP(material_typeID, Prices!$A:$B, 2, FALSE)
            buy_price_formula = f'=VLOOKUP({mat_id},Prices!$A:$B,2,FALSE)'
            worksheet.write(1, col_idx, buy_price_formula, formula_format)
            worksheet.set_column(col_idx, col_idx, 12)
        
        # Skills, Price, and Volume columns
        skills_col = start_col + num_materials
        price_col = skills_col + 1
        volume_col = price_col + 1
        
        worksheet.write(0, skills_col, 'Required Skills', header_format)
        worksheet.write(0, price_col, 'Product Price (ISK)', header_format)
        worksheet.write(0, volume_col, 'Product Volume (m³)', header_format)
        
        # Write empty cells in buy price row (row 1) for non-material columns
        worksheet.write(1, skills_col, '', price_header_format)
        worksheet.write(1, price_col, '', price_header_format)
        worksheet.write(1, volume_col, '', price_header_format)
        
        worksheet.set_column('A:A', 30)  # Product Name
        worksheet.set_column('B:B', 15)  # Product TypeID
        worksheet.set_column('C:C', 20)  # Group
        worksheet.set_column('D:D', 12)  # Output Qty
        worksheet.set_column(skills_col, skills_col, 40)
        worksheet.set_column(price_col, price_col, 18)
        worksheet.set_column(volume_col, volume_col, 18)
        
        # ===== MATERIAL PRICES SHEET =====
        logger.info("Creating Material Prices sheet")
        
        # Create a separate sheet with material prices for easy lookup
        material_prices_data = []
        for mat_id, mat_name in sorted_materials:
            # Get volume (prefer packaged volume if available)
            if mat_id in packaged_volume_lookup:
                mat_volume = packaged_volume_lookup[mat_id]
            else:
                mat_volume = volume_lookup.get(mat_id, 0.0)
            
            material_prices_data.append({
                'Material TypeID': mat_id,
                'Material Name': mat_name,
                'Volume (m³)': mat_volume,
                'Jita Sell Min': prices.get(mat_id, {}).get('sell_min', 0),
                'Jita Sell Avg': prices.get(mat_id, {}).get('sell_avg', 0),
                'Jita Sell Median': prices.get(mat_id, {}).get('sell_median', 0)
            })
        
        mat_prices_df = pd.DataFrame(material_prices_data)
        mat_prices_df.to_excel(writer, sheet_name='Material Prices', index=False)
        
        worksheet = writer.sheets['Material Prices']
        worksheet.set_column('A:A', 15)  # Material TypeID
        worksheet.set_column('B:B', 25)  # Material Name
        worksheet.set_column('C:C', 15)  # Volume
        worksheet.set_column('D:F', 18)  # Prices
        
        # ===== REPROCESSING SHEET =====
        logger.info("Creating Reprocessing sheet with pivot format")
        
        # Collect all unique reprocessing output materials and count usage frequency
        all_reprocess_materials = {}  # materialTypeID -> name
        reprocess_material_usage_count = {}  # materialTypeID -> count of how many items produce it
        
        for outputs_list in reprocessing_df['outputsList']:
            used_materials = set()  # Track unique materials per item
            for output in outputs_list:
                mat_id = output['materialTypeID']
                mat_name = output.get('materialName', 'Unknown')
                if mat_name is None or (isinstance(mat_name, float) and pd.isna(mat_name)):
                    mat_name = 'Unknown'
                all_reprocess_materials[mat_id] = str(mat_name)
                used_materials.add(mat_id)
            
            # Count how many items produce each material
            for mat_id in used_materials:
                reprocess_material_usage_count[mat_id] = reprocess_material_usage_count.get(mat_id, 0) + 1
        
        # Sort materials by usage frequency (most used first), then by name
        sorted_reprocess_materials = sorted(
            all_reprocess_materials.items(),
            key=lambda x: (-reprocess_material_usage_count.get(x[0], 0), str(x[1]))  # Negative for descending order
        )
        
        logger.info(f"Reprocessing material columns sorted by usage frequency (most used first)")
        logger.info(f"Top 10 most common reprocessing outputs:")
        for i, (mat_id, mat_name) in enumerate(sorted_reprocess_materials[:10], 1):
            count = reprocess_material_usage_count.get(mat_id, 0)
            logger.info(f"  {i}. {mat_name}: produced by {count} items")
        
        # Build pivot table for reprocessing
        reprocess_pivot_data = []
        for idx, row in reprocessing_df.iterrows():
            pivot_row = {
                'Item Name': row['itemName'],
                'Item TypeID': row['itemTypeID']
            }
            
            # Initialize all material columns to 0
            for mat_id, mat_name in sorted_reprocess_materials:
                pivot_row[mat_name] = 0
            
            # Fill in actual quantities from outputsList
            for output in row['outputsList']:
                mat_name = output.get('materialName', 'Unknown')
                if mat_name is None or (isinstance(mat_name, float) and pd.isna(mat_name)):
                    mat_name = 'Unknown'
                mat_name = str(mat_name)
                if mat_name in pivot_row:  # Only if column exists
                    pivot_row[mat_name] = int(output['quantity'])
            
            # Add price and volume columns at the end
            item_type_id = row['itemTypeID']
            pivot_row['Item Price (ISK)'] = prices.get(item_type_id, {}).get('sell_min', 0)
            
            # Add volume (prefer packaged volume if available)
            if item_type_id in packaged_volume_lookup:
                pivot_row['Item Volume (m³)'] = packaged_volume_lookup[item_type_id]
            else:
                pivot_row['Item Volume (m³)'] = volume_lookup.get(item_type_id, 0.0)
            
            reprocess_pivot_data.append(pivot_row)
        
        reprocess_pivot_df = pd.DataFrame(reprocess_pivot_data)
        # Write data without headers, starting from row 2 (row 0 = headers, row 1 = buy prices, row 2+ = data)
        reprocess_pivot_df.to_excel(writer, sheet_name='Reprocessing', index=False, startrow=2, header=False)
        
        worksheet = writer.sheets['Reprocessing']
        
        # Write header row (row 0)
        col = 0
        worksheet.write(0, col, 'Item Name', header_format)
        col += 1
        worksheet.write(0, col, 'Item TypeID', header_format)
        col += 1
        
        # Material column headers and buy price row
        num_reprocess_materials = len(sorted_reprocess_materials)
        start_col = 2  # Column C (0-indexed: A=0, B=1, C=2)
        
        for i, (mat_id, mat_name) in enumerate(sorted_reprocess_materials):
            col_idx = start_col + i
            # Write header in row 0
            worksheet.write(0, col_idx, mat_name, header_format)
            # Write buy price formula in row 1 (VLOOKUP from Prices sheet)
            buy_price_formula = f'=VLOOKUP({mat_id},Prices!$A:$B,2,FALSE)'
            worksheet.write(1, col_idx, buy_price_formula, formula_format)
            worksheet.set_column(col_idx, col_idx, 12)
        
        # Price and Volume columns
        price_col = start_col + num_reprocess_materials
        volume_col = price_col + 1
        
        worksheet.write(0, price_col, 'Item Price (ISK)', header_format)
        worksheet.write(0, volume_col, 'Item Volume (m³)', header_format)
        
        # Write empty cells in buy price row (row 1) for non-material columns
        worksheet.write(1, price_col, '', price_header_format)
        worksheet.write(1, volume_col, '', price_header_format)
        
        worksheet.set_column('A:A', 30)  # Item Name
        worksheet.set_column('B:B', 15)  # Item TypeID
        worksheet.set_column(price_col, price_col, 18)  # Item Price
        worksheet.set_column(volume_col, volume_col, 18)  # Item Volume
        
        # Note: "Prices" sheet already created above with all price data
        # "All Prices" sheet removed - use "Prices" sheet instead
        
        # ===== INSTRUCTIONS SHEET =====
        logger.info("Creating Instructions sheet")
        
        # Create instructions as a list of strings (avoiding special characters that Excel interprets as formulas)
        instructions_list = [
            'EVE Online Manufacturing Database',
            '',
            '=== SHEET DESCRIPTIONS ===',
            '',
            'Manufacturing:',
            '  - Each row is a blueprint/product',
            '  - Each material has its own column showing quantity needed',
            '  - Materials show 0 if not used in that blueprint',
            '  - Use this format for easy cost calculations',
            '  - Formula example: =E2*MaterialPrices.C2 (for material in column E)',
            '',
            'Material Prices:',
            '  - Current Jita 4-4 prices for all manufacturing materials',
            '  - Use VLOOKUP or INDEX/MATCH to reference these prices',
            '  - Example: =VLOOKUP(material_name, MaterialPrices!B:C, 2, FALSE)',
            '',
            'Reprocessing:',
            '  - Shows what items break down into when reprocessed',
            '  - Values assume 100% reprocessing efficiency',
            '  - Actual yields: 50-60% depending on skills/station',
            '',
            'All Prices:',
            '  - Raw price data for all items by typeID',
            '  - Includes sell min/avg/median',
            '',
            '=== CALCULATING MANUFACTURING COST ===',
            '',
            'Total Material Cost = Sum(Material Quantity * Material Price)',
            '',
            'For each blueprint row in Manufacturing sheet:',
            '1. For each material column (E onwards), multiply by that material price',
            '2. Sum all material costs',
            '3. Add manufacturing fees (~1-2% of material cost)',
            '4. Compare to Product Price to calculate profit',
            '',
            'Example formula for total cost (assuming prices in row 2 of Material Prices):',
            '  =E2*MaterialPrices.$C$2 + F2*MaterialPrices.$C$3 + G2*MaterialPrices.$C$4 + ...',
            '',
            'Or use SUMPRODUCT for cleaner formula:',
            '  =SUMPRODUCT(E2:N2, MaterialPrices.$C$2:$C$11)',
            '',
            '=== DATA SOURCES ===',
            '',
            'Static Data: Fuzzwork SDE (www.fuzzwork.co.uk)',
            'Market Prices: Fuzzwork Market API (market.fuzzwork.co.uk)',
            'System: Jita 4-4 (system ID: 30000142)',
            '',
            '=== UPDATING PRICES ===',
            '',
            'Prices change frequently, but other data (manufacturing, volumes) rarely change.',
            'To update ONLY prices (fast, ~2-3 minutes):',
            '  1. Close this Excel file',
            '  2. Run: python update_prices.py',
            '     Or double-click: update_prices.bat',
            '  3. Re-open this Excel file',
            '',
            'The VLOOKUP formulas in row 1 will automatically show updated buy prices.',
            '',
            'To update everything (slow, ~5-10 minutes):',
            '  Run: python eve_manufacturing_database.py',
            '',
            'If prices are zero, you can manually update them using:',
            '  - Excel Data > Get Data > From Web',
            '  - Websites: evemarketer.com, evepraisal.com, eve-markets.net',
            '',
            '=== IMPORTANT NOTES ===',
            '',
            '- Material quantities are for ME 0 (unresearched blueprint)',
            '- ME research reduces materials by 1% per level (max ME 10 = -10%)',
            '- Reprocessing values assume 100% efficiency',
            '- Actual reprocessing: 50-60% depending on skills and station',
            '- Prices fetched at: ' + pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
            '',
            '=== PROFIT CALCULATION ===',
            '',
            'Profit per unit = Product Price - Material Cost - Fees - Taxes',
            '',
            'Manufacturing Fee: ~1-2% of material cost (varies by system)',
            'Sales Tax: 2.5-8% of sell price (depends on skills/structure)',
            '',
            'Break-even point: Product Price > Material Cost * 1.10',
            '(Assumes 2% mfg fee + 8% sales tax)',
            '',
            '=== RESOURCES ===',
            '',
            'Fuzzwork: www.fuzzwork.co.uk',
            'EVEMarketer: www.evemarketer.com',
            'EVE University: wiki.eveuniversity.org/Manufacturing',
            'ESI API: esi.evetech.net/ui',
        ]
        
        # Create Instructions sheet manually to avoid formula interpretation issues
        worksheet = workbook.add_worksheet('Instructions')
        
        # Set text format to prevent Excel from interpreting as formulas
        text_format = workbook.add_format({'text_wrap': True})
        worksheet.set_column('A:A', 100, text_format)
        
        # Write cells as text to prevent formula interpretation
        for row_num, text in enumerate(instructions_list):
            # Prefix with apostrophe to force text format for lines that start with formula-like characters
            if text and (text.startswith('=') or text.startswith('+') or (text.startswith('-') and len(text) > 1 and text[1] != ' ')):
                worksheet.write(row_num, 0, "'" + text, text_format)
            else:
                worksheet.write(row_num, 0, text, text_format)
    
    logger.info(f"Excel file created successfully: {output_filename}")
    logger.info(f"Manufacturing sheet has {len(sorted_materials)} material columns")

def main():
    """
    Main function to orchestrate the entire process.
    """
    # Check for command-line arguments
    skip_prices = '--skip-prices' in sys.argv or '-s' in sys.argv
    
    logger.info("Starting EVE Online Manufacturing Database Creator")
    if skip_prices:
        logger.info("Price fetching is SKIPPED (--skip-prices flag used)")
        logger.info("Excel will be created with zero prices. You can add prices manually later.")
    
    try:
        # Step 1: Download and load SDE data
        sde_data = load_sde_data()
        
        # Step 2: Process manufacturing data
        manufacturing_df = process_manufacturing_data(sde_data)
        
        # Step 3: Process reprocessing data
        reprocessing_df = process_reprocessing_data(sde_data)
        
        # Step 4: Create Excel with prices
        # Note: Price fetching may fail partially, but Excel will still be created
        output_file = "EVE_Manufacturing_Database.xlsx"
        
        if skip_prices:
            # Create Excel with zero prices immediately
            logger.info("Creating Excel file without fetching prices...")
            all_type_ids = set()
            all_type_ids.update(manufacturing_df['productTypeID'].unique())
            for materials_list in manufacturing_df['materialsList']:
                for mat in materials_list:
                    all_type_ids.add(mat['materialTypeID'])
            all_type_ids.update(reprocessing_df['itemTypeID'].unique())
            
            prices = {tid: {
                'buy_max': 0,
                'buy_volume': 0,
                'sell_min': 0,
                'sell_avg': 0,
                'sell_median': 0,
                'sell_volume': 0
            } for tid in all_type_ids}
            create_excel_with_prices(manufacturing_df, reprocessing_df, output_file, sde_data=sde_data, prices=prices)
        else:
            create_excel_with_prices(manufacturing_df, reprocessing_df, output_file, sde_data=sde_data)
        
        logger.info("=" * 60)
        logger.info("SUCCESS! Database created successfully")
        logger.info(f"Output file: {output_file}")
        logger.info(f"Total blueprints: {len(manufacturing_df)}")
        logger.info(f"Total reprocessing items: {len(reprocessing_df)}")
        logger.info("=" * 60)
        if skip_prices:
            logger.info("NOTE: Prices were skipped. To add prices:")
            logger.info("  1. Run: python eve_manufacturing_database.py (without --skip-prices)")
            logger.info("  2. Or update just prices: python update_prices.py")
        else:
            logger.info("NOTE: To update prices later (without regenerating everything):")
            logger.info("  1. Close the Excel file")
            logger.info("  2. Run: python update_prices.py")
            logger.info("  3. Re-open the Excel file - prices will auto-update via VLOOKUP")
        logger.info("=" * 60)
        
    except Exception as e:
        logger.error(f"An error occurred: {e}")
        logger.error("The script will attempt to create Excel file anyway if data was processed.")
        raise

if __name__ == "__main__":
    main()

