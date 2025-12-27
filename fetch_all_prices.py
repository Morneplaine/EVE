"""
Fetch all prices from Fuzzwork Market API for all items in the database
This will get prices for all manufacturing products, materials, and reprocessing items
"""

import sys
import json
import pandas as pd
from pathlib import Path
from eve_manufacturing_database import (
    load_sde_data,
    process_manufacturing_data,
    process_reprocessing_data,
    get_fuzzwork_market_prices,
    JITA_SYSTEM_ID
)
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def main():
    logger.info("Starting price fetch for all items")
    logger.info("=" * 60)
    
    # Step 1: Load SDE data
    logger.info("Loading SDE data...")
    sde_data = load_sde_data()
    
    # Step 2: Process manufacturing data
    logger.info("Processing manufacturing data...")
    manufacturing_df = process_manufacturing_data(sde_data)
    
    # Step 3: Process reprocessing data
    logger.info("Processing reprocessing data...")
    reprocessing_df = process_reprocessing_data(sde_data)
    
    # Step 4: Collect all unique type IDs
    logger.info("Collecting all unique type IDs...")
    all_type_ids = set()
    
    # Add product type IDs
    all_type_ids.update(manufacturing_df['productTypeID'].unique())
    logger.info(f"  - {len(manufacturing_df['productTypeID'].unique())} unique products")
    
    # Add material type IDs from manufacturing
    material_count = 0
    for materials_list in manufacturing_df['materialsList']:
        for mat in materials_list:
            all_type_ids.add(mat['materialTypeID'])
            material_count += 1
    logger.info(f"  - {material_count} material entries (may have duplicates)")
    
    # Add item type IDs from reprocessing
    all_type_ids.update(reprocessing_df['itemTypeID'].unique())
    logger.info(f"  - {len(reprocessing_df['itemTypeID'].unique())} unique reprocessing items")
    
    # Convert to sorted list for consistent processing
    all_type_ids = sorted(list(all_type_ids))
    logger.info(f"  - Total unique items: {len(all_type_ids)}")
    logger.info("=" * 60)
    
    # Step 5: Fetch prices from Fuzzwork Market API
    logger.info("Fetching prices from Fuzzwork Market API for Jita...")
    prices = get_fuzzwork_market_prices(all_type_ids, station_id=JITA_SYSTEM_ID, batch_size=100)
    
    # Step 6: Save results to CSV for review
    logger.info("=" * 60)
    logger.info("Saving results...")
    
    # Create DataFrame with results
    price_data = []
    for type_id in all_type_ids:
        if type_id in prices:
            price_info = prices[type_id]
            price_data.append({
                'typeID': type_id,
                'buy_max': price_info['buy_max'],
                'buy_volume': price_info['buy_volume'],
                'sell_min': price_info['sell_min'],
                'sell_volume': price_info['sell_volume']
            })
        else:
            price_data.append({
                'typeID': type_id,
                'buy_max': 0,
                'buy_volume': 0,
                'sell_min': 0,
                'sell_volume': 0
            })
    
    price_df = pd.DataFrame(price_data)
    
    # Save to CSV
    output_file = "fuzzwork_prices_jita.csv"
    price_df.to_csv(output_file, index=False)
    logger.info(f"Prices saved to: {output_file}")
    
    # Also save as JSON for easy programmatic access
    output_json = "fuzzwork_prices_jita.json"
    with open(output_json, 'w') as f:
        json.dump(prices, f, indent=2)
    logger.info(f"Prices saved to: {output_json}")
    
    # Print summary statistics
    logger.info("=" * 60)
    logger.info("Summary Statistics:")
    logger.info(f"  Total items: {len(all_type_ids)}")
    logger.info(f"  Items with buy orders: {len([p for p in prices.values() if p['buy_max'] > 0])}")
    logger.info(f"  Items with sell orders: {len([p for p in prices.values() if p['sell_min'] > 0])}")
    logger.info(f"  Items with no market data: {len([p for p in prices.values() if p['buy_max'] == 0 and p['sell_min'] == 0])}")
    logger.info("=" * 60)
    logger.info("SUCCESS! Price data ready for Excel import")
    logger.info(f"Review the data in: {output_file}")
    logger.info("=" * 60)

if __name__ == "__main__":
    main()

