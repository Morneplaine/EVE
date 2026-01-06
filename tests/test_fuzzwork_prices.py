"""
Test script for Fuzzwork Market API price fetching
Tests the get_fuzzwork_market_prices function with a small sample of items
"""

import sys
from eve_manufacturing_database import get_fuzzwork_market_prices

# Test with a small sample of common items
# These are some common minerals and materials
test_type_ids = [
    34,   # Tritanium
    35,   # Pyerite
    36,   # Mexallon
    37,   # Isogen
    38,   # Nocxium
    39,   # Zydrine
    40,   # Megacyte
]

if __name__ == "__main__":
    print("Testing Fuzzwork Market API for Jita (30000142)")
    print("=" * 60)
    
    # Fetch prices
    prices = get_fuzzwork_market_prices(test_type_ids, station_id=30000142)
    
    print("\nResults:")
    print("=" * 60)
    print(f"{'TypeID':<10} {'Buy Max':<15} {'Buy Volume':<15} {'Sell Min':<15} {'Sell Volume':<15}")
    print("-" * 60)
    
    for type_id in test_type_ids:
        if type_id in prices:
            p = prices[type_id]
            print(f"{type_id:<10} {p['buy_max']:<15.2f} {p['buy_volume']:<15,.0f} {p['sell_min']:<15.2f} {p['sell_volume']:<15,.0f}")
        else:
            print(f"{type_id:<10} {'N/A':<15} {'N/A':<15} {'N/A':<15} {'N/A':<15}")
    
    print("\n" + "=" * 60)
    print("Sample data structure for one item:")
    if test_type_ids[0] in prices:
        import json
        print(json.dumps({test_type_ids[0]: prices[test_type_ids[0]]}, indent=2))

