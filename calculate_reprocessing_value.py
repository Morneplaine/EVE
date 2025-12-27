"""
Calculate reprocessing value for modules.

This module provides functions to calculate the value of reprocessing a module
into minerals, taking into account:
- Reprocessing yield (default 55%)
- Module buy price (highest buy order + markup)
- Mineral sell prices (lowest sell order)
"""

import sqlite3
import pandas as pd
import logging
import sys
from pathlib import Path

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

DATABASE_FILE = "eve_manufacturing.db"


def calculate_reprocessing_value(
    module_type_id=None,
    module_name=None,
    yield_percent=55.0,  # Default: 55% reprocessing yield
    buy_order_markup_percent=10.0,  # Default: 10% markup as safety cushion for market valuation and other selling costs not addressed yet
    num_modules=100,  # Default: 100 modules
    reprocessing_cost_percent=3.37,  # Default: 3.37% base reprocessing cost
    module_price_type='buy_max',  # 'buy_max', 'sell_min', or 'average'
    mineral_price_type='buy_max',  # 'buy_max' or 'sell_min'
    db_file=DATABASE_FILE
):
    """
    Calculate the reprocessing value of a module.
    
    The reprocessing value is calculated as:
    - Module price = selected price type (buy_max, sell_min, or average) * (1 + markup if buy_max)
    - Mineral price = selected price type (buy_max or sell_min) for each mineral
    - Reprocessing cost = total_module_price * (reprocessing_cost_percent / 100) * (yield_percent / 100)
    - Reprocessing value = sum(mineral_quantity * yield * mineral_price) - total_module_price - reprocessing_cost
    
    Args:
        module_type_id (int, optional): TypeID of the module
        module_name (str, optional): Name of the module
        yield_percent (float): Reprocessing yield percentage (default: 55.0)
        buy_order_markup_percent (float): Markup percentage to add to buy_max price (default: 10.0, only used if module_price_type='buy_max')
            This markup serves as a safety cushion for market valuation and other selling costs not addressed yet.
        num_modules (int): Number of modules to reprocess (default: 100)
        reprocessing_cost_percent (float): Reprocessing cost as percentage of buy price (default: 3.37)
        module_price_type (str): Price type for module - 'buy_max', 'sell_min', or 'average' (default: 'buy_max')
        mineral_price_type (str): Price type for minerals - 'buy_max' or 'sell_min' (default: 'buy_max')
        db_file (str): Path to the database file
        
    Returns:
        dict: Dictionary containing:
            - module_type_id: TypeID of the module
            - module_name: Name of the module
            - module_price_type: Price type used for module
            - mineral_price_type: Price type used for minerals
            - module_price_before_markup: Base module price (before markup if applicable)
            - module_price: Final module price per module
            - num_modules: Number of modules being reprocessed
            - total_module_price: Total price for all modules
            - reprocessing_cost: Total reprocessing cost
            - yield_percent: Yield used in calculation
            - reprocessing_outputs: List of dicts with material details (rounded down)
            - total_mineral_value: Total value of reprocessed minerals
            - reprocessing_value: Net value (mineral value - total module price - reprocessing cost)
            - profit_margin_percent: Profit margin percentage
            - error: Error message if any
    """
    
    if not Path(db_file).exists():
        return {
            'error': f"Database file not found: {db_file}"
        }
    
    conn = sqlite3.connect(db_file)
    
    try:
        # Find module by typeID or name
        if module_type_id:
            query = "SELECT typeID, typeName FROM items WHERE typeID = ?"
            params = (module_type_id,)
        elif module_name:
            query = "SELECT typeID, typeName FROM items WHERE typeName = ?"
            params = (module_name,)
        else:
            return {'error': "Either module_type_id or module_name must be provided"}
        
        module_df = pd.read_sql_query(query, conn, params=params)
        
        if len(module_df) == 0:
            return {'error': f"Module not found: {module_type_id or module_name}"}
        
        if len(module_df) > 1:
            return {'error': f"Multiple modules found with name: {module_name}"}
        
        module_type_id = int(module_df.iloc[0]['typeID'])
        module_name = module_df.iloc[0]['typeName']
        
        # Ensure num_modules is an integer
        num_modules = int(num_modules)
        
        # Get reprocessing outputs
        reprocessing_query = """
            SELECT 
                materialTypeID,
                materialName,
                quantity
            FROM reprocessing_outputs
            WHERE itemTypeID = ?
        """
        reprocessing_df = pd.read_sql_query(reprocessing_query, conn, params=(module_type_id,))
        
        if len(reprocessing_df) == 0:
            return {
                'module_type_id': module_type_id,
                'module_name': module_name,
                'error': "This module cannot be reprocessed (no reprocessing outputs found)"
            }
        
        # Get module price based on selected price type
        price_query = "SELECT buy_max, sell_min FROM prices WHERE typeID = ?"
        price_df = pd.read_sql_query(price_query, conn, params=(module_type_id,))
        
        if len(price_df) == 0:
            module_price_before_markup = 0
            module_price = 0
            logger.warning(f"No price data found for {module_name}, using 0")
        else:
            buy_max = float(price_df.iloc[0]['buy_max']) if price_df.iloc[0]['buy_max'] else 0.0
            sell_min = float(price_df.iloc[0]['sell_min']) if price_df.iloc[0]['sell_min'] else 0.0
            
            # Calculate module price based on selected type
            if module_price_type == 'buy_max':
                module_price_before_markup = buy_max
                # Apply markup as safety cushion for market valuation and other selling costs not addressed yet
                module_price = buy_max * (1 + buy_order_markup_percent / 100) if buy_max > 0 else 0
            elif module_price_type == 'sell_min':
                module_price_before_markup = sell_min
                module_price = sell_min  # No markup for sell price
            elif module_price_type == 'average':
                if buy_max > 0 and sell_min > 0:
                    module_price_before_markup = (buy_max + sell_min) / 2
                    module_price = module_price_before_markup
                elif buy_max > 0:
                    module_price_before_markup = buy_max
                    module_price = buy_max
                elif sell_min > 0:
                    module_price_before_markup = sell_min
                    module_price = sell_min
                else:
                    module_price_before_markup = 0
                    module_price = 0
            else:
                logger.warning(f"Invalid module_price_type '{module_price_type}', using 'buy_max'")
                module_price_before_markup = buy_max
                module_price = buy_max * (1 + buy_order_markup_percent / 100) if buy_max > 0 else 0
        
        # Get mineral prices based on selected price type
        material_type_ids = reprocessing_df['materialTypeID'].tolist()
        placeholders = ','.join(['?'] * len(material_type_ids))
        
        if mineral_price_type == 'buy_max':
            price_column = 'buy_max'
        elif mineral_price_type == 'sell_min':
            price_column = 'sell_min'
        else:
            logger.warning(f"Invalid mineral_price_type '{mineral_price_type}', using 'buy_max'")
            price_column = 'buy_max'
        
        mineral_price_query = f"""
            SELECT typeID, buy_max, sell_min
            FROM prices
            WHERE typeID IN ({placeholders})
        """
        mineral_prices_df = pd.read_sql_query(mineral_price_query, conn, params=material_type_ids)
        
        # Create price lookup based on selected type
        mineral_price_lookup = {}
        for _, row in mineral_prices_df.iterrows():
            type_id = int(row['typeID'])
            if price_column == 'buy_max':
                price = float(row['buy_max']) if row['buy_max'] else 0.0
            else:  # sell_min
                price = float(row['sell_min']) if row['sell_min'] else 0.0
            mineral_price_lookup[type_id] = price
        
        # Calculate total module price for all modules
        total_module_price = module_price * num_modules
        
        # Calculate reprocessing cost (base cost percentage × yield percentage)
        # Example: 3.37% × 55% = 1.8535% of buy price
        effective_reprocessing_cost_percent = reprocessing_cost_percent * (yield_percent / 100.0)
        reprocessing_cost = total_module_price * (effective_reprocessing_cost_percent / 100.0)
        
        # Calculate reprocessing value
        yield_multiplier = yield_percent / 100.0
        total_mineral_value = 0.0
        reprocessing_outputs = []
        
        for _, row in reprocessing_df.iterrows():
            material_type_id = int(row['materialTypeID'])
            material_name = row['materialName']
            base_quantity = int(row['quantity'])
            
            # Apply yield and multiply by number of modules
            actual_quantity = base_quantity * yield_multiplier * num_modules
            
            # Round down to integer
            actual_quantity = int(actual_quantity)
            
            # Get mineral price
            mineral_price = mineral_price_lookup.get(material_type_id, 0.0)
            
            # Calculate value for this mineral
            mineral_value = actual_quantity * mineral_price
            
            total_mineral_value += mineral_value
            
            reprocessing_outputs.append({
                'materialTypeID': material_type_id,
                'materialName': material_name,
                'baseQuantity': base_quantity,
                'baseQuantityPerModule': base_quantity * yield_multiplier,  # Per module after yield
                'actualQuantity': actual_quantity,
                'mineralPrice': mineral_price,
                'mineralValue': mineral_value
            })
        
        # Calculate net reprocessing value (mineral value - module price - reprocessing cost)
        reprocessing_value = total_mineral_value - total_module_price - reprocessing_cost
        
        # Calculate profit margin
        if total_module_price > 0:
            profit_margin_percent = (reprocessing_value / total_module_price) * 100
        else:
            profit_margin_percent = float('inf') if reprocessing_value > 0 else 0.0
        
        result = {
            'module_type_id': module_type_id,
            'module_name': module_name,
            'module_price_type': module_price_type,
            'mineral_price_type': mineral_price_type,
            'module_price_before_markup': module_price_before_markup,
            'module_price': module_price,
            'buy_order_markup_percent': buy_order_markup_percent if module_price_type == 'buy_max' else 0,
            'num_modules': num_modules,
            'total_module_price': total_module_price,
            'reprocessing_cost_percent': reprocessing_cost_percent,
            'effective_reprocessing_cost_percent': effective_reprocessing_cost_percent,
            'reprocessing_cost': reprocessing_cost,
            'yield_percent': yield_percent,
            'reprocessing_outputs': reprocessing_outputs,
            'total_mineral_value': total_mineral_value,
            'reprocessing_value': reprocessing_value,
            'profit_margin_percent': profit_margin_percent
        }
        
        return result
        
    except Exception as e:
        logger.error(f"Error calculating reprocessing value: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return {'error': str(e)}
    
    finally:
        conn.close()


def format_reprocessing_result(result):
    """
    Format reprocessing calculation result for display.
    
    Args:
        result (dict): Result from calculate_reprocessing_value
        
    Returns:
        str: Formatted string
    """
    if 'error' in result:
        return f"ERROR: {result['error']}"
    
    output = []
    output = []
    output.append("=" * 60)
    output.append(f"Reprocessing Value Calculation")
    output.append("=" * 60)
    output.append(f"Module: {result['module_name']} (TypeID: {result['module_type_id']})")
    output.append(f"Number of Modules: {result['num_modules']}")
    output.append(f"")
    output.append(f"Price Settings:")
    output.append(f"  Module Price Type: {result['module_price_type']}")
    output.append(f"  Mineral Price Type: {result['mineral_price_type']}")
    output.append(f"")
    
    # Show module price details based on price type
    if result['module_price_type'] == 'buy_max':
        output.append(f"Module Price (before markup): {result.get('module_price_before_markup', 0):,.2f} ISK per module")
        output.append(f"Module Price (after markup):   {result['module_price']:,.2f} ISK per module")
        output.append(f"  (Highest buy order + {result.get('buy_order_markup_percent', 0):.1f}%)")
    elif result['module_price_type'] == 'sell_min':
        output.append(f"Module Price: {result['module_price']:,.2f} ISK per module")
        output.append(f"  (Lowest sell order)")
    elif result['module_price_type'] == 'average':
        output.append(f"Module Price: {result['module_price']:,.2f} ISK per module")
        output.append(f"  (Average of buy_max and sell_min)")
    
    output.append(f"Total Module Price:                 {result['total_module_price']:,.2f} ISK")
    output.append(f"Reprocessing Cost:                  {result['reprocessing_cost']:,.2f} ISK")
    output.append(f"  (Base: {result['reprocessing_cost_percent']:.2f}% × Yield: {result['yield_percent']:.1f}% = {result.get('effective_reprocessing_cost_percent', 0):.4f}%)")
    output.append(f"")
    output.append(f"Reprocessing Yield: {result['yield_percent']:.1f}%")
    output.append(f"Note: Mineral quantities are rounded down to integers")
    output.append(f"")
    output.append("Reprocessing Outputs:")
    output.append("-" * 60)
    output.append(f"{'Mineral':<30} {'Base':>6} {'Per Mod':>8} {'× Mods':>6} {'Rounded':>8} {'Price':>12} {'Value':>15}")
    output.append("-" * 60)
    
    for i, output_mat in enumerate(result['reprocessing_outputs']):
        base_qty = output_mat['baseQuantity']
        per_module = output_mat.get('baseQuantityPerModule', 0)
        num_mods = result['num_modules']
        rounded_qty = int(output_mat['actualQuantity'])
        price = output_mat['mineralPrice']
        value = output_mat['mineralValue']
        
        output.append(
            f"  {output_mat['materialName']:<28} "
            f"{base_qty:6d} "
            f"{per_module:8.2f} "
            f"× {num_mods:3d} "
            f"{rounded_qty:8d} "
            f"{price:12,.2f} "
            f"{value:15,.2f}"
        )
    
    output.append("-" * 60)
    output.append(f"Total Mineral Value: {result['total_mineral_value']:,.2f} ISK")
    output.append(f"Total Module Price:  {result['total_module_price']:,.2f} ISK")
    output.append(f"Reprocessing Cost:   {result['reprocessing_cost']:,.2f} ISK")
    output.append(f"Reprocessing Value:  {result['reprocessing_value']:,.2f} ISK")
    
    if result['profit_margin_percent'] != float('inf'):
        output.append(f"Profit Margin:      {result['profit_margin_percent']:+.2f}%")
    else:
        output.append(f"Profit Margin:      N/A (module buy price is 0)")
    
    output.append("=" * 60)
    
    return "\n".join(output)


def main():
    """Command-line interface for reprocessing value calculation"""
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python calculate_reprocessing_value.py <module_name_or_typeID> [yield_percent] [buy_markup_percent] [num_modules] [reprocessing_cost_percent] [module_price_type] [mineral_price_type]")
        print("")
        print("Price Types:")
        print("  module_price_type: 'buy_max' (default), 'sell_min', or 'average'")
        print("  mineral_price_type: 'buy_max' (default) or 'sell_min'")
        print("")
        print("Examples:")
        print("  python calculate_reprocessing_value.py \"Medium Shield Booster II\"")
        print("  python calculate_reprocessing_value.py 11269 55 10")
        print("  python calculate_reprocessing_value.py \"Medium Shield Booster II\" 60 15 200")
        print("  python calculate_reprocessing_value.py \"Iron Charge S\" 55 10 100 3.37")
        print("  python calculate_reprocessing_value.py \"Gamma L\" 55 10 9 3.37 sell_min sell_min")
        print("  python calculate_reprocessing_value.py \"Gamma L\" 55 10 9 3.37 average buy_max")
        sys.exit(1)
    
    # Parse arguments
    module_arg = sys.argv[1]
    
    # Try to parse as typeID (integer)
    try:
        module_type_id = int(module_arg)
        module_name = None
    except ValueError:
        module_type_id = None
        module_name = module_arg
    
    yield_percent = float(sys.argv[2]) if len(sys.argv) > 2 else 55.0
    buy_markup_percent = float(sys.argv[3]) if len(sys.argv) > 3 else 10.0
    num_modules = int(sys.argv[4]) if len(sys.argv) > 4 else 100
    reprocessing_cost_percent = float(sys.argv[5]) if len(sys.argv) > 5 else 3.37
    module_price_type = sys.argv[6] if len(sys.argv) > 6 else 'buy_max'
    mineral_price_type = sys.argv[7] if len(sys.argv) > 7 else 'buy_max'
    
    # Calculate reprocessing value
    result = calculate_reprocessing_value(
        module_type_id=module_type_id,
        module_name=module_name,
        yield_percent=yield_percent,
        buy_order_markup_percent=buy_markup_percent,
        num_modules=num_modules,
        reprocessing_cost_percent=reprocessing_cost_percent,
        module_price_type=module_price_type,
        mineral_price_type=mineral_price_type
    )
    
    # Display result
    print(format_reprocessing_result(result))
    
    if 'error' in result:
        sys.exit(1)


def analyze_all_modules(
    yield_percent=55.0,
    buy_order_markup_percent=10.0,
    num_modules=100,
    reprocessing_cost_percent=3.37,
    module_price_type='sell_min',  # Default: use sell_min for evaluation
    mineral_price_type='buy_max',
    min_module_price=1.0,  # Minimum module price to filter out unrealistic prices
    max_module_price=100000.0,
    top_n=30,
    db_file=DATABASE_FILE
):
    """
    Analyze reprocessing value for all modules in the database.
    
    Args:
        yield_percent (float): Reprocessing yield percentage (default: 55.0)
        buy_order_markup_percent (float): Markup percentage (default: 10.0)
        num_modules (int): Number of modules to reprocess (default: 100)
        reprocessing_cost_percent (float): Reprocessing cost percentage (default: 3.37)
        module_price_type (str): Price type for module (default: 'buy_max')
        mineral_price_type (str): Price type for minerals (default: 'buy_max')
        max_module_price (float): Maximum module price to include (default: 100000.0)
        top_n (int): Number of top results to return (default: 30)
        db_file (str): Path to the database file
        
    Returns:
        list: List of dicts with top results sorted by return percentage
    """
    logger.info("=" * 60)
    logger.info("Analyzing reprocessing value for all modules")
    logger.info("=" * 60)
    logger.info(f"Parameters:")
    logger.info(f"  Yield: {yield_percent}%")
    logger.info(f"  Markup: {buy_order_markup_percent}%")
    logger.info(f"  Modules per batch: {num_modules}")
    logger.info(f"  Min sell_min price: {min_module_price:,.0f} ISK")
    logger.info(f"  Max sell_min price: {max_module_price:,.0f} ISK")
    logger.info(f"  Module price type: {module_price_type}")
    logger.info(f"  Mineral price type: {mineral_price_type}")
    logger.info(f"  Top N results: {top_n}")
    logger.info("=" * 60)
    
    if not Path(db_file).exists():
        logger.error(f"Database file not found: {db_file}")
        return []
    
    conn = sqlite3.connect(db_file)
    
    try:
        # Get all modules that can be reprocessed
        query = """
            SELECT DISTINCT ro.itemTypeID, ro.itemName
            FROM reprocessing_outputs ro
            ORDER BY ro.itemName
        """
        modules_df = pd.read_sql_query(query, conn)
        
        logger.info(f"Found {len(modules_df)} modules that can be reprocessed")
        logger.info("Calculating reprocessing values...")
        logger.info("This may take several minutes...")
        
        results = []
        processed = 0
        
        for idx, row in modules_df.iterrows():
            module_type_id = int(row['itemTypeID'])
            module_name = row['itemName']
            
            # Get module prices for filtering and display
            price_query = "SELECT buy_max, sell_min FROM prices WHERE typeID = ?"
            price_df = pd.read_sql_query(price_query, conn, params=(module_type_id,))
            
            if len(price_df) == 0:
                continue
            
            buy_max = float(price_df.iloc[0]['buy_max']) if price_df.iloc[0]['buy_max'] else 0.0
            sell_min = float(price_df.iloc[0]['sell_min']) if price_df.iloc[0]['sell_min'] else 0.0
            
            # Filter by sell_min price (max_module_price refers to min sell price)
            if sell_min < min_module_price or sell_min > max_module_price:
                continue
            
            # Calculate reprocessing value
            result = calculate_reprocessing_value(
                module_type_id=module_type_id,
                yield_percent=yield_percent,
                buy_order_markup_percent=buy_order_markup_percent,
                num_modules=num_modules,
                reprocessing_cost_percent=reprocessing_cost_percent,
                module_price_type=module_price_type,
                mineral_price_type=mineral_price_type,
                db_file=db_file
            )
            
            if 'error' in result:
                continue
            
            # Only include if we have valid prices
            if result['module_price'] == 0 or result['total_mineral_value'] == 0:
                continue
            
            # Calculate profit per item and return percentage based on sell_min
            # Expected buy price = buy_max + markup (for comparison)
            expected_buy_price = buy_max * (1 + buy_order_markup_percent / 100) if buy_max > 0 else 0
            
            # Profit per item = (total reprocessing value) / num_modules
            profit_per_item = result['reprocessing_value'] / num_modules if num_modules > 0 else 0
            
            # Return percentage based on sell_min price
            if sell_min > 0:
                return_percent_sell = (profit_per_item / sell_min) * 100
            else:
                return_percent_sell = float('inf') if profit_per_item > 0 else 0.0
            
            results.append({
                'module_name': module_name,
                'module_type_id': module_type_id,
                'expected_buy_price': expected_buy_price,
                'sell_min_price': sell_min,
                'module_price': result['module_price'],  # Price used in calculation
                'total_module_price': result['total_module_price'],
                'total_mineral_value': result['total_mineral_value'],
                'reprocessing_value': result['reprocessing_value'],
                'profit_per_item': profit_per_item,
                'return_percent': result['profit_margin_percent'],  # Based on module_price used
                'return_percent_sell': return_percent_sell,  # Based on sell_min price
                'num_modules': num_modules
            })
            
            processed += 1
            if processed % 100 == 0:
                logger.info(f"Processed {processed}/{len(modules_df)} modules...")
        
        logger.info(f"Analysis complete! Processed {processed} modules")
        
        # Sort by return percentage based on sell_min (descending)
        results.sort(key=lambda x: x['return_percent_sell'], reverse=True)
        
        # Return top N
        top_results = results[:top_n]
        
        return top_results
        
    except Exception as e:
        logger.error(f"Error analyzing modules: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return []
    
    finally:
        conn.close()


def format_analysis_results(results):
    """Format analysis results for display as a table"""
    if not results:
        return "No results found."
    
    output = []
    output.append("=" * 120)
    output.append(f"Top {len(results)} Modules by Return Percentage")
    output.append("=" * 120)
    
    # Table header
    header = f"{'Rank':<6} {'Module Name':<40} {'Buy Price':>12} {'Sell Min':>12} {'Profit/Item':>15} {'Return %':>12}"
    output.append(header)
    output.append("-" * 120)
    
    # Table rows
    for rank, result in enumerate(results, 1):
        # Format return percentage based on sell_min - cap display at 999,999% for readability
        return_pct_sell = result['return_percent_sell']
        if return_pct_sell > 999999:
            return_str = ">999,999%"
        elif return_pct_sell == float('inf'):
            return_str = "N/A"
        else:
            return_str = f"{return_pct_sell:,.2f}%"
        
        # Truncate module name if too long
        module_name = result['module_name']
        if len(module_name) > 38:
            module_name = module_name[:35] + "..."
        
        row = (
            f"{rank:<6} "
            f"{module_name:<40} "
            f"{result['expected_buy_price']:>12,.2f} "
            f"{result['sell_min_price']:>12,.2f} "
            f"{result['profit_per_item']:>15,.2f} "
            f"{return_str:>12}"
        )
        output.append(row)
    
    output.append("=" * 120)
    
    return "\n".join(output)


def analyze_all_modules_main():
    """Command-line interface for analyzing all modules"""
    import sys
    
    # Parse arguments
    yield_percent = float(sys.argv[1]) if len(sys.argv) > 1 else 55.0
    buy_markup_percent = float(sys.argv[2]) if len(sys.argv) > 2 else 10.0
    num_modules = int(sys.argv[3]) if len(sys.argv) > 3 else 100
    reprocessing_cost_percent = float(sys.argv[4]) if len(sys.argv) > 4 else 3.37
    module_price_type = sys.argv[5] if len(sys.argv) > 5 else 'sell_min'
    mineral_price_type = sys.argv[6] if len(sys.argv) > 6 else 'buy_max'
    min_module_price = float(sys.argv[7]) if len(sys.argv) > 7 else 1.0
    max_module_price = float(sys.argv[8]) if len(sys.argv) > 8 else 100000.0
    top_n = int(sys.argv[9]) if len(sys.argv) > 9 else 30
    
    # Run analysis
    results = analyze_all_modules(
        yield_percent=yield_percent,
        buy_order_markup_percent=buy_markup_percent,
        num_modules=num_modules,
        reprocessing_cost_percent=reprocessing_cost_percent,
        module_price_type=module_price_type,
        mineral_price_type=mineral_price_type,
        min_module_price=min_module_price,
        max_module_price=max_module_price,
        top_n=top_n
    )
    
    # Display results
    print(format_analysis_results(results))
    
    if not results:
        sys.exit(1)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == '--analyze-all':
        # Run analysis for all modules
        sys.argv = sys.argv[1:]  # Remove '--analyze-all'
        analyze_all_modules_main()
    else:
        # Run single module analysis
        main()

