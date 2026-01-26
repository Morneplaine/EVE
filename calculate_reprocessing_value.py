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
from assumptions import (
    BROKER_FEE,
    SALES_TAX,
    LISTING_RELIST,
    REPROCESSING_COST,
    DEFAULT_YIELD_PERCENT,
    BUY_ORDER_MARKUP_PERCENT,
    BUY_BUFFER,
    RELIST_DISCOUNT

)

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

DATABASE_FILE = "eve_manufacturing.db"

def buy_order_with_fees(buy_price, broker_fee=BROKER_FEE, sales_tax=SALES_TAX, buy_buffer=BUY_BUFFER, average_relist=LISTING_RELIST,RELIST_DISCOUNT=RELIST_DISCOUNT):
    # we assume that at some point we will have to increse the buy order on average we will increase it by the buffer
    buy_price_loaded = buy_price*(buy_buffer+1)
    # since we may have to relist multiple times there is a relist fee, we do not have to take into account the price increase fee since it is embeded in the buffer but the relist fee is the amount of time we relist but there is a discount based on skills

    relist_fees = buy_price_loaded*(broker_fee/100)*((1-RELIST_DISCOUNT)/100)*average_relist
    #broker fee is broker % times total sale price
    broker_fee= buy_price_loaded*broker_fee/100
    #we add it all
    total_buy_order_cost = buy_price_loaded + relist_fees + broker_fee
    
    return total_buy_order_cost

def sell_order_with_fees(sell_price, sales_tax=SALES_TAX, average_relist=LISTING_RELIST,RELIST_DISCOUNT=RELIST_DISCOUNT):
    # see the mechanism on buy order with fee the difference is we also have to pay the sales tax 
    sell_price_loaded = sell_price*(1+sales_tax/100)
    relist_fees = sell_price_loaded*(broker_fee/100)*((1-RELIST_DISCOUNT)/100)*average_relist
    broker_fee= sell_price_loaded*broker_fee/100
    sales_tax= sell_price_loaded*sales_tax/100
    total_sell_order_cost = sell_price_loaded + relist_fees + broker_fee + sales_tax

    return total_sell_order_cost

def sell_into_buy_order(sell_price, sales_tax=SALES_TAX):
    #when sellling into a buy order there is no s
    return sell_price*(1+sales_tax/100)

def buy_into_sell_order(buy_price):
    #when buying into a sell order there is no sales tax
    return buy_price

def calculate_reprocessing_value(
    module_type_id=None,
    module_name=None,
    yield_percent=DEFAULT_YIELD_PERCENT,  # Default: 55% reprocessing yield
    broker_fee=BROKER_FEE, 
    sales_tax=SALES_TAX, 
    buy_buffer=BUY_BUFFER,
    average_relist=LISTING_RELIST,
    buy_order_markup_percent=BUY_ORDER_MARKUP_PERCENT,  # Markup percentage for buy_max price
    reprocessing_cost_percent=REPROCESSING_COST,  # Default: 3.37% base reprocessing cost
    module_price_type='buy_max',  # 'buy_max' or 'sell_min'
    mineral_price_type='buy_max',  # 'buy_max' or 'sell_min'
    db_file=DATABASE_FILE
):
    
    #when buying at min sell there is no broker fee or listing fee
    #when buying with buy order there is 
    
    """
    Calculate the reprocessing value of a module.
    
    The reprocessing value is calculated as:
    - Module price = selected price type (buy_max or sell_min) * (1 + markup if buy_max)
    - Mineral price = selected price type (buy_max or sell_min) for each mineral
    - Reprocessing cost = total_module_price * (reprocessing_cost_percent / 100) * (yield_percent / 100)
    - Reprocessing value = sum(mineral_quantity * yield * mineral_price) - total_module_price - reprocessing_cost
    
    Args:
        module_type_id (int, optional): TypeID of the module
        module_name (str, optional): Name of the module
        yield_percent (float): Reprocessing yield percentage (default: 55.0)
        buy_order_markup_percent (float): Markup percentage to add to buy_max price (default: 10.0, only used if module_price_type='buy_max')
            This markup serves as a safety cushion for market valuation and other selling costs not addressed yet.
        reprocessing_cost_percent (float): Reprocessing cost as percentage of buy price (default: 3.37)
        module_price_type (str): Price type for module - 'buy_max' or 'sell_min' (default: 'buy_max')
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
            - output_quantity: Number of items needed to obtain reprocessing result (from blueprints table)
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
        
        # ========================================================================
        # GET REPROCESSING OUTPUTS FROM DATABASE
        # ========================================================================
        # Query the database to get all materials that this module reprocesses into.
        # Each row contains:
        #   - materialTypeID: The type ID of the output material
        #   - materialName: Name of the material (e.g., "Morphite", "Tritanium")
        #   - quantity: The batch quantity (total quantity for batch_size items)
        #   - batch_size: Number of items in a batch (e.g., 100 for charges, 1 for modules)
        # ========================================================================
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
                module_price_post_transaction_costs = buy_max*(buy_buffer+1)*(1+broker_fee/100)**(average_relist)
                
            elif module_price_type == 'sell_min':
                module_price_before_markup = sell_min
                module_price_post_transaction_costs= module_price_before_markup
                
            else:
                logger.warning(f"Invalid module_price_type '{module_price_type}', using 'buy_max'")
                module_price_before_markup = buy_max
                module_price_post_transaction_costs = buy_max*(buy_buffer+1)*(1+broker_fee/100)**(average_relist)
                
        
        
        module_price = buy_max * (1 + buy_order_markup_percent / 100) if buy_max > 0 else 0
        # Apply markup as safety cushion for market valuation and other selling costs not addressed yet
        
        
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
        
        
        # Get outputQuantity from blueprints table using productTypeID
        # This is the number of items needed to obtain the reprocessing result
        output_quantity_query = """
            SELECT outputQuantity FROM blueprints WHERE productTypeID = ?
        """
        output_quantity_df = pd.read_sql_query(output_quantity_query, conn, params=(module_type_id,))
        
        if len(output_quantity_df) > 0:
            output_quantity = int(output_quantity_df.iloc[0]['outputQuantity'])
        else:
            output_quantity = 1

        # Calculate total module price for all modules
        total_module_price = module_price * output_quantity
        
        # Calculate reprocessing cost (base cost percentage × yield percentage)
        # Example: 3.37% × 55% = 1.8535% of buy price
        effective_reprocessing_cost_percent = reprocessing_cost_percent * (yield_percent / 100.0)
        reprocessing_cost = total_module_price * (effective_reprocessing_cost_percent / 100.0)
        
        # ========================================================================
        # CALCULATE REPROCESSING OUTPUTS
        # ========================================================================
        # This section calculates the mineral outputs from reprocessing the modules.
        # 
        # Formula:
        #   1. quantity_per_item = batch_quantity / batch_size
        #      (e.g., 30 Morphite / 100 items = 0.3 per item)
        #   
        #   2. quantity_per_module = quantity_per_item * yield_percent / 100
        #      (e.g., 0.3 * 0.55 = 0.165 per module after 55% yield)
        #   
        #   3. actual_quantity = quantity_per_module * output_quantity (rounded down)
        #      (e.g., 0.165 per item * 100 items = 16.5 → rounds to 16)
        #
        # Rounding happens ONLY at the final step to ensure accuracy.
        # ========================================================================
        
        yield_multiplier = yield_percent / 100.0
        total_mineral_value = 0.0
        reprocessing_outputs = []  # List to store all mineral outputs
        
        # Iterate through each material that this module reprocesses into
        for _, row in reprocessing_df.iterrows():
            material_type_id = int(row['materialTypeID'])
            material_name = row['materialName']
            batch_quantity = int(row['quantity'])  # Total quantity for the batch (from database)
            batch_size = int(row.get('batch_size', 1))  # Number of items in the batch (e.g., 100 for charges)
            
            # Step 1: Calculate quantity per single item
            # The batch_quantity in the database represents the base quantity from reprocessing batch_size items
            # Divide by batch_size to get per-item quantity
            # Example: 30 Morphite for batch_size=100 → 30/100 = 0.3 per item
            quantity_per_item = batch_quantity / batch_size if batch_size > 0 else batch_quantity
            
            # Step 2: Apply yield to get quantity per module (after yield)
            # Reprocessing yield reduces the output (default 55%)
            # Example: 0.3 * 0.55 = 0.165 per module
            quantity_per_module = quantity_per_item * yield_multiplier
            
            # Step 3: Calculate total quantity for output_quantity items (keep as float for precision)
            # Multiply per-item quantity by output_quantity (number of items needed for reprocessing)
            # Example: 0.165 per item * 100 items = 16.5
            # Since we're using output_quantity (typically 100), there should be no rounding issues
            actual_quantity_float = quantity_per_module * output_quantity
            
            # Step 4: Round down to integer ONLY at the final step
            # This ensures proper rounding: 16.5 rounds to 16
            # If we rounded per-item first (0.165 → 0), we'd lose precision
            # Note: With output_quantity = 100, there should typically be no rounding needed
            actual_quantity = int(actual_quantity_float)
            
            # Get mineral price from price lookup
            mineral_price = mineral_price_lookup.get(material_type_id, 0.0)
            
            # Calculate total value for this mineral
            # Value = quantity * price
            mineral_value = actual_quantity * mineral_price
            
            # Add to total mineral value
            total_mineral_value += mineral_value
            
            # Store all the calculated data for this material
            reprocessing_outputs.append({
                'materialTypeID': material_type_id,           # Material type ID
                'materialName': material_name,                 # Material name (e.g., "Morphite")
                'batchQuantity': batch_quantity,               # Original batch quantity from database
                'batchSize': batch_size,                       # Batch size (e.g., 100 for charges)
                'quantityPerItem': quantity_per_item,          # Quantity per single item (before yield)
                'baseQuantityPerModule': quantity_per_module,  # Per module after yield (before rounding)
                'actualQuantity': actual_quantity,              # Total for output_quantity after yield and rounding
                'actualQuantityFloat': actual_quantity_float, # Total before rounding (for display/debugging)
                'mineralPrice': mineral_price,                  # Price per unit of this mineral
                'mineralValue': mineral_value                   # Total value (quantity * price)
            })
        
        # Calculate net reprocessing value (mineral value - module price - reprocessing cost)
        reprocessing_value = total_mineral_value - total_module_price - reprocessing_cost
        
        # Calculate profit margin
        if total_module_price > 0:
            profit_margin_percent = (reprocessing_value / total_module_price) * 100
        else:
            profit_margin_percent = float('inf') if reprocessing_value > 0 else 0.0
        
        # Calculate breakeven price (maximum purchase price before markup for 0 profit)
        # For breakeven: total_mineral_value = total_module_price + reprocessing_cost
        # reprocessing_cost = total_module_price * effective_reprocessing_cost_percent / 100
        # So: total_mineral_value = total_module_price * (1 + effective_reprocessing_cost_percent / 100)
        # Therefore: total_module_price_breakeven = total_mineral_value / (1 + effective_reprocessing_cost_percent / 100)
        if total_mineral_value > 0 and output_quantity > 0:
            total_module_price_breakeven = total_mineral_value / (1 + effective_reprocessing_cost_percent / 100.0)
            
            # Calculate per-module price before markup
            # If markup is applied: total_module_price = module_price_before_markup * output_quantity * (1 + markup_percent / 100)
            if module_price_type == 'buy_max' and buy_order_markup_percent > 0:
                module_price_before_markup_breakeven = total_module_price_breakeven / (output_quantity * (1 + buy_order_markup_percent / 100.0))
            else:
                module_price_before_markup_breakeven = total_module_price_breakeven / output_quantity
        else:
            total_module_price_breakeven = 0.0
            module_price_before_markup_breakeven = 0.0
        
        result = {
            'module_type_id': module_type_id,
            'module_name': module_name,
            'module_price_type': module_price_type,
            'mineral_price_type': mineral_price_type,
            'module_price_before_markup': module_price_before_markup,
            'module_price': module_price,
            'buy_order_markup_percent': buy_order_markup_percent if module_price_type == 'buy_max' else 0,
            'output_quantity': output_quantity,
            'total_module_price': total_module_price,
            'reprocessing_cost_percent': reprocessing_cost_percent,
            'effective_reprocessing_cost_percent': effective_reprocessing_cost_percent,
            'reprocessing_cost': reprocessing_cost,
            'yield_percent': yield_percent,
            'reprocessing_outputs': reprocessing_outputs,
            'total_mineral_value': total_mineral_value,
            'reprocessing_value': reprocessing_value,
            'profit_margin_percent': profit_margin_percent,
            'breakeven_total_module_price': total_module_price_breakeven,
            'breakeven_module_price_before_markup': module_price_before_markup_breakeven
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
    output.append(f"Output Quantity: {result['output_quantity']} items (from blueprints table)")
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
    
    output.append(f"Total Module Price:                 {result['total_module_price']:,.2f} ISK")
    output.append(f"Reprocessing Cost:                  {result['reprocessing_cost']:,.2f} ISK")
    output.append(f"  (Base: {result['reprocessing_cost_percent']:.2f}% × Yield: {result['yield_percent']:.1f}% = {result.get('effective_reprocessing_cost_percent', 0):.4f}%)")
    output.append(f"")
    output.append(f"Reprocessing Yield: {result['yield_percent']:.1f}%")
    output.append(f"Note: Mineral quantities are rounded down to integers")
    output.append(f"")
    output.append("Reprocessing Outputs:")
    output.append("-" * 100)
    output.append(f"{'Mineral':<30} {'Batch':>8} {'Batch Size':>10} {'Per Item':>10} {'Per Item':>10} {'× Qty':>6} {'Rounded':>8} {'Price':>12} {'Value':>15}")
    output.append("-" * 100)
    
    for i, output_mat in enumerate(result['reprocessing_outputs']):
        batch_qty = output_mat.get('batchQuantity', output_mat.get('baseQuantity', 0))
        batch_size = output_mat.get('batchSize', 1)
        quantity_per_item = output_mat.get('quantityPerItem', batch_qty / batch_size if batch_size > 0 else 0)
        per_module = output_mat.get('baseQuantityPerModule', 0)
        output_qty = result['output_quantity']
        rounded_qty = int(output_mat['actualQuantity'])
        price = output_mat['mineralPrice']
        value = output_mat['mineralValue']
        
        output.append(
            f"  {output_mat['materialName']:<28} "
            f"{batch_qty:8d} "
            f"× {batch_size:6d} "
            f"{quantity_per_item:10.4f} "
            f"{per_module:10.4f} "
            f"× {output_qty:3d} "
            f"{rounded_qty:8d} "
            f"{price:12,.2f} "
            f"{value:15,.2f}"
        )
    
    output.append("-" * 60)
    
    # Per-item calculations
    output_quantity = result['output_quantity']
    mineral_value_per_item = result['total_mineral_value'] / output_quantity if output_quantity > 0 else 0
    reprocessing_cost_per_item = result['reprocessing_cost'] / output_quantity if output_quantity > 0 else 0
    profit_per_item = mineral_value_per_item - result['module_price'] - reprocessing_cost_per_item
    
    output.append("PER ITEM:")
    output.append(f"  Mineral Value (after yield): {mineral_value_per_item:,.2f} ISK")
    output.append(f"  Module Price:                {result['module_price']:,.2f} ISK")
    output.append(f"  Reprocessing Cost:           {reprocessing_cost_per_item:,.2f} ISK")
    output.append(f"  Profit per Item:             {profit_per_item:+,.2f} ISK")
    
    # Calculate return percentage per item: (mineral_value - module_price) / module_price
    if result['module_price'] > 0:
        net_mineral_value_per_item = mineral_value_per_item - reprocessing_cost_per_item
        return_percent_per_item = ((net_mineral_value_per_item - result['module_price']) / result['module_price']) * 100
        output.append(f"  Return %:                     {return_percent_per_item:+.2f}%")
    else:
        output.append(f"  Return %:                     N/A (module price is 0)")
    
    output.append("")
    output.append(f"TOTAL (for {output_quantity} items):")
    output.append(f"  Total Mineral Value: {result['total_mineral_value']:,.2f} ISK")
    output.append(f"  Total Module Price:  {result['total_module_price']:,.2f} ISK ({result['module_price']:,.2f} × {output_quantity})")
    output.append(f"  Reprocessing Cost:   {result['reprocessing_cost']:,.2f} ISK")
    output.append(f"  Reprocessing Value:  {result['reprocessing_value']:,.2f} ISK")
    
    if result['profit_margin_percent'] != float('inf'):
        output.append(f"  Profit Margin:      {result['profit_margin_percent']:+.2f}%")
    else:
        output.append(f"  Profit Margin:      N/A (module buy price is 0)")
    
    output.append("")
    output.append("BREAKEVEN ANALYSIS:")
    if result.get('breakeven_module_price_before_markup', 0) > 0:
        breakeven_price = result['breakeven_module_price_before_markup']
        current_price = result.get('module_price_before_markup', 0)
        if current_price > 0:
            price_difference = breakeven_price - current_price
            price_difference_percent = (price_difference / current_price) * 100 if current_price > 0 else 0
            output.append(f"  Max Purchase Price (before markup) for breakeven: {breakeven_price:,.2f} ISK per module")
            output.append(f"  Current Price (before markup):                  {current_price:,.2f} ISK per module")
            output.append(f"  Price Difference:                                 {price_difference:+,.2f} ISK ({price_difference_percent:+.2f}%)")
            if result['module_price_type'] == 'buy_max' and result.get('buy_order_markup_percent', 0) > 0:
                breakeven_with_markup = breakeven_price * (1 + result['buy_order_markup_percent'] / 100.0)
                output.append(f"  Max Purchase Price (with {result['buy_order_markup_percent']:.1f}% markup): {breakeven_with_markup:,.2f} ISK per module")
        else:
            output.append(f"  Max Purchase Price (before markup) for breakeven: {breakeven_price:,.2f} ISK per module")
    else:
        output.append(f"  Breakeven price: N/A (no mineral value)")
    
    output.append("=" * 60)
    
    return "\n".join(output)


def main():
    """Command-line interface for reprocessing value calculation"""
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python calculate_reprocessing_value.py <module_name_or_typeID> [yield_percent] [buy_markup_percent] [reprocessing_cost_percent] [module_price_type] [mineral_price_type]")
        print("")
        print("Note: output_quantity is automatically fetched from blueprints table based on productTypeID")
        print("")
        print("Price Types:")
        print("  module_price_type: 'buy_max' (default) or 'sell_min'")
        print("  mineral_price_type: 'buy_max' (default) or 'sell_min'")
        print("")
        print("Examples:")
        print("  python calculate_reprocessing_value.py \"Medium Shield Booster II\"")
        print("  python calculate_reprocessing_value.py 11269 55 10")
        print("  python calculate_reprocessing_value.py \"Medium Shield Booster II\" 60 15")
        print("  python calculate_reprocessing_value.py \"Iron Charge S\" 55 10 3.37")
        print("  python calculate_reprocessing_value.py \"Gamma L\" 55 10 3.37 sell_min sell_min")
        print("  python calculate_reprocessing_value.py \"Gamma L\" 55 10 3.37 sell_min buy_max")
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
    
    yield_percent = float(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_YIELD_PERCENT
    buy_markup_percent = float(sys.argv[3]) if len(sys.argv) > 3 else BUY_ORDER_MARKUP_PERCENT
    reprocessing_cost_percent = float(sys.argv[4]) if len(sys.argv) > 4 else REPROCESSING_COST
    module_price_type = sys.argv[5] if len(sys.argv) > 5 else 'buy_max'
    mineral_price_type = sys.argv[6] if len(sys.argv) > 6 else 'buy_max'
    
    # Calculate reprocessing value
    result = calculate_reprocessing_value(
        module_type_id=module_type_id,
        module_name=module_name,
        yield_percent=yield_percent,
        buy_order_markup_percent=buy_markup_percent,
        reprocessing_cost_percent=reprocessing_cost_percent,
        module_price_type=module_price_type,
        mineral_price_type=mineral_price_type
    )
    
    # Display result
    print(format_reprocessing_result(result))
    
    if 'error' in result:
        sys.exit(1)


def analyze_all_modules(
    yield_percent=DEFAULT_YIELD_PERCENT,
    buy_order_markup_percent=BUY_ORDER_MARKUP_PERCENT,
    reprocessing_cost_percent=REPROCESSING_COST,
    module_price_type='sell_min',  # Default: use sell_min for evaluation
    mineral_price_type='buy_max',
    min_module_price=1.0,  # Minimum module price to filter out unrealistic prices
    max_module_price=100000.0,
    top_n=30,
    excluded_module_ids=None,  # Set of module type IDs to exclude
    db_file=DATABASE_FILE
):
    """
    Analyze reprocessing value for all modules in the database.
    
    Args:
        yield_percent (float): Reprocessing yield percentage (default: 55.0)
        buy_order_markup_percent (float): Markup percentage (default: 10.0)
        reprocessing_cost_percent (float): Reprocessing cost percentage (default: 3.37)
        module_price_type (str): Price type for module (default: 'buy_max')
        mineral_price_type (str): Price type for minerals (default: 'buy_max')
        max_module_price (float): Maximum module price to include (default: 100000.0)
        top_n (int): Number of top results to return (default: 30)
        excluded_module_ids (set): Set of module type IDs to exclude (default: None)
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
    logger.info(f"  Note: output_quantity fetched from blueprints table")
    logger.info(f"  Min sell_min price: {min_module_price:,.0f} ISK")
    logger.info(f"  Max sell_min price: {max_module_price:,.0f} ISK")
    logger.info(f"  Module price type: {module_price_type}")
    logger.info(f"  Mineral price type: {mineral_price_type}")
    logger.info(f"  Top N results: {top_n}")
    logger.info("=" * 60)
    
    if not Path(db_file).exists():
        logger.error(f"Database file not found: {db_file}")
        return []
    
    # Initialize excluded_module_ids if None
    if excluded_module_ids is None:
        excluded_module_ids = set()
    
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
        if excluded_module_ids:
            logger.info(f"Excluding {len(excluded_module_ids)} modules")
        logger.info("Calculating reprocessing values...")
        logger.info("This may take several minutes...")
        
        results = []
        processed = 0
        
        for idx, row in modules_df.iterrows():
            module_type_id = int(row['itemTypeID'])
            module_name = row['itemName']
            
            # Skip excluded modules
            if module_type_id in excluded_module_ids:
                continue
            
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
            
            # Calculate profit per item and return percentage
            # Expected buy price = buy_max + markup (for comparison)
            expected_buy_price = buy_max * (1 + buy_order_markup_percent / 100) if buy_max > 0 else 0
            
            # For per-item calculations:
            # - Mineral value per item = total_mineral_value / output_quantity
            # - Module price per item = module_price (already per item)
            # - Reprocessing cost per item = reprocessing_cost / output_quantity
            # - Profit per item = (mineral_value_per_item - module_price - reprocessing_cost_per_item)
            output_quantity = result.get('output_quantity', 1)
            mineral_value_per_item = result['total_mineral_value'] / output_quantity if output_quantity > 0 else 0
            reprocessing_cost_per_item = result['reprocessing_cost'] / output_quantity if output_quantity > 0 else 0
            profit_per_item = mineral_value_per_item - result['module_price'] - reprocessing_cost_per_item
            
            # Return percentage = (mineral_sell_price - buy_price) / buy_price * 100
            # Where mineral_sell_price = mineral_value_per_item (net of reprocessing cost)
            # And buy_price = module_price (the price we use to buy the module)
            if result['module_price'] > 0:
                # Net mineral value per item (after reprocessing cost)
                net_mineral_value_per_item = mineral_value_per_item - reprocessing_cost_per_item
                return_percent = ((net_mineral_value_per_item - result['module_price']) / result['module_price']) * 100
            else:
                return_percent = float('inf') if profit_per_item > 0 else 0.0
            
            results.append({
                'module_name': module_name,
                'module_type_id': module_type_id,
                'expected_buy_price': expected_buy_price,
                'sell_min_price': sell_min,
                'module_price': result['module_price'],  # Price used in calculation (per item)
                'total_module_price': result['total_module_price'],
                'total_mineral_value': result['total_mineral_value'],
                'reprocessing_value': result['reprocessing_value'],
                'profit_per_item': profit_per_item,  # Profit per single item
                'return_percent': return_percent,  # (mineral_sell_price - buy_price) / buy_price
                'output_quantity': output_quantity
            })
            
            processed += 1
            if processed % 100 == 0:
                logger.info(f"Processed {processed}/{len(modules_df)} modules...")
        
        logger.info(f"Analysis complete! Processed {processed} modules")
        
        # Sort by return percentage (descending)
        results.sort(key=lambda x: x['return_percent'], reverse=True)
        
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
        # Format return percentage - cap display at 999,999% for readability
        return_pct = result['return_percent']
        if return_pct > 999999:
            return_str = ">999,999%"
        elif return_pct == float('inf'):
            return_str = "N/A"
        else:
            return_str = f"{return_pct:,.2f}%"
        
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
    yield_percent = float(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_YIELD_PERCENT
    buy_markup_percent = float(sys.argv[2]) if len(sys.argv) > 2 else BUY_ORDER_MARKUP_PERCENT
    reprocessing_cost_percent = float(sys.argv[3]) if len(sys.argv) > 3 else REPROCESSING_COST
    module_price_type = sys.argv[4] if len(sys.argv) > 4 else 'sell_min'
    mineral_price_type = sys.argv[5] if len(sys.argv) > 5 else 'buy_max'
    min_module_price = float(sys.argv[6]) if len(sys.argv) > 6 else 1.0
    max_module_price = float(sys.argv[7]) if len(sys.argv) > 7 else 100000.0
    top_n = int(sys.argv[8]) if len(sys.argv) > 8 else 30
    
    # Run analysis
    results = analyze_all_modules(
        yield_percent=yield_percent,
        buy_order_markup_percent=buy_markup_percent,
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

