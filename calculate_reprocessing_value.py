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


def ensure_input_quantity_cache_table(conn):
    """Ensure the input_quantity_cache table exists"""
    # Check if table exists and if it has typeName column
    cursor = conn.execute("""
        SELECT name FROM sqlite_master 
        WHERE type='table' AND name='input_quantity_cache'
    """)
    table_exists = cursor.fetchone() is not None
    
    if table_exists:
        # Check if typeName column exists
        cursor = conn.execute("PRAGMA table_info(input_quantity_cache)")
        columns = [row[1] for row in cursor.fetchall()]
        if 'typeName' not in columns:
            # Add typeName column to existing table
            conn.execute("ALTER TABLE input_quantity_cache ADD COLUMN typeName TEXT")
            # Populate typeName for existing rows
            conn.execute("""
                UPDATE input_quantity_cache 
                SET typeName = (SELECT typeName FROM items WHERE items.typeID = input_quantity_cache.typeID)
            """)
            conn.commit()
    else:
        # Create new table with typeName column
        conn.execute("""
            CREATE TABLE IF NOT EXISTS input_quantity_cache (
                typeID INTEGER PRIMARY KEY,
                typeName TEXT NOT NULL,
                input_quantity INTEGER NOT NULL,
                source TEXT NOT NULL,
                needs_review INTEGER DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (typeID) REFERENCES items(typeID)
            )
        """)
        conn.commit()
    
    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_input_quantity_cache_review 
        ON input_quantity_cache(needs_review)
    """)
    conn.commit()


def get_input_quantity(conn, type_id):
    """
    Get input_quantity for an item, using cache, blueprints, or group-based lookup.
    
    This function should only be called for items that exist in the items table.
    It processes items in this order:
    1. Check cache (fast lookup)
    2. Check blueprints table directly (for this specific item)
    3. If no blueprint, lookup by group:
       - Get all items in the same group
       - Check if all items with blueprints have the same outputQuantity (consensus)
       - If no consensus, use most frequent quantity (needs review)
       - If no blueprints in group, use default 1 (needs review)
    
    Returns:
        tuple: (input_quantity, source, needs_review)
            - input_quantity: The quantity value
            - source: 'blueprint', 'group_consensus', 'group_most_frequent', or 'default'
            - needs_review: 1 if needs manual review, 0 otherwise
    """
    # First, check cache
    cache_query = "SELECT input_quantity, source, needs_review FROM input_quantity_cache WHERE typeID = ?"
    cache_df = pd.read_sql_query(cache_query, conn, params=(type_id,))
    
    if len(cache_df) > 0:
        return (
            int(cache_df.iloc[0]['input_quantity']),
            cache_df.iloc[0]['source'],
            int(cache_df.iloc[0]['needs_review'])
        )
    
    # Get item name and groupID for caching and group lookup
    item_query = "SELECT typeName, groupID FROM items WHERE typeID = ?"
    item_df = pd.read_sql_query(item_query, conn, params=(type_id,))
    
    if len(item_df) == 0:
        # Item not in items table - this shouldn't happen if called correctly
        logger.warning(f"Item typeID {type_id} not found in items table")
        item_name = f"Item_{type_id}"
        input_quantity = 1
        conn.execute("""
            INSERT OR REPLACE INTO input_quantity_cache (typeID, typeName, input_quantity, source, needs_review)
            VALUES (?, ?, ?, 'default', 1)
        """, (type_id, item_name, input_quantity))
        conn.commit()
        return (input_quantity, 'default', 1)
    
    item_name = item_df.iloc[0]['typeName']
    group_id = int(item_df.iloc[0]['groupID']) if item_df.iloc[0]['groupID'] else None
    
    # Second, check blueprints table directly for this specific item
    blueprint_query = "SELECT outputQuantity FROM blueprints WHERE productTypeID = ?"
    blueprint_df = pd.read_sql_query(blueprint_query, conn, params=(type_id,))
    
    if len(blueprint_df) > 0:
        input_quantity = int(blueprint_df.iloc[0]['outputQuantity'])
        # Cache the result
        conn.execute("""
            INSERT OR REPLACE INTO input_quantity_cache (typeID, typeName, input_quantity, source, needs_review)
            VALUES (?, ?, ?, 'blueprint', 0)
        """, (type_id, item_name, input_quantity))
        conn.commit()
        return (input_quantity, 'blueprint', 0)
    
    # Third, lookup by group (if no blueprint found for this item)
    if group_id is None:
        # No groupID, use default
        input_quantity = 1
        conn.execute("""
            INSERT OR REPLACE INTO input_quantity_cache (typeID, typeName, input_quantity, source, needs_review)
            VALUES (?, ?, ?, 'default', 1)
        """, (type_id, item_name, input_quantity))
        conn.commit()
        return (input_quantity, 'default', 1)
    
    # Get all items in the same group
    group_items_query = "SELECT typeID FROM items WHERE groupID = ?"
    group_items_df = pd.read_sql_query(group_items_query, conn, params=(group_id,))
    
    if len(group_items_df) == 0:
        # No items in group, use default
        input_quantity = 1
        conn.execute("""
            INSERT OR REPLACE INTO input_quantity_cache (typeID, typeName, input_quantity, source, needs_review)
            VALUES (?, ?, ?, 'default', 1)
        """, (type_id, item_name, input_quantity))
        conn.commit()
        return (input_quantity, 'default', 1)
    
    # Get outputQuantity for all items in the group that have blueprints
    group_type_ids = group_items_df['typeID'].tolist()
    placeholders = ','.join(['?'] * len(group_type_ids))
    group_blueprints_query = f"""
        SELECT productTypeID, outputQuantity 
        FROM blueprints 
        WHERE productTypeID IN ({placeholders})
    """
    group_blueprints_df = pd.read_sql_query(group_blueprints_query, conn, params=group_type_ids)
    
    if len(group_blueprints_df) == 0:
        # No blueprints in group, use default
        input_quantity = 1
        conn.execute("""
            INSERT OR REPLACE INTO input_quantity_cache (typeID, typeName, input_quantity, source, needs_review)
            VALUES (?, ?, ?, 'default', 1)
        """, (type_id, item_name, input_quantity))
        conn.commit()
        return (input_quantity, 'default', 1)
    
    # Check if all items with blueprints have the same outputQuantity
    output_quantities = group_blueprints_df['outputQuantity'].unique()
    
    if len(output_quantities) == 1:
        # All items in group have the same outputQuantity - consensus
        input_quantity = int(output_quantities[0])
        conn.execute("""
            INSERT OR REPLACE INTO input_quantity_cache (typeID, typeName, input_quantity, source, needs_review)
            VALUES (?, ?, ?, 'group_consensus', 0)
        """, (type_id, item_name, input_quantity))
        conn.commit()
        return (input_quantity, 'group_consensus', 0)
    
    # No consensus - use most frequent quantity
    quantity_counts = group_blueprints_df['outputQuantity'].value_counts()
    most_frequent_quantity = int(quantity_counts.index[0])
    input_quantity = most_frequent_quantity
    
    # Cache with needs_review flag
    conn.execute("""
        INSERT OR REPLACE INTO input_quantity_cache (typeID, typeName, input_quantity, source, needs_review)
        VALUES (?, ?, ?, 'group_most_frequent', 1)
    """, (type_id, item_name, input_quantity))
    conn.commit()
    
    return (input_quantity, 'group_most_frequent', 1)


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

def sell_order_with_fees(sell_price, broker_fee=BROKER_FEE, sales_tax=SALES_TAX, average_relist=LISTING_RELIST, sell_buffer=BUY_BUFFER, RELIST_DISCOUNT=RELIST_DISCOUNT):
    # see the mechanism on buy order with fee the difference is we also have to pay the sales tax 
    sell_price_loaded = sell_price*(1-sell_buffer/100)
    relist_fees = sell_price_loaded*(broker_fee/100)*((1-RELIST_DISCOUNT)/100)*average_relist
    broker_fee_amount = sell_price_loaded*broker_fee/100
    sales_tax_amount = sell_price_loaded*sales_tax/100
    total_sell_order_realised = sell_price_loaded - relist_fees - broker_fee_amount - sales_tax_amount

    return total_sell_order_realised

def sell_into_buy_order(sell_price, sales_tax=SALES_TAX):
    #when sellling into a buy order there is no s
    return sell_price*(1-sales_tax/100)

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
    module_price_type='buy_immediate',  # 'buy_immediate' or 'buy_offer'
    mineral_price_type='sell_immediate',  # 'sell_immediate' or 'sell_offer'
 
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
            - input_quantity: Number of items needed to obtain reprocessing result (from blueprints table)
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
    
    # Ensure cache table exists
    ensure_input_quantity_cache_table(conn)
    
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
            module_price_post_transaction_costs = 0
            logger.warning(f"No price data found for {module_name}, using 0")
        else:
            buy_max = float(price_df.iloc[0]['buy_max']) if price_df.iloc[0]['buy_max'] else 0.0
            sell_min = float(price_df.iloc[0]['sell_min']) if price_df.iloc[0]['sell_min'] else 0.0
            
            # Calculate module price based on selected type
            if module_price_type == 'buy_offer':
                module_price_before_markup = buy_max
                module_price_post_transaction_costs = buy_order_with_fees(buy_max)
                
            elif module_price_type == 'buy_immediate':
                module_price_before_markup = sell_min
                module_price_post_transaction_costs= buy_into_sell_order(sell_min)
                
            else:
                logger.warning(f"Invalid module_price_type '{module_price_type}', using 'buy_immediate'")
                module_price_before_markup = sell_min
                module_price_post_transaction_costs = buy_into_sell_order(sell_min)
                
        
        
        
        # Apply markup as safety cushion for market valuation and other selling costs not addressed yet
        
        
        material_type_ids = reprocessing_df['materialTypeID'].tolist()
        placeholders = ','.join(['?'] * len(material_type_ids))
        
        if mineral_price_type == 'sell_immediate':
            price_column = 'buy_max'
        elif mineral_price_type == 'sell_offer':
            price_column = 'sell_min'
        else:
            logger.warning(f"Invalid mineral_price_type '{mineral_price_type}', using 'sell_immediate'")
            price_column = 'buy_max'
        
        mineral_price_query = f"""
            SELECT typeID, buy_max, sell_min
            FROM prices
            WHERE typeID IN ({placeholders})
        """
        mineral_prices_df = pd.read_sql_query(mineral_price_query, conn, params=material_type_ids)
        
        # Create price lookup based on selected type
        mineral_price_lookup = {}
        mineral_price_lookup_after_costs = {}
        for _, row in mineral_prices_df.iterrows():
            type_id = int(row['typeID'])
            if price_column == 'buy_max':
                price = float(row['buy_max']) if row['buy_max'] else 0.0
                price_after_costs = sell_into_buy_order(price)
            else:  # sell_min
                price = float(row['sell_min']) if row['sell_min'] else 0.0
                price_after_costs = sell_order_with_fees(price)
            mineral_price_lookup[type_id] = price
            mineral_price_lookup_after_costs[type_id] = price_after_costs
        
        
        # Get input_quantity - first check cache, then blueprints, then group lookup
        input_quantity, source, needs_review = get_input_quantity(conn, module_type_id)

        # Calculate total module price for all modules
        total_module_price = module_price_post_transaction_costs * input_quantity
        
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
        #   1.
        #   
        #   2. quantity_per_module = quantity_per_item * yield_percent / 100
        #      (e.g., 0.3 * 0.55 = 0.165 per module after 55% yield)
        #   
        #   3. actual_quantity = quantity_per_module * input_quantity (rounded down)
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
            batch_quantity = int(row['quantity'])  # Quantity from database (per item for reprocessing)
             # For reprocessing, quantity is already per item
            
            # Step 1: we apply the yiel to the amount we get from reprocessing formula
            # Example: 30  item reprocessed * 0.55 = 16.5 per item after 55% yield
            actual_quantity = batch_quantity * yield_multiplier
            
            
            
            # Get mineral price from price lookup
            mineral_price = mineral_price_lookup.get(material_type_id, 0.0)
            mineral_price_after_costs = mineral_price_lookup_after_costs.get(material_type_id, 0.0)
            
            
            
            # Calculate the value for this type of mineral after reprocessing
            mineral_value = actual_quantity * mineral_price_after_costs
            
            # Add to total mineral value
            total_mineral_value += mineral_value
            
            # Store all the calculated data for this material
            reprocessing_outputs.append({
                'materialTypeID': material_type_id,           # Material type ID
                'materialName': material_name,                 # Material name (e.g., "Morphite")
                'batchQuantity': batch_quantity,               # Original quantity from database                      # Per item after yield (before rounding)
                'QuantityAfterYield': actual_quantity,              # Total for input_quantity after yield and rounding # Total before rounding (for display/debugging)
                'mineralPrice': mineral_price,   
                'mineralPriceAfterCosts': mineral_price_after_costs, # Price per unit of this mineral after costs
                'mineralValue': mineral_value                   # Total value (quantity * price) for all input_quantity items
            })
        
        
        # Calculate net reprocessing value (mineral value - module price - reprocessing cost)
        reprocessing_value = total_mineral_value - total_module_price - reprocessing_cost
        
        # Calculate profit margin
        if total_module_price > 0:
            profit_margin_percent = ((reprocessing_value / total_module_price) -1) *100
        else:
            profit_margin_percent = "na" #not applicable
        
        # Calculate breakeven price (maximum purchase price for 0 profit)
        
        if total_mineral_value > 0 and input_quantity > 0:
            job_income_after_costs = total_mineral_value - reprocessing_cost
            income_per_item = job_income_after_costs / input_quantity
            cost_factor = module_price_post_transaction_costs/module_price_before_markup
            module_price_breakeven = income_per_item / cost_factor

        else:
            module_price_breakeven = "na"
        
        
        result = {
            'module_type_id': module_type_id,
            'module_name': module_name,
            'module_price_type': module_price_type,
            'mineral_price_type': mineral_price_type,
            'module_price': module_price_before_markup,
            'module_price_after_costs': module_price_post_transaction_costs,
            'input_quantity': input_quantity,
            'input_quantity_source': source,
            'input_quantity_needs_review': bool(needs_review),
            'total_module_cost_per_job': total_module_price,
            'reprocessing_cost_percent': reprocessing_cost_percent,
            'effective_reprocessing_cost_percent': effective_reprocessing_cost_percent,
            'reprocessing_cost_per_job': reprocessing_cost,
            'yield_percent': yield_percent,
            'reprocessing_outputs': reprocessing_outputs,
            
            'total_mineral_value_per_job_after_costs': total_mineral_value,
            'reprocessing_value_per_job_after_costs': reprocessing_value,
            'profit_margin_percent': profit_margin_percent,
            'breakeven_module_price': module_price_breakeven
            
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
    output.append("=" * 60)
    output.append(f"Reprocessing Value Calculation")
    output.append("=" * 60)
    output.append(f"Module: {result['module_name']} (TypeID: {result['module_type_id']})")
    input_qty = result['input_quantity']
    source = result.get('input_quantity_source', 'unknown')
    needs_review = result.get('input_quantity_needs_review', False)
    source_desc = {
        'blueprint': 'from blueprint',
        'group_consensus': 'from group consensus',
        'group_most_frequent': 'from group most frequent (needs review)',
        'default': 'default (needs review)'
    }.get(source, source)
    review_marker = " ⚠ NEEDS REVIEW" if needs_review else ""
    output.append(f"Input Quantity: {input_qty} items (needed to reprocess) - {source_desc}{review_marker}")
    output.append(f"")
    output.append(f"Price Settings:")
    output.append(f"  Module Price Type: {result['module_price_type']}")
    output.append(f"  Mineral Price Type: {result['mineral_price_type']}")
    output.append(f"")
    
    # Show module price details
    output.append(f"Module Price (base):        {result['module_price']:,.2f} ISK per module")
    output.append(f"Module Price (after costs):  {result['module_price_after_costs']:,.2f} ISK per module")
    output.append(f"")
    output.append(f"Total Module Cost per Job:        {result['total_module_cost_per_job']:,.2f} ISK")
    output.append(f"Reprocessing Cost per Job:         {result['reprocessing_cost_per_job']:,.2f} ISK")
    output.append(f"  (Base: {result['reprocessing_cost_percent']:.2f}% × Yield: {result['yield_percent']:.1f}% = {result.get('effective_reprocessing_cost_percent', 0):.4f}%)")
    output.append(f"")
    output.append(f"Reprocessing Yield: {result['yield_percent']:.1f}%")
    output.append(f"")
    
    output.append("Reprocessing Outputs:")
    output.append("-" * 100)
    output.append(f"{'Mineral':<30} {'Qty Before Yield':>18} {'Qty After Yield':>18} {'Price':>12} {'Price After Costs':>18} {'Value':>15}")
    output.append("-" * 100)
    
    for i, output_mat in enumerate(result['reprocessing_outputs']):
        batch_qty = output_mat.get('batchQuantity', 0)
        qty_after_yield = output_mat.get('QuantityAfterYield', 0)
        price = output_mat.get('mineralPrice', 0)
        price_after_costs = output_mat.get('mineralPriceAfterCosts', 0)
        value = output_mat.get('mineralValue', 0)
        
        output.append(
            f"  {output_mat['materialName']:<28} "
            f"{batch_qty:18.2f} "
            f"{qty_after_yield:18.2f} "
            f"{price:12,.2f} "
            f"{price_after_costs:18,.2f} "
            f"{value:15,.2f}"
        )
    
    output.append("-" * 60)
    
    # Per-item calculations
    input_quantity = result['input_quantity']
    total_mineral_value = result['total_mineral_value_per_job_after_costs']
    reprocessing_cost = result['reprocessing_cost_per_job']
    total_module_cost = result['total_module_cost_per_job']
    reprocessing_value = result['reprocessing_value_per_job_after_costs']
    
    mineral_value_per_item = total_mineral_value / input_quantity if input_quantity > 0 else 0
    reprocessing_cost_per_item = reprocessing_cost / input_quantity if input_quantity > 0 else 0
    module_cost_per_item = total_module_cost / input_quantity if input_quantity > 0 else 0
    profit_per_item = mineral_value_per_item - module_cost_per_item - reprocessing_cost_per_item
    
    output.append("PER ITEM:")
    output.append(f"  Mineral Value (after costs): {mineral_value_per_item:,.2f} ISK")
    output.append(f"  Module Cost:                 {module_cost_per_item:,.2f} ISK")
    output.append(f"  Reprocessing Cost:           {reprocessing_cost_per_item:,.2f} ISK")
    output.append(f"  Profit per Item:             {profit_per_item:+,.2f} ISK")
    
    # Calculate return percentage per item
    if module_cost_per_item > 0:
        net_mineral_value_per_item = mineral_value_per_item - reprocessing_cost_per_item
        return_percent_per_item = ((net_mineral_value_per_item - module_cost_per_item) / module_cost_per_item) * 100
        output.append(f"  Return %:                     {return_percent_per_item:+.2f}%")
    else:
        output.append(f"  Return %:                     N/A (module cost is 0)")
    
    output.append("")
    output.append(f"TOTAL PER JOB (for {input_quantity} items):")
    output.append(f"  Total Mineral Value (after costs): {total_mineral_value:,.2f} ISK")
    output.append(f"  Total Module Cost:                  {total_module_cost:,.2f} ISK ({module_cost_per_item:,.2f} × {input_quantity})")
    output.append(f"  Reprocessing Cost:                  {reprocessing_cost:,.2f} ISK")
    output.append(f"  Net Profit per Job:                 {reprocessing_value:,.2f} ISK")
    
    profit_margin = result.get('profit_margin_percent', 'na')
    if profit_margin != 'na' and profit_margin != float('inf'):
        output.append(f"  Profit Margin:                     {profit_margin:+.2f}%")
    else:
        output.append(f"  Profit Margin:                     N/A")
    
    output.append("")
    output.append("BREAKEVEN ANALYSIS:")
    breakeven_price = result.get('breakeven_module_price', 'na')
    if breakeven_price != 'na' and breakeven_price != 0:
        current_price = result.get('module_price', 0)
        if current_price > 0:
            price_difference = breakeven_price - current_price
            price_difference_percent = (price_difference / current_price) * 100 if current_price > 0 else 0
            output.append(f"  Max Purchase Price for breakeven: {breakeven_price:,.2f} ISK per module")
            output.append(f"  Current Price:                     {current_price:,.2f} ISK per module")
            output.append(f"  Price Difference:                  {price_difference:+,.2f} ISK ({price_difference_percent:+.2f}%)")
        else:
            output.append(f"  Max Purchase Price for breakeven: {breakeven_price:,.2f} ISK per module")
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
        print("Note: input_quantity is automatically fetched from blueprints table based on productTypeID")
        print("")
        print("Price Types:")
        print("  module_price_type: 'buy_immediate' (default) or 'buy_offer'")
        print("  mineral_price_type: 'sell_immediate' (default) or 'sell_offer'")
        print("")
        print("Examples:")
        print("  python calculate_reprocessing_value.py \"Medium Shield Booster II\"")
        print("  python calculate_reprocessing_value.py 11269 55 10")
        print("  python calculate_reprocessing_value.py \"Medium Shield Booster II\" 60 15")
        print("  python calculate_reprocessing_value.py \"Iron Charge S\" 55 10 3.37")
        print("  python calculate_reprocessing_value.py \"Gamma L\" 55 10 3.37 buy_immediate sell_immediate")
        print("  python calculate_reprocessing_value.py \"Gamma L\" 55 10 3.37 buy_offer sell_offer")
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
    sort_by='return',  # 'return' or 'profit'
    item_source_filter='all',  # 'all', 'blueprint', or 'group_consensus' (faster when restricted)
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
    logger.info(f"  Note: input_quantity fetched from blueprints table")
    price_basis = 'buy_max' if module_price_type == 'buy_offer' else 'sell_min'
    logger.info(f"  Min {price_basis} price: {min_module_price:,.0f} ISK")
    logger.info(f"  Max {price_basis} price: {max_module_price:,.0f} ISK")
    logger.info(f"  Module price type: {module_price_type}")
    logger.info(f"  Mineral price type: {mineral_price_type}")
    logger.info(f"  Top N results: {top_n}")
    logger.info(f"  Item source filter: {item_source_filter}")
    logger.info("=" * 60)
    
    if not Path(db_file).exists():
        logger.error(f"Database file not found: {db_file}")
        return []
    
    # Initialize excluded_module_ids if None
    if excluded_module_ids is None:
        excluded_module_ids = set()
    
    conn = sqlite3.connect(db_file)
    
    # Ensure cache table exists
    ensure_input_quantity_cache_table(conn)
    
    try:
        # Get all modules that can be reprocessed, excluding certain high-level categories
        # Optionally restrict to blueprint or group_consensus items (faster).
        # CategoryIDs excluded (from items.categoryID):
        # 25, 91, 1, 2, 3, 4, 5, 17, 29, 14, 9, 10, 11, 16, 20, 2100, 2118, 24, 26, 30, 350001
        if item_source_filter in ('blueprint', 'group_consensus'):
            query = """
                SELECT DISTINCT ro.itemTypeID, ro.itemName
                FROM reprocessing_outputs ro
                JOIN items i ON ro.itemTypeID = i.typeID
                JOIN input_quantity_cache c ON ro.itemTypeID = c.typeID
                WHERE i.categoryID NOT IN (
                    25, 91, 1, 2, 3, 4, 5, 17, 29, 14, 9, 10, 11, 16, 20, 2100, 2118, 24, 26, 30, 350001
                )
                AND c.source = ?
                ORDER BY ro.itemName
            """
            modules_df = pd.read_sql_query(query, conn, params=(item_source_filter,))
        else:
            query = """
                SELECT DISTINCT ro.itemTypeID, ro.itemName
                FROM reprocessing_outputs ro
                JOIN items i ON ro.itemTypeID = i.typeID
                WHERE i.categoryID NOT IN (
                    25, 91, 1, 2, 3, 4, 5, 17, 29, 14, 9, 10, 11, 16, 20, 2100, 2118, 24, 26, 30, 350001
                )
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
            
            # Filter by appropriate price based on module_price_type
            # - If buying via buy orders (buy_offer), use highest buy order (buy_max)
            # - If buying immediately (buy_immediate), use lowest sell order (sell_min)
            filter_price = buy_max if module_price_type == 'buy_offer' else sell_min
            if filter_price < min_module_price or filter_price > max_module_price:
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
            if result['module_price'] == 0 or result['total_mineral_value_per_job_after_costs'] == 0:
                continue
            
            # Calculate profit per item and return percentage
            # Expected buy price = buy_max + markup (for comparison)
            expected_buy_price = buy_max * (1 + buy_order_markup_percent / 100) if buy_max > 0 else 0
            
            # For per-item calculations:
            # - Mineral value per item = total_mineral_value_per_job_after_costs / input_quantity
            # - Module price per item = module_price (base price per item)
            # - Reprocessing cost per item = reprocessing_cost_per_job / input_quantity
            # - Profit per item = (mineral_value_per_item - module_price_after_costs - reprocessing_cost_per_item)
            input_quantity = result.get('input_quantity', 1)
            total_mineral_value = result['total_mineral_value_per_job_after_costs']
            total_module_cost = result['total_module_cost_per_job']
            reprocessing_cost = result['reprocessing_cost_per_job']
            module_price_base = result['module_price']
            module_price_after_costs = result.get('module_price_after_costs', module_price_base)
            
            mineral_value_per_item = total_mineral_value / input_quantity if input_quantity > 0 else 0
            reprocessing_cost_per_item = reprocessing_cost / input_quantity if input_quantity > 0 else 0
            module_cost_per_item = total_module_cost / input_quantity if input_quantity > 0 else 0
            profit_per_item = mineral_value_per_item - module_cost_per_item - reprocessing_cost_per_item
            
            # Return percentage = (mineral_sell_price - buy_price) / buy_price * 100
            # Where mineral_sell_price = mineral_value_per_item (net of reprocessing cost)
            # And buy_price = module_price_after_costs (the price we use to buy the module)
            if module_price_after_costs > 0:
                # Net mineral value per item (after reprocessing cost)
                net_mineral_value_per_item = mineral_value_per_item - reprocessing_cost_per_item
                return_percent = ((net_mineral_value_per_item - module_price_after_costs) / module_price_after_costs) * 100
            else:
                return_percent = float('inf') if profit_per_item > 0 else 0.0
            
            results.append({
                'module_name': module_name,
                'module_type_id': module_type_id,
                'expected_buy_price': expected_buy_price,
                'sell_min_price': sell_min,
                'module_price': module_price_base,  # Base price used in calculation (per item)
                'module_price_after_costs': module_price_after_costs,  # Price after transaction costs
                'total_module_cost': total_module_cost,
                'total_mineral_value': total_mineral_value,
                'reprocessing_value': result['reprocessing_value_per_job_after_costs'],
                'profit_per_item': profit_per_item,  # Profit per single item
                'return_percent': return_percent,  # (mineral_sell_price - buy_price) / buy_price
                'input_quantity': input_quantity,
                'input_quantity_source': result.get('input_quantity_source', 'unknown'),
                'breakeven_module_price': result.get('breakeven_module_price', 'na')
            })
            
            processed += 1
            if processed % 100 == 0:
                logger.info(f"Processed {processed}/{len(modules_df)} modules...")
        
        logger.info(f"Analysis complete! Processed {processed} modules")
        
        # Sort results
        if sort_by == 'profit':
            logger.info("Sorting results by profit per item (descending)")
            results.sort(key=lambda x: x.get('profit_per_item', 0), reverse=True)
        else:
            logger.info("Sorting results by return percentage (descending)")
            results.sort(key=lambda x: x.get('return_percent', 0), reverse=True)
        
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
    header = (
        f"{'Rank':<6} "
        f"{'Module Name':<40} "
        f"{'Buy Price':>12} "
        f"{'Sell Min':>12} "
        f"{'Profit/Item':>15} "
        f"{'Return %':>12} "
        f"{'Breakeven Max Buy':>18}"
    )
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
        
        breakeven_price = result.get('breakeven_module_price', 'na')
        if isinstance(breakeven_price, (int, float)) and breakeven_price not in (0, float('inf')):
            breakeven_str = f"{breakeven_price:,.2f}"
        else:
            breakeven_str = "N/A"
        
        row = (
            f"{rank:<6} "
            f"{module_name:<40} "
            f"{result['expected_buy_price']:>12,.2f} "
            f"{result['sell_min_price']:>12,.2f} "
            f"{result['profit_per_item']:>15,.2f} "
            f"{return_str:>12} "
            f"{breakeven_str:>18}"
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
        top_n=top_n,
        sort_by='return'
    )
    
    # Display results
    print(format_analysis_results(results))
    
    if not results:
        sys.exit(1)


if __name__ == "__main__":
    # if len(sys.argv) > 1 and sys.argv[1] == '--analyze-all':
    #     # Run analysis for all modules
    #     sys.argv = sys.argv[1:]  # Remove '--analyze-all'
    #     analyze_all_modules_main()
    # else:
    #     # Run single module analysis
    #     main()

    result = calculate_reprocessing_value(module_name="Javelin L", yield_percent=55, buy_order_markup_percent=10, reprocessing_cost_percent=3.37)
    print(result)