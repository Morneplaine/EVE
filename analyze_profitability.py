"""
Analyze manufacturing profitability based on:
- Character skills (ME reduction)
- Available resources/inventory
- Market prices
- Manufacturing fees and taxes

This script finds the most profitable items to manufacture.
"""

import sqlite3
import pandas as pd
import logging
from pathlib import Path

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

DB_FILE = "eve_manufacturing.db"

# Manufacturing settings (adjustable)
MANUFACTURING_FEE_PERCENT = 0.02  # 2% manufacturing fee
SALES_TAX_PERCENT = 0.08  # 8% sales tax (worst case)
ME_LEVEL = 0  # Material Efficiency level (0-10, reduces material cost by 1% per level)

def calculate_material_cost(blueprint_type_id, me_level, conn):
    """Calculate total material cost for a blueprint at given ME level"""
    # Get materials
    materials_query = """
        SELECT 
            mm.materialTypeID,
            mm.materialName,
            mm.quantity,
            p.sell_min as material_price
        FROM manufacturing_materials mm
        LEFT JOIN prices p ON mm.materialTypeID = p.typeID
        WHERE mm.blueprintTypeID = ?
    """
    materials_df = pd.read_sql_query(materials_query, conn, params=(blueprint_type_id,))
    
    if len(materials_df) == 0:
        return 0
    
    # Apply ME reduction (1% per level, max 10%)
    me_reduction = 1.0 - (me_level * 0.01)
    me_reduction = max(0.9, me_reduction)  # Cap at 10% reduction
    
    total_cost = 0
    for _, mat in materials_df.iterrows():
        quantity = int(mat['quantity']) * me_reduction
        price = float(mat['material_price'] or 0)
        total_cost += quantity * price
    
    return total_cost

def check_skills_met(blueprint_type_id, conn):
    """Check if character has required skills"""
    skills_query = """
        SELECT 
            ms.skillID,
            ms.skillName,
            ms.level as required_level,
            COALESCE(cs.level, 0) as character_level
        FROM manufacturing_skills ms
        LEFT JOIN character_skills cs ON ms.skillID = cs.skillID
        WHERE ms.blueprintTypeID = ?
    """
    skills_df = pd.read_sql_query(skills_query, conn, params=(blueprint_type_id,))
    
    if len(skills_df) == 0:
        return True, []  # No skills required
    
    missing_skills = []
    for _, skill in skills_df.iterrows():
        if skill['character_level'] < skill['required_level']:
            missing_skills.append(f"{skill['skillName']} {int(skill['required_level'])}")
    
    return len(missing_skills) == 0, missing_skills

def check_resources_available(blueprint_type_id, conn):
    """Check if enough resources are available in inventory
    
    Returns:
        tuple: (can_make, max_units, missing_materials)
    """
    materials_query = """
        SELECT 
            mm.materialTypeID,
            mm.materialName,
            mm.quantity as required,
            COALESCE(inv.quantity, 0) as available
        FROM manufacturing_materials mm
        LEFT JOIN inventory inv ON mm.materialTypeID = inv.typeID
        WHERE mm.blueprintTypeID = ?
    """
    materials_df = pd.read_sql_query(materials_query, conn, params=(blueprint_type_id,))
    
    if len(materials_df) == 0:
        return True, float('inf'), []  # No materials required
    
    # Find minimum number of units we can make
    max_units = float('inf')
    missing_materials = []
    
    for _, mat in materials_df.iterrows():
        required = int(mat['required'])
        available = int(mat['available'])
        
        if available < required:
            missing_materials.append(f"{mat['materialName']}: need {required}, have {available}")
            max_units = 0
        else:
            units_possible = available // required
            max_units = min(max_units, units_possible)
    
    can_make = max_units > 0 if max_units != float('inf') else True
    return can_make, max_units if max_units != float('inf') else None, missing_materials

def analyze_profitability(me_level=0, filter_skills=True, filter_resources=False, min_profit=0):
    """
    Analyze profitability of all blueprints
    
    Args:
        me_level: Material Efficiency level (0-10)
        filter_skills: Only show blueprints where skills are met
        filter_resources: Only show blueprints where resources are available
        min_profit: Minimum profit per unit to show
    """
    logger.info("=" * 60)
    logger.info("Analyzing Manufacturing Profitability")
    logger.info("=" * 60)
    logger.info(f"ME Level: {me_level}")
    logger.info(f"Manufacturing Fee: {MANUFACTURING_FEE_PERCENT*100}%")
    logger.info(f"Sales Tax: {SALES_TAX_PERCENT*100}%")
    logger.info("=" * 60)
    
    conn = sqlite3.connect(DB_FILE)
    
    try:
        # Get all blueprints
        blueprints_query = """
            SELECT 
                b.blueprintTypeID,
                b.productTypeID,
                b.productName,
                b.groupName,
                b.outputQuantity,
                p.sell_min as product_price
            FROM blueprints b
            LEFT JOIN prices p ON b.productTypeID = p.typeID
            ORDER BY b.productName
        """
        blueprints_df = pd.read_sql_query(blueprints_query, conn)
        
        results = []
        
        for _, bp in blueprints_df.iterrows():
            blueprint_type_id = bp['blueprintTypeID']
            product_type_id = bp['productTypeID']
            
            # Check skills
            skills_met, missing_skills = check_skills_met(blueprint_type_id, conn)
            if filter_skills and not skills_met:
                continue
            
            # Check resources
            if filter_resources:
                resources_ok, max_units, missing_materials = check_resources_available(blueprint_type_id, conn)
                if not resources_ok or max_units == 0:
                    continue
            else:
                max_units = None
                missing_materials = []
            
            # Calculate costs
            material_cost = calculate_material_cost(blueprint_type_id, me_level, conn)
            manufacturing_fee = material_cost * MANUFACTURING_FEE_PERCENT
            total_cost = material_cost + manufacturing_fee
            
            # Calculate revenue
            product_price = float(bp['product_price'] or 0)
            revenue_per_unit = product_price * (1 - SALES_TAX_PERCENT)
            revenue_total = revenue_per_unit * int(bp['outputQuantity'])
            
            # Calculate profit
            profit_per_unit = revenue_per_unit - (total_cost / int(bp['outputQuantity']))
            profit_total = revenue_total - total_cost
            profit_margin = (profit_per_unit / revenue_per_unit * 100) if revenue_per_unit > 0 else 0
            
            # ROI
            roi = (profit_total / total_cost * 100) if total_cost > 0 else 0
            
            if profit_per_unit >= min_profit:
                results.append({
                    'Product Name': bp['productName'],
                    'Group': bp['groupName'],
                    'Output Qty': int(bp['outputQuantity']),
                    'Product Price': product_price,
                    'Material Cost': material_cost,
                    'Manufacturing Fee': manufacturing_fee,
                    'Total Cost': total_cost,
                    'Revenue (after tax)': revenue_total,
                    'Profit per Unit': profit_per_unit,
                    'Total Profit': profit_total,
                    'Profit Margin %': profit_margin,
                    'ROI %': roi,
                    'Skills Met': 'Yes' if skills_met else f"No: {', '.join(missing_skills)}",
                    'Max Units (resources)': max_units if max_units is not None else 'N/A'
                })
        
        if not results:
            logger.warning("No profitable blueprints found with current filters")
            return
        
        # Create DataFrame and sort by profit
        results_df = pd.DataFrame(results)
        results_df = results_df.sort_values('Total Profit', ascending=False)
        
        # Display results
        logger.info(f"\nTop 20 Most Profitable Items:\n")
        logger.info("=" * 120)
        
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', None)
        pd.set_option('display.max_colwidth', 30)
        
        print(results_df.head(20).to_string(index=False))
        
        # Save to CSV
        output_file = "profitability_analysis.csv"
        results_df.to_csv(output_file, index=False)
        logger.info(f"\nFull results saved to: {output_file}")
        
        # Summary statistics
        logger.info("\n" + "=" * 60)
        logger.info("Summary Statistics:")
        logger.info(f"  Total profitable items: {len(results_df)}")
        logger.info(f"  Average profit per unit: {results_df['Profit per Unit'].mean():.2f} ISK")
        logger.info(f"  Average ROI: {results_df['ROI %'].mean():.2f}%")
        logger.info(f"  Best profit: {results_df.iloc[0]['Total Profit']:.2f} ISK ({results_df.iloc[0]['Product Name']})")
        logger.info("=" * 60)
        
    except Exception as e:
        logger.error(f"Error analyzing profitability: {e}")
        import traceback
        logger.error(traceback.format_exc())
        raise
    finally:
        conn.close()

def main():
    import sys
    
    # Parse command line arguments
    me_level = 0
    filter_skills = True
    filter_resources = False
    min_profit = 0
    
    if len(sys.argv) > 1:
        me_level = int(sys.argv[1])
    if len(sys.argv) > 2:
        filter_skills = sys.argv[2].lower() == 'true'
    if len(sys.argv) > 3:
        filter_resources = sys.argv[3].lower() == 'true'
    if len(sys.argv) > 4:
        min_profit = float(sys.argv[4])
    
    analyze_profitability(
        me_level=me_level,
        filter_skills=filter_skills,
        filter_resources=filter_resources,
        min_profit=min_profit
    )

if __name__ == "__main__":
    main()

