"""
Generate Excel file from SQLite database
This creates a static Excel file with all the data, similar to the original script.
"""

import sqlite3
import pandas as pd
import xlsxwriter
import logging
from pathlib import Path

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

DB_FILE = "eve_manufacturing.db"
OUTPUT_FILE = "EVE_Manufacturing_Database.xlsx"

def generate_excel():
    """Generate Excel file from database"""
    logger.info("=" * 60)
    logger.info("Generating Excel file from database")
    logger.info("=" * 60)
    
    conn = sqlite3.connect(DB_FILE)
    
    try:
        # Load data from database
        logger.info("Loading data from database...")
        
        # Manufacturing data
        manufacturing_query = """
            SELECT 
                b.blueprintTypeID,
                b.productTypeID,
                b.productName,
                b.groupName,
                b.outputQuantity
            FROM blueprints b
            ORDER BY b.productName
        """
        manufacturing_df = pd.read_sql_query(manufacturing_query, conn)
        
        # Get materials for each blueprint
        materials_query = """
            SELECT 
                mm.blueprintTypeID,
                mm.materialTypeID,
                mm.materialName,
                mm.quantity
            FROM manufacturing_materials mm
        """
        materials_df = pd.read_sql_query(materials_query, conn)
        
        # Get skills for each blueprint
        skills_query = """
            SELECT 
                ms.blueprintTypeID,
                ms.skillID,
                ms.skillName,
                ms.level
            FROM manufacturing_skills ms
        """
        skills_df = pd.read_sql_query(skills_query, conn)
        
        # Reprocessing data
        reprocessing_query = """
            SELECT 
                ro.itemTypeID,
                ro.itemName,
                ro.materialTypeID,
                ro.materialName,
                ro.quantity
            FROM reprocessing_outputs ro
            ORDER BY ro.itemName
        """
        reprocessing_df = pd.read_sql_query(reprocessing_query, conn)
        
        # Prices
        prices_query = "SELECT * FROM prices"
        prices_df = pd.read_sql_query(prices_query, conn)
        
        # Items (for volumes)
        items_query = "SELECT typeID, typeName, volume, packaged_volume FROM items"
        items_df = pd.read_sql_query(items_query, conn)
        
        logger.info(f"Loaded {len(manufacturing_df)} blueprints")
        logger.info(f"Loaded {len(reprocessing_df)} reprocessing items")
        logger.info(f"Loaded {len(prices_df)} prices")
        
        # Create pivot format for manufacturing
        logger.info("Creating manufacturing pivot table...")
        
        # Get all unique materials and count usage
        material_usage = materials_df.groupby(['materialTypeID', 'materialName']).size().reset_index(name='usage_count')
        material_usage = material_usage.sort_values(['usage_count', 'materialName'], ascending=[False, True])
        sorted_materials = list(zip(material_usage['materialTypeID'], material_usage['materialName']))
        
        # Build manufacturing pivot
        mfg_pivot_data = []
        for _, bp in manufacturing_df.iterrows():
            pivot_row = {
                'Product Name': bp['productName'],
                'Product TypeID': bp['productTypeID'],
                'Group': bp['groupName'],
                'Output Qty': bp['outputQuantity']
            }
            
            # Initialize all material columns to 0
            for mat_id, mat_name in sorted_materials:
                pivot_row[mat_name] = 0
            
            # Fill in actual quantities
            bp_materials = materials_df[materials_df['blueprintTypeID'] == bp['blueprintTypeID']]
            for _, mat in bp_materials.iterrows():
                mat_name = mat['materialName']
                if mat_name in pivot_row:
                    pivot_row[mat_name] = int(mat['quantity'])
            
            # Add skills
            bp_skills = skills_df[skills_df['blueprintTypeID'] == bp['blueprintTypeID']]
            if len(bp_skills) > 0:
                skills_list = [f"{row['skillName']} {int(row['level'])}" for _, row in bp_skills.iterrows()]
                pivot_row['Required Skills'] = " | ".join(skills_list)
            else:
                pivot_row['Required Skills'] = "None"
            
            # Add price and volume
            product_type_id = bp['productTypeID']
            price_row = prices_df[prices_df['typeID'] == product_type_id]
            if len(price_row) > 0:
                pivot_row['Product Price (ISK)'] = float(price_row.iloc[0]['sell_min'])
            else:
                pivot_row['Product Price (ISK)'] = 0
            
            item_row = items_df[items_df['typeID'] == product_type_id]
            if len(item_row) > 0:
                packaged_vol = item_row.iloc[0]['packaged_volume']
                if packaged_vol and packaged_vol > 0:
                    pivot_row['Product Volume (m³)'] = float(packaged_vol)
                else:
                    pivot_row['Product Volume (m³)'] = float(item_row.iloc[0]['volume'] or 0)
            else:
                pivot_row['Product Volume (m³)'] = 0
            
            mfg_pivot_data.append(pivot_row)
        
        mfg_pivot_df = pd.DataFrame(mfg_pivot_data)
        
        # Create reprocessing pivot (similar process)
        logger.info("Creating reprocessing pivot table...")
        
        reprocess_material_usage = reprocessing_df.groupby(['materialTypeID', 'materialName']).size().reset_index(name='usage_count')
        reprocess_material_usage = reprocess_material_usage.sort_values(['usage_count', 'materialName'], ascending=[False, True])
        sorted_reprocess_materials = list(zip(reprocess_material_usage['materialTypeID'], reprocess_material_usage['materialName']))
        
        reprocess_pivot_data = []
        for item_id in reprocessing_df['itemTypeID'].unique():
            item_data = reprocessing_df[reprocessing_df['itemTypeID'] == item_id]
            item_name = item_data.iloc[0]['itemName']
            
            pivot_row = {
                'Item Name': item_name,
                'Item TypeID': item_id
            }
            
            # Initialize all material columns to 0
            for mat_id, mat_name in sorted_reprocess_materials:
                pivot_row[mat_name] = 0
            
            # Fill in actual quantities
            for _, mat in item_data.iterrows():
                mat_name = mat['materialName']
                if mat_name in pivot_row:
                    pivot_row[mat_name] = int(mat['quantity'])
            
            # Add price and volume
            price_row = prices_df[prices_df['typeID'] == item_id]
            if len(price_row) > 0:
                pivot_row['Item Price (ISK)'] = float(price_row.iloc[0]['sell_min'])
            else:
                pivot_row['Item Price (ISK)'] = 0
            
            item_row = items_df[items_df['typeID'] == item_id]
            if len(item_row) > 0:
                packaged_vol = item_row.iloc[0]['packaged_volume']
                if packaged_vol and packaged_vol > 0:
                    pivot_row['Item Volume (m³)'] = float(packaged_vol)
                else:
                    pivot_row['Item Volume (m³)'] = float(item_row.iloc[0]['volume'] or 0)
            else:
                pivot_row['Item Volume (m³)'] = 0
            
            reprocess_pivot_data.append(pivot_row)
        
        reprocess_pivot_df = pd.DataFrame(reprocess_pivot_data)
        
        # Create Excel file
        logger.info(f"Writing Excel file: {OUTPUT_FILE}")
        
        with pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter') as writer:
            # Manufacturing sheet
            mfg_pivot_df.to_excel(writer, sheet_name='Manufacturing', index=False)
            worksheet = writer.sheets['Manufacturing']
            worksheet.set_column('A:A', 30)  # Product Name
            worksheet.set_column('B:B', 15)  # Product TypeID
            worksheet.set_column('C:C', 20)  # Group
            worksheet.set_column('D:D', 12)  # Output Qty
            
            # Material columns
            num_materials = len(sorted_materials)
            for i in range(num_materials):
                col_idx = 4 + i
                worksheet.set_column(col_idx, col_idx, 12)
            
            # Skills, Price, Volume columns
            skills_col = 4 + num_materials
            price_col = skills_col + 1
            volume_col = price_col + 1
            worksheet.set_column(skills_col, skills_col, 40)
            worksheet.set_column(price_col, price_col, 18)
            worksheet.set_column(volume_col, volume_col, 18)
            
            # Reprocessing sheet
            reprocess_pivot_df.to_excel(writer, sheet_name='Reprocessing', index=False)
            worksheet = writer.sheets['Reprocessing']
            worksheet.set_column('A:A', 30)
            worksheet.set_column('B:B', 15)
            
            num_reprocess_materials = len(sorted_reprocess_materials)
            for i in range(num_reprocess_materials):
                col_idx = 2 + i
                worksheet.set_column(col_idx, col_idx, 12)
            
            price_col = 2 + num_reprocess_materials
            volume_col = price_col + 1
            worksheet.set_column(price_col, price_col, 18)
            worksheet.set_column(volume_col, volume_col, 18)
            
            # Prices sheet
            prices_df.to_excel(writer, sheet_name='Prices', index=False)
            worksheet = writer.sheets['Prices']
            worksheet.set_column('A:A', 15)
            worksheet.set_column('B:B', 30)
            for col in ['C', 'D', 'E', 'F', 'G', 'H']:
                worksheet.set_column(f'{col}:{col}', 18)
        
        logger.info("=" * 60)
        logger.info(f"SUCCESS! Excel file created: {OUTPUT_FILE}")
        logger.info("=" * 60)
        
    except Exception as e:
        logger.error(f"Error generating Excel: {e}")
        import traceback
        logger.error(traceback.format_exc())
        raise
    finally:
        conn.close()

if __name__ == "__main__":
    generate_excel()

