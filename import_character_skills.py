"""
Import character skills into the database
Supports:
1. Manual CSV import (typeID, skillName, level)
2. ESI API import (requires access token)
3. Manual entry via command line
"""

import sqlite3
import pandas as pd
import csv
import logging
import sys
from pathlib import Path

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

DB_FILE = "eve_manufacturing.db"

def import_from_csv(csv_file):
    """Import skills from CSV file
    
    CSV format should be:
    typeID,skillName,level
    3300,Science,5
    3301,Research,4
    """
    logger.info(f"Importing skills from CSV: {csv_file}")
    
    conn = sqlite3.connect(DB_FILE)
    
    try:
        # Read CSV
        skills_df = pd.read_csv(csv_file)
        
        # Validate columns
        required_cols = ['typeID', 'level']
        if not all(col in skills_df.columns for col in required_cols):
            raise ValueError(f"CSV must contain columns: {required_cols}")
        
        # If skillName not provided, try to get from items table
        if 'skillName' not in skills_df.columns:
            items_query = "SELECT typeID, typeName FROM items WHERE typeID IN ({})".format(
                ','.join(['?'] * len(skills_df))
            )
            items_df = pd.read_sql_query(items_query, conn, params=skills_df['typeID'].tolist())
            skills_df = skills_df.merge(items_df, on='typeID', how='left')
            skills_df['skillName'] = skills_df['typeName']
            skills_df = skills_df.drop('typeName', axis=1)
        
        # Clear existing skills
        conn.execute("DELETE FROM character_skills")
        
        # Insert new skills
        for _, row in skills_df.iterrows():
            conn.execute("""
                INSERT OR REPLACE INTO character_skills (skillID, skillName, level)
                VALUES (?, ?, ?)
            """, (
                int(row['typeID']),
                str(row.get('skillName', 'Unknown')),
                int(row['level'])
            ))
        
        conn.commit()
        logger.info(f"Imported {len(skills_df)} skills")
        
    except Exception as e:
        logger.error(f"Error importing from CSV: {e}")
        conn.rollback()
        raise
    finally:
        conn.close()

def import_from_esi(access_token, character_id):
    """Import skills from ESI API
    
    Requires:
    - access_token: ESI OAuth access token
    - character_id: Character ID
    """
    import requests
    
    logger.info(f"Importing skills from ESI for character {character_id}")
    
    # ESI endpoint for character skills
    url = f"https://esi.evetech.net/latest/characters/{character_id}/skills/"
    headers = {
        'Authorization': f'Bearer {access_token}'
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        skills_data = response.json()
        
        conn = sqlite3.connect(DB_FILE)
        
        try:
            # Clear existing skills
            conn.execute("DELETE FROM character_skills")
            
            # Get skill names from items table
            skill_ids = [skill['skill_id'] for skill in skills_data.get('skills', [])]
            if skill_ids:
                items_query = "SELECT typeID, typeName FROM items WHERE typeID IN ({})".format(
                    ','.join(['?'] * len(skill_ids))
                )
                items_df = pd.read_sql_query(items_query, conn, params=skill_ids)
                items_dict = dict(zip(items_df['typeID'], items_df['typeName']))
            else:
                items_dict = {}
            
            # Insert skills
            for skill in skills_data.get('skills', []):
                skill_id = skill['skill_id']
                skill_level = skill.get('active_skill_level', 0)
                skill_name = items_dict.get(skill_id, f"Skill {skill_id}")
                
                conn.execute("""
                    INSERT OR REPLACE INTO character_skills (skillID, skillName, level)
                    VALUES (?, ?, ?)
                """, (skill_id, skill_name, skill_level))
            
            conn.commit()
            logger.info(f"Imported {len(skills_data.get('skills', []))} skills from ESI")
            
        except Exception as e:
            logger.error(f"Error saving skills to database: {e}")
            conn.rollback()
            raise
        finally:
            conn.close()
            
    except Exception as e:
        logger.error(f"Error fetching from ESI: {e}")
        raise

def manual_entry():
    """Interactive manual entry of skills"""
    logger.info("Manual skill entry (type 'done' when finished)")
    
    conn = sqlite3.connect(DB_FILE)
    
    try:
        # Clear existing skills
        conn.execute("DELETE FROM character_skills")
        
        skills = []
        while True:
            skill_input = input("Enter skill (typeID or skillName, level) or 'done': ").strip()
            
            if skill_input.lower() == 'done':
                break
            
            try:
                parts = skill_input.split(',')
                if len(parts) == 2:
                    skill_id_or_name = parts[0].strip()
                    level = int(parts[1].strip())
                    
                    # Try to resolve skill name to typeID
                    if skill_id_or_name.isdigit():
                        skill_id = int(skill_id_or_name)
                        skill_name_query = "SELECT typeName FROM items WHERE typeID = ?"
                        result = conn.execute(skill_name_query, (skill_id,)).fetchone()
                        skill_name = result[0] if result else f"Skill {skill_id}"
                    else:
                        # Look up by name
                        skill_id_query = "SELECT typeID FROM items WHERE typeName LIKE ? LIMIT 1"
                        result = conn.execute(skill_id_query, (f"%{skill_id_or_name}%",)).fetchone()
                        if result:
                            skill_id = result[0]
                            skill_name = skill_id_or_name
                        else:
                            logger.warning(f"Skill not found: {skill_id_or_name}")
                            continue
                    
                    skills.append((skill_id, skill_name, level))
                    logger.info(f"Added: {skill_name} (level {level})")
                else:
                    logger.warning("Format: skillName,level or typeID,level")
            except ValueError:
                logger.warning("Invalid level, must be a number")
        
        # Insert skills
        for skill_id, skill_name, level in skills:
            conn.execute("""
                INSERT OR REPLACE INTO character_skills (skillID, skillName, level)
                VALUES (?, ?, ?)
            """, (skill_id, skill_name, level))
        
        conn.commit()
        logger.info(f"Imported {len(skills)} skills")
        
    except Exception as e:
        logger.error(f"Error in manual entry: {e}")
        conn.rollback()
        raise
    finally:
        conn.close()

def main():
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python import_character_skills.py csv <file.csv>")
        print("  python import_character_skills.py esi <access_token> <character_id>")
        print("  python import_character_skills.py manual")
        return
    
    mode = sys.argv[1].lower()
    
    if mode == 'csv':
        if len(sys.argv) < 3:
            logger.error("Please provide CSV file path")
            return
        import_from_csv(sys.argv[2])
    elif mode == 'esi':
        if len(sys.argv) < 4:
            logger.error("Please provide access_token and character_id")
            return
        import_from_esi(sys.argv[2], sys.argv[3])
    elif mode == 'manual':
        manual_entry()
    else:
        logger.error(f"Unknown mode: {mode}")

if __name__ == "__main__":
    main()

