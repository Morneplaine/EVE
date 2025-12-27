"""
Check if SDE (Static Data Export) files have been updated on Fuzzwork.
This script compares local file sizes with remote file sizes to detect updates.
"""

import requests
from pathlib import Path
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

FUZZWORK_BASE = "https://www.fuzzwork.co.uk/dump/latest/"
DATA_DIR = Path("eve_data")

# Required CSV files from Fuzzwork SDE
REQUIRED_FILES = [
    "invTypes.csv.bz2",
    "invGroups.csv.bz2",
    "industryActivityMaterials.csv.bz2",
    "industryActivityProducts.csv.bz2",
    "industryActivitySkills.csv.bz2",
    "invTypeMaterials.csv.bz2",
    "industryActivity.csv.bz2",
    "invVolumes.csv.bz2",
]

def check_file_update(filename):
    """
    Check if a remote file is newer than the local file.
    
    Returns:
        tuple: (needs_update, local_size, remote_size, status_message)
    """
    local_path = DATA_DIR / filename
    url = FUZZWORK_BASE + filename
    
    # Check if local file exists
    local_exists = local_path.exists()
    local_size = local_path.stat().st_size if local_exists else 0
    
    try:
        # Try HEAD request first (lighter)
        response = requests.head(url, timeout=10, allow_redirects=True)
        response.raise_for_status()
        
        # Get content length
        remote_size_str = response.headers.get('Content-Length')
        last_modified = response.headers.get('Last-Modified', 'Unknown')
        
        # If Content-Length not available, try GET with range request
        if not remote_size_str:
            # Use range request to get just the first byte to check if file exists
            headers = {'Range': 'bytes=0-0'}
            range_response = requests.get(url, headers=headers, timeout=10, allow_redirects=True)
            if range_response.status_code == 206:  # Partial content
                # Get actual content length from Content-Range header
                content_range = range_response.headers.get('Content-Range', '')
                if '/' in content_range:
                    remote_size_str = content_range.split('/')[-1]
                else:
                    # Fallback: assume file exists but size unknown
                    remote_size_str = 'unknown'
            else:
                remote_size_str = '0'
        
        try:
            remote_size = int(remote_size_str) if remote_size_str != 'unknown' else 0
        except (ValueError, TypeError):
            remote_size = 0
        
        if not local_exists:
            return (True, 0, remote_size if remote_size > 0 else 'unknown', f"File not found locally - needs download")
        
        if remote_size == 0 or remote_size_str == 'unknown':
            # Can't determine remote size, assume update needed if local exists
            return (None, local_size, 'unknown', f"Could not determine remote size (local: {local_size:,} bytes)")
        
        if local_size != remote_size:
            return (True, local_size, remote_size, f"Size mismatch: local={local_size:,} bytes, remote={remote_size:,} bytes")
        
        return (False, local_size, remote_size, f"Up to date (size: {local_size:,} bytes, modified: {last_modified})")
        
    except requests.exceptions.RequestException as e:
        logger.warning(f"Could not check {filename}: {e}")
        return (None, local_size, 0, f"Error checking remote file: {e}")

def main():
    """Check all SDE files for updates"""
    logger.info("=" * 60)
    logger.info("Checking for SDE (Static Data Export) Updates")
    logger.info("=" * 60)
    logger.info("")
    
    DATA_DIR.mkdir(exist_ok=True)
    
    needs_update = []
    up_to_date = []
    errors = []
    
    for filename in REQUIRED_FILES:
        logger.info(f"Checking {filename}...")
        update_needed, local_size, remote_size, status = check_file_update(filename)
        
        if update_needed is True:
            needs_update.append((filename, local_size, remote_size, status))
            logger.info(f"  ⚠ UPDATE AVAILABLE: {status}")
        elif update_needed is False:
            up_to_date.append((filename, status))
            logger.info(f"  ✓ {status}")
        else:
            errors.append((filename, status))
            logger.warning(f"  ✗ {status}")
        logger.info("")
    
    # Summary
    logger.info("=" * 60)
    logger.info("SUMMARY")
    logger.info("=" * 60)
    logger.info(f"Files up to date: {len(up_to_date)}")
    logger.info(f"Files needing update: {len(needs_update)}")
    logger.info(f"Errors: {len(errors)}")
    logger.info("")
    
    if needs_update:
        logger.info("Files that need updating:")
        for filename, local_size, remote_size, status in needs_update:
            logger.info(f"  - {filename}: {status}")
        logger.info("")
        logger.info("To update, run: python build_database.py")
        logger.info("(This will download updated files and rebuild the database)")
    else:
        logger.info("All SDE files are up to date!")
        logger.info("No database rebuild needed.")
    
    logger.info("=" * 60)

if __name__ == "__main__":
    main()

