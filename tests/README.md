# Test Scripts and Utilities

This folder contains test scripts, debugging utilities, and one-off scripts used during development to verify functionality and check data.

## Files in this folder

### Test Scripts
- `test_tremor_calc.py` - Test calculation for Tremor L module
- `test_tremor_100.py` - Test Tremor L calculation with 100 modules
- `test_fuzzwork_prices.py` - Test Fuzzwork market price API integration

### Check/Debug Scripts
- `check_tremor.py` - Check Tremor L reprocessing data in database
- `check_morphite.py` - Check Morphite calculation for Tremor L
- `check_missile_batch.py` - Check missile batch size data
- `check_missile_plasma_batch.py` - Check plasma missile batch size data
- `check_reprocessing_structure.py` - Check reprocessing table structure
- `check_sde_updates.py` - Check SDE data updates

### Fix Scripts
- `fix_tremor_batch.py` - Fix Tremor L batch_size in database

## Note

These scripts are **not part of the core application**. They are:
- Development/debugging tools
- One-off data verification scripts
- Temporary test scripts

They can be safely ignored when reviewing the main codebase.

