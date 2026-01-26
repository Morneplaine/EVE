"""
Main Assumptions and Configuration Parameters
=============================================
This file contains the main assumptions used throughout the EVE manufacturing
and reprocessing calculations. Edit these values to adjust calculations.

All percentages are in their decimal form (e.g., 1.37 means 1.37%)
"""

# Transaction Costs
BROKER_FEE = 1.37  # Broker fee percentage when placing buy orders
SALES_TAX = 3.5    # Sales tax percentage when selling items
LISTING_RELIST = 3  # Average number of times an order needs to be relisted

# Reprocessing Costs
REPROCESSING_COST = 3.37  # Base reprocessing cost percentage

# Other commonly used defaults
DEFAULT_YIELD_PERCENT = 55.0  # Default reprocessing yield percentage
BUY_ORDER_MARKUP_PERCENT = 10.0  # Default markup percentage for buy orders
BUY_BUFFER = 0.1  # Buffer percentage for buy orders
RELIST_DISCOUNT = 80
