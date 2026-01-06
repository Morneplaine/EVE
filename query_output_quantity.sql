-- Query to get outputQuantity from blueprints table given a module name
-- This query returns the number of items produced by a blueprint for a given product name

SELECT outputQuantity 
FROM blueprints 
WHERE productName = ?;

-- Example usage in Python:
-- output_quantity_df = pd.read_sql_query(
--     "SELECT outputQuantity FROM blueprints WHERE productName = ?",
--     conn,
--     params=(module_name,)
-- )
-- 
-- if len(output_quantity_df) > 0:
--     output_quantity = int(output_quantity_df.iloc[0]['outputQuantity'])
-- else:
--     output_quantity = None  # No blueprint found for this module


