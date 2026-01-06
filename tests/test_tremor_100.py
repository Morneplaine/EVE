from calculate_reprocessing_value import calculate_reprocessing_value

result = calculate_reprocessing_value(
    module_name='Tremor L',
    num_modules=100,
    yield_percent=55.0,
    buy_order_markup_percent=10.0,
    module_price_type='buy_max',
    mineral_price_type='buy_max'
)

print("Tremor L Calculation (100 modules):")
print(f"Module Price (before markup): {result.get('module_price_before_markup', 0):,.2f}")
print(f"Module Price (after markup): {result.get('module_price', 0):,.2f}")
print(f"Total Module Price: {result.get('total_module_price', 0):,.2f}")
print(f"\nExpected:")
print(f"  Cost = 332.8 × 100 + markup = 33,280 + (33,280 × 10%) = 36,608")
print(f"\nOutputs:")
for m in result.get('reprocessing_outputs', []):
    print(f"  {m['materialName']}: {m['actualQuantity']}")

print(f"\nExpected outputs:")
print(f"  Morphite: 16")
print(f"  Fernite Carbide: 3300")
print(f"  Fullerides: 825")

