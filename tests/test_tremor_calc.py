from calculate_reprocessing_value import calculate_reprocessing_value

result = calculate_reprocessing_value(
    module_name='Tremor L',
    num_modules=100,
    yield_percent=55.0,
    module_price_type='buy_max',
    mineral_price_type='buy_max'
)

print("Tremor L Calculation (100 modules):")
print(f"Module Price: {result.get('module_price', 0):,.2f}")
print(f"Total Module Price: {result.get('total_module_price', 0):,.2f}")
print("\nOutputs:")
for m in result.get('reprocessing_outputs', []):
    print(f"  {m['materialName']}: {m['actualQuantity']} @ {m['mineralPrice']:,.2f} = {m['mineralValue']:,.2f}")
print(f"\nTotal Mineral Value: {result.get('total_mineral_value', 0):,.2f}")

# Expected values from user:
print("\nExpected (from user):")
print("  Total Module Price: 36,608")
print("  Morphite: 30 × 19,620 = 588,600")
print("  Fernite Carbide: 6,000 × 47.05 = 282,300")
print("  Fullerides: 1,500 × 802.1 = 1,203,150")
print("  Total: 2,074,050")


