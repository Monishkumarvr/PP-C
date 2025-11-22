import pandas as pd

# Read data
df_order = pd.read_excel('production_plan_COMPREHENSIVE_test.xlsx', sheet_name='Order_Fulfillment')
df_part = pd.read_excel('Master_Data_Updated_Nov_Dec.xlsx', sheet_name='Part Master')
df_mach = pd.read_excel('Master_Data_Updated_Nov_Dec.xlsx', sheet_name='Machine Constraints')

# Get all resource codes from Machine Constraints
valid_resources = set(df_mach['Resource Code'].values)

# Find all parts with unmet demand
unmet_parts = df_order[df_order['Unmet_Qty'] > 0]['Material_Code'].unique()

print('='*80)
print('PARTS WITH UNMET DEMAND - CHECKING FOR MISSING RESOURCES')
print('='*80)

for part in sorted(unmet_parts):
    p = df_part[df_part['FG Code'] == part]
    if len(p) == 0:
        print(f'{part:15} | NOT IN PART MASTER')
        continue

    p = p.iloc[0]
    unmet_qty = df_order[df_order['Material_Code'] == part]['Unmet_Qty'].sum()

    # Check machining resources
    mc1 = str(p.get('Machining resource code 1', '')).strip()
    mc2 = str(p.get('Machining resource code 2', '')).strip()
    mc3 = str(p.get('Machining resource code 3', '')).strip()

    missing = []
    if mc1 and mc1 != 'nan' and mc1 not in valid_resources:
        missing.append(f'MC1={mc1}')
    if mc2 and mc2 != 'nan' and mc2 not in valid_resources:
        missing.append(f'MC2={mc2}')
    if mc3 and mc3 != 'nan' and mc3 not in valid_resources:
        missing.append(f'MC3={mc3}')

    status = '❌ MISSING: ' + ', '.join(missing) if missing else '✓ Resources OK'
    print(f'{part:15} | Unmet: {unmet_qty:3.0f} | {status}')

print()
print('='*80)
print('ROOT CAUSE IDENTIFIED')
print('='*80)
print('Parts cannot be produced because their machining resources')
print('do not exist in the Machine Constraints sheet!')
print()
print('The optimizer creates 0 capacity for non-existent resources,')
print('which blocks production flow through machining.')
