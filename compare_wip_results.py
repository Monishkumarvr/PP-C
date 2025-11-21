import pandas as pd
import sys

# Compare utilization across different WIP levels
wip_levels = ['10pct', '30pct', '50pct', '100pct']
results = {}

print("\n" + "="*100)
print("WIP SENSITIVITY ANALYSIS - UTILIZATION COMPARISON")
print("="*100)

for wip in wip_levels:
    try:
        if wip == '100pct':
            filename = 'production_plan_COMPREHENSIVE_test.xlsx'
        else:
            filename = f'production_plan_COMPREHENSIVE_{wip}_WIP.xlsx'
        
        ws = pd.read_excel(filename, sheet_name='Weekly_Summary')
        util_cols = [c for c in ws.columns if 'Util_%' in c]
        
        results[wip] = {}
        for col in util_cols:
            stage = col.replace('_Util_%', '')
            avg = ws[col].mean()
            max_val = ws[col].max()
            results[wip][stage] = {'avg': avg, 'max': max_val}
        
        print(f"\n✅ Loaded: {wip.replace('pct', '%')} WIP")
    except Exception as e:
        print(f"\n⚠️  Not found: {wip.replace('pct', '%')} WIP - {e}")

# Create comparison table
if results:
    print("\n" + "="*100)
    print("AVERAGE UTILIZATION COMPARISON")
    print("="*100)
    print(f"{'Stage':20s}", end='')
    for wip in wip_levels:
        if wip in results:
            print(f"{wip.replace('pct', '%'):>12s}", end='')
    print()
    print("-"*100)
    
    # Get all stages
    all_stages = set()
    for wip_data in results.values():
        all_stages.update(wip_data.keys())
    
    for stage in sorted(all_stages):
        print(f"{stage:20s}", end='')
        for wip in wip_levels:
            if wip in results and stage in results[wip]:
                avg = results[wip][stage]['avg']
                print(f"{avg:11.1f}%", end='')
            else:
                print(f"{'N/A':>12s}", end='')
        print()
    
    print("\n" + "="*100)
    print("MAXIMUM UTILIZATION COMPARISON")
    print("="*100)
    print(f"{'Stage':20s}", end='')
    for wip in wip_levels:
        if wip in results:
            print(f"{wip.replace('pct', '%'):>12s}", end='')
    print()
    print("-"*100)
    
    for stage in sorted(all_stages):
        print(f"{stage:20s}", end='')
        for wip in wip_levels:
            if wip in results and stage in results[wip]:
                max_val = results[wip][stage]['max']
                print(f"{max_val:11.1f}%", end='')
            else:
                print(f"{'N/A':>12s}", end='')
        print()

print("\n" + "="*100)
