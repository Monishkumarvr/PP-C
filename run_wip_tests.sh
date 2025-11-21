#!/bin/bash

echo "Running WIP sensitivity analysis..."

# Function to run optimizer with specific WIP percentage
run_wip_test() {
    WIP_PCT=$1
    echo ""
    echo "============================================"
    echo "Testing with ${WIP_PCT}% WIP"
    echo "============================================"
    
    # Update the file path in production_plan_test.py
    sed -i "s/Master_Data_Updated_Nov_Dec_[0-9]*pct_WIP.xlsx/Master_Data_Updated_Nov_Dec_${WIP_PCT}pct_WIP.xlsx/g" production_plan_test.py
    
    # Run the optimizer
    python3 production_plan_test.py > /tmp/opt_${WIP_PCT}pct.log 2>&1
    
    # Save output file with unique name
    cp production_plan_COMPREHENSIVE_test.xlsx production_plan_COMPREHENSIVE_${WIP_PCT}pct_WIP.xlsx
    
    # Extract utilization summary
    python3 << PYTHON_EOF
import pandas as pd

ws = pd.read_excel('production_plan_COMPREHENSIVE_${WIP_PCT}pct_WIP.xlsx', sheet_name='Weekly_Summary')
util_cols = [c for c in ws.columns if 'Util' in c and c != 'Week']

print("\n${WIP_PCT}% WIP Results:")
print("="*80)
for col in util_cols:
    avg = ws[col].mean()
    max_val = ws[col].max()
    print(f"{col:20s}: Avg={avg:6.1f}%, Max={max_val:6.1f}%")
PYTHON_EOF
}

# Run tests
run_wip_test 30
run_wip_test 50

# Restore original
sed -i "s/Master_Data_Updated_Nov_Dec_[0-9]*pct_WIP.xlsx/Master_Data_Updated_Nov_Dec.xlsx/g" production_plan_test.py

echo ""
echo "============================================"
echo "WIP Sensitivity Analysis Complete"
echo "============================================"
