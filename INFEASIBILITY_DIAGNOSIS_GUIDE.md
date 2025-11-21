# Infeasibility Diagnosis and Fix Guide

## Overview

This guide explains how to diagnose and fix optimization infeasibility issues in the Production Planning system.

## Quick Start

When optimization fails with "Not Solved" or "Infeasible" status:

```bash
# Step 1: Diagnose the issues
python3 diagnose_infeasibility.py

# Step 2: Automatically fix the issues
python3 fix_infeasibility.py

# Step 3: Verify the fixes
python3 diagnose_infeasibility.py

# Step 4: Run optimization with fixed data
# Update production_plan_test.py to use: Master_Data_Updated_Nov_Dec_FIXED.xlsx
python3 production_plan_test.py
```

## What These Scripts Do

### 1. `diagnose_infeasibility.py`

**Purpose**: Identifies root causes of optimization infeasibility.

**Checks Performed**:

| Check | What it Does | Critical Issues Detected |
|-------|--------------|-------------------------|
| **Order Validity** | Verifies all ordered parts exist in Part Master | Missing parts that can't be produced |
| **Capacity Analysis** | Compares demand vs available capacity by stage | Capacity shortfalls (>100% utilization) |
| **Lead Time Analysis** | Checks if delivery dates are achievable | Past-due orders, insufficient lead time |
| **WIP Mapping** | Validates WIP inventory mapping to parts | Unmapped WIP (usually non-critical) |
| **Box Capacity** | Checks mould box capacity constraints | Box size capacity exceeded |

**Output**: Detailed diagnostic report with:
- Critical issues (causes infeasibility)
- Warnings (may cause suboptimal results)
- Recommendations for fixes

**Example Output**:
```
❌ CRITICAL ISSUES (5):
1. [Order Validity] 1 parts in orders NOT in Part Master
2. [Lead Time] 100 orders cannot meet delivery date
3. [Box Capacity] Box size 1050X750 capacity exceeded (126.4%)
4. [Box Capacity] Box size 400X625 capacity exceeded (190.5%)
5. [Box Capacity] Box size 750X500 capacity exceeded (110.8%)
```

### 2. `fix_infeasibility.py`

**Purpose**: Automatically fixes the identified critical issues.

**Fixes Applied**:

#### Fix 1: Delivery Date Adjustments
- **What**: Adjusts past-due or insufficient lead time delivery dates
- **How**: Calculates minimum lead time (casting → grinding → machining → painting + passive times)
- **Action**: Moves delivery date forward to be feasible (earliest possible date + 1 day buffer)

#### Fix 2: Box Capacity Increases
- **What**: Increases mould box weekly capacity where exceeded
- **How**: Calculates total demand over planning horizon
- **Action**: Sets new capacity to 150% of demand to provide buffer

#### Fix 3: Invalid Parts Removal
- **What**: Removes orders for parts not in Part Master
- **How**: Filters orders against valid Part Master FG Codes
- **Action**: Deletes order lines for invalid parts

**Safety Features**:
- ✅ Creates backup before modifications (`Master_Data_*_BACKUP.xlsx`)
- ✅ Creates new fixed file (doesn't overwrite original)
- ✅ Provides detailed summary of all changes made

**Example Output**:
```
✅ Fixes Applied:
   1. Adjusted 100 delivery dates to be feasible
   2. Increased capacity for 3 box sizes
   3. Removed 1 orders for invalid parts

✓ Fixed data saved to: Master_Data_Updated_Nov_Dec_FIXED.xlsx
```

## Common Infeasibility Causes

### 1. **Lead Time Conflicts** (Most Common)

**Symptom**: Orders with delivery dates in the past or too close to current date.

**Root Causes**:
- Sales Order sheet not updated with current date context
- Rush orders accepted without checking production lead time
- Planning date (`CURRENT_DATE`) not updated in config

**Example**:
```
Part: LKK-033
Delivery: 2025-11-01
Current Date: 2025-11-22
Issue: Delivery is 21 days in the PAST!
```

**Fix Options**:
- **Automatic**: Run `fix_infeasibility.py` (adjusts dates forward)
- **Manual**: Update delivery dates in Sales Order sheet to realistic future dates
- **Config**: Update `CURRENT_DATE` in `ProductionConfig` class if planning for past

### 2. **Capacity Constraints Exceeded**

**Symptom**: Stage or box capacity utilization >100%.

**Root Causes**:
- More orders than production capacity can handle
- Machine capacity data not up to date
- Box capacity limits too conservative

**Example**:
```
Box Size: 400X625
Demand: 400 boxes
Capacity: 210 boxes (over 7 weeks)
Utilization: 190.5% ❌ OVERFLOW
```

**Fix Options**:
- **Automatic**: Run `fix_infeasibility.py` (increases capacity by 50%)
- **Manual**:
  - Increase weekly capacity in Mould Box Capacity sheet
  - Add overtime/extra shifts in Machine Constraints
  - Negotiate later delivery dates to spread demand
  - Decline some orders

### 3. **Invalid Parts in Orders**

**Symptom**: Orders exist for parts not defined in Part Master.

**Root Causes**:
- Typos in Material Code
- New parts ordered before being added to Part Master
- Part Master out of sync with ERP system

**Example**:
```
Part: KBS-013
Orders: 24 units
Issue: Part not found in Part Master FG Code column
```

**Fix Options**:
- **Automatic**: Run `fix_infeasibility.py` (removes invalid orders)
- **Manual**:
  - Add missing part to Part Master sheet with all required data
  - Or fix typo in Sales Order sheet Material Code
  - Or remove the order if it was entered in error

### 4. **WIP Mapping Issues** (Usually Non-Critical)

**Symptom**: WIP items don't map to Part Master CS Code or FG Code.

**Root Causes**:
- WIP for obsolete parts
- Part naming convention changes
- Sample parts or one-off items

**Impact**: Usually just a warning. WIP for unmapped parts is ignored (doesn't cause infeasibility).

**Fix Options**:
- Add missing CS Codes to Part Master
- Or remove obsolete WIP from Stage WIP sheet
- Or update WIP CastingItem/Material Code to match Part Master

## Step-by-Step Fix Process

### Step 1: Run Diagnostics

```bash
python3 diagnose_infeasibility.py
```

**Review the output carefully**:
- How many CRITICAL issues?
- Which categories (Order Validity, Lead Time, Capacity)?
- Are the issues data entry errors or real constraints?

### Step 2: Decide Fix Strategy

**Option A: Automatic Fix (Recommended for first-time)**
```bash
python3 fix_infeasibility.py
```

**Option B: Manual Fix (If you want more control)**
- Open `Master_Data_Updated_Nov_Dec.xlsx`
- Fix issues based on diagnostic report
- Save and re-run diagnostics to verify

**Option C: Hybrid (Fix some manually, auto-fix rest)**
- Fix critical issues manually (e.g., add missing parts to Part Master)
- Run auto-fix for bulk adjustments (e.g., date adjustments)

### Step 3: Verify Fixes

```bash
# Run diagnostics on fixed file
python3 diagnose_infeasibility.py
```

**Expected output**:
```
✅ NO ISSUES FOUND - Optimization should be feasible!
```

Or only non-critical warnings remaining:
```
⚠️  WARNINGS (1):
1. [WIP Mapping] 12 CS WIP parts not mapped
```

### Step 4: Update Optimization to Use Fixed Data

**Edit your optimization script** (e.g., `production_plan_test.py`):

```python
# Find this line:
master_file = 'Master_Data_Updated_Nov_Dec.xlsx'

# Change to:
master_file = 'Master_Data_Updated_Nov_Dec_FIXED.xlsx'
```

Or specify in command line:
```bash
python3 production_plan_test.py --master-file Master_Data_Updated_Nov_Dec_FIXED.xlsx
```

### Step 5: Run Optimization

```bash
python3 production_plan_test.py
```

**Monitor for success**:
- Look for "Solver completed: Optimal" status
- Check that all orders are fulfilled (100% fulfillment)
- Verify output files are generated

## Advanced: Understanding the Fixes

### Delivery Date Adjustment Logic

The fix script calculates minimum lead time as:

```
Total Lead Time = Sum of:
  - Casting cycle time (active)
  - Cooling time (passive)
  - Shakeout time (passive)
  - Vacuum time (passive)
  - Grinding cycle time (active)
  - Machining cycle times (3 stages, sequential, active)
  - Painting cycle times (3 stages, sequential, active)
  - Painting dry times (3 stages, sequential, passive)

Plus 20% buffer for setup, material handling, etc.
```

**New Delivery Date** = `CURRENT_DATE + ceiling(Total Lead Time) + 1 day`

### Box Capacity Increase Logic

```
Current Weekly Capacity = X boxes/week
Planning Horizon = W weeks
Total Current Capacity = X * W boxes

Demand = D boxes (over planning horizon)

If D > (X * W):
    New Weekly Capacity = ceiling(D / W * 1.5)
    # 1.5 factor provides 50% buffer
```

## Troubleshooting

### Issue: "Still infeasible after auto-fix"

**Possible Causes**:
1. **Model complexity**: Daily optimization with 300K+ variables may time out
   - **Solution**: Use weekly optimization (`production_plan_test.py`) instead of daily

2. **Stage seriality conflicts**: Upstream stages blocking downstream
   - **Solution**: Check that capacity increases were applied to all constrained stages

3. **Solver timeout**: Not enough time to find solution
   - **Solution**: Increase solver timeout from 600s to 1800s (30 min)

### Issue: "Diagnostic script crashes"

**Common Causes**:
- **Duplicate FG Codes**: Fixed automatically now (keeps first occurrence)
- **Missing Excel sheets**: Check all 5 required sheets exist
- **Date parsing errors**: Verify date format in Sales Order (YYYY-MM-DD)

### Issue: "Auto-fix too aggressive"

**Examples**:
- Moved delivery dates too far forward
- Increased capacity more than realistically possible

**Solution**: Use manual fix approach:
1. Restore from backup: `cp Master_Data_*_BACKUP.xlsx Master_Data_*.xlsx`
2. Manually adjust only the most critical issues
3. Run diagnostics again to check remaining issues

## Integration with Optimization Workflow

### Recommended Workflow

```
┌─────────────────────────────────────┐
│ 1. Update Master Data Excel         │
│    - Sales Orders                    │
│    - Part Master                     │
│    - WIP                             │
└──────────────┬──────────────────────┘
               │
               ▼
┌─────────────────────────────────────┐
│ 2. Run Diagnostics                  │
│    python diagnose_infeasibility.py │
└──────────────┬──────────────────────┘
               │
          ┌────┴─────┐
          │ Issues?  │
          └────┬─────┘
               │
        ┌──────┴───────┐
        │ YES          │ NO
        ▼              ▼
┌───────────────┐  ┌────────────────┐
│ 3. Auto-Fix   │  │ 4. Run         │
│    OR         │  │    Optimization│
│    Manual Fix │  └────────────────┘
└───────┬───────┘
        │
        ▼
┌───────────────────────────────────┐
│ 5. Verify with Diagnostics Again │
└───────────┬───────────────────────┘
            │
            ▼
┌──────────────────────────────────┐
│ 6. Run Optimization with Fixed   │
│    Master Data                   │
└──────────────────────────────────┘
```

### Automated Daily Workflow

For daily scheduled runs, create a wrapper script:

```bash
#!/bin/bash
# daily_production_plan.sh

# 1. Diagnose
python3 diagnose_infeasibility.py > diagnostic_log.txt

# 2. Check if critical issues found
if grep -q "CRITICAL ISSUES" diagnostic_log.txt; then
    echo "Critical issues found. Auto-fixing..."
    python3 fix_infeasibility.py

    # Update master file reference
    export MASTER_FILE="Master_Data_Updated_Nov_Dec_FIXED.xlsx"
else
    export MASTER_FILE="Master_Data_Updated_Nov_Dec.xlsx"
fi

# 3. Run optimization
python3 production_plan_test.py

# 4. Generate reports
python3 production_plan_executive_test7sheets.py
python3 run_decision_support.py

# 5. Send email notification
python3 send_report_email.py
```

## Files Generated

| File | Description | When Created |
|------|-------------|--------------|
| `Master_Data_*_BACKUP.xlsx` | Backup of original data | Before auto-fix |
| `Master_Data_*_FIXED.xlsx` | Fixed master data | After auto-fix |
| `diagnostic_log.txt` | Diagnostic report | Manual save (redirect output) |

## Best Practices

### 1. **Always Run Diagnostics First**
- Don't blindly run optimization
- Diagnose → Fix → Verify → Optimize

### 2. **Review Auto-Fixes Before Using**
- Check if delivery date changes are acceptable
- Verify capacity increases are realistic
- Confirm invalid part removals are correct

### 3. **Maintain Data Quality**
- Keep Part Master up to date
- Update Sales Orders with current context
- Clean up obsolete WIP periodically

### 4. **Document Manual Fixes**
- Keep notes on why manual adjustments were made
- Track recurring issues (may indicate process problems)

### 5. **Backup Before Fixing**
- Auto-fix creates backup automatically
- For manual fixes: save a copy first

### 6. **Version Control for Master Data**
- Use git to track Master Data changes
- Tag versions before major runs
- Keep audit trail of data modifications

## FAQ

**Q: Why does the optimizer fail with "Not Solved" even when diagnostics show no issues?**

A: The optimizer may time out due to model complexity. Daily granularity creates 300K+ variables. Consider:
- Using weekly optimization instead
- Reducing planning horizon
- Increasing solver timeout
- Simplifying constraints (e.g., remove vacuum time)

**Q: Can I use the fixed data file as-is for production?**

A: Review the changes first:
- Delivery date changes may need customer approval
- Capacity increases should be validated with operations
- Removed orders should be communicated to sales

**Q: What if my actual delivery dates are in the past (historical data)?**

A: Update `CURRENT_DATE` in `ProductionConfig`:
```python
CURRENT_DATE = datetime(2024, 10, 1)  # Historical planning date
```

**Q: How often should I run diagnostics?**

A: Run diagnostics:
- Before every optimization run (recommended)
- After updating Sales Orders
- After major Part Master changes
- Weekly as part of automated workflow

**Q: Can I customize the fix logic?**

A: Yes! Edit `fix_infeasibility.py`:
- Adjust lead time buffer (currently 20%)
- Change capacity increase factor (currently 50%)
- Modify date adjustment strategy
- Add custom validation rules

## Related Documentation

- **CLAUDE.md**: Overall project documentation
- **DECISION_SUPPORT_IMPLEMENTATION.md**: Decision support system details
- **production_plan_test.py**: Core optimization engine

## Changelog

### 2025-11-21: Initial Version
- Created diagnostic and auto-fix scripts
- Documented common infeasibility causes
- Added integration with optimization workflow

---

*Last updated: 2025-11-21*
