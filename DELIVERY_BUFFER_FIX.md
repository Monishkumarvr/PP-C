# Delivery Buffer Fix - Week 2+ Capacity Utilization

## Issue Discovered

**Date**: 2025-11-22  
**Reported By**: User investigation of UBC-035 partial fulfillment

### Problem Statement

The optimizer was leaving demand unmet even when production capacity was available in Week 2 and beyond. 

**Example Case**:
- Part: UBC-035
- Order: 2526100367 (60 units, due 2025-11-01)
- Result with buffer=1:
  - Delivered: 54 units (Week 1)
  - Unmet: 6 units (90% fulfillment)
- KVCAD30MC001 utilization:
  - Week 1: 107.9% (FULL)
  - Week 2: 54.1% (room for 88+ units!)

**User's Key Question**:
> "But it should be produced in next week or next week or next to next week right?"

This was absolutely correct - the optimizer should use available future capacity.

## Root Cause Analysis

### The Constraint Triangle

Three interacting constraints prevented Week 2+ production:

1. **Delivery Window Constraint** (line 1333-1336):
   ```python
   DELIVERY_BUFFER_WEEKS = 1  # ±1 week from due date
   # For order due Week 1:
   # - Can deliver: Weeks 1-2
   # - Cannot deliver: Week 3+ (delivery_ub = 0)
   ```

2. **Lead Time Constraint** (line 1747-1752):
   ```python
   # Cumulative delivery by Week W requires casting by Week (W - lead_time)
   # For UBC-035 (lead_time = 1 week):
   # - Week 1 delivery requires Week 0 casting (impossible → uses WIP)
   # - Week 2 delivery requires Week 1 casting
   # - Week 3 delivery requires Week 2 casting
   ```

3. **Capacity Constraint**:
   - Week 1 KVCAD30MC001: 18,640 min used / 17,280 min capacity = 107.9% FULL
   - Week 2 KVCAD30MC001: 9,340 min used / 17,280 min capacity = 54.1% available

### Why 6 Units Were Unmet

```
Week 1 delivery:
  ✓ Uses WIP (46 units) - no Week 0 casting needed
  ✗ Cannot produce more (Week 1 capacity full at 107.9%)

Week 2 delivery:
  ✗ Requires Week 1 casting (capacity full)

Week 3 delivery:
  ✓ Requires Week 2 casting (capacity available!)
  ✗ BLOCKED by delivery window (buffer=1 only allows Weeks 1-2)
```

Result: 6 units left unmet, optimizer pays 60M penalty instead of using available Week 2 capacity.

## The Fix

### Code Change

**File**: `production_plan_test.py`  
**Line**: 72

```python
# BEFORE:
self.DELIVERY_BUFFER_WEEKS = 1  # Tight window (±1 week)

# AFTER:
self.DELIVERY_BUFFER_WEEKS = 2  # Allow ±2 weeks to use available capacity
```

### Impact

**Before (buffer = 1)**:
- Unmet demand: 21 units across 5 parts
- Fulfillment rate: 99.4%
- UBC-035: 54/60 units (90%)

**After (buffer = 2)**:
- Unmet demand: **0 units**
- Fulfillment rate: **100.0%**
- UBC-035: **60/60 units (100%)**
- Late orders: 9 (avg 4 days late)

### Solution Mechanics

With buffer=2, variant windows expand:

```python
# Order due Week 1:
# Before: Window = Weeks 1-2 (due-1 to due+1)
# After:  Window = Weeks 1-3 (due-2 to due+2)
```

Now the optimizer can:
1. Deliver 46 units in Week 1 (from WIP)
2. Cast 6 units in Week 1
3. Process through machining/painting in Week 2
4. **Deliver 6 units in Week 3** (2 weeks late)

Cost comparison:
- Unmet penalty: 6 × 10,000,000 = 60,000,000
- Late penalty: 6 × 150,000 × 2 = 1,800,000
- **Savings: 58.2M** (97% reduction)

## Business Logic

### Why Buffer=2 is Optimal

1. **Capacity Utilization**: Uses available Week 2+ capacity instead of leaving it idle
2. **Customer Service**: 100% fulfillment vs 99.4%
3. **Cost**: Late delivery (150K/week) << Unmet demand (10M)
4. **Realistic**: Real manufacturing can often deliver 1-2 weeks late

### Trade-offs

| Metric | Buffer=1 | Buffer=2 | Change |
|--------|----------|----------|--------|
| Fulfillment | 99.4% | 100.0% | +0.6% ✓ |
| On-time rate | ~94% | 93.1% | -0.9% |
| Late orders | 3 | 9 | +6 |
| Avg lateness | ~3 days | 4 days | +1 day |
| Unmet units | 21 | 0 | -21 ✓ |

**Conclusion**: Small increase in lateness (1 day avg) is acceptable trade-off for 100% fulfillment.

## Technical Details

### Variant Window Calculation

**File**: `production_plan_test.py`  
**Lines**: 747-770

```python
buffer = max(0, int(self.config.DELIVERY_BUFFER_WEEKS))

for _, row in self.sales_order.iterrows():
    committed_week = min(self.config.PLANNING_WEEKS,
                        int(days_diff / 7) + 1)
    
    earliest_week = max(1, committed_week - buffer)
    latest_week = committed_week + buffer
    
    variant_windows[variant] = (earliest_week, latest_week)
```

### Delivery Variable Bounds

**File**: `production_plan_test.py`  
**Lines**: 1327-1339

```python
window_start, window_end = self.variant_windows.get(
    variant, (1, self.planning_weeks_count)
)

for w in self.all_weeks:
    if window_start <= w <= window_end:
        delivery_ub = demand_up  # Can deliver
    else:
        delivery_ub = 0  # BLOCKED outside window
```

### Lead Time Constraint

**File**: `production_plan_test.py`  
**Lines**: 1747-1752

```python
L = max(self.config.MIN_LEAD_TIME_WEEKS, 
        int(self.params[part]['lead_time_weeks']))

for w in self.all_weeks:
    wL = max(0, w - L)
    # Cumulative delivery by week w requires casting by week (w-L)
    self.model += (
        pulp.lpSum(self.x_delivery[(v, t)] for t in self.all_weeks if t <= w)
        <= total_wip + 
           pulp.lpSum(self.x_casting[(v, t)] for t in self.weeks if 1 <= t <= wL)
    )
```

## Lessons Learned

### Key Insights

1. **Tight constraints can prevent optimal capacity usage**
   - Even with high unmet penalty (10M), optimizer cannot violate hard constraints
   - Delivery window constraint (delivery_ub=0) is a hard block

2. **Interaction between constraints matters**
   - Delivery window + Lead time + Capacity = Complex interaction
   - Individual constraints may look reasonable, but together they can block solutions

3. **User feedback is invaluable**
   - User correctly identified: "It should be produced in next week..."
   - This insight led to discovering the root cause

4. **Testing edge cases**
   - Partial fulfillment scenarios reveal constraint interactions
   - Week-by-week capacity analysis essential for debugging

### Debugging Process

1. ✓ Identified specific part with unmet demand (UBC-035)
2. ✓ Checked Week 2 capacity (54% - plenty available!)
3. ✓ Examined variant window (Weeks 1-2 only)
4. ✓ Traced lead time constraint (Week 3 delivery needs Week 2 casting)
5. ✓ Found delivery_ub=0 for Week 3 (blocking constraint)
6. ✓ Increased buffer to allow Week 3 delivery
7. ✓ Verified fix (100% fulfillment achieved)

## Future Considerations

### Potential Enhancements

1. **Dynamic Buffer**: Adjust buffer based on capacity utilization
   ```python
   if capacity_utilization > 90%:
       buffer = 3  # More flexibility when tight
   else:
       buffer = 2  # Standard
   ```

2. **Buffer by Customer**: VIP customers get tighter buffer (on-time priority)
   ```python
   buffer = customer_config.get('delivery_buffer', 2)
   ```

3. **Cost-Based Buffer**: Optimize buffer size based on late vs unmet costs
   ```python
   # Find minimum buffer that achieves 100% fulfillment
   for buffer in range(1, 5):
       if fulfillment >= 99.9%:
           break
   ```

### Monitoring

Track these metrics after deployment:
- Fulfillment rate by week
- Late delivery frequency and duration
- Week 2+ capacity utilization
- Customer satisfaction with late deliveries

## References

**Commit**: 7ed03a1 (2025-11-22)  
**Files Modified**: `production_plan_test.py` (line 72)  
**Issue**: Partial fulfillment despite available Week 2+ capacity  
**Resolution**: Increase DELIVERY_BUFFER_WEEKS from 1 to 2

---
*Document created: 2025-11-22*  
*Author: Claude (AI Assistant)*
