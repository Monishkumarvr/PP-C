# Daily-Level Optimization Implementation Plan

## Overview
Convert weekly optimization to daily-level planning to ensure feasible daily schedules that respect daily capacity constraints.

## Key Changes

### 1. Planning Horizon
**Before:** 30 weeks (~6 months)
**After:** 180 days (~6 months)

**Calculation:**
```python
PLANNING_DAYS = (latest_order_date - CURRENT_DATE).days + BUFFER_DAYS
all_days = [CURRENT_DATE + timedelta(days=i) for i in range(PLANNING_DAYS)]
working_days = [d for d in all_days if is_working_day(d)]
```

### 2. Decision Variables
**Before:**
```python
x_casting[(variant, week)] for week in weeks
```

**After:**
```python
x_casting[(variant, day)] for day in working_days
```

### 3. Capacity Constraints
**Before (Weekly):**
```python
sum(x_casting[(v, w)] * cycle_time for v in variants) <= WEEKLY_CAPACITY
```

**After (Daily):**
```python
sum(x_casting[(v, d)] * cycle_time for v in variants) <= DAILY_CAPACITY
# where DAILY_CAPACITY = WEEKLY_CAPACITY / 6
```

### 4. Flow Constraints (Lead Times)
**Before:**
```python
casting[w] -> grinding[w+1]  # Next week
```

**After:**
```python
casting[d] -> grinding[d + cooling_days]  # Specific days later
```

### 5. Demand Constraints
**Before:**
```python
sum(delivery[(v, w)] for w in weeks) = demand[v]
delivery_week = variant_deadline_week
```

**After:**
```python
sum(delivery[(v, d)] for d in days) = demand[v]
delivery_day = variant_deadline_day (or within window)
```

## Implementation Strategy

### Phase 1: New Configuration
- Create `DailyProductionConfig` class
- Add daily capacity parameters
- Add lead time parameters (in days)

### Phase 2: Daily Demand Calculator
- Convert weekly demand to daily delivery windows
- Handle WIP at daily granularity

### Phase 3: Daily Optimization Model
- Create `DailyOptimizationModel` class
- Implement daily decision variables
- Add daily constraints

### Phase 4: Results Analyzer
- Extract daily results
- Aggregate to weekly for reporting compatibility

### Phase 5: Testing & Validation
- Compare results with weekly model
- Validate feasibility
- Performance benchmarking

## Expected Performance

**Problem Size:**
- Weekly: ~1,000 part-week variants × 30 weeks × 9 stages = ~270,000 variables
- Daily: ~1,000 part-day variants × 150 working days × 9 stages = ~1,350,000 variables

**Solve Time Estimate:**
- Weekly: 30-60 seconds
- Daily: 3-10 minutes (acceptable for overnight batch runs)

## Compatibility

- Keep weekly optimization as `production_plan_test.py`
- New daily optimization as `production_plan_daily.py`
- Both generate compatible output formats
- Executive reports work with either

## Files to Create/Modify

**New Files:**
- `production_plan_daily.py` - Main daily optimization
- `DAILY_OPTIMIZATION_PLAN.md` - This file

**Modified Files:**
- None initially (keep weekly version intact)

**Future:**
- Add flag to choose weekly vs daily in Streamlit app
- Unified reporting for both approaches
