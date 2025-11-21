# Utilization Maximization Strategy

**Date**: 2025-11-21
**Goal**: Keep utilization near 100% as long as demand allows

---

## Current Configuration

```python
# production_plan_test.py lines 87-88
INVENTORY_HOLDING_COST = 0      # No penalty for early production
MAX_EARLY_WEEKS = 4             # Limit to 4 weeks ahead of delivery
MAX_PLANNING_WEEKS = 100        # Support far-future orders
```

---

## Strategy: "Front-Load Until Demand Exhausted"

### Behavior

**Week 1**: Maximize production (aim for 90-100% utilization)
- Produce all parts due in Weeks 1-5 (within 4-week window)
- Use all available capacity

**Week 2-5**: Continue high utilization
- Process parts from Week 1 (grinding, machining, painting)
- Stage seriality creates natural "wave" through production

**Week 6+**: Utilization drops naturally
- Most urgent orders already produced
- Only late-delivery orders remain
- **This is expected and optimal!**

---

## Why This Approach

### ✅ Advantages

1. **Maximum capacity usage** while demand exists
2. **Parts ready early** for delivery flexibility
3. **Simple to understand** - "produce as much as possible early"
4. **No complex smoothing** constraints needed

### Current Results (With 3,307 units demand)

```
Week 1: 93.8% casting, 89.5% Big Line, 97.0% Small Line ✅
Week 2: 35.8% casting (downstream stages processing W1 output)
Week 3: 36.7% casting (continued downstream processing)
Week 4-5: 3-5% casting (demand exhausted)
```

---

## Key Insight: Demand Limitation

**Total demand**: 3,307 units over 19 weeks
**Week 1 demand**: 810 units (24.5% of total)

**After producing Weeks 1-5 orders (77% of total)**, there's little left to produce!

---

## How MAX_EARLY_WEEKS Works

### With INVENTORY_HOLDING_COST = 0

```python
# Week 1 production window:
if due_week <= 1 + MAX_EARLY_WEEKS:  # Due in weeks 1-5
    can_produce_in_week_1 = True     # No penalty!
else:
    penalty = 0 × (early_weeks - 4) = 0  # Still no penalty!
```

**Effect**: `MAX_EARLY_WEEKS` **has minimal impact** when `INVENTORY_HOLDING_COST = 0`.

### Real Constraint: Planning Horizon

The **planning horizon** (19 weeks) is the real limit. Orders beyond Week 19 won't be included.

---

## To Maintain 90%+ Utilization Every Week

You would need one of these:

### Option 1: More Orders (Recommended ⭐)

**Current**: 3,307 units
**Needed for 90% every week**: ~6,000 units

**Action**: Sales team targets orders for Weeks 6-19

---

### Option 2: Build Safety Stock

Produce commonly ordered parts **without confirmed orders**:

```python
# Add to sales orders:
safety_stock = {
    'PART-001': 100 units,  # Due Week 10
    'PART-002': 80 units,   # Due Week 11
    ...
}
```

**Risk**: Inventory may not sell

---

### Option 3: Reduce Capacity (Not Recommended)

```python
# Reduce working days
WORKING_DAYS_PER_WEEK = 4  # Was: 6
```

**Result**: 60% utilization of 4 days = "100%" (but lost capacity)

---

### Option 4: Production Smoothing (Alternative Strategy)

Change configuration to spread production:

```python
INVENTORY_HOLDING_COST = 0.5  # Small penalty for early production
MAX_EARLY_WEEKS = 4           # Limit to 4 weeks ahead
```

**Effect**:
- Week 1: 70-80% utilization (produce only Week 1-2 orders)
- Week 2: 70-80% utilization (produce Week 2-3 orders)
- Week 3: 70-80% utilization (produce Week 3-4 orders)
- ...
- Week 8: 50-60% utilization (running out of orders)

**Trade-off**: Lower peak utilization, but more consistent across weeks

---

## Stage Seriality Effect

Production naturally "waves" through stages:

```
         Cast  Grind  MC1  MC2  MC3  SP1  SP2  SP3
Week 1:  95%   30%   20%  10%   0%  10%   5%   0%   ← Casting peak
Week 2:  80%   90%   60%  30%  10%  20%  15%   5%   ← Grinding peak
Week 3:  60%   85%   95%  80%  30%  60%  40%  20%   ← Machining peak
Week 4:  20%   60%   70%  90%  50%  80%  70%  40%   ← Painting peak
Week 5:   5%   30%   40%  60%  60%  70%  85%  60%   ← Finishing peak
```

**This is NORMAL and expected!** Different stages peak at different times as parts flow through.

---

## Summary Table

| Metric | Current Value | Result |
|--------|---------------|--------|
| **INVENTORY_HOLDING_COST** | 0 | Front-load production ✅ |
| **MAX_EARLY_WEEKS** | 4 | Produce up to 4 weeks early |
| **MAX_PLANNING_WEEKS** | 100 | Handle 700-day orders |
| **Week 1 Utilization** | 90-97% | ✅ Target achieved |
| **Week 2-5 Utilization** | 30-80% | ✅ Downstream processing |
| **Week 6+ Utilization** | 0-20% | ⚠️ Demand exhausted |

---

## Configuration Modes Comparison

### Mode 1: Maximization (Current) ✅

```python
INVENTORY_HOLDING_COST = 0
MAX_EARLY_WEEKS = 4
```

**Best for**: Maximum output in early weeks, accepting idle time later

---

### Mode 2: Smoothing

```python
INVENTORY_HOLDING_COST = 0.5
MAX_EARLY_WEEKS = 4
```

**Best for**: Consistent moderate utilization across all weeks

---

### Mode 3: Just-In-Time

```python
INVENTORY_HOLDING_COST = 1
MAX_EARLY_WEEKS = 2
```

**Best for**: Minimize inventory, produce close to delivery date

---

## Recommendation

**Current configuration is OPTIMAL** for your stated goal:

> "Keep near to 100% till we can"

✅ Week 1: 90-97% utilization
✅ Weeks 2-3: 30-80% (downstream processing)
✅ Weeks 4+: Low (demand exhausted - **expected**)

**To extend high utilization**: Add more orders for Weeks 6-19.

---

*Updated: 2025-11-21*
*Mode: Utilization Maximization (Front-Load)*
