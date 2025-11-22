# Capacity Constraint Investigation - Findings

## User Questions

1. **"When CAPACITY_BONUS = -0.1, why did Big Line and Small Line max out at only 65%? Is it molten metal tons or number of boxes?"**

2. **"Are older orders still delivered?"**

3. **"Implement daily production optimization with Indian calendar holidays"**

---

## Answer 1: Why Capacity Maxes at 65%?

### THE LIMITING CONSTRAINT: **TONNAGE** (800 tons/week)

**Week 2 Analysis** (highest utilization week):
```
Big Line:        64.9% utilization (84 / 86 hours)
Small Line:      56.2% utilization (73 / 86 hours)
Total Tonnage:   781.0 tons
Tonnage Limit:   800 tons/week
Tonnage Usage:   97.6% ← THIS IS THE BOTTLENECK!
```

**Finding:** The 800-ton weekly casting limit is hit at 97.6%, which prevents the lines from going above 65% time utilization.

**Why?**
- Big Line produced: 537.1 tons
- Small Line produced: 244.0 tons
- Total: 781 tons (97.6% of 800-ton limit)
- Lines could run more hours, but ran out of tonnage capacity!

---

## Answer 2: Box Capacity - CRITICAL BUG FOUND!

### ⚠️  BOX CAPACITY CONSTRAINTS ARE BEING VIOLATED!

**Actual box usage vs limits:**

| Box Size | Weekly Limit | Worst Week Usage | Violation |
|----------|--------------|------------------|-----------|
| **400X625** | 30 boxes | **108 boxes** | **360%** ❌ |
| **750X500** | 90 boxes | **285 boxes** | **316%** ❌ |
| **1050X750** | 60 boxes | **204 boxes** | **340%** ❌ |
| **400X500** | 30 boxes | **58 boxes** | **193%** ❌ |
| 650X750 | 180 boxes | 133 boxes | 74% ✓ |
| 750X750 | 60 boxes | 58 boxes | 97% ✓ |

**Examples:**
- 400X625: Using 108 boxes/week when limit is 30 (3.6× over!)
- 750X500: Using 285 boxes/week when limit is 90 (3.2× over!)

**ROOT CAUSE:** Box capacity constraints in the optimizer are not properly enforced. The model is ignoring these limits!

**This needs to be fixed** - box capacity violations make the schedule infeasible in practice.

---

## Answer 3: Past-Due Orders - ALL FULFILLED ✓

**Status:** ✅ All past-due orders (before 2025-11-21) are 100% fulfilled

```
Total past-due orders: 65 orders
Total ordered: 648 units
Total delivered: 648 units
Total unmet: 0 units
Fulfillment: 100.0%
```

**Sample deliveries:**
- Week -20 (UBC-062): 2/2 units delivered ✓
- Week -1 (UBC-035): 60/60 units delivered ✓
- Week 0 (ARM-004-N): 5/5 units delivered ✓

**Conclusion:** Older orders ARE being delivered correctly!

---

## Summary of Constraints

### Constraint Hierarchy (What Actually Limits Production)

1. **TONNAGE (800 tons/week)** ← **PRIMARY BOTTLENECK**
   - Week 2: 97.6% utilized
   - Prevents lines from exceeding ~65% time utilization
   - This is the constraint limiting capacity to 65%

2. **BOX CAPACITY** ← **BEING VIOLATED (BUG!)**
   - Should limit production but currently ignored
   - Violations up to 360% of capacity
   - Needs urgent fix

3. **LINE HOURS (86 hours/week per line)**
   - Not limiting (only 65% used when tonnage hits 100%)
   - Could go higher if tonnage allows

4. **DEMAND**
   - Only 3,307 units needed total
   - With current demand profile, can't utilize more capacity

---

## Recommendations

### Immediate Actions

1. **Fix box capacity constraint enforcement**
   - Current violations make schedule infeasible
   - Need to debug why box constraints aren't working

2. **Consider increasing tonnage limit** (if foundry can handle it)
   - Current: 800 tons/week
   - If increased to 1,000 tons/week → Could achieve 80%+ line utilization
   - Would require more molten metal capacity

3. **Implement daily optimization** (next task)
   - More precise scheduling
   - Better capacity utilization within weeks
   - Account for daily holidays

### Long-term Improvements

1. **Optimize box capacity allocation**
   - Some boxes severely over-subscribed (400X625 at 360%)
   - Others under-utilized (650X750 at 74%)
   - May need to adjust box size mix or redesign patterns

2. **Balance tonnage across weeks**
   - Week 2: 781 tons (98%)
   - Week 3: 547 tons (68%)
   - Could redistribute to balance better

---

## Next Steps

Per user request: **Implement daily production optimization with Indian calendar holidays**

This will:
- Schedule production day-by-day (not week-by-week)
- Respect Indian holidays (Sundays + national holidays)
- Provide more precise capacity utilization
- Enable better short-term planning

---

*Analysis date: 2025-11-22*  
*Key finding: Tonnage limit (800 tons/week) is the primary constraint at 97.6%*  
*Critical bug: Box capacity constraints being violated up to 360%*  
*Past-due orders: 100% fulfilled (648/648 units)*
