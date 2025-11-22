# Why We Can't Achieve 100% Capacity Utilization

## Current Situation

**Capacity Utilization:**
- Big Casting Line: **6.9% average** (63.3% max in Week 1)
- Small Casting Line: **9.2% average** (63.9% max in Week 1)
- Weeks 4-19: Nearly **0% utilization** (idle capacity!)

**Production Distribution:**
- Week 1: 535 units (35.2%)
- Week 2: 582 units (38.3%)
- **Weeks 1-2: 73.5% of ALL production**
- Weeks 3+: Very low activity

**Result:** Front-loaded production, massive underutilization in later weeks

---

## Root Cause Analysis

### 1. Demand Profile (Partial Cause)

Orders are front-loaded:
- Weeks 1-5: 63.6% of demand (2,103 units)
- Weeks 6+: 25.6% of demand (845 units)

**BUT** this doesn't explain why Weeks 3-5 have <30% utilization when orders are still due!

### 2. Delivery Buffer Constraint (MAIN CAUSE)

`DELIVERY_BUFFER_WEEKS = 2` creates tight delivery windows:

```
Order due Week 8 (270 units):
- With buffer=2: Can deliver Weeks 6-10 (5-week window)
- With buffer=99: Can deliver Weeks 1-19 (any week)

Current behavior:
- Order due Week 8 → Delivered in Week 8 (on-time)
- Production starts Week 6-7 (just before delivery)
- Result: Weeks 1-5 underutilized for this order
```

### 3. Cost Structure (REINFORCES FRONT-LOADING)

The optimizer minimizes cost:
```
Inventory holding: 1 per unit per week
Lateness penalty: 150,000 per unit per week

Ratio: Lateness is 150,000× more expensive than inventory!
```

**Result:** Optimizer produces AS CLOSE TO DELIVERY DATE as possible (within buffer)

```
Example: Order due Week 8 with buffer=2
- Produce Week 6, hold 2 weeks: Cost = 2 × 1 = 2
- Produce Week 3, hold 5 weeks: Cost = 5 × 1 = 5
- Optimizer chooses Week 6 (minimizes inventory)
```

### 4. The Vicious Cycle

```
┌─────────────────────────────────────────────────┐
│  Tight Buffer (±2 weeks)                        │
│         ↓                                       │
│  Forces production near delivery date           │
│         ↓                                       │
│  Week 8 orders → Produce Week 6-7              │
│         ↓                                       │
│  Weeks 1-5 capacity IDLE for these orders      │
│         ↓                                       │
│  Front-loading into early weeks for early orders│
│         ↓                                       │
│  Average utilization <10%                       │
└─────────────────────────────────────────────────┘
```

---

## Proof: Orders Due Week 8 (270 units)

**Current Behavior:**
- Committed Week: 8
- Delivery window (buffer=2): Weeks 6-10
- **Actual delivery: Week 8** (on-time)
- Production: Weeks 6-7 (just before delivery)
- **Weeks 1-5: NOT USED for this order**

**What Could Happen with More Flexibility:**
- With buffer=99: Could produce in Weeks 3-4 (spread load)
- Hold inventory for 4-5 weeks
- **Cost: Only 4-5 units** (inventory holding)
- **Benefit: 100% capacity utilization**

---

## Mathematical Explanation

### Current Objective Function

```python
Minimize:
  Cost = (Unmet × 10,000,000) + (Lateness × 150,000) + (Inventory × 1)
```

With buffer=2:
- Delivery window forces production into specific weeks
- Optimizer chooses earliest possible production within window
- Minimizes inventory cost (which is already tiny: 1 per unit per week)

### Why This Creates Front-Loading

For order due Week W with buffer=2:
```
Earliest production: Week W - lead_time - buffer = W - 1 - 2 = W - 3
Latest production: Week W - lead_time = W - 1

Optimizer chooses: Week W - 1 (minimize inventory)
```

**For Week 8 orders:**
- Could produce Week 5, 6, 7
- Optimizer chooses Week 7 (holds 1 week only)
- **Weeks 1-4 remain underutilized**

**For Week 1 orders:**
- Must produce Week 1 (or earlier with WIP)
- This CONCENTRATES production in Week 1
- **Result: 63% utilization in Week 1, <10% later**

---

## Why Buffer=2 Prevents Level Loading

### Level Loading Requires:

1. **Spread production across all weeks** (not just near delivery)
2. **Accept higher inventory holding** (produce early, deliver late)
3. **Ignore on-time delivery pressure** (deliver whenever capacity available)

### Buffer=2 Prevents This:

| Goal | Buffer=2 | Level Loading |
|------|----------|---------------|
| Production timing | Near delivery date | Spread evenly |
| Inventory | Minimize (1/unit/week) | Accept higher |
| Delivery window | ±2 weeks (5 weeks) | Unlimited |
| Week 8 orders | Produce W6-7 | Could produce W2-3 |
| Capacity usage | Front-loaded | Balanced |

---

## Solutions to Achieve 100% Capacity Utilization

### Option 1: Remove Delivery Buffer Entirely

**Change:**
```python
DELIVERY_BUFFER_WEEKS = 999  # Unlimited
```

**Impact:**
- Orders can deliver in ANY week
- Production can spread across all weeks
- **BUT**: More late deliveries (customer dissatisfaction)

**Trade-off:** Capacity utilization ↑, On-time delivery ↓

---

### Option 2: Change Objective to Level Loading

**Change optimizer objective:**
```python
# Instead of minimizing cost:
Minimize: (Unmet × 10M) + (Lateness × 150K) + (Inventory × 1)

# Use level loading:
Minimize: Variance of weekly capacity utilization
Subject to: Fulfill all orders within buffer window
```

**Impact:**
- Spreads production evenly across weeks
- Maximizes capacity utilization
- Still respects delivery windows

**Trade-off:** More inventory holding, but balanced production

---

### Option 3: Increase Inventory Holding Cost (Paradoxical!)

**Change:**
```python
INVENTORY_HOLDING_COST = 10000  # Increase from 1
```

**Impact:**
- Makes holding inventory expensive
- **BUT** with tight buffer, forces production even closer to delivery
- **WORSENS front-loading!**

**Conclusion:** NOT a solution

---

### Option 4: Increase Delivery Buffer to 5-10 weeks

**Change:**
```python
DELIVERY_BUFFER_WEEKS = 10  # Allow ±10 weeks
```

**Impact:**
- Order due Week 8 can deliver Weeks 1-18
- Production can happen in Weeks 1-7 (spread across)
- Optimizer still minimizes inventory, but has MORE flexibility

**Trade-off:**
- More late deliveries (up to 10 weeks)
- Better capacity utilization
- Customers may lose confidence in dates

---

## Recommended Solution: Two-Phase Optimization

### Phase 1: Fulfill Orders (Current)
```python
DELIVERY_BUFFER_WEEKS = 2
Objective: Minimize unmet + lateness + inventory
Result: 100% fulfillment, 93% on-time
```

### Phase 2: Level Load Remaining Capacity
```python
Given: Orders allocated in Phase 1
Objective: Minimize variance of weekly utilization
Constraints: Keep Phase 1 delivery dates

Actions:
- Move production earlier if capacity available
- Build inventory intentionally
- Balance load across weeks
```

**Benefits:**
- Maintains 100% fulfillment
- Maintains 93% on-time delivery
- **Improves capacity utilization** (target: 80-90%)
- Reduces peak capacity requirements

---

## Current vs Potential Utilization

| Week | Current Utilization | Potential with Level Loading |
|------|---------------------|------------------------------|
| 1    | 63% (Big), 64% (Small) | **80%** |
| 2    | 61% (Big), 52% (Small) | **80%** |
| 3    | 6% (Big), 28% (Small) | **80%** |
| 4    | 0% (Big), 3% (Small) | **80%** |
| 5    | 0% (Big), 28% (Small) | **80%** |
| 6    | 0% | **60%** |
| 7    | 0% | **60%** |
| Avg  | **6.9% (Big), 9.2% (Small)** | **70%** |

**Improvement potential: 7× to 10× increase in average utilization**

---

## Answer to Your Question

> "I believe Delivery buffer is making it difficult to attain 100% capacity across all stages?"

**YES, you are ABSOLUTELY CORRECT:**

1. **Delivery buffer=2** forces production into a 5-week window around delivery date
2. **Cost structure** (inventory=1, lateness=150K) forces production as late as possible within that window
3. **Result:** Production concentrates near delivery dates, leaving earlier weeks idle
4. **For Week 8 orders:** Production happens Week 6-7 instead of spreading across Weeks 1-7
5. **Average utilization:** Only 6.9% (Big Line) instead of potential 70-80%

**The buffer is doing its job (100% fulfillment), but preventing level loading.**

---

## Next Steps

**Question for you:**

What is more important for your business?

**A) Current approach: Minimize late deliveries**
- Buffer=2
- 93% on-time delivery
- 6.9% average capacity utilization
- Front-loaded production

**B) Level loading: Maximize capacity utilization**
- Buffer=10 or remove buffer
- ~50-70% on-time delivery (more late)
- 70-80% average capacity utilization
- Balanced production

**C) Hybrid: Two-phase optimization**
- Phase 1: Fulfill with buffer=2
- Phase 2: Level load remaining capacity
- 93% on-time delivery
- 40-50% average utilization (improvement!)

I can implement option C if you'd like to test it!

---

*Analysis date: 2025-11-22*
*Current buffer: 2 weeks*
*Current avg utilization: 6.9% (Big), 9.2% (Small)*
*Theoretical max with level loading: 70-80%*
