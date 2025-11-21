# WIP Sensitivity Analysis - Impact on Capacity Utilization

## Objective
Test if reducing WIP inventory can force the optimizer to increase production capacity utilization to meet demand.

## Scenarios Tested
1. **10% WIP**: Reduced all WIP to 10% of original (238 units total)
2. **30% WIP**: Reduced all WIP to 30% of original (715 units total)
3. **50% WIP**: Reduced all WIP to 50% of original (1,192 units total)
4. **100% WIP**: Original WIP levels (2,383 units total)

## Key Findings

### Casting Utilization: **NO CHANGE**
- **Result**: 50% utilization across ALL WIP levels (10%, 30%, 50%, 100%)
- **Conclusion**: WIP level does NOT affect casting utilization
- **Root Cause**: Casting is limited by other factors (demand, part mix, flow constraints)

### Grinding Utilization: **INCREASES WITH WIP**
| WIP Level | Average Utilization | Change from 10% |
|-----------|---------------------|-----------------|
| 10%       | 67.4%               | Baseline        |
| 30%       | 69.4%               | +2.0%           |
| 50%       | 71.8%               | +4.4%           |
| 100%      | 77.0%               | +9.6%           |

- **Trend**: Higher WIP → Higher grinding utilization
- **Peak**: All scenarios hit 100% grinding in some weeks
- **Why**: More WIP (especially CS WIP) provides more raw material for grinding

### Downstream Stages: **SLIGHT INCREASE WITH WIP**
All downstream stages show minor improvements with higher WIP:
- MC1: 15.0% → 16.9% (+1.9%)
- MC2: 10.2% → 11.7% (+1.5%)
- MC3: 1.9% → 2.7% (+0.8%)
- SP1/SP2: 11.8% → 14.1% (+2.3%)
- SP3: 5.8% → 7.8% (+2.0%)

## Critical Insights

### 1. **Casting Bottleneck is NOT WIP-Related**
The fact that casting stays at exactly 50% regardless of WIP level proves that:
- WIP is NOT limiting casting capacity
- The bottleneck is elsewhere:
  - **Part mix limitation** (only 28 distinct parts being produced)
  - **Demand constraint** (already producing 2.08x demand)
  - **Production variable bounds** (10x demand cap per variant)

### 2. **WIP Helps Downstream, Not Upstream**
- **CS WIP** (casting inventory) helps grinding utilize more capacity
- **GR/MC/SP WIP** helps downstream stages
- But WIP doesn't help casting itself - it's the source, not the consumer

### 3. **The 50% Barrier**
Casting is stuck at 50% because:
1. **Capacity calculation bug**: Displays 50% but actual is 75% (see previous analysis)
2. **Real constraint**: Even at true 75%, limited by:
   - Available part mix (28 parts vs 203 in master)
   - Demand already exceeded (2.08x overproduction)
   - Diminishing returns on further overproduction

## Recommendations

### To Increase Casting Utilization:
1. **Fix capacity display bug** (priority: immediate)
2. **Increase part variety** in production mix (more customer orders)
3. **Relax production bounds** beyond 10x demand (questionable business value)
4. **Add more orders** to the system (sales-driven solution)

### WIP Strategy:
- **Higher WIP IS beneficial** for overall plant utilization
- **100% WIP maintains 77% grinding** vs 67% at 10% WIP
- **But WIP doesn't solve casting bottleneck** - that requires addressing root causes above

## Conclusion
WIP reduction experiment proves that **WIP is NOT the limiting factor for casting utilization**. The true limiters are:
1. Display bug (50% vs actual 75%)
2. Limited part mix (28 active parts)
3. Demand constraint (already 2x overproduction)
4. Business model constraints (how much speculative inventory makes sense?)

**Action**: Fix capacity display bug to show true utilization, then evaluate if 75% is acceptable given current order book.
