# Decision Support System - User Guide

This guide explains how to read and use the `production_plan_DECISION_SUPPORT.xlsx` output file.

## Quick Start

```bash
# 1. Run optimizer first
python production_plan_test.py

# 2. Generate decision support analysis
python run_decision_support.py
```

This creates:
- `production_plan_DECISION_SUPPORT.xlsx` (11 analysis sheets)
- `ATP_INPUT.xlsx` (template for checking new orders)

---

## Sheet-by-Sheet Guide

### 1. Executive Summary

**Purpose**: Quick overview of system health and key metrics.

| Metric | What It Means |
|--------|---------------|
| Total Orders | Number of sales orders in the system |
| Fulfillment Rate | % of total quantity that can be delivered |
| On-Time Rate | % of orders delivered by due date |
| Critical Bottlenecks | Resources at 100%+ utilization |
| High-Risk Orders | Orders needing immediate attention |

**Action**: If Fulfillment Rate < 100% or Critical Bottlenecks > 0, check sheets 2-4.

---

### 2. Bottleneck Analysis

**Purpose**: Shows which resources are overloaded and blocking production.

| Column | Description |
|--------|-------------|
| Week | Planning week (W1, W2, etc.) |
| Resource | Machine/line name |
| Utilization_% | Current load (>100% = overloaded) |
| Overflow | Amount exceeding capacity |
| Severity | Critical (≥100%), High (≥95%), Medium (≥85%) |

**How to Read**:
- **Critical (red)**: Resource is overloaded - production will be delayed
- **High (orange)**: Near capacity - any new orders will cause delays
- **Medium (yellow)**: Approaching limits - monitor closely

**Action**: For Critical/High resources, consider:
- Add overtime shifts
- Outsource to subcontractors
- Reschedule lower-priority orders

---

### 3. Bottleneck Summary

**Purpose**: Aggregated view of overflow by resource across all weeks.

| Column | Description |
|--------|-------------|
| Resource | Machine/line name |
| Total_Overflow | Total excess demand across all weeks |
| Weeks_Constrained | Number of weeks with capacity issues |
| Max_Overflow | Worst week overflow amount |

**Action**: Focus improvement efforts on resources with highest Total_Overflow.

---

### 4. Order Risk

**Purpose**: Every order classified by risk level.

| Column | Description |
|--------|-------------|
| Order_ID | Sales order number |
| Part_Code | Material/part code |
| Customer | Customer name |
| Due_Week | Delivery due week |
| Ordered_Qty | Quantity ordered |
| Planned_Qty | Quantity that can be produced |
| Fulfillment_% | Planned/Ordered × 100 |
| Risk_Level | Critical, High, Medium, Low |

**Risk Levels**:
- **Critical**: Past due with unmet quantity OR 0% fulfillment
- **High**: Due within 1 week with unmet quantity OR <50% fulfillment
- **Medium**: Due within 2 weeks with unmet quantity OR <100% fulfillment
- **Low**: Fully fulfilled, no issues

**Action**:
- Critical orders → Immediate escalation to management
- High orders → Contact customer about potential delays
- Medium orders → Monitor and prepare contingency

---

### 5. Risk by Customer

**Purpose**: Customer-level risk summary for account management.

| Column | Description |
|--------|-------------|
| Customer | Customer name |
| Total_Orders | Number of orders |
| Critical_Orders | Orders at critical risk |
| High_Orders | Orders at high risk |
| Avg_Fulfillment_% | Average fulfillment across orders |

**Action**: Prioritize communication with customers having Critical_Orders > 0.

---

### 6. Risk by Week

**Purpose**: Weekly risk distribution for capacity planning.

| Column | Description |
|--------|-------------|
| Week | Delivery week |
| Total_Orders | Orders due that week |
| Critical | Critical risk orders |
| High | High risk orders |
| Total_Risk_Orders | Critical + High |

**Action**: Weeks with high Total_Risk_Orders need overtime or order rescheduling.

---

### 7. Capacity Forecast

**Purpose**: Weekly capacity outlook across all resources.

| Column | Description |
|--------|-------------|
| Week | Planning week |
| Avg_Utilization_% | Average utilization across resources |
| Constrained_Resources | Number of resources at ≥85% |
| Status | Critical, Tight, OK |

**Status Meanings**:
- **Critical**: 3+ resources constrained
- **Tight**: 1-2 resources constrained
- **OK**: All resources have capacity

**Action**: Avoid adding new orders in Critical weeks.

---

### 8. Capacity by Resource

**Purpose**: Available capacity matrix (Resource × Week).

**Format**: Each cell shows available capacity with status indicator.
- `150` = 150 units available
- `50 (TIGHT)` = Only 50 units, approaching limit
- `0 (FULL)` = No capacity available

**How to Use**:
1. Find your part's required resources (casting line, machining, painting)
2. Check if capacity exists in the requested delivery week
3. If not, look for the first week with available capacity

---

### 9. ATP New Orders (Available-to-Promise)

**Purpose**: Feasibility check for potential new orders.

| Column | Description |
|--------|-------------|
| Part_Code | Part/material code |
| Requested_Qty | Quantity requested |
| Requested_Week | Desired delivery week |
| Feasible | Yes/No |
| Earliest_Delivery | First week order can be delivered |
| Delay_Weeks | Gap from requested date |
| Limiting_Resource | Resource blocking the order |
| Confidence | High/Medium/Low reliability |
| Notes | Additional information |

**How to Quote Customers**:
- **Feasible = Yes**: Quote the requested delivery date
- **Feasible = No**: Quote the `Earliest_Delivery` week or decline the order

---

### 10. Recommendations

**Purpose**: Prioritized list of actionable suggestions.

| Column | Description |
|--------|-------------|
| Priority | 1 = Most urgent |
| Category | Bottleneck, Risk, Optimization |
| Recommendation | Suggested action |
| Impact | Expected benefit |
| Effort | Implementation difficulty |

**Action**: Start with Priority 1 recommendations and work down the list.

---

### 11. Action Plan

**Purpose**: Immediate actions for top recommendations.

| Column | Description |
|--------|-------------|
| Recommendation | Parent recommendation |
| Action | Specific task to complete |
| Owner | Suggested responsible role |
| Deadline | Suggested completion date |

**Action**: Assign owners and track completion of each action item.

---

## ATP Workflow: Checking New Order Feasibility

### When a Customer Requests a New Order

**Step 1: Create/Edit Input File**

First run creates `ATP_INPUT.xlsx` automatically. Edit it with the potential order:

| Part_Code | Qty | Requested_Week |
|-----------|-----|----------------|
| PART-001  | 100 | 5 |
| PART-002  | 200 | 6 |

- **Part_Code**: Must match a code from Part Master
- **Qty**: Requested quantity (integer)
- **Requested_Week**: Desired delivery week (1, 2, 3... or W1, W2, W3...)

**Step 2: Run Analysis**

```bash
python run_decision_support.py
```

**Step 3: Check Results**

Open `production_plan_DECISION_SUPPORT.xlsx`, go to sheet `9_ATP_NEW_ORDERS`:

```
Part_Code | Qty | Requested_Week | Feasible | Earliest_Delivery | Limiting_Resource
PART-001  | 100 | W5             | No       | W7                | Big_Line
PART-002  | 200 | W6             | Yes      | W6                | None
```

**Step 4: Respond to Customer**

| Result | Customer Response |
|--------|-------------------|
| Feasible = Yes | "We can deliver by [Requested_Week]" |
| Feasible = No | "Earliest delivery is [Earliest_Delivery]" OR "We cannot accept this order" |

### Understanding ATP Results

**Feasible = Yes**
- All resources have capacity in the requested week
- Quote the requested delivery date confidently

**Feasible = No, Small Delay (1-2 weeks)**
- Capacity exists but not in requested week
- Offer the `Earliest_Delivery` date to customer
- Consider if order is worth expediting (overtime)

**Feasible = No, Large Delay (3+ weeks)**
- Significant capacity constraints
- May need to decline or negotiate smaller quantity
- Check `Limiting_Resource` to understand bottleneck

### Confidence Levels

| Confidence | Meaning |
|------------|---------|
| High | Order can definitely be fulfilled as indicated |
| Medium | Some uncertainty, monitor closely |
| Low | Part not found or major constraints exist |

---

## Common Scenarios

### Scenario 1: Customer Wants 500 Units by Week 5

1. Add to `ATP_INPUT.xlsx`: `PART-XYZ, 500, 5`
2. Run `python run_decision_support.py`
3. Check result:
   - Feasible = Yes → Confirm Week 5
   - Feasible = No, Earliest = W6 → Offer Week 6
   - Feasible = No, Earliest = W10 → Consider declining

### Scenario 2: Sales Team Wants to Know Available Capacity

1. Open sheet `8_CAPACITY_BY_RESOURCE`
2. Find the resources for the customer's part type
3. Look for weeks with available capacity (no TIGHT or FULL labels)
4. Proactively offer those weeks to customers

### Scenario 3: Production Manager Needs to Prioritize

1. Open sheet `4_ORDER_RISK`
2. Filter by `Risk_Level = Critical`
3. These orders need immediate attention
4. Check sheet `10_RECOMMENDATIONS` for suggested actions

### Scenario 4: Finding Why an Order Can't Be Fulfilled

1. Check sheet `9_ATP_NEW_ORDERS` for the `Limiting_Resource`
2. Go to sheet `2_BOTTLENECK_ANALYSIS`
3. Find that resource and see the severity
4. Check sheet `10_RECOMMENDATIONS` for how to resolve

---

## Tips for Best Results

1. **Update Input Data Regularly**: Re-run optimization when orders or WIP change
2. **Check ATP Before Committing**: Always verify capacity before promising delivery
3. **Monitor Bottlenecks Weekly**: Address constraints before they become critical
4. **Use Recommendations**: The system suggests specific actions based on analysis
5. **Track Risk Trends**: If risk increases week-over-week, investigate root cause

---

## Troubleshooting

### "Part not found in Part_Parameters"
- Ensure Part_Code matches exactly (case-sensitive)
- Check if part exists in Part Master sheet
- Verify the comprehensive output was generated with current data

### All orders show as infeasible
- Check if comprehensive output exists and is current
- Verify capacity data in Weekly_Summary sheet
- Look for extremely high utilization across all resources

### ATP shows wrong earliest delivery
- The calculation is based on available capacity slots
- If capacity is overestimated, results will be optimistic
- Cross-check with actual shop floor capacity

---

*Generated for Decision Support System v1.0*
