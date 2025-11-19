# Production Planning Optimization System

A manufacturing production planning optimization system for multi-stage foundry/casting operations. Uses linear programming to generate optimal production schedules while respecting capacity constraints, delivery deadlines, and work-in-progress inventory.

## Features

- **8-Stage Production Optimization**: Casting, Grinding, MC1-MC3, SP1-SP3
- **Linear Programming Engine**: PuLP/CBC solver for optimal scheduling
- **WIP-Aware Planning**: Automatically skips stages based on existing inventory
- **Capacity Constraint Management**: Machine hours, mould box limits, line capacity
- **Decision Support System**: Bottleneck analysis, risk classification, ATP calculator
- **Executive Reporting**: 10-sheet Excel dashboard with daily schedules
- **Indian Holiday Calendar**: Automatic working day calculations

## Current Performance

- 100% order fulfillment rate
- 100% on-time delivery rate

## Installation

```bash
pip install pandas numpy openpyxl pulp holidays
```

## Quick Start

### 1. Run Optimization

```bash
python production_plan_test.py
```
Creates `production_plan_COMPREHENSIVE_test.xlsx`

### 2. Generate Executive Reports

```bash
python production_plan_executive_test7sheets.py
```
Creates `production_plan_EXECUTIVE_test.xlsx` (10 sheets)

### 3. Generate Decision Support Analysis

```bash
python run_decision_support.py
```
Creates:
- `production_plan_DECISION_SUPPORT.xlsx` (11 sheets)
- `ATP_INPUT.xlsx` (template for checking new orders)

## Checking New Order Feasibility

1. Run `python run_decision_support.py` (creates template)
2. Edit `ATP_INPUT.xlsx` with potential orders:
   ```
   Part_Code | Qty | Requested_Week
   PART-001  | 100 | 5
   ```
3. Re-run `python run_decision_support.py`
4. Check `9_ATP_NEW_ORDERS` sheet for results

## Output Files

| File | Description |
|------|-------------|
| `production_plan_COMPREHENSIVE_test.xlsx` | Detailed optimizer output |
| `production_plan_EXECUTIVE_test.xlsx` | Executive dashboard (10 sheets) |
| `production_plan_DECISION_SUPPORT.xlsx` | Decision support analysis (11 sheets) |

## Decision Support Sheets

1. Executive Summary
2. Bottleneck Analysis
3. Bottleneck Summary
4. Order Risk
5. Risk by Customer
6. Risk by Week
7. Capacity Forecast
8. Capacity by Resource
9. ATP New Orders
10. Recommendations
11. Action Plan

## Input Data

Update `Master_Data_*.xlsx` with:
- Part Master specifications
- Sales Orders with delivery dates
- Current WIP inventory (by stage)
- Machine capacity constraints
- Mould box capacity

## Documentation

- [CLAUDE.md](CLAUDE.md) - Technical documentation and AI assistant guidelines
- [DECISION_SUPPORT_USER_GUIDE.md](DECISION_SUPPORT_USER_GUIDE.md) - User guide for reading Excel output
- [DECISION_SUPPORT_IMPLEMENTATION.md](DECISION_SUPPORT_IMPLEMENTATION.md) - System design document

## Project Structure

```
PP-C/
├── production_plan_test.py              # Core optimization engine
├── production_plan_executive_test7sheets.py  # Executive reporting
├── run_decision_support.py              # Decision support entry point
├── decision_support/                    # Decision support modules
│   ├── bottleneck_analyzer.py
│   ├── atp_calculator.py
│   ├── order_risk_dashboard.py
│   └── recommendations_engine.py
├── Master_Data_*.xlsx                   # Input data
└── *.xlsx                               # Output reports
```

## Manufacturing Stages

1. **Casting** - Moulding lines (Big/Small), tonnage-based
2. **Grinding** - Post-casting finishing
3. **MC1/MC2/MC3** - Machining stages
4. **SP1/SP2/SP3** - Painting stages (Primer/Intermediate/Top Coat)

## Technology Stack

- Python 3.x
- pandas, numpy - Data manipulation
- openpyxl - Excel read/write
- PuLP - Linear programming
- holidays - Indian calendar

## License

Internal use only.

---

*Last updated: 2025-11-19*
