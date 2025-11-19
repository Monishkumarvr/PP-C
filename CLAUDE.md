# CLAUDE.md - Production Planning Optimization System

## Project Overview

This repository contains a **Manufacturing Production Planning Optimization System** designed for multi-stage manufacturing environments (casting/foundry operations). The system uses linear programming to optimize production schedules across 8 manufacturing stages while respecting capacity constraints, delivery deadlines, and work-in-progress (WIP) inventory.

**Primary Use Case**: Enterprise production planning for foundry/casting manufacturing facilities in India.

**Current Performance**: 100% fulfillment, 100% on-time delivery rate.

## Repository Structure

```
PP-C/
├── production_plan_test.py                    # Core optimization engine (~3,100 lines)
├── production_plan_executive_test7sheets.py   # Executive reporting module (~3,600 lines)
├── run_decision_support.py                    # Decision Support System entry point
├── decision_support/                          # Decision Support System modules
│   ├── __init__.py
│   ├── bottleneck_analyzer.py                 # Identify capacity constraints
│   ├── atp_calculator.py                      # Available-to-Promise calculator
│   ├── order_risk_dashboard.py                # Risk classification by order
│   └── recommendations_engine.py              # Actionable suggestions generator
├── Master_Data_Updated_Nov_Dec.xlsx           # Input: Part master, sales orders, constraints
├── Master_Data_Optimised_Updated__2_.xlsx     # Input: Alternative/updated master data
├── production_plan_COMPREHENSIVE_test.xlsx    # Output: Detailed optimization results
├── production_plan_EXECUTIVE_test.xlsx        # Output: Executive dashboard reports (10 sheets)
├── production_plan_DECISION_SUPPORT.xlsx      # Output: Decision support analysis (11 sheets)
├── ATP_INPUT.xlsx                             # Input: User-defined orders for ATP check
├── DECISION_SUPPORT_IMPLEMENTATION.md         # Decision support system design doc
└── CLAUDE.md                                  # This file
```

## Core Modules

### 1. `production_plan_test.py` - Comprehensive Optimization Engine

**Purpose**: Core linear programming optimization engine that generates production schedules.

**Key Classes**:

- `ProductionConfig` - Central configuration (dates, penalties, capacities, OEE)
- `ProductionCalendar` - Indian holiday calendar integration
- `ComprehensiveDataLoader` - Excel data ingestion and validation
- `WIPDemandCalculator` - Net demand calculation with stage-wise WIP skip logic
- `ComprehensiveParameterBuilder` - Part parameter extraction with routing awareness
- `MachineResourceManager` - Machine capacity management
- `BoxCapacityManager` - Mould box capacity constraints
- `ComprehensiveOptimizationModel` - PuLP linear programming model
- `DailyScheduleGenerator` - Weekly-to-daily schedule distribution

**Manufacturing Stages Modeled**:
1. Casting (tonnage-based, moulding lines)
2. Grinding
3. MC1 - Machining Stage 1
4. MC2 - Machining Stage 2
5. MC3 - Machining Stage 3
6. SP1 - Painting Stage 1 (Primer)
7. SP2 - Painting Stage 2 (Intermediate)
8. SP3 - Painting Stage 3 (Top Coat)

### 2. `production_plan_executive_test7sheets.py` - Executive Reporting Module

**Purpose**: Generates formatted Excel reports with utilization dashboards and daily schedules.

**Key Classes**:

- `ProductionCalendar` - Working day calculations
- `MasterDataEnricher` - Enriches schedules with master data (cycle times, machines)
- `DailyProductionInventoryTracker` - Matrix format daily production and inventory snapshots
- `FixedExecutiveReportGenerator` - Multi-sheet Excel report generation

**Output Sheets Generated** (10 sheets):
1. Executive Dashboard - KPIs and stage summaries
2. Master Schedule - Weekly production plan
3. Delivery Tracker - Order fulfillment status
4. Bottleneck Alerts - Capacity constraints
5. Capacity Overview - Utilization by stage
6. Material Flow - WIP movement tracking
7. Daily Schedule - Aggregate daily totals with calendar dates
8. Part-Level Daily Schedule - Detailed machine assignments
9. **Daily Production** - Matrix format (Date × Part-Stage columns)
10. **Daily Inventory** - Matrix format (Date × Part-WIP columns)

### 3. `run_decision_support.py` - Decision Support System

**Purpose**: Analyzes optimizer results to provide actionable insights for production planning decisions.

**Key Classes** (in `decision_support/` module):

- `BottleneckAnalyzer` - Identifies capacity constraints (>85% utilization)
- `ATPCalculator` - Calculates Available-to-Promise for new orders
- `OrderRiskAnalyzer` - Classifies orders by risk level (Critical/High/Medium/Low)
- `RecommendationsEngine` - Generates actionable suggestions based on analysis

**Output Sheets Generated** (11 sheets):
1. **Executive Summary** - Key metrics and KPIs at a glance
2. **Bottleneck Analysis** - Resources operating above 85% utilization
3. **Bottleneck Summary** - Aggregated overflow by resource
4. **Order Risk** - All orders with risk classification
5. **Risk by Customer** - Customer-level risk summary
6. **Risk by Week** - Weekly risk distribution
7. **Capacity Forecast** - Weekly capacity outlook
8. **Capacity by Resource** - Available capacity matrix (Resource × Week)
9. **ATP New Orders** - Feasibility check for potential new orders
10. **Recommendations** - Prioritized actionable suggestions
11. **Action Plan** - Immediate actions (top 2 per recommendation)

**Risk Classification Criteria**:
- **Critical**: Past due with unmet qty, or 0% fulfillment
- **High**: Due within 1 week with unmet qty, or <50% fulfillment
- **Medium**: Due within 2 weeks with unmet qty, or <100% fulfillment
- **Low**: Fully fulfilled or no risk factors

**Bottleneck Severity Thresholds**:
- **Critical**: ≥100% utilization
- **High**: ≥95% utilization
- **Medium**: ≥85% utilization

### Checking New Order Feasibility (ATP)

The ATP feature allows checking if potential new orders can be fulfilled:

**Workflow**:
1. Run `python run_decision_support.py` (creates `ATP_INPUT.xlsx` template)
2. Edit `ATP_INPUT.xlsx` with potential orders:
   - `Part_Code`: Part/material code from Part Master
   - `Qty`: Requested quantity
   - `Requested_Week`: Desired delivery week
3. Re-run `python run_decision_support.py`
4. Check sheet `9_ATP_NEW_ORDERS` for results:
   - `Feasible`: Yes/No
   - `Earliest_Delivery`: When order can be delivered
   - `Delay_Weeks`: Gap from requested date
   - `Limiting_Resource`: Which resource blocks the order
   - `Confidence`: High/Medium/Low

**Example ATP Result**:
```
Part_Code | Qty | Requested_Week | Feasible | Earliest_Delivery | Limiting_Resource
PART-001  | 100 | W5             | No       | W7                | Big_Line
```

## Technology Stack

### Dependencies (Required)

```python
# Core
pandas                 # Data manipulation and analysis
numpy                  # Numerical operations
openpyxl               # Excel file read/write

# Optimization
pulp                   # Linear Programming (PuLP with CBC solver)

# Utilities
holidays               # Indian holiday calendar (holidays.India)
datetime               # Date/time operations
collections            # defaultdict for data structures
```

### Installation

```bash
pip install pandas numpy openpyxl pulp holidays
```

## Data Flow

```
┌─────────────────┐
│  Master Data    │ (Part Master, Sales Orders, Constraints, WIP, Box Capacity)
│  Excel File     │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Data Loader    │ Validates orders against Part Master
│  & Validation   │ Calculates dynamic planning horizon
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  WIP Calculator │ Stage-wise WIP skip logic
│  & Demand Split │ Net demand by part-week variant
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  LP Optimizer   │ PuLP/CBC solver
│  (PuLP Model)   │ Minimizes: penalties + lateness + inventory
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Output Gen     │ Comprehensive Excel + Executive Reports
│  Excel Reports  │ Daily schedules with machine assignments
└─────────────────┘
```

## Key Algorithms & Concepts

### 1. Stage Seriality Constraints
- MC1 → MC2 → MC3 (machining must complete before next stage)
- SP1 → SP2 → SP3 (painting stages sequential)
- Casting → Grinding → Machining → Painting → Delivery

### 2. WIP Skip Logic
Parts with existing WIP inventory skip earlier stages:
- FG (Finished Goods) → Skip all production
- SP (Painting WIP) → Skip Casting, Grinding, Machining
- MC (Machining WIP) → Skip Casting, Grinding
- GR (Grinding WIP) → Skip Casting
- CS (Casting WIP) → Skip Casting

### 3. Capacity Constraints
- **Mould box capacity**: Units per box size per week
- **Machine hours**: Resource-specific weekly hours with OEE
- **Vacuum line capacity**: Big/Small line hours with penalty factor

### 4. Optimization Objective
Minimize total cost:
- `UNMET_DEMAND_PENALTY` (200,000) - High cost for unfulfilled orders
- `LATENESS_PENALTY` (150,000) - Cost per week late (increased to prioritize on-time delivery)
- `INVENTORY_HOLDING_COST` (1) - Per unit per week
- `SETUP_PENALTY` (5) - Pattern changeover cost

### 5. Part-Specific Timing
- Cooling time (hours) between casting and grinding
- Shakeout time (hours) for part separation
- Vacuum time for vacuum-cast parts
- Dry times between painting stages

### 6. Passive vs Active Time
**Critical distinction for capacity planning:**

| Time Type | Examples | Behavior |
|-----------|----------|----------|
| **Active** | Cycle time (casting, grinding, machining, painting) | Consumes machine capacity |
| **Passive** | Cooling, shakeout, drying between paint coats | Affects lead time only, not capacity |

Machines can work on the next batch while the previous batch cools/dries.

### 7. WIP Handling
WIP is treated as **initial inventory** that flows through the production pipeline:

| WIP Stage | Entry Point | Skips |
|-----------|-------------|-------|
| **FG** | Directly fulfills orders | All production |
| **SP** | Directly fulfills orders | Casting, Grinding, Machining |
| **MC** | Enters at painting | Casting, Grinding, Machining |
| **GR** | Enters at machining | Casting, Grinding |
| **CS** | Enters at grinding | Casting |

**WIP Allocation**: FG + SP WIP is allocated proportionally across all orders for the same part.

### 8. Just-in-Time Production
The optimizer produces **close to delivery date** to minimize inventory:
- Orders for the same part with different due dates are produced in separate batches
- Production is split across weeks to match order schedules
- WIP is consumed in earliest-due-date-first order

## Excel Input Format

### Required Sheets in Master Data Excel:

1. **Part Master** - Part specifications
   - `FG Code`, `CS Code` - Part identifiers
   - `Standard unit wt.`, `Bunch Wt.` - Weights
   - `Box Size`, `Box Quantity` - Packaging
   - `Moulding Line` - Casting line assignment
   - `Casting Cycle time (min)`, `Grinding Cycle time (min)` - Process times
   - `Machining resource code 1/2/3`, `Machining Cycle time 1/2/3 (min)`
   - `Painting Resource code 1/2/3`, `Painting Cycle time 1/2/3 (min)`
   - `Vacuum Time (hrs)`, `Cooling Time (hrs)`, `Shakeout Time (hrs)`

2. **Sales Order** - Customer orders
   - `Material Code` - FG Code
   - `Balance Qty` - Quantity to produce
   - `Comitted Delivery Date` - Due date

3. **Machine Constraints** - Resource capacities
   - `Resource Code`, `Resource Name`, `Operation Name`
   - `No Of Resource`, `Available Hours per Day`, `No of Shift`

4. **Stage WIP** - Work in progress inventory
   - `CastingItem` or `Material Code`
   - `FG`, `SP`, `MC`, `GR`, `CS` - Quantities at each stage

5. **Mould Box Capacity** - Casting constraints
   - `Box_Size`, `Weekly_Capacity`

## Configuration Parameters

Key parameters in `ProductionConfig` class:

```python
CURRENT_DATE = datetime(2025, 10, 1)      # Planning start date
MAX_PLANNING_WEEKS = 30                    # Maximum horizon
PLANNING_BUFFER_WEEKS = 2                  # Buffer beyond latest order
OEE = 0.90                                 # Overall Equipment Effectiveness
WORKING_DAYS_PER_WEEK = 6                  # Mon-Sat
WEEKLY_OFF_DAY = 6                         # Sunday (0=Mon, 6=Sun)
PATTERN_CHANGE_TIME_MIN = 18               # Mould changeover time
UNMET_DEMAND_PENALTY = 200000              # Penalty for unfulfilled demand
LATENESS_PENALTY = 150000                  # Per week late (prioritize on-time)
INVENTORY_HOLDING_COST = 1                 # Per unit per week
```

## Usage

### Running the Optimization

```bash
# Run comprehensive optimization
python production_plan_test.py

# Generate executive reports (after running optimization)
python production_plan_executive_test7sheets.py

# Generate decision support analysis (after running optimization)
python run_decision_support.py
```

### Typical Workflow

1. Update `Master_Data_*.xlsx` with current:
   - Part Master specifications
   - Sales Orders with delivery dates
   - Current WIP inventory
   - Machine capacity constraints

2. Run `production_plan_test.py` to generate optimized schedule
   - Creates `production_plan_COMPREHENSIVE_test.xlsx`

3. Run `production_plan_executive_test7sheets.py` for reports
   - Creates `production_plan_EXECUTIVE_test.xlsx`

4. Run `run_decision_support.py` for decision support analysis
   - Creates `production_plan_DECISION_SUPPORT.xlsx`
   - Creates `ATP_INPUT.xlsx` (template for checking new orders)

### Checking New Orders

To check if new orders can be accepted:

1. Run `python run_decision_support.py` once (creates template)
2. Edit `ATP_INPUT.xlsx` with potential orders
3. Re-run `python run_decision_support.py`
4. Open `production_plan_DECISION_SUPPORT.xlsx`, sheet `9_ATP_NEW_ORDERS`
5. Check results:
   - **Feasible = Yes**: Quote the requested delivery date
   - **Feasible = No**: Quote the `Earliest_Delivery` date or decline

## Code Conventions

### Naming Conventions
- Classes: `PascalCase` (e.g., `ComprehensiveOptimizationModel`)
- Functions/methods: `snake_case` (e.g., `calculate_net_demand_with_stages`)
- Constants: `UPPER_SNAKE_CASE` (e.g., `UNMET_DEMAND_PENALTY`)
- Variables: `snake_case` (e.g., `weekly_hours`)

### Stage Abbreviations
- `CS` - Casting
- `GR` - Grinding
- `MC` / `MC1/2/3` - Machining stages
- `SP` / `SP1/2/3` - Painting stages (Surface Painting)
- `FG` - Finished Goods

### DataFrame Column Patterns
- `*_Units` - Quantity in units
- `*_Tons` - Weight in metric tons
- `*_Util_%` - Utilization percentage
- `*_Hours` - Time in hours
- `*_min` - Time in minutes

### Error Handling Pattern
```python
def _safe_float(self, value):
    try:
        return float(value) if pd.notna(value) else 0.0
    except Exception:
        return 0.0
```

### Console Output Style
- Use emojis for status: ✓ (success), ⚠ (warning), ❌ (error)
- Progress indicators with descriptive headers
- Section separators with `"="*80`

## AI Assistant Guidelines

### When Modifying Code

1. **Preserve LP Model Integrity**: Changes to constraint formulations must maintain feasibility
2. **Maintain Stage Seriality**: Do not break the sequential stage dependencies
3. **Validate Data Types**: Always use `_safe_float()` / `_safe_int()` for Excel data
4. **Test with Real Data**: Ensure changes work with the provided Excel files

### Common Modification Points

1. **Adding new stages**: Update both `production_plan_test.py` (model) and `production_plan_executive_test7sheets.py` (reports)
2. **Changing penalties**: Modify `ProductionConfig` class
3. **Adding constraints**: Add to `ComprehensiveOptimizationModel.build_model()`
4. **New report sheets**: Add methods in `FixedExecutiveReportGenerator`

### Important Considerations

- **Holiday Calendar**: Uses `holidays.India` - adjust for different countries
- **Planning Horizon**: Dynamically calculated from latest order + buffer
- **Capacity Limits**: Auto-adjusted from optimizer output to avoid >100% utilization
- **WIP Mapping**: Uses Part Master `CS Code → FG Code` mapping

### Debugging Tips

1. Check `Weekly_Summary` sheet for aggregated optimizer output
2. Verify `Part Master` has all ordered parts
3. Ensure `Stage WIP` uses correct `Material Code` / `CastingItem` format
4. Review console output for validation warnings

## Limitations & Known Issues

1. **No CI/CD**: Manual execution only
2. **No automated tests**: Validation via Excel output inspection
3. **No dependency management**: No `requirements.txt` or `pyproject.toml`
4. **Single-solver**: Uses CBC solver only (no solver selection)
5. **In-memory processing**: Large datasets may cause memory issues

## Recent Fixes (Nov 2025)

1. **WIP-covered orders**: Orders fully covered by WIP now create variants for delivery tracking
2. **Stage skip calculation**: Fixed cascade reduction (WIP properly reduces upstream requirements)
3. **Part_Fulfillment reporting**: Now includes FG+SP WIP in delivered count
4. **Order_Fulfillment WIP allocation**: Proportionally allocates WIP to individual orders
5. **Painting constraints**: Removed dry time from capacity (passive, not active time)
6. **Week number calculation**: Fixed boundary calculation (day 7 = week 2, not week 1)
7. **Daily Production & Inventory Tracker**: New matrix format sheets for shop floor visibility
   - `9_DAILY_PRODUCTION`: Date rows × Part-Stage columns (CS, GR, MC, SP)
   - `10_DAILY_INVENTORY`: Date rows × Part-WIP columns (FG, SP, MC, GR, CS)
   - Two-row headers with part names spanning stage columns
   - Frozen panes for easy navigation

## Decision Support System Status

The system now includes a **decision support module** that answers key planning questions:

| Question | Module | Status |
|----------|--------|--------|
| "Why can't we fulfill these orders?" | `bottleneck_analyzer.py` | ✅ Implemented |
| "Can we take this new order?" | `atp_calculator.py` | ✅ Implemented |
| "Which orders need attention?" | `order_risk_dashboard.py` | ✅ Implemented |
| "What should we do about it?" | `recommendations_engine.py` | ✅ Implemented |
| "What if we add capacity/overtime?" | `scenario_analyzer.py` | 🔲 Future |

### Current Outputs

- **Bottleneck Report**: Resources at >85% utilization with overflow amounts
- **ATP Response**: Feasibility, earliest delivery, limiting resource for new orders
- **Risk Dashboard**: Critical/High/Medium/Low classification by order, customer, week
- **Recommendations**: Prioritized suggestions with action items
- **Capacity Forecast**: Available capacity by resource by week

### Future Enhancements

| Phase | Feature | Purpose |
|-------|---------|---------|
| 5 | `scenario_analyzer.py` | What-if analysis (add capacity, overtime) |
| 6 | Interactive Dashboard | Web-based visualization |
| 7 | Automated Alerts | Email/SMS for critical risks |

## Technical Improvements (Suggestions)

1. Add `requirements.txt` for dependency management
2. Implement unit tests for core functions
3. Add command-line arguments for configuration
4. Create `.gitignore` for output files
5. Add logging instead of print statements
6. Implement solver timeout configuration
7. Add data validation schemas
8. Refactor for dependency injection (enable scenario analysis)

## Contact & Support

This is an internal manufacturing optimization tool. For issues:
1. Check input Excel format against expected schema
2. Verify all required sheets are present
3. Ensure Part Master contains all ordered parts
4. Review console output for specific validation errors

---

*Last updated: 2025-11-19 (Decision Support System v1.0)*
