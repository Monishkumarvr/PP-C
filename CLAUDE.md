# CLAUDE.md - Production Planning Optimization System

## Project Overview

This repository contains a **Manufacturing Production Planning Optimization System** designed for multi-stage manufacturing environments (casting/foundry operations). The system uses linear programming to optimize production schedules across 8 manufacturing stages while respecting capacity constraints, delivery deadlines, and work-in-progress (WIP) inventory.

**Primary Use Case**: Enterprise production planning for foundry/casting manufacturing facilities in India.

**Current Performance**: 100% fulfillment, 100% on-time delivery rate.

## Repository Structure

```
PP-C/
├── production_plan_test.py                    # Core optimization engine (~3,100 lines)
├── production_plan_executive_test7sheets.py   # Executive reporting module (~3,030 lines)
├── Master_Data_Updated_Nov_Dec.xlsx           # Input: Part master, sales orders, constraints
├── Master_Data_Optimised_Updated__2_.xlsx     # Input: Alternative/updated master data
├── production_plan_COMPREHENSIVE_test.xlsx    # Output: Detailed optimization results
├── production_plan_EXECUTIVE_test.xlsx        # Output: Executive dashboard reports
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
- `FixedExecutiveReportGenerator` - Multi-sheet Excel report generation

**Output Sheets Generated**:
- Executive Dashboard per stage (8 sheets)
- Daily Schedule (aggregate with calendar dates)
- Part-Level Daily Schedule (detailed machine assignments)

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

## Future Roadmap: Decision Support System

The current tool generates optimal schedules. Future development should transform it into a **decision support system** that answers:

1. **"Why can't we fulfill these orders?"** → Bottleneck Analysis
2. **"Can we take this new order?"** → Available-to-Promise (ATP) Calculator
3. **"Which orders need attention?"** → Order Risk Dashboard
4. **"What should we do about it?"** → Recommendations Engine

### Planned Modules

| Phase | Module | Purpose |
|-------|--------|---------|
| 1 | `bottleneck_analyzer.py` | Identify resources blocking fulfillment |
| 2 | `atp_calculator.py` | Check feasibility of new orders |
| 3 | `order_risk_dashboard.py` | Classify orders by risk level |
| 4 | `recommendations_engine.py` | Generate actionable suggestions |
| 5 | `scenario_analyzer.py` | What-if analysis (add capacity, overtime) |

### Key Outputs Needed

- **Bottleneck Report**: Which resources block which orders
- **ATP Response**: Earliest delivery date for new orders
- **Risk Dashboard**: Critical/High/Medium/Low order classification
- **Recommendations**: Overtime, outsourcing, rescheduling suggestions
- **Capacity Forecast**: Available capacity by resource by week

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

*Last updated: 2025-11-18*
