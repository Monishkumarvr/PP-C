# CLAUDE.md - Production Planning Optimization System

## Project Overview

This repository contains a **Manufacturing Production Planning Optimization System** designed for multi-stage manufacturing environments (casting/foundry operations). The system uses linear programming to optimize production schedules across 8 manufacturing stages while respecting capacity constraints, delivery deadlines, and work-in-progress (WIP) inventory.

**Primary Use Case**: Enterprise production planning for foundry/casting manufacturing facilities in India.

**Current Performance**: 100% fulfillment, 100% on-time delivery rate, **100% capacity utilization** on Casting and Grinding stages (PUSH model).

## Repository Structure

```
PP-C/
├── production_plan_test.py                      # Weekly optimization engine (~3,100 lines)
├── production_plan_executive_test7sheets.py     # Weekly executive reports (~3,600 lines)
├── production_plan_daily.py                     # Daily optimization engine (NEW)
├── production_plan_daily_executive.py           # Daily executive reports (NEW)
├── Master_Data_Updated_Nov_Dec.xlsx             # Input: Part master, sales orders, constraints
├── Master_Data_Optimised_Updated__2_.xlsx       # Input: Alternative/updated master data
├── production_plan_COMPREHENSIVE_test.xlsx      # Output: Weekly detailed results
├── production_plan_EXECUTIVE_test.xlsx          # Output: Weekly executive reports (10 sheets)
├── production_plan_daily_comprehensive.xlsx     # Output: Daily detailed results (NEW)
├── production_plan_daily_EXECUTIVE.xlsx         # Output: Daily executive reports (NEW)
├── DECISION_SUPPORT_IMPLEMENTATION.md           # Future decision support system design
└── CLAUDE.md                                    # This file
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

### 3. `production_plan_daily.py` - Daily Optimization Engine (NEW)

**Purpose**: Day-level linear programming optimization with Hybrid PUSH-PULL inventory management.

**Key Features**:
- **Daily granularity**: Optimizes production by calendar day (not week)
- **Hybrid PUSH-PULL**: Combines order-driven (PULL) with capacity-driven (PUSH) production
  - PULL: Produces to fulfill specific orders
  - PUSH: Uses idle capacity to process WIP through stages (speculative processing)
- **Inventory balance constraints**: Tracks daily WIP flow through CS → GR → MC → SP → FG stages
- **Stage seriality**: Enforces MC1→MC2→MC3 and SP1→SP2→SP3 sequencing
- **Hour-based capacity**: Uses cycle times to calculate machine hours (not just units/tons)

**Key Differences from Weekly Optimizer**:
- Working day calendar awareness (skips Sundays, holidays)
- Daily demand variants (orders split by delivery date)
- Inventory tracking variables for each stage
- Multi-machine capacity calculation

**Configuration Parameters** (different from weekly):
```python
LATENESS_PENALTY_PER_DAY = 20000      # Daily penalty (vs 150k/week)
INVENTORY_HOLDING_COST_PER_DAY = 0.14 # Daily holding cost
MAX_EARLY_DAYS = 56                    # 8 weeks buffer before inventory penalty
WIP_INVENTORY_COST_PER_DAY = 0.05      # Cost to hold WIP at intermediate stages
```

**Inventory Variables**:
- `inv_cs[(part, day)]` - Casting WIP inventory
- `inv_gr[(part, day)]` - Grinding WIP inventory
- `inv_mc[(part, day)]` - Machining WIP inventory
- `inv_fg[(part, day)]` - Finished goods (includes SP WIP + FG, packing assumed instant)

### 4. `production_plan_daily_executive.py` - Daily Executive Reports (NEW)

**Purpose**: Generate executive-level reports from daily optimization results with multi-machine capacity awareness.

**Key Features**:
- **Hour-based utilization**: Shows capacity % using actual cycle times
- **Multi-machine capacity calculation**: Reads Machine Constraints to compute total available hours
  - Formula: `Available_Hours = Num_Machines × Hours_Per_Shift × Shifts × OEE(0.9)`
  - Casting: 43.2 hrs/day (2 vacuum lines)
  - Grinding: 252 hrs/day (35 hand grinding machines)
  - Machining: 475.2 hrs/day (shared across MC1/MC2/MC3)
  - Painting: SP1/SP2 = 72 hrs/day each (2 primer booths), SP3 = 36 hrs/day (1 top coat booth)

**Key Classes**:
- `ProductionCalendar` - Working day calculations with Indian holidays
- `MasterDataEnricher` - Enriches schedules with cycle times and machine assignments
- `DailyExecutiveReportGenerator` - Generates 10-sheet Excel report

**Capacity Calculation Method**:
```python
def _calculate_available_hours(self):
    # Reads from Master_Data Machine Constraints sheet
    # Maps resources to stages (Casting, Grinding, MC1-3, SP1-3)
    # Returns dict: {'Casting': 43.2, 'Grinding': 252.0, ...}
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

### 4. Optimization Objective (PUSH Model - November 2025)

**PUSH Model**: Maximize capacity utilization while meeting delivery requirements.

Minimize total cost (with production rewards):
- `UNMET_DEMAND_PENALTY` (200,000) - CRITICAL: High cost for unfulfilled orders
- `LATENESS_PENALTY` (150,000) - CRITICAL: Cost per week late (prioritize on-time delivery)
- **PRODUCTION_REWARD** (-0.1) - **NEW**: Negative cost (reward) per unit produced to encourage capacity utilization
- **INVENTORY_HOLDING_COST** - **REMOVED**: No penalty for early production or inventory (allows unlimited inventory buildup)
- `SETUP_PENALTY` (5) - Pattern changeover cost

**Key Changes from PULL to PUSH Model**:
1. **Removed inventory holding penalties** → Allows unlimited early production
2. **Added production maximization incentive** → Encourages using 100% capacity
3. **Increased production variable upper bounds** → 10x demand per week (vs. limiting to exact demand)
4. **Allowed early delivery** → Can ship before due date without penalty

**Results**: Achieves 100% utilization on Casting and Grinding (bottleneck stages) across all weeks.

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
# Weekly Optimizer (legacy - week-level granularity)
python production_plan_test.py
python production_plan_executive_test7sheets.py

# Daily Optimizer (NEW - day-level granularity with Hybrid PUSH-PULL)
python production_plan_daily.py
python production_plan_daily_executive.py
```

### Typical Workflow

#### Weekly Optimizer:
1. Update `Master_Data_*.xlsx` with current:
   - Part Master specifications
   - Sales Orders with delivery dates
   - Current WIP inventory
   - Machine capacity constraints

2. Run `production_plan_test.py` to generate optimized schedule
   - Creates `production_plan_COMPREHENSIVE_test.xlsx`

3. Run `production_plan_executive_test7sheets.py` for reports
   - Creates `production_plan_EXECUTIVE_test.xlsx`

#### Daily Optimizer (Recommended for better granularity):
1. Update same `Master_Data_*.xlsx` file (shared input)

2. Run `production_plan_daily.py` to generate daily schedule
   - Creates `production_plan_daily_comprehensive.xlsx`
   - Solver time: ~8-10 minutes for 105 days, 245 orders

3. Run `production_plan_daily_executive.py` for reports
   - Creates `production_plan_daily_EXECUTIVE.xlsx`
   - Includes hour-based utilization with multi-machine capacity

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

## Recent Updates (Nov 2025)

### Weekly Optimizer Fixes:
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

### Daily Optimizer Implementation (NEW):
8. **Hybrid PUSH-PULL System**: Implemented inventory balance constraints for speculative WIP processing
   - **File**: `production_plan_daily.py`
   - **Features**:
     - Inventory variables by part and day: `inv_cs`, `inv_gr`, `inv_mc`, `inv_fg`
     - Stage seriality constraints: MC1→MC2→MC3, SP1→SP2→SP3
     - Inventory balance: `inv[d] = inv[d-1] + production - consumption`
     - WIP inventory holding cost (uniform 0.05/day - **needs graduated costs**)
   - **Status**: ⚠️ Works but underutilizes capacity (casting often 0%)
   - **Next**: Implement graduated WIP costs (see DECISION_SUPPORT_IMPLEMENTATION.md §5.5)

9. **Cycle Time Key Fixes**: Corrected parameter key names for hour calculation
   - `grinding_cycle` → `grind_cycle`
   - `mc1_cycle/mc2_cycle/mc3_cycle` → `mach_cycles[0/1/2]`
   - `sp1_cycle/sp2_cycle/sp3_cycle` → `paint_cycles[0/1/2]`
   - Applied to both optimization constraints and reporting

10. **Multi-Machine Capacity Calculation**: Hour-based utilization with realistic capacity limits
    - **File**: `production_plan_daily_executive.py`
    - **Method**: `_calculate_available_hours()` reads Machine Constraints sheet
    - **Results**:
      - Casting: 3.7% avg, 27.8% peak (43.2 hrs/day available)
      - Grinding: 15.4% avg, 100% peak (252 hrs/day available) ← **Bottleneck identified**
      - MC1: 2.1% avg, 43.9% peak (475.2 hrs/day available)
      - SP1: 4.1% avg, 78.7% peak (72 hrs/day available)
    - All utilizations now realistic (≤100% except bottlenecks)

11. **SP WIP vs FG Clarification**: Packing treated as instantaneous
    - **SP WIP** = Painted but unpacked parts (ready to pack)
    - **FG** = Packed parts (ready to ship)
    - **Model**: SP WIP + FG combined as `inv_fg` (packing non-bottleneck)
    - **Flow**: Casting → Grinding → Machining → Painting → SP WIP → Packing (instant) → FG → Delivery

### PUSH Model Transformation (Nov 21, 2025):
12. **Full Capacity PUSH Model**: Transformed weekly optimizer from PULL (just-in-time) to PUSH (capacity-driven)
    - **File**: `production_plan_test.py`
    - **Objective Changed**:
      - ❌ **REMOVED**: Inventory holding cost penalty (was 1 per unit per week)
      - ✅ **ADDED**: Production maximization reward (-0.1 per unit produced)
      - ✅ **KEPT**: Unmet demand penalty (200,000) and lateness penalty (150,000) for delivery requirements
    - **Production Bounds Changed**:
      - **Before**: Limited to exact demand quantity (cast_ub = demand)
      - **After**: Allow 10x overproduction (cast_ub = 10 * demand) to enable PUSH production
    - **Delivery Constraints Changed**:
      - **Before**: Restricted to specific delivery window (window_start to window_end)
      - **After**: Allow early delivery in any week up to due date
    - **Results Achieved**:
      - ✅ **Casting**: 100% utilization (Big Line + Small Line) across all 16 weeks
      - ✅ **Grinding**: 100% utilization across all 16 weeks
      - ✅ **Machining**: 10-26% (MC1), 2-20% (MC2), 0-9% (MC3) - bottleneck limited
      - ✅ **Painting**: 4-26% (SP1), 3-33% (SP2), 0-21% (SP3) - bottleneck limited
      - ✅ **Order Fulfillment**: 100% fulfillment, 100% on-time delivery maintained
    - **Minimum Utilization Constraints**: Initially attempted 95% minimum constraints per stage, but caused infeasibility - **disabled** in favor of production reward approach

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

*Last updated: 2025-11-21*
*Weekly Optimizer: Transformed to Full Capacity PUSH Model - achieving 100% utilization on Casting and Grinding*
*Daily Optimizer: Implemented Hybrid PUSH-PULL with hour-based multi-machine capacity tracking*
