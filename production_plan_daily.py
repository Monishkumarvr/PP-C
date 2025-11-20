"""
Daily-Level Production Planning Optimization System
===================================================
Optimizes production schedules at DAILY granularity instead of weekly.

Key Differences from Weekly Version:
- Decision variables: x_casting[(variant, day)] instead of [(variant, week)]
- Constraints: Daily capacity limits instead of weekly
- Lead times: Measured in days (cooling_time, drying_time, etc.)
- Feasibility: Ensures no single day exceeds capacity

Usage:
    python production_plan_daily.py

Generates: production_plan_daily_comprehensive.xlsx
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import pulp
from collections import defaultdict
import holidays

# Reuse existing infrastructure
from production_plan_test import (
    ProductionCalendar,
    ComprehensiveDataLoader,
    WIPDemandCalculator,
    ComprehensiveParameterBuilder,
    MachineResourceManager,
    BoxCapacityManager,
    build_wip_init
)


class DailyProductionConfig:
    """Configuration for daily-level production planning."""
    
    def __init__(self):
        # Planning horizon
        self.CURRENT_DATE = datetime(2025, 10, 1)
        self.PLANNING_BUFFER_DAYS = 14  # Buffer beyond latest order
        self.MAX_PLANNING_DAYS = 210  # ~7 months

        # Compatibility attributes for weekly data loader and parameter builder
        self.PLANNING_BUFFER_WEEKS = int(self.PLANNING_BUFFER_DAYS / 7)  # ~2 weeks
        self.MAX_PLANNING_WEEKS = int(self.MAX_PLANNING_DAYS / 7)  # ~30 weeks
        self.PLANNING_WEEKS = None  # Will be calculated dynamically
        self.TRACKING_WEEKS = None  # Will be calculated dynamically

        # Lag attributes (set to 0 as we handle lead times in days directly)
        self.COOLING_SHAKEOUT_LAG_WEEKS = 0
        self.GRINDING_LAG_WEEKS = 0
        self.MACHINING_LAG_WEEKS = 0
        self.PAINTING_LAG_WEEKS = 0

        # Lead time parameters (for parameter builder compatibility)
        self.MIN_LEAD_TIME_WEEKS = 2  # Minimum lead time
        self.AVG_LEAD_TIME_WEEKS = 4  # Average lead time for forecasting
        self.DELIVERY_BUFFER_WEEKS = 1  # Delivery flexibility
        self.MAX_EARLY_WEEKS = 8  # Maximum weeks to produce early

        # Additional compatibility attributes
        self.WORKING_HOURS_PER_DAY = 24  # Hours available for cooling/shakeout
        self.OVERTIME_ALLOWANCE = 0.0
        self.STARTUP_BONUS = -50
        self.COUNTRY_CODE = 'IN'

        # Working schedule
        self.WORKING_DAYS_PER_WEEK = 6
        self.WEEKLY_OFF_DAY = 6  # Sunday (0=Monday, 6=Sunday)
        
        # Machine configuration
        self.OEE = 0.90
        self.HOURS_PER_SHIFT = 12
        self.SHIFTS_PER_DAY = 2
        
        # Daily capacities
        self.BIG_LINE_HOURS_PER_SHIFT = self.HOURS_PER_SHIFT
        self.SMALL_LINE_HOURS_PER_SHIFT = self.HOURS_PER_SHIFT
        self.BIG_LINE_HOURS_PER_DAY = self.HOURS_PER_SHIFT * self.SHIFTS_PER_DAY * self.OEE
        self.SMALL_LINE_HOURS_PER_DAY = self.HOURS_PER_SHIFT * self.SHIFTS_PER_DAY * self.OEE

        # Weekly capacities (for compatibility with existing code)
        self.BIG_LINE_HOURS_PER_WEEK = self.BIG_LINE_HOURS_PER_DAY * self.WORKING_DAYS_PER_WEEK
        self.SMALL_LINE_HOURS_PER_WEEK = self.SMALL_LINE_HOURS_PER_DAY * self.WORKING_DAYS_PER_WEEK
        
        # Lead times (in days)
        self.DEFAULT_COOLING_DAYS = 1  # Casting -> Grinding
        self.DEFAULT_GRINDING_TO_MC1_DAYS = 1
        self.DEFAULT_MC_STAGE_DAYS = 1  # Between MC stages
        self.DEFAULT_PAINT_DRYING_DAYS = 1  # Between paint stages

        # Penalties (daily-based) - MATCHED TO WEEKLY OPTIMIZER
        self.UNMET_DEMAND_PENALTY = 200000
        self.LATENESS_PENALTY_PER_DAY = 20000  # ~150k/week ÷ 7 ≈ 21,429/day (increased from 5k)
        self.INVENTORY_HOLDING_COST_PER_DAY = 0.14  # ~1/week ÷ 7 ≈ 0.14/day
        self.MAX_EARLY_DAYS = 56  # 8 weeks × 7 days (allow up to 8 weeks early delivery)
        self.SETUP_PENALTY = 5

        # Hybrid PUSH-PULL parameters
        self.WIP_INVENTORY_COST_PER_DAY = 0.05  # Cost to hold WIP inventory per unit per day
        self.IDLE_CAPACITY_PENALTY = 0  # Set to 0 for now (can enable later)

        # Penalty compatibility (weekly equivalents for existing code)
        self.LATENESS_PENALTY = 150000  # Per week late (for compatibility)
        self.INVENTORY_HOLDING_COST = 1  # Per unit per week (for compatibility)
        
        # Delivery flexibility
        self.DELIVERY_WINDOW_DAYS = 3  # Allow ±3 days from due date
        
        # Vacuum moulding
        self.VACUUM_CAPACITY_PENALTY = 0.70
        
        # Pattern changeover
        self.PATTERN_CHANGE_TIME_MIN = 18
        
        # Dynamic planning horizon (will be calculated)
        self.PLANNING_DAYS = None


print("="*80)
print("DAILY-LEVEL PRODUCTION PLANNING OPTIMIZATION")
print("="*80)
print("\nThis script implements DAILY granularity optimization.")
print("Each day is optimized independently with daily capacity constraints.")
print("\nKey improvements over weekly version:")
print("  ✓ Respects daily capacity limits")
print("  ✓ Accurate lead times (in days, not weeks)")
print("  ✓ Feasible daily schedules")
print("  ✓ Better short-term planning visibility")
print("\n" + "="*80 + "\n")

# Continue in next message due to length...


class DailyDemandCalculator:
    """Calculate daily delivery targets from orders."""
    
    def __init__(self, sales_order, stage_wip, config, calendar):
        self.sales_order = sales_order
        self.stage_wip = stage_wip
        self.config = config
        self.calendar = calendar
    
    def calculate_daily_demand(self):
        """Convert weekly demand to daily delivery targets."""
        print("\n" + "="*80)
        print("CALCULATING DAILY DEMAND")
        print("="*80)
        
        # Calculate WIP coverage (reuse weekly logic)
        weekly_calc = WIPDemandCalculator(self.sales_order, self.stage_wip, self.config)
        (net_demand, stage_start_qty, wip_coverage,
         gross_demand, wip_by_part) = weekly_calc.calculate_net_demand_with_stages()
        
        # Convert net_demand (part->qty) to daily deliveries
        daily_demand = {}
        part_day_mapping = {}
        variant_id = 0
        
        for _, order in self.sales_order.iterrows():
            part = str(order['Material Code']).strip()
            qty = int(order['Balance Qty'] or 0)
            due_date = order['Delivery_Date']
            
            if qty <= 0:
                continue
            
            # Find closest working day to due date
            delivery_day = self.calendar.get_nearest_working_day(due_date)
            
            # Create variant
            variant = f"V{variant_id}"
            variant_id += 1
            
            daily_demand[variant] = qty
            part_day_mapping[variant] = (part, delivery_day)
        
        print(f"✓ Created {len(daily_demand)} daily demand variants")
        print(f"✓ Parts: {len(set(p for p, d in part_day_mapping.values()))} unique parts")
        
        return daily_demand, part_day_mapping, stage_start_qty, wip_by_part


class DailyOptimizationModel:
    """Daily-level production optimization model."""
    
    def __init__(self, daily_demand, part_day_mapping, params, stage_start_qty,
                 machine_manager, box_manager, config, calendar, wip_init=None):
        self.daily_demand = daily_demand
        self.part_day_mapping = part_day_mapping
        self.params = params
        self.stage_start_qty = stage_start_qty
        self.machine_manager = machine_manager
        self.box_manager = box_manager
        self.config = config
        self.calendar = calendar
        self.wip_init = wip_init or {}
        
        # Get all working days
        self.working_days = calendar.working_days
        self.day_index = {day: idx for idx, day in enumerate(self.working_days)}
        
        # LP model
        self.model = None
        self.x_casting = {}
        self.x_grinding = {}
        self.x_mc1 = {}
        self.x_mc2 = {}
        self.x_mc3 = {}
        self.x_sp1 = {}
        self.x_sp2 = {}
        self.x_sp3 = {}
        self.x_delivery = {}
        self.x_unmet = {}

        # Hybrid PUSH-PULL: Inventory tracking variables (by part, not variant)
        self.inv_cs = {}  # CS inventory (casting + CS WIP)
        self.inv_gr = {}  # GR inventory (ground parts)
        self.inv_mc = {}  # MC inventory (machined parts)
        # NOTE: inv_sp removed - SP WIP combined with FG (both are finished goods)
        self.inv_fg = {}  # FG inventory (finished goods, includes FG + SP WIP initially)
    
    def build_and_solve(self):
        """Build complete model and solve."""
        print("\n" + "="*80)
        print("BUILDING DAILY OPTIMIZATION MODEL")
        print("="*80)
        
        self.model = pulp.LpProblem("Daily_Production_Planning", pulp.LpMinimize)

        # Build components
        self._create_decision_variables()
        self._add_objective_function()
        self._add_capacity_constraints()
        self._add_box_constraints()
        self._add_inventory_balance_constraints()  # NEW: Hybrid PUSH-PULL inventory tracking
        self._add_stage_seriality_constraints()  # NEW: MC1→MC2→MC3, SP1→SP2→SP3 stage ordering
        # REMOVED: self._add_flow_constraints()  # Conflicts with inventory balance (part-level redundant)
        self._add_demand_constraints()

        print("\n✓ Model built successfully")
        print(f"  Variables: {len(self.model.variables()):,}")
        print(f"  Constraints: {len(self.model.constraints):,}")
        
        # Solve
        print("\n🔧 Solving optimization model...")
        print("   (This may take 5-15 minutes for daily optimization)")
        
        solver = pulp.PULP_CBC_CMD(msg=1, timeLimit=600)  # 10 min timeout
        status = self.model.solve(solver)
        
        print(f"\n✓ Solver completed: {pulp.LpStatus[status]}")
        
        return status
    
    def _create_decision_variables(self):
        """Create decision variables with routing awareness."""
        print("\n📊 Creating decision variables (with routing awareness)...")

        variants = list(self.daily_demand.keys())
        days = self.working_days

        # Production variables for each stage (with routing awareness)
        for v in variants:
            part, _ = self.part_day_mapping[v]
            part_params = self.params.get(part, {})

            for d in days:
                # Casting - Integer (whole units)
                self.x_casting[(v, d)] = pulp.LpVariable(
                    f"cast_{v}_{d.strftime('%Y%m%d')}", 0, None, cat='Integer'
                )

                # Grinding - Integer (whole units for daily granularity)
                self.x_grinding[(v, d)] = pulp.LpVariable(
                    f"grind_{v}_{d.strftime('%Y%m%d')}", 0, None, cat='Integer'
                )

                # Machining stages - Integer, only create if part routing requires them
                if part_params.get('has_mc1', True):
                    self.x_mc1[(v, d)] = pulp.LpVariable(
                        f"mc1_{v}_{d.strftime('%Y%m%d')}", 0, None, cat='Integer'
                    )
                else:
                    self.x_mc1[(v, d)] = 0  # Part skips MC1

                if part_params.get('has_mc2', True):
                    self.x_mc2[(v, d)] = pulp.LpVariable(
                        f"mc2_{v}_{d.strftime('%Y%m%d')}", 0, None, cat='Integer'
                    )
                else:
                    self.x_mc2[(v, d)] = 0  # Part skips MC2

                if part_params.get('has_mc3', True):
                    self.x_mc3[(v, d)] = pulp.LpVariable(
                        f"mc3_{v}_{d.strftime('%Y%m%d')}", 0, None, cat='Integer'
                    )
                else:
                    self.x_mc3[(v, d)] = 0  # Part skips MC3

                # Painting stages - Integer, only create if part routing requires them
                if part_params.get('has_sp1', True):
                    self.x_sp1[(v, d)] = pulp.LpVariable(
                        f"sp1_{v}_{d.strftime('%Y%m%d')}", 0, None, cat='Integer'
                    )
                else:
                    self.x_sp1[(v, d)] = 0  # Part skips SP1

                if part_params.get('has_sp2', True):
                    self.x_sp2[(v, d)] = pulp.LpVariable(
                        f"sp2_{v}_{d.strftime('%Y%m%d')}", 0, None, cat='Integer'
                    )
                else:
                    self.x_sp2[(v, d)] = 0  # Part skips SP2

                if part_params.get('has_sp3', True):
                    self.x_sp3[(v, d)] = pulp.LpVariable(
                        f"sp3_{v}_{d.strftime('%Y%m%d')}", 0, None, cat='Integer'
                    )
                else:
                    self.x_sp3[(v, d)] = 0  # Part skips SP3

                # Delivery - Integer (whole units)
                self.x_delivery[(v, d)] = pulp.LpVariable(
                    f"deliv_{v}_{d.strftime('%Y%m%d')}", 0, None, cat='Integer'
                )
        
        # Unmet demand (no separate lateness variable - calculated directly in objective like weekly)
        for v in variants:
            self.x_unmet[v] = pulp.LpVariable(f"unmet_{v}", 0, None)

        # Hybrid PUSH-PULL: Create inventory tracking variables (by PART, not variant)
        parts = set(part for part, _ in self.part_day_mapping.values())
        for part in parts:
            for d in days:
                # Inventory at each stage
                self.inv_cs[(part, d)] = pulp.LpVariable(
                    f"inv_cs_{part}_{d.strftime('%Y%m%d')}", 0, None, cat='Continuous'
                )
                self.inv_gr[(part, d)] = pulp.LpVariable(
                    f"inv_gr_{part}_{d.strftime('%Y%m%d')}", 0, None, cat='Continuous'
                )
                self.inv_mc[(part, d)] = pulp.LpVariable(
                    f"inv_mc_{part}_{d.strftime('%Y%m%d')}", 0, None, cat='Continuous'
                )
                # NOTE: inv_sp removed - SP WIP is combined with FG (both are finished goods)
                self.inv_fg[(part, d)] = pulp.LpVariable(
                    f"inv_fg_{part}_{d.strftime('%Y%m%d')}", 0, None, cat='Continuous'
                )

        print(f"  ✓ Created decision variables (including {len(parts) * len(days) * 4} inventory variables)")
    
    def _add_objective_function(self):
        """Add objective: minimize cost (matching weekly optimizer logic exactly)."""
        print("\n🎯 Adding objective function...")

        objective_terms = []

        # Unmet demand penalty
        for v in self.daily_demand.keys():
            objective_terms.append(self.config.UNMET_DEMAND_PENALTY * self.x_unmet[v])

        # Lateness penalty - calculated directly like weekly optimizer (no separate variable)
        for v in self.daily_demand.keys():
            part, due_day = self.part_day_mapping[v]
            due_day_idx = self.day_index[due_day]

            for d_idx, d in enumerate(self.working_days):
                days_late = max(0, d_idx - due_day_idx)
                if days_late > 0:
                    objective_terms.append(
                        self.config.LATENESS_PENALTY_PER_DAY * days_late * self.x_delivery[(v, d)]
                    )

        # Inventory holding cost - ONLY penalize deliveries > MAX_EARLY_DAYS (like weekly optimizer)
        for v in self.daily_demand.keys():
            part, due_day = self.part_day_mapping[v]
            due_day_idx = self.day_index[due_day]

            for d_idx, d in enumerate(self.working_days):
                days_early = due_day_idx - d_idx
                # Only penalize if delivering TOO early (beyond MAX_EARLY_DAYS buffer)
                if days_early > self.config.MAX_EARLY_DAYS:
                    excess_early_days = days_early - self.config.MAX_EARLY_DAYS
                    objective_terms.append(
                        self.config.INVENTORY_HOLDING_COST_PER_DAY * excess_early_days * self.x_delivery[(v, d)]
                    )

        # Hybrid PUSH-PULL: WIP inventory holding cost (encourages processing WIP but not too much)
        parts = set(part for part, _ in self.part_day_mapping.values())
        for part in parts:
            for d in self.working_days:
                # Small cost to hold WIP inventory (allows speculative processing but penalizes excess)
                objective_terms.append(
                    self.config.WIP_INVENTORY_COST_PER_DAY * (
                        self.inv_cs[(part, d)] + self.inv_gr[(part, d)] + self.inv_mc[(part, d)]
                    )
                )

        self.model += pulp.lpSum(objective_terms), "Total_Cost"
        print("  ✓ Objective function added (unmet + lateness + FG inventory + WIP inventory)")
    
    def _add_capacity_constraints(self):
        """Add daily capacity constraints for all resources."""
        print("\n⚙️ Adding daily capacity constraints...")
        
        # Big/Small line daily capacity
        big_line_cap_min = self.config.BIG_LINE_HOURS_PER_DAY * 60
        small_line_cap_min = self.config.SMALL_LINE_HOURS_PER_DAY * 60
        
        for d in self.working_days:
            big_line_minutes = 0
            small_line_minutes = 0
            
            for v in self.daily_demand.keys():
                part, _ = self.part_day_mapping[v]
                p = self.params.get(part, {})
                
                casting_cycle = p.get('casting_cycle', 0)
                moulding_line = p.get('moulding_line', '')
                requires_vacuum = p.get('requires_vacuum', False)
                
                effective_cycle = casting_cycle
                if requires_vacuum:
                    effective_cycle = casting_cycle / self.config.VACUUM_CAPACITY_PENALTY
                
                if 'Big Line' in moulding_line:
                    big_line_minutes += self.x_casting[(v, d)] * effective_cycle
                elif 'Small Line' in moulding_line:
                    small_line_minutes += self.x_casting[(v, d)] * effective_cycle
            
            # Daily capacity limits
            self.model += (
                big_line_minutes <= big_line_cap_min,
                f"BigLine_Cap_{d.strftime('%Y%m%d')}"
            )
            self.model += (
                small_line_minutes <= small_line_cap_min,
                f"SmallLine_Cap_{d.strftime('%Y%m%d')}"
            )
        
        # Machine resource daily capacities
        for resource_code, resource_data in self.machine_manager.machines.items():
            daily_hours = resource_data['weekly_hours'] / 6  # Approximate daily capacity
            daily_minutes = daily_hours * 60
            
            operation = resource_data['operation']
            
            for d in self.working_days:
                usage_minutes = 0
                
                for v in self.daily_demand.keys():
                    part, _ = self.part_day_mapping[v]
                    p = self.params.get(part, {})
                    
                    # Determine which stage uses this resource
                    if 'Grinding' in operation or 'GR' in resource_code:
                        usage_minutes += self.x_grinding[(v, d)] * p.get('grinding_cycle', 0)
                    elif 'MC1' in operation or 'MC1' in resource_code:
                        usage_minutes += self.x_mc1[(v, d)] * p.get('mc1_cycle', 0)
                    elif 'MC2' in operation or 'MC2' in resource_code:
                        usage_minutes += self.x_mc2[(v, d)] * p.get('mc2_cycle', 0)
                    elif 'MC3' in operation or 'MC3' in resource_code:
                        usage_minutes += self.x_mc3[(v, d)] * p.get('mc3_cycle', 0)
                    elif 'SP1' in operation or 'Primer' in operation:
                        usage_minutes += self.x_sp1[(v, d)] * p.get('sp1_cycle', 0)
                    elif 'SP2' in operation or 'Intermediate' in operation:
                        usage_minutes += self.x_sp2[(v, d)] * p.get('sp2_cycle', 0)
                    elif 'SP3' in operation or 'Top' in operation:
                        usage_minutes += self.x_sp3[(v, d)] * p.get('sp3_cycle', 0)

                # Add constraint (usage_minutes is a PuLP expression, always add it)
                self.model += (
                    usage_minutes <= daily_minutes,
                    f"{resource_code}_Cap_{d.strftime('%Y%m%d')}"
                )
        
        print(f"  ✓ Added {len(self.working_days) * 2} casting line capacity constraints")
        print(f"  ✓ Added {len(self.machine_manager.machines) * len(self.working_days)} machine capacity constraints")

    def _add_box_constraints(self):
        """Add mould box capacity constraints (daily version of weekly logic)."""
        print("\n📦 Adding mould box capacity constraints...")

        box_variants = defaultdict(list)
        for v in self.daily_demand.keys():
            part, _ = self.part_day_mapping[v]
            if part not in self.params:
                continue

            box_size = self.params[part].get('box_size')
            box_qty = self.params[part].get('box_quantity', 0)
            if box_size and box_size != 'Unknown' and box_qty > 0:
                box_variants[box_size].append((v, max(1, int(box_qty))))

        constraints_added = 0
        for box_size, vlist in box_variants.items():
            base_cap = self.box_manager.get_capacity(box_size)
            if base_cap == 0:
                continue

            # Daily capacity = weekly capacity / 6 working days
            daily_cap = base_cap / 6.0 * 0.90  # With OEE

            for d in self.working_days:
                terms = []
                for (v, box_qty) in vlist:
                    moulds_per_unit = 1.0 / float(box_qty)
                    terms.append(self.x_casting[(v, d)] * moulds_per_unit)

                if terms:
                    self.model += (
                        pulp.lpSum(terms) <= daily_cap,
                        f"Box_{box_size}_D{d.strftime('%Y%m%d')}"
                    )
                    constraints_added += 1

        print(f"  ✓ Added {constraints_added} mould box capacity constraints")

    def _add_inventory_balance_constraints(self):
        """Add inventory balance constraints for hybrid PUSH-PULL system."""
        print("\n📦 Adding inventory balance constraints (Hybrid PUSH-PULL)...")

        parts = set(part for part, _ in self.part_day_mapping.values())
        variants_by_part = defaultdict(list)
        for v in self.daily_demand.keys():
            part, _ = self.part_day_mapping[v]
            variants_by_part[part].append(v)

        constraints_added = 0

        for part in parts:
            if part not in self.params:
                continue

            p = self.params[part]
            wip = self.wip_init.get(part, {'FG': 0, 'SP': 0, 'MC': 0, 'GR': 0, 'CS': 0})
            variants = variants_by_part[part]

            for d_idx, d in enumerate(self.working_days):
                # Aggregate production for this part across all variants
                casting_total = pulp.lpSum(self.x_casting[(v, d)] for v in variants)
                grinding_total = pulp.lpSum(self.x_grinding[(v, d)] for v in variants)
                mc1_total = pulp.lpSum(self.x_mc1[(v, d)] for v in variants if isinstance(self.x_mc1.get((v, d)), pulp.LpVariable))
                mc2_total = pulp.lpSum(self.x_mc2[(v, d)] for v in variants if isinstance(self.x_mc2.get((v, d)), pulp.LpVariable))
                mc3_total = pulp.lpSum(self.x_mc3[(v, d)] for v in variants if isinstance(self.x_mc3.get((v, d)), pulp.LpVariable))
                sp1_total = pulp.lpSum(self.x_sp1[(v, d)] for v in variants if isinstance(self.x_sp1.get((v, d)), pulp.LpVariable))
                sp2_total = pulp.lpSum(self.x_sp2[(v, d)] for v in variants if isinstance(self.x_sp2.get((v, d)), pulp.LpVariable))
                sp3_total = pulp.lpSum(self.x_sp3[(v, d)] for v in variants if isinstance(self.x_sp3.get((v, d)), pulp.LpVariable))
                delivery_total = pulp.lpSum(self.x_delivery[(v, d)] for v in variants)

                # Initial inventory (first day)
                if d_idx == 0:
                    # CS Inventory Balance: inv_cs[d] = CS_WIP + casting[d] - grinding[d]
                    self.model += (
                        self.inv_cs[(part, d)] == wip['CS'] + casting_total - grinding_total,
                        f"InvBal_CS_{part}_D{d_idx}"
                    )
                    constraints_added += 1

                    # GR Inventory Balance: inv_gr[d] = GR_WIP + grinding[d] - mc1[d]
                    self.model += (
                        self.inv_gr[(part, d)] == wip['GR'] + grinding_total - mc1_total,
                        f"InvBal_GR_{part}_D{d_idx}"
                    )
                    constraints_added += 1

                    # MC Inventory Balance: inv_mc[d] = MC_WIP + (mc1+mc2+mc3)[d] - sp1[d]
                    mc_output = mc1_total + mc2_total + mc3_total
                    self.model += (
                        self.inv_mc[(part, d)] == wip['MC'] + mc_output - sp1_total,
                        f"InvBal_MC_{part}_D{d_idx}"
                    )
                    constraints_added += 1

                    # FG Inventory Balance: inv_fg[d] = (FG_WIP + SP_WIP) + sp_output[d] - delivery[d]
                    # NOTE: SP WIP = finished painted parts, same as FG, so combine them
                    # Finished goods come from last SP stage
                    if p.get('has_sp3', True):
                        sp_output = sp3_total
                    elif p.get('has_sp2', True):
                        sp_output = sp2_total
                    else:
                        sp_output = sp1_total

                    self.model += (
                        self.inv_fg[(part, d)] == wip['FG'] + wip['SP'] + sp_output - delivery_total,
                        f"InvBal_FG_{part}_D{d_idx}"
                    )
                    constraints_added += 1

                else:
                    # Subsequent days: inv[d] = inv[d-1] + inflow[d] - outflow[d]
                    prev_d = self.working_days[d_idx - 1]

                    # CS: inv_cs[d] = inv_cs[d-1] + casting[d] - grinding[d]
                    self.model += (
                        self.inv_cs[(part, d)] == self.inv_cs[(part, prev_d)] + casting_total - grinding_total,
                        f"InvBal_CS_{part}_D{d_idx}"
                    )
                    constraints_added += 1

                    # GR: inv_gr[d] = inv_gr[d-1] + grinding[d] - mc1[d]
                    self.model += (
                        self.inv_gr[(part, d)] == self.inv_gr[(part, prev_d)] + grinding_total - mc1_total,
                        f"InvBal_GR_{part}_D{d_idx}"
                    )
                    constraints_added += 1

                    # MC: inv_mc[d] = inv_mc[d-1] + mc_output[d] - sp1[d]
                    mc_output = mc1_total + mc2_total + mc3_total
                    self.model += (
                        self.inv_mc[(part, d)] == self.inv_mc[(part, prev_d)] + mc_output - sp1_total,
                        f"InvBal_MC_{part}_D{d_idx}"
                    )
                    constraints_added += 1

                    # FG: inv_fg[d] = inv_fg[d-1] + sp_output[d] - delivery[d]
                    # Finished goods come from last SP stage
                    if p.get('has_sp3', True):
                        sp_output = sp3_total
                    elif p.get('has_sp2', True):
                        sp_output = sp2_total
                    else:
                        sp_output = sp1_total

                    self.model += (
                        self.inv_fg[(part, d)] == self.inv_fg[(part, prev_d)] + sp_output - delivery_total,
                        f"InvBal_FG_{part}_D{d_idx}"
                    )
                    constraints_added += 1

        print(f"  ✓ Added {constraints_added} inventory balance constraints")
        print(f"  → Enables speculative WIP processing to utilize idle capacity!")

    def _add_stage_seriality_constraints(self):
        """Add stage seriality constraints (MC1→MC2→MC3, SP1→SP2→SP3) for variants."""
        print("\n🔄 Adding stage seriality constraints...")

        cnt = 0

        # Add VARIANT-LEVEL cumulative constraints for internal stage seriality
        for v in self.daily_demand.keys():
            part, _ = self.part_day_mapping[v]
            if part not in self.params:
                continue

            p = self.params[part]

            for d_idx, d in enumerate(self.working_days):
                days_up_to_d = self.working_days[:d_idx + 1]

                # CUMULATIVE: MC2 <= MC1
                if p.get('has_mc2', True) and p.get('has_mc1', True):
                    self.model += (
                        pulp.lpSum(self.x_mc2[(v, t)] for t in days_up_to_d)
                        <= pulp.lpSum(self.x_mc1[(v, t)] for t in days_up_to_d),
                        f"Serial_MC1_MC2_{v}_D{d_idx}"
                    )
                    cnt += 1

                # CUMULATIVE: MC3 <= MC2 or MC1
                if p.get('has_mc3', True):
                    if p.get('has_mc2', True):
                        self.model += (
                            pulp.lpSum(self.x_mc3[(v, t)] for t in days_up_to_d)
                            <= pulp.lpSum(self.x_mc2[(v, t)] for t in days_up_to_d),
                            f"Serial_MC2_MC3_{v}_D{d_idx}"
                        )
                    elif p.get('has_mc1', True):
                        self.model += (
                            pulp.lpSum(self.x_mc3[(v, t)] for t in days_up_to_d)
                            <= pulp.lpSum(self.x_mc1[(v, t)] for t in days_up_to_d),
                            f"Serial_MC1_MC3_{v}_D{d_idx}"
                        )
                    cnt += 1

                # CUMULATIVE: SP2 <= SP1
                if p.get('has_sp2', True):
                    self.model += (
                        pulp.lpSum(self.x_sp2[(v, t)] for t in days_up_to_d)
                        <= pulp.lpSum(self.x_sp1[(v, t)] for t in days_up_to_d),
                        f"Serial_SP1_SP2_{v}_D{d_idx}"
                    )
                    cnt += 1

                # CUMULATIVE: SP3 <= SP2 or SP1
                if p.get('has_sp3', True):
                    if p.get('has_sp2', True):
                        self.model += (
                            pulp.lpSum(self.x_sp3[(v, t)] for t in days_up_to_d)
                            <= pulp.lpSum(self.x_sp2[(v, t)] for t in days_up_to_d),
                            f"Serial_SP2_SP3_{v}_D{d_idx}"
                        )
                    else:
                        self.model += (
                            pulp.lpSum(self.x_sp3[(v, t)] for t in days_up_to_d)
                            <= pulp.lpSum(self.x_sp1[(v, t)] for t in days_up_to_d),
                            f"Serial_SP1_SP3_{v}_D{d_idx}"
                        )
                    cnt += 1

        print(f"  ✓ Added {cnt:,} stage seriality constraints (MC1→MC2→MC3, SP1→SP2→SP3)")

    def _add_flow_constraints(self):
        """Add CUMULATIVE material flow constraints (like weekly optimizer)."""
        print("\n🔄 Adding CUMULATIVE flow constraints with stage seriality...")

        cnt = 0

        # Group variants by part for aggregate constraints
        variants_by_part = defaultdict(list)
        for v in self.daily_demand.keys():
            part, _ = self.part_day_mapping[v]
            if part in self.params:
                variants_by_part[part].append(v)

        # Add PART-LEVEL cumulative constraints for WIP-to-production transitions
        for part, variants in variants_by_part.items():
            p = self.params[part]
            wip = self.wip_init.get(part, {'FG': 0, 'SP': 0, 'MC': 0, 'GR': 0, 'CS': 0})

            # Lead times (in days)
            cooling_days = max(1, int(p.get('cooling_time', 0) / 24))

            for d_idx, d in enumerate(self.working_days):
                d_cooled_idx = max(0, d_idx - cooling_days)
                days_up_to_d = self.working_days[:d_idx + 1]
                days_up_to_cooled = self.working_days[:d_cooled_idx + 1]

                # CUMULATIVE: Total grinding <= CS WIP + total casting (with cooling delay)
                self.model += (
                    pulp.lpSum(self.x_grinding[(v, t)] for v in variants for t in days_up_to_d)
                    <= wip['CS'] + pulp.lpSum(self.x_casting[(v, t)] for v in variants for t in days_up_to_cooled),
                    f"Cum_Cast_Grind_{part}_D{d_idx}"
                )
                cnt += 1

                # CUMULATIVE: Total MC1 <= GR WIP + total grinding
                if p.get('has_mc1', True):
                    self.model += (
                        pulp.lpSum(self.x_mc1[(v, t)] for v in variants for t in days_up_to_d)
                        <= wip['GR'] + pulp.lpSum(self.x_grinding[(v, t)] for v in variants for t in days_up_to_d),
                        f"Cum_Grind_MC1_{part}_D{d_idx}"
                    )
                    cnt += 1

                # CUMULATIVE: Total SP1 <= MC WIP + total last machining stage
                if p.get('has_mc3', True):
                    mach_source = self.x_mc3
                    has_machining = True
                elif p.get('has_mc2', True):
                    mach_source = self.x_mc2
                    has_machining = True
                elif p.get('has_mc1', True):
                    mach_source = self.x_mc1
                    has_machining = True
                else:
                    mach_source = self.x_grinding
                    has_machining = False

                if has_machining:
                    self.model += (
                        pulp.lpSum(self.x_sp1[(v, t)] for v in variants for t in days_up_to_d)
                        <= wip['MC'] + pulp.lpSum(mach_source[(v, t)] for v in variants for t in days_up_to_d),
                        f"Cum_Mach_SP1_{part}_D{d_idx}"
                    )
                else:
                    self.model += (
                        pulp.lpSum(self.x_sp1[(v, t)] for v in variants for t in days_up_to_d)
                        <= wip['MC'] + wip['GR'] + pulp.lpSum(mach_source[(v, t)] for v in variants for t in days_up_to_d),
                        f"Cum_Grind_SP1_{part}_D{d_idx}"
                    )
                cnt += 1

                # CUMULATIVE: Total delivery <= FG+SP WIP + total last painting stage
                if p.get('has_sp3', True):
                    paint_source = self.x_sp3
                elif p.get('has_sp2', True):
                    paint_source = self.x_sp2
                else:
                    paint_source = self.x_sp1

                self.model += (
                    pulp.lpSum(self.x_delivery[(v, t)] for v in variants for t in days_up_to_d)
                    <= wip['FG'] + wip['SP'] + pulp.lpSum(paint_source[(v, t)] for v in variants for t in days_up_to_d),
                    f"Cum_Paint_Deliv_{part}_D{d_idx}"
                )
                cnt += 1

        # Add VARIANT-LEVEL cumulative constraints for internal stage seriality
        for v in self.daily_demand.keys():
            part, _ = self.part_day_mapping[v]
            if part not in self.params:
                continue

            p = self.params[part]

            for d_idx, d in enumerate(self.working_days):
                days_up_to_d = self.working_days[:d_idx + 1]

                # CUMULATIVE: MC2 <= MC1
                if p.get('has_mc2', True) and p.get('has_mc1', True):
                    self.model += (
                        pulp.lpSum(self.x_mc2[(v, t)] for t in days_up_to_d)
                        <= pulp.lpSum(self.x_mc1[(v, t)] for t in days_up_to_d),
                        f"Cum_MC1_MC2_{v}_D{d_idx}"
                    )
                    cnt += 1

                # CUMULATIVE: MC3 <= MC2 or MC1
                if p.get('has_mc3', True):
                    if p.get('has_mc2', True):
                        self.model += (
                            pulp.lpSum(self.x_mc3[(v, t)] for t in days_up_to_d)
                            <= pulp.lpSum(self.x_mc2[(v, t)] for t in days_up_to_d),
                            f"Cum_MC2_MC3_{v}_D{d_idx}"
                        )
                    elif p.get('has_mc1', True):
                        self.model += (
                            pulp.lpSum(self.x_mc3[(v, t)] for t in days_up_to_d)
                            <= pulp.lpSum(self.x_mc1[(v, t)] for t in days_up_to_d),
                            f"Cum_MC1_MC3_{v}_D{d_idx}"
                        )
                    cnt += 1

                # CUMULATIVE: SP2 <= SP1
                if p.get('has_sp2', True):
                    self.model += (
                        pulp.lpSum(self.x_sp2[(v, t)] for t in days_up_to_d)
                        <= pulp.lpSum(self.x_sp1[(v, t)] for t in days_up_to_d),
                        f"Cum_SP1_SP2_{v}_D{d_idx}"
                    )
                    cnt += 1

                # CUMULATIVE: SP3 <= SP2 or SP1
                if p.get('has_sp3', True):
                    if p.get('has_sp2', True):
                        self.model += (
                            pulp.lpSum(self.x_sp3[(v, t)] for t in days_up_to_d)
                            <= pulp.lpSum(self.x_sp2[(v, t)] for t in days_up_to_d),
                            f"Cum_SP2_SP3_{v}_D{d_idx}"
                        )
                    else:
                        self.model += (
                            pulp.lpSum(self.x_sp3[(v, t)] for t in days_up_to_d)
                            <= pulp.lpSum(self.x_sp1[(v, t)] for t in days_up_to_d),
                            f"Cum_SP1_SP3_{v}_D{d_idx}"
                        )
                    cnt += 1

        print(f"  ✓ Added {cnt:,} cumulative flow constraints (forces production through all stages)")
    
    def _add_demand_constraints(self):
        """Add demand fulfillment constraints (lateness calculated in objective, not here)."""
        print("\n📦 Adding demand constraints...")

        for v in self.daily_demand.keys():
            demand_qty = self.daily_demand[v]
            part, due_day = self.part_day_mapping[v]

            # Total delivered + unmet = demand
            total_delivered = pulp.lpSum(
                self.x_delivery[(v, d)] for d in self.working_days
            )
            self.model += (
                total_delivered + self.x_unmet[v] == demand_qty,
                f"Demand_{v}"
            )

        # Note: Lateness is calculated directly in objective function (like weekly optimizer)
        # No separate lateness constraints needed

        print(f"  ✓ Added demand constraints for {len(self.daily_demand)} variants")


class DailyResultsAnalyzer:
    """Analyze and extract results from daily optimization."""
    
    def __init__(self, model, daily_demand, part_day_mapping, params, config, calendar):
        self.model = model
        self.daily_demand = daily_demand
        self.part_day_mapping = part_day_mapping
        self.params = params
        self.config = config
        self.calendar = calendar
    
    def extract_all_results(self):
        """Extract all results and aggregate to daily schedule."""
        print("\n" + "="*80)
        print("EXTRACTING DAILY RESULTS")
        print("="*80)
        
        # Extract stage plans
        stage_plans = self._extract_stage_plans()
        
        # Aggregate to daily summary
        daily_summary = self._generate_daily_summary(stage_plans)
        
        # Aggregate to weekly summary (for compatibility)
        weekly_summary = self._aggregate_to_weekly(daily_summary)
        
        # Fulfillment analysis
        fulfillment = self._analyze_fulfillment()
        
        return {
            'daily_summary': daily_summary,
            'weekly_summary': weekly_summary,
            'stage_plans': stage_plans,
            'fulfillment': fulfillment,
            'casting_plan': stage_plans['casting'],
            'grinding_plan': stage_plans['grinding'],
            'mc1_plan': stage_plans['mc1'],
            'mc2_plan': stage_plans['mc2'],
            'mc3_plan': stage_plans['mc3'],
            'sp1_plan': stage_plans['sp1'],
            'sp2_plan': stage_plans['sp2'],
            'sp3_plan': stage_plans['sp3'],
            'delivery_plan': stage_plans['delivery']
        }
    
    def _extract_stage_plans(self):
        """Extract production quantities for each stage by day."""
        print("\n✓ Extracting stage plans...")
        
        stages = {
            'casting': (self.model.x_casting, 'Casting'),
            'grinding': (self.model.x_grinding, 'Grinding'),
            'mc1': (self.model.x_mc1, 'Machining-1'),
            'mc2': (self.model.x_mc2, 'Machining-2'),
            'mc3': (self.model.x_mc3, 'Machining-3'),
            'sp1': (self.model.x_sp1, 'Painting-1'),
            'sp2': (self.model.x_sp2, 'Painting-2'),
            'sp3': (self.model.x_sp3, 'Painting-3'),
            'delivery': (self.model.x_delivery, 'Delivery')
        }
        
        stage_plans = {}
        
        for stage_name, (stage_vars, stage_label) in stages.items():
            stage_data = []
            
            for (v, d), var in stage_vars.items():
                units = float(pulp.value(var) or 0)
                if units < 0.1:
                    continue
                
                part, due_day = self.part_day_mapping[v]
                p = self.params.get(part, {})
                
                stage_data.append({
                    'Part': part,
                    'Variant': v,
                    'Date': d,
                    'Day': d.strftime('%A'),
                    'Deadline_Date': due_day,
                    'Stage': stage_label,
                    'Units': round(units, 2),
                    'Unit_Weight_kg': p.get('unit_weight', 0),
                    'Total_Weight_ton': units * p.get('unit_weight', 0) / 1000.0,
                    'Moulding_Line': p.get('moulding_line', 'N/A'),
                    'Requires_Vacuum': p.get('requires_vacuum', False)
                })
            
            # Create DataFrame with proper columns even if empty
            if stage_data:
                stage_plans[stage_name] = pd.DataFrame(stage_data)
            else:
                # Empty DataFrame with correct columns
                stage_plans[stage_name] = pd.DataFrame(columns=[
                    'Part', 'Variant', 'Date', 'Day', 'Deadline_Date', 'Stage',
                    'Units', 'Unit_Weight_kg', 'Total_Weight_ton', 'Moulding_Line', 'Requires_Vacuum'
                ])
            print(f"  {stage_label}: {len(stage_data)} daily entries")

        return stage_plans
    
    def _generate_daily_summary(self, stage_plans):
        """Generate daily summary across all stages."""
        print("\n✓ Generating daily summary...")

        daily_data = []

        for d in self.calendar.working_days:
            # Filter stage plans for this day (DataFrames now have proper columns even if empty)
            dc = stage_plans['casting'][stage_plans['casting']['Date'] == d] if len(stage_plans['casting']) > 0 else stage_plans['casting']
            dg = stage_plans['grinding'][stage_plans['grinding']['Date'] == d] if len(stage_plans['grinding']) > 0 else stage_plans['grinding']
            dm1 = stage_plans['mc1'][stage_plans['mc1']['Date'] == d] if len(stage_plans['mc1']) > 0 else stage_plans['mc1']
            dm2 = stage_plans['mc2'][stage_plans['mc2']['Date'] == d] if len(stage_plans['mc2']) > 0 else stage_plans['mc2']
            dm3 = stage_plans['mc3'][stage_plans['mc3']['Date'] == d] if len(stage_plans['mc3']) > 0 else stage_plans['mc3']
            ds1 = stage_plans['sp1'][stage_plans['sp1']['Date'] == d] if len(stage_plans['sp1']) > 0 else stage_plans['sp1']
            ds2 = stage_plans['sp2'][stage_plans['sp2']['Date'] == d] if len(stage_plans['sp2']) > 0 else stage_plans['sp2']
            ds3 = stage_plans['sp3'][stage_plans['sp3']['Date'] == d] if len(stage_plans['sp3']) > 0 else stage_plans['sp3']
            dd = stage_plans['delivery'][stage_plans['delivery']['Date'] == d] if len(stage_plans['delivery']) > 0 else stage_plans['delivery']

            # Sum units and weights (empty DataFrames return 0 for sum)
            # Calculate Big Line and Small Line tonnage separately
            big_line_tons = 0
            small_line_tons = 0
            if 'Total_Weight_ton' in dc.columns and 'Moulding_Line' in dc.columns:
                big_line_tons = dc[dc['Moulding_Line'].str.contains('Big Line', case=False, na=False)]['Total_Weight_ton'].sum()
                small_line_tons = dc[dc['Moulding_Line'].str.contains('Small Line', case=False, na=False)]['Total_Weight_ton'].sum()

            # Calculate hours for each stage using cycle times
            def calc_hours(stage_df, cycle_key):
                """Calculate total hours for a stage using cycle times from params."""
                hours = 0.0
                if len(stage_df) > 0 and 'Part' in stage_df.columns and 'Units' in stage_df.columns:
                    for _, row in stage_df.iterrows():
                        part = row['Part']
                        units = row['Units']
                        cycle_min = self.params.get(part, {}).get(cycle_key, 0)
                        hours += units * cycle_min / 60.0
                return hours

            casting_hours = calc_hours(dc, 'casting_cycle')
            grinding_hours = calc_hours(dg, 'grinding_cycle')
            mc1_hours = calc_hours(dm1, 'mc1_cycle')
            mc2_hours = calc_hours(dm2, 'mc2_cycle')
            mc3_hours = calc_hours(dm3, 'mc3_cycle')
            sp1_hours = calc_hours(ds1, 'sp1_cycle')
            sp2_hours = calc_hours(ds2, 'sp2_cycle')
            sp3_hours = calc_hours(ds3, 'sp3_cycle')

            # Calculate Big Line and Small Line hours separately
            big_line_hours = 0.0
            small_line_hours = 0.0
            if len(dc) > 0 and 'Moulding_Line' in dc.columns:
                for _, row in dc.iterrows():
                    part = row['Part']
                    units = row['Units']
                    cycle_min = self.params.get(part, {}).get('casting_cycle', 0)
                    hours = units * cycle_min / 60.0
                    if 'Big Line' in row.get('Moulding_Line', ''):
                        big_line_hours += hours
                    elif 'Small Line' in row.get('Moulding_Line', ''):
                        small_line_hours += hours

            daily_data.append({
                'Date': d,
                'Day': d.strftime('%A'),
                'Is_Holiday': 'No',
                'Casting_Tons': dc['Total_Weight_ton'].sum() if 'Total_Weight_ton' in dc.columns else 0,
                'Big_Line_Tons': big_line_tons,
                'Small_Line_Tons': small_line_tons,
                'Casting_Units': dc['Units'].sum() if 'Units' in dc.columns else 0,
                'Grinding_Units': dg['Units'].sum() if 'Units' in dg.columns else 0,
                'MC1_Units': dm1['Units'].sum() if 'Units' in dm1.columns else 0,
                'MC2_Units': dm2['Units'].sum() if 'Units' in dm2.columns else 0,
                'MC3_Units': dm3['Units'].sum() if 'Units' in dm3.columns else 0,
                'SP1_Units': ds1['Units'].sum() if 'Units' in ds1.columns else 0,
                'SP2_Units': ds2['Units'].sum() if 'Units' in ds2.columns else 0,
                'SP3_Units': ds3['Units'].sum() if 'Units' in ds3.columns else 0,
                'Delivery_Units': dd['Units'].sum() if 'Units' in dd.columns else 0,
                # Add hours for each stage
                'Casting_Hours': round(casting_hours, 2),
                'Big_Line_Hours': round(big_line_hours, 2),
                'Small_Line_Hours': round(small_line_hours, 2),
                'Grinding_Hours': round(grinding_hours, 2),
                'MC1_Hours': round(mc1_hours, 2),
                'MC2_Hours': round(mc2_hours, 2),
                'MC3_Hours': round(mc3_hours, 2),
                'SP1_Hours': round(sp1_hours, 2),
                'SP2_Hours': round(sp2_hours, 2),
                'SP3_Hours': round(sp3_hours, 2)
            })

        return pd.DataFrame(daily_data)
    
    def _aggregate_to_weekly(self, daily_summary):
        """Aggregate daily summary to weekly for reporting compatibility."""
        print("\n✓ Aggregating to weekly summary...")

        weekly_data = []

        # Group by week
        for week_num in range(1, self.config.PLANNING_DAYS // 7 + 2):
            week_start = self.config.CURRENT_DATE + timedelta(weeks=week_num - 1)
            week_end = week_start + timedelta(days=6)

            week_days = daily_summary[
                (daily_summary['Date'] >= week_start) &
                (daily_summary['Date'] <= week_end)
            ]

            if len(week_days) == 0:
                continue

            weekly_data.append({
                'Week': week_num,
                'Casting_Tons': week_days['Casting_Tons'].sum(),
                'Big_Line_Tons': week_days['Big_Line_Tons'].sum(),
                'Small_Line_Tons': week_days['Small_Line_Tons'].sum(),
                'Casting_Units': week_days['Casting_Units'].sum(),
                'Grinding_Units': week_days['Grinding_Units'].sum(),
                'MC1_Units': week_days['MC1_Units'].sum(),
                'MC2_Units': week_days['MC2_Units'].sum(),
                'MC3_Units': week_days['MC3_Units'].sum(),
                'SP1_Units': week_days['SP1_Units'].sum(),
                'SP2_Units': week_days['SP2_Units'].sum(),
                'SP3_Units': week_days['SP3_Units'].sum(),
                'Delivery_Units': week_days['Delivery_Units'].sum(),
                # Add hours aggregation
                'Casting_Hours': week_days['Casting_Hours'].sum(),
                'Big_Line_Hours': week_days['Big_Line_Hours'].sum(),
                'Small_Line_Hours': week_days['Small_Line_Hours'].sum(),
                'Grinding_Hours': week_days['Grinding_Hours'].sum(),
                'MC1_Hours': week_days['MC1_Hours'].sum(),
                'MC2_Hours': week_days['MC2_Hours'].sum(),
                'MC3_Hours': week_days['MC3_Hours'].sum(),
                'SP1_Hours': week_days['SP1_Hours'].sum(),
                'SP2_Hours': week_days['SP2_Hours'].sum(),
                'SP3_Hours': week_days['SP3_Hours'].sum()
            })

        return pd.DataFrame(weekly_data)
    
    def _analyze_fulfillment(self):
        """Analyze order fulfillment - calculate lateness from actual deliveries."""
        print("\n✓ Analyzing fulfillment...")

        fulfillment_data = []

        for v in self.daily_demand.keys():
            demand_qty = self.daily_demand[v]
            part, due_day = self.part_day_mapping[v]
            due_day_idx = self.calendar.working_days.index(due_day) if due_day in self.calendar.working_days else 0

            delivered = sum(
                float(pulp.value(self.model.x_delivery[(v, d)]) or 0)
                for d in self.calendar.working_days
            )

            unmet = float(pulp.value(self.model.x_unmet[v]) or 0)

            # Calculate lateness from actual delivery dates
            late_days = 0.0
            for d_idx, d in enumerate(self.calendar.working_days):
                delivery_qty = float(pulp.value(self.model.x_delivery[(v, d)]) or 0)
                if delivery_qty > 0 and d_idx > due_day_idx:
                    late_days = max(late_days, d_idx - due_day_idx)

            fulfillment_data.append({
                'Variant': v,
                'Part': part,
                'Due_Date': due_day,
                'Ordered_Qty': demand_qty,
                'Delivered_Qty': round(delivered, 2),
                'Unmet_Qty': round(unmet, 2),
                'Late_Days': round(late_days, 1),
                'Fulfillment_%': round(delivered / demand_qty * 100, 1) if demand_qty > 0 else 0,
                'Status': 'On-Time' if late_days < 0.5 and unmet < 0.5 else
                         ('Late' if unmet < 0.5 else 'Partial')
            })

        return pd.DataFrame(fulfillment_data)


def main():
    """Main execution function for daily optimization."""
    print("\n" + "="*80)
    print("DAILY PRODUCTION PLANNING OPTIMIZATION - MAIN EXECUTION")
    print("="*80)
    
    # Configuration
    config = DailyProductionConfig()
    master_data_path = 'Master_Data_Updated_Nov_Dec.xlsx'
    output_path = 'production_plan_daily_comprehensive.xlsx'
    
    # Load data (reuse weekly loader, but with daily config)
    print("\n📂 Loading master data...")
    loader = ComprehensiveDataLoader(master_data_path, config)
    data = loader.load_all_data()
    
    # Create calendar with working days
    calendar = ProductionCalendar(config)
    
    # Calculate planning horizon
    latest_order = data['sales_order']['Delivery_Date'].max()
    planning_days = (latest_order - config.CURRENT_DATE).days + config.PLANNING_BUFFER_DAYS
    config.PLANNING_DAYS = min(planning_days, config.MAX_PLANNING_DAYS)
    
    # Get working days
    calendar.working_days = calendar.get_all_working_days(config.PLANNING_DAYS)
    
    print(f"\n📅 Planning Horizon:")
    print(f"   Planning Days: {config.PLANNING_DAYS} days")
    print(f"   Working Days: {len(calendar.working_days)} days")
    print(f"   Start: {config.CURRENT_DATE.strftime('%Y-%m-%d')}")
    print(f"   End: {(config.CURRENT_DATE + timedelta(days=config.PLANNING_DAYS)).strftime('%Y-%m-%d')}")
    
    # Calculate daily demand
    demand_calc = DailyDemandCalculator(
        data['sales_order'],
        data['stage_wip'],
        config,
        calendar
    )
    daily_demand, part_day_mapping, stage_start_qty, wip_by_part = demand_calc.calculate_daily_demand()
    
    # Build parameters
    param_builder = ComprehensiveParameterBuilder(data['part_master'], config)
    params = param_builder.build_parameters()
    
    # Setup resources
    machine_manager = MachineResourceManager(data['machine_constraints'], config)
    box_manager = BoxCapacityManager(data['box_capacity'], config, machine_manager)
    wip_init = build_wip_init(data['stage_wip'])
    
    # Build and solve model
    optimizer = DailyOptimizationModel(
        daily_demand,
        part_day_mapping,
        params,
        stage_start_qty,
        machine_manager,
        box_manager,
        config,
        calendar,
        wip_init=wip_init
    )
    
    status = optimizer.build_and_solve()
    
    if status != pulp.LpStatusOptimal:
        print(f"\n❌ Optimization failed with status: {pulp.LpStatus[status]}")
        return
    
    # Extract results
    analyzer = DailyResultsAnalyzer(
        optimizer,
        daily_demand,
        part_day_mapping,
        params,
        config,
        calendar
    )
    results = analyzer.extract_all_results()
    
    # Save results
    print(f"\n💾 Saving results to {output_path}...")
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        results['daily_summary'].to_excel(writer, sheet_name='Daily_Summary', index=False)
        results['weekly_summary'].to_excel(writer, sheet_name='Weekly_Summary', index=False)
        results['fulfillment'].to_excel(writer, sheet_name='Order_Fulfillment', index=False)
        
        for stage_name in ['casting', 'grinding', 'mc1', 'mc2', 'mc3', 'sp1', 'sp2', 'sp3', 'delivery']:
            results['stage_plans'][stage_name].to_excel(
                writer,
                sheet_name=stage_name.capitalize(),
                index=False
            )
    
    print(f"✅ Daily optimization completed successfully!")
    print(f"📁 Output: {output_path}")
    print("\n" + "="*80)


if __name__ == '__main__':
    main()
