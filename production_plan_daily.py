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

        # Compatibility attributes for weekly data loader
        self.PLANNING_BUFFER_WEEKS = int(self.PLANNING_BUFFER_DAYS / 7)  # ~2 weeks
        self.MAX_PLANNING_WEEKS = int(self.MAX_PLANNING_DAYS / 7)  # ~30 weeks

        # Working schedule
        self.WORKING_DAYS_PER_WEEK = 6
        self.WEEKLY_OFF_DAY = 6  # Sunday (0=Monday, 6=Sunday)
        
        # Machine configuration
        self.OEE = 0.90
        self.HOURS_PER_SHIFT = 12
        self.SHIFTS_PER_DAY = 2
        
        # Daily capacities (derived from weekly / 6)
        self.BIG_LINE_HOURS_PER_DAY = self.HOURS_PER_SHIFT * self.SHIFTS_PER_DAY * self.OEE
        self.SMALL_LINE_HOURS_PER_DAY = self.HOURS_PER_SHIFT * self.SHIFTS_PER_DAY * self.OEE
        
        # Lead times (in days)
        self.DEFAULT_COOLING_DAYS = 1  # Casting -> Grinding
        self.DEFAULT_GRINDING_TO_MC1_DAYS = 1
        self.DEFAULT_MC_STAGE_DAYS = 1  # Between MC stages
        self.DEFAULT_PAINT_DRYING_DAYS = 1  # Between paint stages
        
        # Penalties
        self.UNMET_DEMAND_PENALTY = 200000
        self.LATENESS_PENALTY_PER_DAY = 5000  # Per day late (was 150k/week)
        self.INVENTORY_HOLDING_COST_PER_DAY = 0.15  # Per unit per day (was 1/week)
        self.SETUP_PENALTY = 5
        
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
        self.x_late_days = {}
        
        # Inventory tracking
        self.inv_cs = {}
        self.inv_gr = {}
        self.inv_mc = {}
        self.inv_sp = {}
        self.inv_fg = {}
    
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
        self._add_flow_constraints()
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
        """Create all decision variables."""
        print("\n📊 Creating decision variables...")
        
        variants = list(self.daily_demand.keys())
        days = self.working_days
        
        # Production variables for each stage
        for v in variants:
            for d in days:
                self.x_casting[(v, d)] = pulp.LpVariable(f"cast_{v}_{d.strftime('%Y%m%d')}", 0, None)
                self.x_grinding[(v, d)] = pulp.LpVariable(f"grind_{v}_{d.strftime('%Y%m%d')}", 0, None)
                self.x_mc1[(v, d)] = pulp.LpVariable(f"mc1_{v}_{d.strftime('%Y%m%d')}", 0, None)
                self.x_mc2[(v, d)] = pulp.LpVariable(f"mc2_{v}_{d.strftime('%Y%m%d')}", 0, None)
                self.x_mc3[(v, d)] = pulp.LpVariable(f"mc3_{v}_{d.strftime('%Y%m%d')}", 0, None)
                self.x_sp1[(v, d)] = pulp.LpVariable(f"sp1_{v}_{d.strftime('%Y%m%d')}", 0, None)
                self.x_sp2[(v, d)] = pulp.LpVariable(f"sp2_{v}_{d.strftime('%Y%m%d')}", 0, None)
                self.x_sp3[(v, d)] = pulp.LpVariable(f"sp3_{v}_{d.strftime('%Y%m%d')}", 0, None)
                self.x_delivery[(v, d)] = pulp.LpVariable(f"deliv_{v}_{d.strftime('%Y%m%d')}", 0, None)
        
        # Unmet demand and lateness
        for v in variants:
            self.x_unmet[v] = pulp.LpVariable(f"unmet_{v}", 0, None)
            self.x_late_days[v] = pulp.LpVariable(f"late_{v}", 0, None)
        
        # Inventory variables
        for v in variants:
            part, _ = self.part_day_mapping[v]
            for d in days:
                self.inv_cs[(part, d)] = pulp.LpVariable(f"inv_cs_{part}_{d.strftime('%Y%m%d')}", 0, None)
                self.inv_gr[(part, d)] = pulp.LpVariable(f"inv_gr_{part}_{d.strftime('%Y%m%d')}", 0, None)
                self.inv_mc[(part, d)] = pulp.LpVariable(f"inv_mc_{part}_{d.strftime('%Y%m%d')}", 0, None)
                self.inv_sp[(part, d)] = pulp.LpVariable(f"inv_sp_{part}_{d.strftime('%Y%m%d')}", 0, None)
                self.inv_fg[(part, d)] = pulp.LpVariable(f"inv_fg_{part}_{d.strftime('%Y%m%d')}", 0, None)
        
        print(f"  ✓ Created {len(self.model.variables())} decision variables")
    
    def _add_objective_function(self):
        """Add objective: minimize cost."""
        print("\n🎯 Adding objective function...")
        
        # Unmet demand penalty
        unmet_cost = pulp.lpSum(
            self.config.UNMET_DEMAND_PENALTY * self.x_unmet[v]
            for v in self.daily_demand.keys()
        )
        
        # Lateness penalty
        late_cost = pulp.lpSum(
            self.config.LATENESS_PENALTY_PER_DAY * self.x_late_days[v]
            for v in self.daily_demand.keys()
        )
        
        # Inventory holding cost
        inv_cost = pulp.lpSum(
            self.config.INVENTORY_HOLDING_COST_PER_DAY * (
                self.inv_cs[(part, d)] + self.inv_gr[(part, d)] + 
                self.inv_mc[(part, d)] + self.inv_sp[(part, d)] + self.inv_fg[(part, d)]
            )
            for part in set(p for p, _ in self.part_day_mapping.values())
            for d in self.working_days
        )
        
        self.model += unmet_cost + late_cost + inv_cost, "Total_Cost"
        print("  ✓ Objective function added")
    
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
                
                if usage_minutes > 0:
                    self.model += (
                        usage_minutes <= daily_minutes,
                        f"{resource_code}_Cap_{d.strftime('%Y%m%d')}"
                    )
        
        print(f"  ✓ Added {len(self.working_days) * 2} casting line capacity constraints")
        print(f"  ✓ Added {len(self.machine_manager.machines) * len(self.working_days)} machine capacity constraints")
    
    def _add_flow_constraints(self):
        """Add material flow constraints with daily lead times."""
        print("\n🔄 Adding flow constraints with daily lead times...")
        
        for v in self.daily_demand.keys():
            part, _ = self.part_day_mapping[v]
            p = self.params.get(part, {})
            
            # Get WIP initialization
            wip = self.wip_init.get(part, {})
            start_qty = self.stage_start_qty.get(v, {})
            
            # Determine which stages this variant needs
            need_casting = start_qty.get('casting', 0) > 0
            need_grinding = start_qty.get('grinding', 0) > 0
            need_mc1 = start_qty.get('mc1', 0) > 0
            need_mc2 = start_qty.get('mc2', 0) > 0
            need_mc3 = start_qty.get('mc3', 0) > 0
            need_sp1 = start_qty.get('sp1', 0) > 0
            need_sp2 = start_qty.get('sp2', 0) > 0
            need_sp3 = start_qty.get('sp3', 0) > 0
            
            # Lead times (in days)
            cooling_days = max(1, int(p.get('cooling_time', 0) / 24))  # Convert hours to days
            grinding_to_mc1_days = 1
            mc_stage_days = 1
            paint_drying_days = max(1, int(p.get('sp1_dry_time', 0) / 24))  # SP1->SP2 drying
            
            for d_idx, d in enumerate(self.working_days):
                # Casting -> Grinding (after cooling)
                if need_casting and need_grinding and d_idx >= cooling_days:
                    casting_day = self.working_days[d_idx - cooling_days]
                    self.model += (
                        self.x_grinding[(v, d)] <= self.x_casting[(v, casting_day)],
                        f"Flow_Cast_Grind_{v}_{d.strftime('%Y%m%d')}"
                    )
                
                # Grinding -> MC1
                if need_grinding and need_mc1 and d_idx >= grinding_to_mc1_days:
                    grinding_day = self.working_days[d_idx - grinding_to_mc1_days]
                    self.model += (
                        self.x_mc1[(v, d)] <= self.x_grinding[(v, grinding_day)],
                        f"Flow_Grind_MC1_{v}_{d.strftime('%Y%m%d')}"
                    )
                
                # MC1 -> MC2 -> MC3 (sequential)
                if need_mc1 and need_mc2 and d_idx >= mc_stage_days:
                    mc1_day = self.working_days[d_idx - mc_stage_days]
                    self.model += (
                        self.x_mc2[(v, d)] <= self.x_mc1[(v, mc1_day)],
                        f"Flow_MC1_MC2_{v}_{d.strftime('%Y%m%d')}"
                    )
                
                if need_mc2 and need_mc3 and d_idx >= mc_stage_days:
                    mc2_day = self.working_days[d_idx - mc_stage_days]
                    self.model += (
                        self.x_mc3[(v, d)] <= self.x_mc2[(v, mc2_day)],
                        f"Flow_MC2_MC3_{v}_{d.strftime('%Y%m%d')}"
                    )
                
                # MC3 or GR -> SP1 (depending on routing)
                source_stage = self.x_mc3[(v, d)] if need_mc3 else self.x_grinding[(v, d)]
                if need_sp1:
                    self.model += (
                        self.x_sp1[(v, d)] <= source_stage,
                        f"Flow_toSP1_{v}_{d.strftime('%Y%m%d')}"
                    )
                
                # SP1 -> SP2 -> SP3 (with drying time)
                if need_sp1 and need_sp2 and d_idx >= paint_drying_days:
                    sp1_day = self.working_days[d_idx - paint_drying_days]
                    self.model += (
                        self.x_sp2[(v, d)] <= self.x_sp1[(v, sp1_day)],
                        f"Flow_SP1_SP2_{v}_{d.strftime('%Y%m%d')}"
                    )
                
                if need_sp2 and need_sp3 and d_idx >= paint_drying_days:
                    sp2_day = self.working_days[d_idx - paint_drying_days]
                    self.model += (
                        self.x_sp3[(v, d)] <= self.x_sp2[(v, sp2_day)],
                        f"Flow_SP2_SP3_{v}_{d.strftime('%Y%m%d')}"
                    )
                
                # Last stage -> Delivery
                final_stage = self.x_sp3[(v, d)] if need_sp3 else (
                    self.x_mc3[(v, d)] if need_mc3 else self.x_grinding[(v, d)]
                )
                self.model += (
                    self.x_delivery[(v, d)] <= final_stage,
                    f"Flow_toDel_{v}_{d.strftime('%Y%m%d')}"
                )
        
        print("  ✓ Added flow constraints for all stages")
    
    def _add_demand_constraints(self):
        """Add demand fulfillment and delivery timing constraints."""
        print("\n📦 Adding demand and delivery constraints...")
        
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
            
            # Calculate lateness (days late)
            due_day_idx = self.day_index[due_day]
            
            for d_idx, d in enumerate(self.working_days):
                if d_idx > due_day_idx:
                    days_late = d_idx - due_day_idx
                    self.model += (
                        self.x_late_days[v] >= self.x_delivery[(v, d)] * days_late / demand_qty,
                        f"Late_{v}_{d.strftime('%Y%m%d')}"
                    )
        
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
            
            stage_plans[stage_name] = pd.DataFrame(stage_data)
            print(f"  {stage_label}: {len(stage_data)} daily entries")
        
        return stage_plans
    
    def _generate_daily_summary(self, stage_plans):
        """Generate daily summary across all stages."""
        print("\n✓ Generating daily summary...")
        
        daily_data = []
        
        for d in self.calendar.working_days:
            dc = stage_plans['casting'][stage_plans['casting']['Date'] == d]
            dg = stage_plans['grinding'][stage_plans['grinding']['Date'] == d]
            dm1 = stage_plans['mc1'][stage_plans['mc1']['Date'] == d]
            dm2 = stage_plans['mc2'][stage_plans['mc2']['Date'] == d]
            dm3 = stage_plans['mc3'][stage_plans['mc3']['Date'] == d]
            ds1 = stage_plans['sp1'][stage_plans['sp1']['Date'] == d]
            ds2 = stage_plans['sp2'][stage_plans['sp2']['Date'] == d]
            ds3 = stage_plans['sp3'][stage_plans['sp3']['Date'] == d]
            dd = stage_plans['delivery'][stage_plans['delivery']['Date'] == d]
            
            daily_data.append({
                'Date': d,
                'Day': d.strftime('%A'),
                'Is_Holiday': 'No',
                'Casting_Tons': dc['Total_Weight_ton'].sum(),
                'Casting_Units': dc['Units'].sum(),
                'Grinding_Units': dg['Units'].sum(),
                'MC1_Units': dm1['Units'].sum(),
                'MC2_Units': dm2['Units'].sum(),
                'MC3_Units': dm3['Units'].sum(),
                'SP1_Units': ds1['Units'].sum(),
                'SP2_Units': ds2['Units'].sum(),
                'SP3_Units': ds3['Units'].sum(),
                'Delivery_Units': dd['Units'].sum()
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
                'Casting_Units': week_days['Casting_Units'].sum(),
                'Grinding_Units': week_days['Grinding_Units'].sum(),
                'MC1_Units': week_days['MC1_Units'].sum(),
                'MC2_Units': week_days['MC2_Units'].sum(),
                'MC3_Units': week_days['MC3_Units'].sum(),
                'SP1_Units': week_days['SP1_Units'].sum(),
                'SP2_Units': week_days['SP2_Units'].sum(),
                'SP3_Units': week_days['SP3_Units'].sum(),
                'Delivery_Units': week_days['Delivery_Units'].sum()
            })
        
        return pd.DataFrame(weekly_data)
    
    def _analyze_fulfillment(self):
        """Analyze order fulfillment."""
        print("\n✓ Analyzing fulfillment...")
        
        fulfillment_data = []
        
        for v in self.daily_demand.keys():
            demand_qty = self.daily_demand[v]
            part, due_day = self.part_day_mapping[v]
            
            delivered = sum(
                float(pulp.value(self.model.x_delivery[(v, d)]) or 0)
                for d in self.calendar.working_days
            )
            
            unmet = float(pulp.value(self.model.x_unmet[v]) or 0)
            late_days = float(pulp.value(self.model.x_late_days[v]) or 0)
            
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
