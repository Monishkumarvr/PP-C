"""
Infeasibility Diagnostic Tool for Production Planning
====================================================

This script diagnoses why the production planning optimization is infeasible.
It checks:
1. Capacity vs Demand mismatches
2. Lead time vs Delivery date conflicts
3. WIP mapping issues
4. Box capacity constraints
5. Seriality constraint conflicts
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from collections import defaultdict
import warnings
warnings.filterwarnings('ignore')

class InfeasibilityDiagnostics:
    """Comprehensive diagnostics for optimization infeasibility"""

    def __init__(self, master_file='Master_Data_Updated_Nov_Dec_FIXED_V2.xlsx'):
        self.master_file = master_file
        self.issues = []
        self.warnings_list = []

        # Configuration
        self.CURRENT_DATE = datetime(2025, 11, 22)
        self.OEE = 0.90
        self.WORKING_DAYS_PER_WEEK = 6

        print("="*80)
        print("INFEASIBILITY DIAGNOSTICS")
        print("="*80)
        print()

    def load_data(self):
        """Load all required data"""
        print("📂 Loading master data...")

        try:
            self.part_master = pd.read_excel(self.master_file, sheet_name='Part Master')
            self.sales_orders = pd.read_excel(self.master_file, sheet_name='Sales Order')
            self.machine_constraints = pd.read_excel(self.master_file, sheet_name='Machine Constraints')
            self.stage_wip = pd.read_excel(self.master_file, sheet_name='Stage WIP')
            self.box_capacity = pd.read_excel(self.master_file, sheet_name='Mould Box Capacity')

            print(f"  ✓ Part Master: {len(self.part_master)} parts")
            print(f"  ✓ Sales Orders: {len(self.sales_orders)} orders")
            print(f"  ✓ Machine Constraints: {len(self.machine_constraints)} resources")
            print(f"  ✓ Stage WIP: {len(self.stage_wip)} items")
            print(f"  ✓ Box Capacity: {len(self.box_capacity)} box sizes")
            print()

        except Exception as e:
            print(f"❌ Error loading data: {e}")
            raise

    def check_order_validity(self):
        """Check if all ordered parts exist in Part Master"""
        print("="*80)
        print("1. ORDER VALIDITY CHECK")
        print("="*80)

        valid_parts = set(self.part_master['FG Code'].dropna().astype(str))
        ordered_parts = set(self.sales_orders['Material Code'].dropna().astype(str))

        invalid_parts = ordered_parts - valid_parts

        if invalid_parts:
            self.issues.append({
                'severity': 'CRITICAL',
                'category': 'Order Validity',
                'issue': f'{len(invalid_parts)} parts in orders NOT in Part Master',
                'details': list(invalid_parts)
            })
            print(f"❌ CRITICAL: {len(invalid_parts)} invalid parts found:")
            for part in sorted(invalid_parts)[:10]:
                orders = self.sales_orders[self.sales_orders['Material Code'] == part]
                print(f"    • {part}: {orders['Balance Qty'].sum():.0f} units")
            if len(invalid_parts) > 10:
                print(f"    ... and {len(invalid_parts) - 10} more")
        else:
            print("✓ All ordered parts exist in Part Master")

        print()

    def check_capacity_vs_demand(self):
        """Check if capacity is sufficient for demand"""
        print("="*80)
        print("2. CAPACITY vs DEMAND ANALYSIS")
        print("="*80)

        # Filter valid orders
        valid_parts = set(self.part_master['FG Code'].dropna().astype(str))
        valid_orders = self.sales_orders[
            self.sales_orders['Material Code'].isin(valid_parts)
        ].copy()

        # Calculate planning horizon
        valid_orders['Comitted Delivery Date'] = pd.to_datetime(
            valid_orders['Comitted Delivery Date'], errors='coerce'
        )
        valid_orders = valid_orders.dropna(subset=['Comitted Delivery Date'])

        latest_delivery = valid_orders['Comitted Delivery Date'].max()
        planning_days = (latest_delivery - self.CURRENT_DATE).days + 14  # 2 week buffer

        print(f"📅 Planning horizon: {planning_days} days ({planning_days/7:.1f} weeks)")
        print(f"   Start: {self.CURRENT_DATE.strftime('%Y-%m-%d')}")
        print(f"   End: {latest_delivery.strftime('%Y-%m-%d')}")
        print()

        # Calculate total demand by stage
        demand_by_stage = self._calculate_demand_by_stage(valid_orders)

        # Calculate available capacity by stage
        capacity_by_stage = self._calculate_available_capacity(planning_days)

        # Compare
        print("Stage-wise Capacity Analysis:")
        print("-" * 80)
        print(f"{'Stage':<20} {'Demand (hrs)':<15} {'Capacity (hrs)':<15} {'Utilization':<12} {'Status'}")
        print("-" * 80)

        for stage in ['Casting', 'Grinding', 'MC1', 'MC2', 'MC3', 'SP1', 'SP2', 'SP3']:
            demand_hrs = demand_by_stage.get(stage, 0)
            capacity_hrs = capacity_by_stage.get(stage, 0)

            if capacity_hrs > 0:
                utilization = (demand_hrs / capacity_hrs) * 100
                status = "✓" if utilization <= 100 else "❌ OVERFLOW"

                if utilization > 100:
                    self.issues.append({
                        'severity': 'CRITICAL',
                        'category': 'Capacity',
                        'issue': f'{stage} capacity exceeded',
                        'details': f'Need {demand_hrs:.0f} hrs, have {capacity_hrs:.0f} hrs ({utilization:.1f}%)'
                    })
                elif utilization > 90:
                    self.warnings_list.append({
                        'severity': 'WARNING',
                        'category': 'Capacity',
                        'issue': f'{stage} near capacity limit',
                        'details': f'{utilization:.1f}% utilization'
                    })

                print(f"{stage:<20} {demand_hrs:<15.0f} {capacity_hrs:<15.0f} {utilization:>10.1f}%  {status}")
            else:
                print(f"{stage:<20} {demand_hrs:<15.0f} {'N/A':<15} {'N/A':<12}  ⚠️")
                if demand_hrs > 0:
                    self.issues.append({
                        'severity': 'CRITICAL',
                        'category': 'Capacity',
                        'issue': f'{stage} has demand but no capacity',
                        'details': f'Need {demand_hrs:.0f} hrs but no resources defined'
                    })

        print("-" * 80)
        print()

    def _calculate_demand_by_stage(self, orders):
        """Calculate total hours needed by stage"""
        demand = defaultdict(float)

        # Check for duplicates first
        fg_codes = self.part_master['FG Code'].dropna()
        if fg_codes.duplicated().any():
            duplicates = fg_codes[fg_codes.duplicated()].unique()
            print(f"⚠️  WARNING: Found {len(duplicates)} duplicate FG Codes in Part Master")
            print(f"   Using first occurrence for each duplicate")
            # Keep first occurrence only
            part_master_dedup = self.part_master.drop_duplicates(subset=['FG Code'], keep='first')
            part_master_dict = part_master_dedup.set_index('FG Code').to_dict('index')
        else:
            part_master_dict = self.part_master.set_index('FG Code').to_dict('index')

        for _, order in orders.iterrows():
            part = order['Material Code']
            qty = order['Balance Qty']

            if part not in part_master_dict:
                continue

            pm = part_master_dict[part]

            # Casting
            casting_time = self._safe_float(pm.get('Casting Cycle time (min)', 0))
            if casting_time > 0:
                demand['Casting'] += (qty * casting_time) / 60

            # Grinding
            grinding_time = self._safe_float(pm.get('Grinding Cycle time (min)', 0))
            if grinding_time > 0:
                demand['Grinding'] += (qty * grinding_time) / 60

            # Machining
            for i in range(1, 4):
                mc_time = self._safe_float(pm.get(f'Machining Cycle time {i} (min)', 0))
                if mc_time > 0:
                    demand[f'MC{i}'] += (qty * mc_time) / 60

            # Painting
            for i in range(1, 4):
                sp_time = self._safe_float(pm.get(f'Painting Cycle time {i} (min)', 0))
                if sp_time > 0:
                    demand[f'SP{i}'] += (qty * sp_time) / 60

        return demand

    def _calculate_available_capacity(self, planning_days):
        """Calculate total available hours by stage"""
        capacity = defaultdict(float)

        working_days = int(planning_days * (self.WORKING_DAYS_PER_WEEK / 7))

        for _, machine in self.machine_constraints.iterrows():
            operation = str(machine.get('Operation Name', '')).strip()
            num_resources = self._safe_int(machine.get('No Of Resource', 1))
            hrs_per_day = self._safe_float(machine.get('Available Hours per Day', 8))
            num_shifts = self._safe_int(machine.get('No of Shift', 1))

            daily_hrs = num_resources * hrs_per_day * num_shifts * self.OEE
            total_hrs = daily_hrs * working_days

            # Map operation to stage
            if 'Casting' in operation or 'Moulding' in operation:
                capacity['Casting'] += total_hrs
            elif 'Grinding' in operation:
                capacity['Grinding'] += total_hrs
            elif 'Machining' in operation:
                # Distribute across MC stages
                for i in range(1, 4):
                    capacity[f'MC{i}'] += total_hrs / 3
            elif 'Painting' in operation or 'SP' in operation:
                # Distribute across SP stages
                for i in range(1, 4):
                    capacity[f'SP{i}'] += total_hrs / 3

        return capacity

    def check_lead_time_feasibility(self):
        """Check if lead times allow meeting delivery dates"""
        print("="*80)
        print("3. LEAD TIME vs DELIVERY DATE ANALYSIS")
        print("="*80)

        valid_parts = set(self.part_master['FG Code'].dropna().astype(str))
        valid_orders = self.sales_orders[
            self.sales_orders['Material Code'].isin(valid_parts)
        ].copy()

        valid_orders['Comitted Delivery Date'] = pd.to_datetime(
            valid_orders['Comitted Delivery Date'], errors='coerce'
        )
        valid_orders = valid_orders.dropna(subset=['Comitted Delivery Date'])

        # Handle duplicates
        part_master_dedup = self.part_master.drop_duplicates(subset=['FG Code'], keep='first')
        part_master_dict = part_master_dedup.set_index('FG Code').to_dict('index')

        backlog_orders = []
        rush_orders = []

        for _, order in valid_orders.iterrows():
            part = order['Material Code']
            delivery_date = order['Comitted Delivery Date']

            if part not in part_master_dict:
                continue

            pm = part_master_dict[part]

            # Calculate minimum lead time (all active times + passive times)
            total_lead_days = 0

            # Casting
            casting_time = self._safe_float(pm.get('Casting Cycle time (min)', 0))
            cooling_time = self._safe_float(pm.get('Cooling Time (hrs)', 0))
            shakeout_time = self._safe_float(pm.get('Shakeout Time (hrs)', 0))
            vacuum_time = self._safe_float(pm.get('Vacuum Time (hrs)', 0))

            total_lead_days += (casting_time / 60 + cooling_time + shakeout_time + vacuum_time) / 24

            # Grinding
            grinding_time = self._safe_float(pm.get('Grinding Cycle time (min)', 0))
            total_lead_days += (grinding_time / 60) / 24

            # Machining (sequential)
            for i in range(1, 4):
                mc_time = self._safe_float(pm.get(f'Machining Cycle time {i} (min)', 0))
                total_lead_days += (mc_time / 60) / 24

            # Painting (sequential with dry times)
            for i in range(1, 4):
                sp_time = self._safe_float(pm.get(f'Painting Cycle time {i} (min)', 0))
                dry_time = self._safe_float(pm.get(f'Dry Time {i} (hrs)', 0))
                total_lead_days += (sp_time / 60 + dry_time) / 24

            # Add buffer for material handling, setup, etc.
            total_lead_days = total_lead_days * 1.2  # 20% buffer

            # Calculate available time
            available_days = (delivery_date - self.CURRENT_DATE).days

            if available_days < 0:
                # Backlog order (past due) - this is OK, will incur lateness penalty
                backlog_orders.append({
                    'Part': part,
                    'Qty': order['Balance Qty'],
                    'Delivery': delivery_date.strftime('%Y-%m-%d'),
                    'Days_Late': -available_days
                })
            elif available_days < total_lead_days:
                # Rush order - insufficient lead time but not past due yet
                rush_orders.append({
                    'Part': part,
                    'Qty': order['Balance Qty'],
                    'Delivery': delivery_date.strftime('%Y-%m-%d'),
                    'Available_Days': available_days,
                    'Required_Days': total_lead_days,
                    'Shortfall_Days': total_lead_days - available_days
                })

        # Report backlog orders (INFO - not an error)
        if backlog_orders:
            print(f"📊 Found {len(backlog_orders)} backlog orders (past due):")
            print("   → These will be produced ASAP and delivered late (with penalties)")
            print()
            print(f"   {'Part':<15} {'Qty':<8} {'Original Due':<15} {'Days Late'}")
            print("   " + "-" * 60)

            for order in backlog_orders[:10]:
                print(f"   {order['Part']:<15} {order['Qty']:<8.0f} {order['Delivery']:<15} {order['Days_Late']:.0f}")

            if len(backlog_orders) > 10:
                print(f"   ... and {len(backlog_orders) - 10} more backlog orders")

            print()
            print("   ℹ️  This is not an error - optimizer handles backlogs automatically")
            print()

        # Report rush orders (WARNING - might be tight but not infeasible)
        if rush_orders:
            print(f"⚠️  Found {len(rush_orders)} rush orders (tight lead time but not past due):")
            print()
            print(f"   {'Part':<15} {'Qty':<8} {'Delivery':<12} {'Avail':<8} {'Need':<8} {'Short':<8}")
            print("   " + "-" * 70)

            for order in rush_orders[:10]:
                print(f"   {order['Part']:<15} {order['Qty']:<8.0f} {order['Delivery']:<12} "
                      f"{order['Available_Days']:<8.0f} {order['Required_Days']:<8.1f} "
                      f"{order['Shortfall_Days']:<8.1f}")

            if len(rush_orders) > 10:
                print(f"   ... and {len(rush_orders) - 10} more rush orders")

            print()
            print("   ⚠️  These may be produced/delivered slightly late")
            print()

            self.warnings_list.append({
                'severity': 'WARNING',
                'category': 'Lead Time',
                'issue': f'{len(rush_orders)} rush orders with tight lead time',
                'details': rush_orders[:20]
            })

        if not backlog_orders and not rush_orders:
            print("✓ All orders have sufficient lead time")
            print()

        print()

    def check_wip_mapping(self):
        """Check WIP mapping to Part Master"""
        print("="*80)
        print("4. WIP MAPPING VALIDATION")
        print("="*80)

        # Check FG WIP mapping
        fg_wip = self.stage_wip.dropna(subset=['FG'])
        fg_codes = set(self.part_master['FG Code'].dropna().astype(str))

        unmapped_fg = []
        for _, row in fg_wip.iterrows():
            material_code = str(row.get('Material Code', '')).strip()
            if material_code and material_code not in fg_codes:
                unmapped_fg.append(material_code)

        # Check CS WIP mapping
        cs_wip = self.stage_wip.dropna(subset=['CS'])
        cs_codes = set(self.part_master['CS Code'].dropna().astype(str))

        unmapped_cs = []
        for _, row in cs_wip.iterrows():
            casting_item = str(row.get('CastingItem', '')).strip()
            if casting_item and casting_item not in cs_codes:
                unmapped_cs.append(casting_item)

        if unmapped_fg:
            print(f"⚠️  {len(unmapped_fg)} FG WIP parts not in Part Master:")
            for part in unmapped_fg[:10]:
                print(f"    • {part}")
            if len(unmapped_fg) > 10:
                print(f"    ... and {len(unmapped_fg) - 10} more")
            print()

            self.warnings_list.append({
                'severity': 'WARNING',
                'category': 'WIP Mapping',
                'issue': f'{len(unmapped_fg)} FG WIP parts not mapped',
                'details': unmapped_fg[:20]
            })

        if unmapped_cs:
            print(f"⚠️  {len(unmapped_cs)} CS WIP parts not in Part Master:")
            for part in unmapped_cs[:10]:
                print(f"    • {part}")
            if len(unmapped_cs) > 10:
                print(f"    ... and {len(unmapped_cs) - 10} more")
            print()

            self.warnings_list.append({
                'severity': 'WARNING',
                'category': 'WIP Mapping',
                'issue': f'{len(unmapped_cs)} CS WIP parts not mapped',
                'details': unmapped_cs[:20]
            })

        if not unmapped_fg and not unmapped_cs:
            print("✓ All WIP items mapped correctly")
            print()

    def check_box_capacity(self):
        """Check mould box capacity constraints"""
        print("="*80)
        print("5. MOULD BOX CAPACITY ANALYSIS")
        print("="*80)

        valid_parts = set(self.part_master['FG Code'].dropna().astype(str))
        valid_orders = self.sales_orders[
            self.sales_orders['Material Code'].isin(valid_parts)
        ].copy()

        # Calculate demand by box size
        box_demand = defaultdict(float)
        part_master_dedup = self.part_master.drop_duplicates(subset=['FG Code'], keep='first')
        part_master_dict = part_master_dedup.set_index('FG Code').to_dict('index')

        for _, order in valid_orders.iterrows():
            part = order['Material Code']
            qty = order['Balance Qty']

            if part not in part_master_dict:
                continue

            pm = part_master_dict[part]
            box_size = str(pm.get('Box Size', '')).strip()
            box_qty = self._safe_float(pm.get('Box Quantity', 1))

            if box_size and box_qty > 0:
                boxes_needed = qty / box_qty
                box_demand[box_size] += boxes_needed

        # Get capacity
        box_capacity_dict = {}
        for _, row in self.box_capacity.iterrows():
            box_size = str(row.get('Box_Size', '')).strip()
            weekly_cap = self._safe_float(row.get('Weekly_Capacity', 0))
            box_capacity_dict[box_size] = weekly_cap

        # Calculate planning weeks
        valid_orders['Comitted Delivery Date'] = pd.to_datetime(
            valid_orders['Comitted Delivery Date'], errors='coerce'
        )
        valid_orders = valid_orders.dropna(subset=['Comitted Delivery Date'])
        latest_delivery = valid_orders['Comitted Delivery Date'].max()
        planning_weeks = int((latest_delivery - self.CURRENT_DATE).days / 7) + 2

        print(f"Box Capacity Analysis (over {planning_weeks} weeks):")
        print("-" * 70)
        print(f"{'Box Size':<15} {'Demand (boxes)':<18} {'Capacity':<15} {'Utilization':<12} {'Status'}")
        print("-" * 70)

        for box_size in sorted(box_demand.keys()):
            demand = box_demand[box_size]
            capacity = box_capacity_dict.get(box_size, 0) * planning_weeks

            if capacity > 0:
                utilization = (demand / capacity) * 100
                status = "✓" if utilization <= 100 else "❌ OVERFLOW"

                if utilization > 100:
                    self.issues.append({
                        'severity': 'CRITICAL',
                        'category': 'Box Capacity',
                        'issue': f'Box size {box_size} capacity exceeded',
                        'details': f'Need {demand:.0f} boxes, have {capacity:.0f} ({utilization:.1f}%)'
                    })

                print(f"{box_size:<15} {demand:<18.0f} {capacity:<15.0f} {utilization:>10.1f}%  {status}")
            else:
                print(f"{box_size:<15} {demand:<18.0f} {'N/A':<15} {'N/A':<12}  ⚠️")
                self.issues.append({
                    'severity': 'CRITICAL',
                    'category': 'Box Capacity',
                    'issue': f'Box size {box_size} has no capacity defined',
                    'details': f'Need {demand:.0f} boxes but capacity is 0'
                })

        print("-" * 70)
        print()

    def generate_report(self):
        """Generate final diagnostic report"""
        print("="*80)
        print("DIAGNOSTIC SUMMARY")
        print("="*80)
        print()

        if not self.issues and not self.warnings_list:
            print("✅ NO ISSUES FOUND - Optimization should be feasible!")
            print()
            print("If optimization still fails, check:")
            print("  • Solver timeout (increase limit)")
            print("  • Model complexity (consider weekly instead of daily)")
            print("  • Constraint interactions (stage seriality + capacity)")
            return

        # Critical issues
        critical = [i for i in self.issues if i['severity'] == 'CRITICAL']
        if critical:
            print(f"❌ CRITICAL ISSUES ({len(critical)}):")
            print("-" * 80)
            for i, issue in enumerate(critical, 1):
                print(f"{i}. [{issue['category']}] {issue['issue']}")
                print(f"   {issue['details']}")
                print()

        # Warnings
        if self.warnings_list:
            print(f"⚠️  WARNINGS ({len(self.warnings_list)}):")
            print("-" * 80)
            for i, warning in enumerate(self.warnings_list, 1):
                print(f"{i}. [{warning['category']}] {warning['issue']}")
                print(f"   {warning['details']}")
                print()

        # Recommendations
        print("="*80)
        print("RECOMMENDATIONS")
        print("="*80)
        print()

        if critical:
            print("🔧 Immediate Actions:")
            print()

            capacity_issues = [i for i in critical if i['category'] == 'Capacity']
            if capacity_issues:
                print("1. CAPACITY SHORTFALL:")
                print("   • Add overtime shifts")
                print("   • Add temporary resources")
                print("   • Negotiate extended delivery dates")
                print("   • Consider outsourcing overflow work")
                print()

            lead_time_issues = [i for i in critical if i['category'] == 'Lead Time']
            if lead_time_issues:
                print("2. LEAD TIME CONFLICTS:")
                print("   • Prioritize orders with sufficient lead time")
                print("   • Negotiate later delivery dates for rush orders")
                print("   • Use expedited processes where possible")
                print()

            box_issues = [i for i in critical if i['category'] == 'Box Capacity']
            if box_issues:
                print("3. BOX CAPACITY CONSTRAINTS:")
                print("   • Increase mould box availability")
                print("   • Review box allocation strategy")
                print("   • Consider alternative mould sizes")
                print()

            order_issues = [i for i in critical if i['category'] == 'Order Validity']
            if order_issues:
                print("4. INVALID PARTS IN ORDERS:")
                print("   • Add missing parts to Part Master")
                print("   • OR remove invalid orders from Sales Order sheet")
                print("   • Verify part code consistency")
                print()

    def _safe_float(self, value):
        """Safely convert to float"""
        try:
            return float(value) if pd.notna(value) and value != '' else 0.0
        except:
            return 0.0

    def _safe_int(self, value):
        """Safely convert to int"""
        try:
            return int(float(value)) if pd.notna(value) and value != '' else 0
        except:
            return 0

    def run_full_diagnostics(self):
        """Run all diagnostic checks"""
        try:
            self.load_data()
            self.check_order_validity()
            self.check_capacity_vs_demand()
            self.check_lead_time_feasibility()
            self.check_wip_mapping()
            self.check_box_capacity()
            self.generate_report()

            print("="*80)
            print("DIAGNOSTICS COMPLETE")
            print("="*80)

        except Exception as e:
            print(f"\n❌ Diagnostic failed with error: {e}")
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    diagnostics = InfeasibilityDiagnostics()
    diagnostics.run_full_diagnostics()
