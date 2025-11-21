"""
Corrected Infeasibility Fix (v2)
=================================

IMPORTANT: This version correctly handles backlog orders by NOT adjusting
their delivery dates. Backlog orders are legitimate late orders that should
be produced ASAP and delivered late (with lateness penalties).

This script ONLY fixes:
1. Box capacity constraints (increases where needed)
2. Invalid parts (removes orders for parts not in Part Master)

It does NOT adjust delivery dates for backlogs.
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import shutil
from pathlib import Path

class CorrectedInfeasibilityFixer:
    """Fix only actual infeasibility issues, preserve backlog dates"""

    def __init__(self, master_file='Master_Data_Updated_Nov_Dec.xlsx'):
        self.master_file = master_file
        self.backup_file = master_file.replace('.xlsx', '_BACKUP_V2.xlsx')
        self.CURRENT_DATE = datetime(2025, 11, 22)
        self.fixes_applied = []

        print("="*80)
        print("CORRECTED INFEASIBILITY FIX (v2)")
        print("="*80)
        print()
        print("This version preserves backlog order dates.")
        print("Backlog orders will be optimized as late deliveries with penalties.")
        print()

    def create_backup(self):
        """Create backup of original file"""
        print("📦 Creating backup...")
        shutil.copy2(self.master_file, self.backup_file)
        print(f"   ✓ Backup created: {self.backup_file}")
        print()

    def load_data(self):
        """Load all data"""
        print("📂 Loading data...")
        self.part_master = pd.read_excel(self.master_file, sheet_name='Part Master')
        self.sales_orders = pd.read_excel(self.master_file, sheet_name='Sales Order')
        self.machine_constraints = pd.read_excel(self.master_file, sheet_name='Machine Constraints')
        self.stage_wip = pd.read_excel(self.master_file, sheet_name='Stage WIP')
        self.box_capacity = pd.read_excel(self.master_file, sheet_name='Mould Box Capacity')
        print("   ✓ Data loaded")
        print()

    def analyze_backlog(self):
        """Analyze backlog orders (don't fix them, just report)"""
        print("="*80)
        print("BACKLOG ANALYSIS (NO CHANGES)")
        print("="*80)

        self.sales_orders['Comitted Delivery Date'] = pd.to_datetime(
            self.sales_orders['Comitted Delivery Date'], errors='coerce'
        )

        backlog = self.sales_orders[
            self.sales_orders['Comitted Delivery Date'] < self.CURRENT_DATE
        ].copy()

        if len(backlog) > 0:
            total_backlog_qty = backlog['Balance Qty'].sum()
            print(f"📊 Found {len(backlog)} backlog orders ({total_backlog_qty:.0f} units)")
            print()
            print("   These orders will be:")
            print("   • Produced as soon as possible")
            print("   • Delivered late (incurring lateness penalties)")
            print("   • Optimized alongside future orders")
            print()
            print("   Top 10 backlog orders:")
            print(f"   {'Part':<15} {'Qty':<8} {'Original Due':<15} {'Days Late'}")
            print("   " + "-"*60)

            backlog['days_late'] = (self.CURRENT_DATE - backlog['Comitted Delivery Date']).dt.days

            for _, order in backlog.nlargest(10, 'days_late').iterrows():
                part = str(order['Material Code'])[:15]
                qty = order['Balance Qty']
                due = order['Comitted Delivery Date'].strftime('%Y-%m-%d')
                days = order['days_late']
                print(f"   {part:<15} {qty:<8.0f} {due:<15} {days:.0f}")

            print()
            print(f"   ℹ️  Keeping original dates - optimizer will handle lateness")
        else:
            print("   ✓ No backlog orders (all orders are future-dated)")

        print()

    def fix_box_capacity(self):
        """Increase box capacity where needed"""
        print("="*80)
        print("FIX 1: BOX CAPACITY INCREASES")
        print("="*80)

        valid_parts = set(self.part_master['FG Code'].dropna().astype(str))
        valid_orders = self.sales_orders[
            self.sales_orders['Material Code'].isin(valid_parts)
        ].copy()

        # Calculate demand by box size
        box_demand = {}
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
                box_demand[box_size] = box_demand.get(box_size, 0) + boxes_needed

        # Calculate planning weeks
        valid_orders['Comitted Delivery Date'] = pd.to_datetime(
            valid_orders['Comitted Delivery Date'], errors='coerce'
        )
        valid_orders = valid_orders.dropna(subset=['Comitted Delivery Date'])
        latest_delivery = valid_orders['Comitted Delivery Date'].max()
        planning_weeks = int((latest_delivery - self.CURRENT_DATE).days / 7) + 2

        print(f"📅 Planning horizon: {planning_weeks} weeks")
        print()

        # Check and fix capacity
        boxes_fixed = 0

        for idx, row in self.box_capacity.iterrows():
            box_size = str(row.get('Box_Size', '')).strip()
            weekly_cap = self._safe_float(row.get('Weekly_Capacity', 0))
            total_capacity = weekly_cap * planning_weeks

            demand = box_demand.get(box_size, 0)
            utilization = (demand / total_capacity * 100) if total_capacity > 0 else 0

            if demand > total_capacity:
                # Increase capacity by 50% over demand
                new_weekly_cap = int((demand / planning_weeks) * 1.5)
                self.box_capacity.at[idx, 'Weekly_Capacity'] = new_weekly_cap

                print(f"   • {box_size}: {weekly_cap:.0f} → {new_weekly_cap:.0f} boxes/week")
                print(f"     Reason: Demand {demand:.0f} > Capacity {total_capacity:.0f} ({utilization:.1f}%)")
                boxes_fixed += 1

        if boxes_fixed == 0:
            print("   ✓ No box capacity increases needed")
        else:
            print()
            print(f"✓ Increased capacity for {boxes_fixed} box sizes")
            self.fixes_applied.append(f"Increased capacity for {boxes_fixed} box sizes")

        print()

    def fix_invalid_parts(self):
        """Remove orders for invalid parts"""
        print("="*80)
        print("FIX 2: INVALID PARTS REMOVAL")
        print("="*80)

        valid_parts = set(self.part_master['FG Code'].dropna().astype(str))
        ordered_parts = set(self.sales_orders['Material Code'].dropna().astype(str))
        invalid_parts = ordered_parts - valid_parts

        if invalid_parts:
            # Remove invalid orders
            original_count = len(self.sales_orders)
            self.sales_orders = self.sales_orders[
                self.sales_orders['Material Code'].isin(valid_parts)
            ]
            removed_count = original_count - len(self.sales_orders)

            print(f"   ✓ Removed {removed_count} order lines for {len(invalid_parts)} invalid parts:")
            for part in sorted(invalid_parts):
                print(f"      • {part}")

            self.fixes_applied.append(f"Removed {removed_count} orders for invalid parts")
        else:
            print("   ✓ No invalid parts found")

        print()

    def save_fixed_data(self):
        """Save fixed data to Excel"""
        print("="*80)
        print("SAVING FIXED DATA")
        print("="*80)

        output_file = self.master_file.replace('.xlsx', '_FIXED_V2.xlsx')

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            self.part_master.to_excel(writer, sheet_name='Part Master', index=False)
            self.sales_orders.to_excel(writer, sheet_name='Sales Order', index=False)
            self.machine_constraints.to_excel(writer, sheet_name='Machine Constraints', index=False)
            self.stage_wip.to_excel(writer, sheet_name='Stage WIP', index=False)
            self.box_capacity.to_excel(writer, sheet_name='Mould Box Capacity', index=False)

        print(f"✓ Fixed data saved to: {output_file}")
        print()

        return output_file

    def generate_summary(self, output_file):
        """Print summary of fixes"""
        print("="*80)
        print("FIX SUMMARY")
        print("="*80)
        print()

        if self.fixes_applied:
            print("✅ Fixes Applied:")
            for i, fix in enumerate(self.fixes_applied, 1):
                print(f"   {i}. {fix}")
        else:
            print("ℹ️  No fixes needed")

        print()
        print("✅ Preserved:")
        print("   • All backlog order dates (will incur lateness penalties)")
        print("   • All future order dates (unchanged)")
        print()

        print("="*80)
        print("NEXT STEPS")
        print("="*80)
        print()
        print("1. Review the fixed data file:")
        print(f"   {output_file}")
        print()
        print("2. Update your optimization script:")
        print(f"   master_file = '{output_file}'")
        print()
        print("3. Run optimization:")
        print("   python production_plan_test.py")
        print()
        print("4. Backlog orders will:")
        print("   • Be produced ASAP")
        print("   • Show as late deliveries")
        print("   • Incur lateness penalties in objective function")
        print()

    def _safe_float(self, value):
        """Safely convert to float"""
        try:
            return float(value) if pd.notna(value) and value != '' else 0.0
        except:
            return 0.0

    def run_all_fixes(self):
        """Run all fixes"""
        try:
            self.create_backup()
            self.load_data()
            self.analyze_backlog()  # Analyze but don't fix
            self.fix_box_capacity()
            self.fix_invalid_parts()
            output_file = self.save_fixed_data()
            self.generate_summary(output_file)

        except Exception as e:
            print(f"\n❌ Fix failed with error: {e}")
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1:
        master_file = sys.argv[1]
        print(f"Using master file: {master_file}")
        print()
        fixer = CorrectedInfeasibilityFixer(master_file=master_file)
    else:
        fixer = CorrectedInfeasibilityFixer()

    fixer.run_all_fixes()
