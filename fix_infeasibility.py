"""
Automatic Fix for Infeasibility Issues
======================================

This script automatically fixes the identified infeasibility issues by:
1. Adjusting past-due delivery dates to feasible dates
2. Increasing box capacity where needed
3. Removing invalid parts from orders (or flagging them)

BACKUP: Creates backup of original file before modifications
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import shutil
from pathlib import Path

class InfeasibilityFixer:
    """Automatically fix infeasibility issues"""

    def __init__(self, master_file='Master_Data_Updated_Nov_Dec.xlsx'):
        self.master_file = master_file
        self.backup_file = master_file.replace('.xlsx', '_BACKUP.xlsx')
        self.CURRENT_DATE = datetime(2025, 11, 22)
        self.fixes_applied = []

        print("="*80)
        print("INFEASIBILITY AUTO-FIX")
        print("="*80)
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

    def fix_delivery_dates(self):
        """Fix past-due and insufficient lead time delivery dates"""
        print("="*80)
        print("FIX 1: DELIVERY DATE ADJUSTMENTS")
        print("="*80)

        valid_parts = set(self.part_master['FG Code'].dropna().astype(str))
        part_master_dedup = self.part_master.drop_duplicates(subset=['FG Code'], keep='first')
        part_master_dict = part_master_dedup.set_index('FG Code').to_dict('index')

        # Parse delivery dates
        self.sales_orders['Comitted Delivery Date'] = pd.to_datetime(
            self.sales_orders['Comitted Delivery Date'], errors='coerce'
        )

        orders_fixed = 0

        for idx, order in self.sales_orders.iterrows():
            part = str(order.get('Material Code', '')).strip()
            delivery_date = order.get('Comitted Delivery Date')

            if pd.isna(delivery_date) or part not in valid_parts:
                continue

            if part not in part_master_dict:
                continue

            pm = part_master_dict[part]

            # Calculate minimum lead time
            total_lead_days = self._calculate_lead_time(pm)

            # Calculate available time
            available_days = (delivery_date - self.CURRENT_DATE).days

            # Check if feasible
            if available_days < total_lead_days:
                # Adjust delivery date
                new_delivery_date = self.CURRENT_DATE + timedelta(days=int(total_lead_days) + 1)

                self.sales_orders.at[idx, 'Comitted Delivery Date'] = new_delivery_date

                orders_fixed += 1
                if orders_fixed <= 10:
                    print(f"   • {part}: {delivery_date.strftime('%Y-%m-%d')} → "
                          f"{new_delivery_date.strftime('%Y-%m-%d')} "
                          f"(+{int(total_lead_days - available_days)} days)")

        if orders_fixed > 10:
            print(f"   ... and {orders_fixed - 10} more orders")

        print()
        print(f"✓ Fixed {orders_fixed} delivery dates")
        self.fixes_applied.append(f"Adjusted {orders_fixed} delivery dates to be feasible")
        print()

    def fix_box_capacity(self):
        """Increase box capacity where needed"""
        print("="*80)
        print("FIX 2: BOX CAPACITY INCREASES")
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

        # Check and fix capacity
        boxes_fixed = 0

        for idx, row in self.box_capacity.iterrows():
            box_size = str(row.get('Box_Size', '')).strip()
            weekly_cap = self._safe_float(row.get('Weekly_Capacity', 0))
            total_capacity = weekly_cap * planning_weeks

            demand = box_demand.get(box_size, 0)

            if demand > total_capacity:
                # Increase capacity by 50% over demand
                new_weekly_cap = int((demand / planning_weeks) * 1.5)
                self.box_capacity.at[idx, 'Weekly_Capacity'] = new_weekly_cap

                print(f"   • {box_size}: {weekly_cap:.0f} → {new_weekly_cap:.0f} boxes/week "
                      f"(+{((new_weekly_cap/weekly_cap - 1)*100):.1f}%)")
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
        print("FIX 3: INVALID PARTS REMOVAL")
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

        output_file = self.master_file.replace('.xlsx', '_FIXED.xlsx')

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
        print("="*80)
        print("NEXT STEPS")
        print("="*80)
        print()
        print("1. Review the fixed data file:")
        print(f"   {output_file}")
        print()
        print("2. If satisfied, update your optimization script to use:")
        print(f"   master_file = '{output_file}'")
        print()
        print("3. Run the optimization again:")
        print("   python production_plan_test.py")
        print()
        print("4. Original backup available at:")
        print(f"   {self.backup_file}")
        print()

    def _calculate_lead_time(self, pm):
        """Calculate minimum lead time for a part"""
        total_days = 0

        # Casting + passive times
        casting_time = self._safe_float(pm.get('Casting Cycle time (min)', 0))
        cooling_time = self._safe_float(pm.get('Cooling Time (hrs)', 0))
        shakeout_time = self._safe_float(pm.get('Shakeout Time (hrs)', 0))
        vacuum_time = self._safe_float(pm.get('Vacuum Time (hrs)', 0))
        total_days += (casting_time / 60 + cooling_time + shakeout_time + vacuum_time) / 24

        # Grinding
        grinding_time = self._safe_float(pm.get('Grinding Cycle time (min)', 0))
        total_days += (grinding_time / 60) / 24

        # Machining (sequential)
        for i in range(1, 4):
            mc_time = self._safe_float(pm.get(f'Machining Cycle time {i} (min)', 0))
            total_days += (mc_time / 60) / 24

        # Painting (sequential with dry times)
        for i in range(1, 4):
            sp_time = self._safe_float(pm.get(f'Painting Cycle time {i} (min)', 0))
            dry_time = self._safe_float(pm.get(f'Dry Time {i} (hrs)', 0))
            total_days += (sp_time / 60 + dry_time) / 24

        # Add 20% buffer
        return total_days * 1.2

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
            self.fix_delivery_dates()
            self.fix_box_capacity()
            self.fix_invalid_parts()
            output_file = self.save_fixed_data()
            self.generate_summary(output_file)

        except Exception as e:
            print(f"\n❌ Fix failed with error: {e}")
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    fixer = InfeasibilityFixer()
    fixer.run_all_fixes()
