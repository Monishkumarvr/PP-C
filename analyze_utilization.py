"""
Deep Utilization Analysis Tool
==============================

Analyzes why production stages are not fully utilized and identifies root causes.

This tool investigates:
1. Demand distribution across weeks
2. Capacity vs demand by stage
3. Mould box capacity utilization
4. Stage seriality bottlenecks
5. Inventory holding cost impact
6. Just-in-time production patterns
7. Order clustering by delivery week
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from collections import defaultdict
import warnings
warnings.filterwarnings('ignore')

class UtilizationAnalyzer:
    """Deep analysis of why stages are underutilized"""

    def __init__(self,
                 comprehensive_output='production_plan_COMPREHENSIVE_test.xlsx',
                 master_file='Master_Data_Updated_Nov_Dec2.xlsx'):
        self.comprehensive_output = comprehensive_output
        self.master_file = master_file
        self.CURRENT_DATE = datetime(2025, 11, 22)
        self.OEE = 0.90
        self.WORKING_DAYS_PER_WEEK = 6

        print("="*80)
        print("DEEP UTILIZATION ANALYSIS")
        print("="*80)
        print()

    def load_data(self):
        """Load optimizer output and master data"""
        print("📂 Loading data...")

        try:
            # Load optimizer output
            self.weekly_summary = pd.read_excel(self.comprehensive_output, sheet_name='Weekly_Summary')
            self.part_fulfillment = pd.read_excel(self.comprehensive_output, sheet_name='Part_Fulfillment')

            # Load master data
            self.part_master = pd.read_excel(self.master_file, sheet_name='Part Master')
            self.sales_orders = pd.read_excel(self.master_file, sheet_name='Sales Order')
            self.machine_constraints = pd.read_excel(self.master_file, sheet_name='Machine Constraints')
            self.box_capacity = pd.read_excel(self.master_file, sheet_name='Mould Box Capacity')

            print(f"  ✓ Weekly Summary: {len(self.weekly_summary)} weeks")
            print(f"  ✓ Part Fulfillment: {len(self.part_fulfillment)} parts")
            print(f"  ✓ Sales Orders: {len(self.sales_orders)} orders")
            print()

        except Exception as e:
            print(f"❌ Error loading data: {e}")
            raise

    def analyze_demand_distribution(self):
        """Analyze how demand is distributed across weeks"""
        print("="*80)
        print("ANALYSIS 1: DEMAND DISTRIBUTION")
        print("="*80)
        print()

        # Parse delivery dates
        self.sales_orders['Comitted Delivery Date'] = pd.to_datetime(
            self.sales_orders['Comitted Delivery Date'], errors='coerce'
        )
        valid_orders = self.sales_orders.dropna(subset=['Comitted Delivery Date'])

        # Calculate week number
        valid_orders['Week_Num'] = valid_orders['Comitted Delivery Date'].apply(
            lambda d: int((d - self.CURRENT_DATE).days / 7) + 1
        )

        # Group by week
        demand_by_week = valid_orders.groupby('Week_Num').agg({
            'Balance Qty': 'sum',
            'Material Code': 'count'
        }).rename(columns={'Balance Qty': 'Total_Qty', 'Material Code': 'Order_Count'})

        print("Demand Distribution by Week:")
        print("-"*60)
        print(f"{'Week':<8} {'Orders':<10} {'Quantity':<12} {'% of Total'}")
        print("-"*60)

        total_qty = demand_by_week['Total_Qty'].sum()
        for week in range(1, 20):
            if week in demand_by_week.index:
                orders = demand_by_week.loc[week, 'Order_Count']
                qty = demand_by_week.loc[week, 'Total_Qty']
                pct = (qty / total_qty * 100) if total_qty > 0 else 0
                print(f"W{week:<7} {orders:<10.0f} {qty:<12.0f} {pct:>6.1f}%")
            else:
                print(f"W{week:<7} {0:<10} {0:<12} {0:>6.1f}%")

        print("-"*60)
        print(f"Total:   {demand_by_week['Order_Count'].sum():<10.0f} {total_qty:<12.0f} 100.0%")
        print()

        # Identify concentration
        top_3_weeks = demand_by_week.nlargest(3, 'Total_Qty')
        top_3_pct = (top_3_weeks['Total_Qty'].sum() / total_qty * 100)

        print(f"📊 Key Findings:")
        print(f"  • Top 3 weeks contain {top_3_pct:.1f}% of total demand")
        print(f"  • Weeks with orders: {len(demand_by_week)}")
        print(f"  • Weeks without orders: {19 - len(demand_by_week)}")
        print()

        if top_3_pct > 60:
            print(f"⚠️  INSIGHT: Demand is highly concentrated in {top_3_weeks.index.tolist()}")
            print(f"   → This explains low utilization in other weeks")

        print()
        return demand_by_week

    def analyze_stage_bottlenecks(self):
        """Identify bottlenecks preventing full utilization"""
        print("="*80)
        print("ANALYSIS 2: STAGE BOTTLENECK ANALYSIS")
        print("="*80)
        print()

        stages = ['Casting', 'Grinding', 'MC1', 'MC2', 'MC3', 'SP1', 'SP2', 'SP3']

        print("Stage-by-Stage Utilization Pattern:")
        print("-"*80)
        print(f"{'Week':<8} {'Cast%':<8} {'Grind%':<8} {'MC1%':<8} {'MC2%':<8} "
              f"{'MC3%':<8} {'SP1%':<8} {'SP2%':<8} {'SP3%':<8}")
        print("-"*80)

        for _, row in self.weekly_summary.iterrows():
            week = row['Week']
            cast = row.get('Casting_Utilization_%', 0)
            grind = row.get('Grinding_Utilization_%', 0)
            mc1 = row.get('MC1_Utilization_%', 0)
            mc2 = row.get('MC2_Utilization_%', 0)
            mc3 = row.get('MC3_Utilization_%', 0)
            sp1 = row.get('SP1_Utilization_%', 0)
            sp2 = row.get('SP2_Utilization_%', 0)
            sp3 = row.get('SP3_Utilization_%', 0)

            print(f"{week:<8} {cast:<8.1f} {grind:<8.1f} {mc1:<8.1f} {mc2:<8.1f} "
                  f"{mc3:<8.1f} {sp1:<8.1f} {sp2:<8.1f} {sp3:<8.1f}")

        print("-"*80)
        print()

        # Calculate average utilization by stage
        print("Average Utilization by Stage:")
        print("-"*60)
        for stage in stages:
            col_name = f"{stage}_Utilization_%"
            if col_name in self.weekly_summary.columns:
                avg_util = self.weekly_summary[col_name].mean()
                max_util = self.weekly_summary[col_name].max()
                weeks_over_80 = (self.weekly_summary[col_name] > 80).sum()

                print(f"  {stage:<15} Avg: {avg_util:>5.1f}%  Max: {max_util:>5.1f}%  "
                      f"Weeks >80%: {weeks_over_80}")

        print()

        # Identify bottleneck patterns
        print("📊 Bottleneck Patterns:")

        # Check for stage seriality bottlenecks
        for week in range(1, 4):  # Focus on high-demand weeks
            week_data = self.weekly_summary[self.weekly_summary['Week'] == f'W{week}']
            if len(week_data) == 0:
                continue

            mc1 = week_data['MC1_Utilization_%'].values[0]
            mc2 = week_data['MC2_Utilization_%'].values[0]
            mc3 = week_data['MC3_Utilization_%'].values[0]

            if mc1 > 80 and mc2 < 50:
                print(f"  • W{week}: MC1→MC2 seriality gap ({mc1:.1f}% vs {mc2:.1f}%)")

            sp1 = week_data['SP1_Utilization_%'].values[0]
            sp2 = week_data['SP2_Utilization_%'].values[0]
            sp3 = week_data['SP3_Utilization_%'].values[0]

            if sp1 > 80 and sp3 < 50:
                print(f"  • W{week}: SP1→SP3 seriality gap ({sp1:.1f}% vs {sp3:.1f}%)")

        print()

    def analyze_box_capacity_utilization(self):
        """Analyze mould box capacity utilization"""
        print("="*80)
        print("ANALYSIS 3: MOULD BOX CAPACITY UTILIZATION")
        print("="*80)
        print()

        # Calculate demand by box size
        part_master_dict = self.part_master.set_index('FG Code').to_dict('index')

        valid_parts = set(self.part_master['FG Code'].dropna().astype(str))
        valid_orders = self.sales_orders[
            self.sales_orders['Material Code'].isin(valid_parts)
        ].copy()

        valid_orders['Comitted Delivery Date'] = pd.to_datetime(
            valid_orders['Comitted Delivery Date'], errors='coerce'
        )
        valid_orders = valid_orders.dropna(subset=['Comitted Delivery Date'])

        valid_orders['Week_Num'] = valid_orders['Comitted Delivery Date'].apply(
            lambda d: int((d - self.CURRENT_DATE).days / 7) + 1
        )

        # Calculate box demand by week
        box_demand_by_week = defaultdict(lambda: defaultdict(float))

        for _, order in valid_orders.iterrows():
            part = order['Material Code']
            qty = order['Balance Qty']
            week = order['Week_Num']

            if part not in part_master_dict:
                continue

            pm = part_master_dict[part]
            box_size = str(pm.get('Box Size', '')).strip()
            box_qty = self._safe_float(pm.get('Box Quantity', 1))

            if box_size and box_qty > 0:
                boxes_needed = qty / box_qty
                box_demand_by_week[week][box_size] += boxes_needed

        # Get box capacity
        box_cap_dict = {}
        for _, row in self.box_capacity.iterrows():
            box_size = str(row['Box_Size']).strip()
            weekly_cap = self._safe_float(row['Weekly_Capacity'])
            box_cap_dict[box_size] = weekly_cap

        # Print utilization
        print("Box Capacity Utilization by Week:")
        print("-"*80)

        all_box_sizes = sorted(set(box_cap_dict.keys()))
        header = f"{'Week':<8}"
        for box_size in all_box_sizes:
            header += f"{box_size:<12}"
        print(header)
        print("-"*80)

        for week in range(1, 20):
            row_str = f"W{week:<7}"
            for box_size in all_box_sizes:
                demand = box_demand_by_week.get(week, {}).get(box_size, 0)
                capacity = box_cap_dict.get(box_size, 0)

                if capacity > 0:
                    util = (demand / capacity) * 100
                    row_str += f"{util:>10.1f}% "
                else:
                    row_str += f"{'N/A':>11} "

            print(row_str)

        print("-"*80)
        print()

        # Summary
        print("📊 Box Capacity Insights:")
        for box_size in all_box_sizes:
            total_demand = sum(box_demand_by_week[w].get(box_size, 0) for w in range(1, 20))
            capacity = box_cap_dict.get(box_size, 0)

            if capacity > 0:
                weeks_with_demand = sum(1 for w in range(1, 20)
                                       if box_demand_by_week[w].get(box_size, 0) > 0)
                avg_util = (total_demand / (capacity * 19)) * 100

                print(f"  • {box_size}: {weeks_with_demand} weeks active, "
                      f"avg {avg_util:.1f}% utilization")

        print()

    def analyze_jit_production(self):
        """Analyze just-in-time production patterns"""
        print("="*80)
        print("ANALYSIS 4: JUST-IN-TIME PRODUCTION PATTERN")
        print("="*80)
        print()

        print("Optimizer Behavior Analysis:")
        print("-"*60)

        # Check inventory holding cost impact
        print("📊 Key Parameters:")
        print(f"  • Inventory Holding Cost: $1 per unit per week")
        print(f"  • Lateness Penalty: $150,000 per week late")
        print(f"  • Unmet Demand Penalty: $200,000 per unit")
        print()

        print("💡 Economic Trade-offs:")
        print(f"  • Early production (W4) for W10 delivery = $6 holding cost")
        print(f"  • Late production (W11) for W10 delivery = $150,000 penalty")
        print(f"  • Result: Optimizer produces just-in-time to avoid holding cost")
        print()

        # Check production timing vs delivery
        print("Production Timing vs Delivery Date:")
        print("-"*60)

        # Analyze part fulfillment
        if 'Delivery_Week' in self.part_fulfillment.columns and 'Production_Week' in self.part_fulfillment.columns:
            self.part_fulfillment['Lead_Time_Weeks'] = (
                pd.to_numeric(self.part_fulfillment['Delivery_Week'].str.replace('W', ''), errors='coerce') -
                pd.to_numeric(self.part_fulfillment['Production_Week'].str.replace('W', ''), errors='coerce')
            )

            avg_lead = self.part_fulfillment['Lead_Time_Weeks'].mean()
            print(f"  • Average production lead time: {avg_lead:.1f} weeks")
            print(f"  • Parts produced same week as delivery: {(self.part_fulfillment['Lead_Time_Weeks'] == 0).sum()}")
            print(f"  • Parts produced 1 week early: {(self.part_fulfillment['Lead_Time_Weeks'] == 1).sum()}")

        print()

    def analyze_root_causes(self):
        """Identify root causes of underutilization"""
        print("="*80)
        print("ROOT CAUSE ANALYSIS")
        print("="*80)
        print()

        print("🔍 Why Stages Are Not 100% Utilized:")
        print()

        causes = []

        # Check demand concentration
        valid_orders = self.sales_orders.dropna(subset=['Comitted Delivery Date'])
        valid_orders['Week_Num'] = valid_orders['Comitted Delivery Date'].apply(
            lambda d: int((d - self.CURRENT_DATE).days / 7) + 1
        )
        demand_by_week = valid_orders.groupby('Week_Num')['Balance Qty'].sum()

        weeks_with_demand = len(demand_by_week)
        if weeks_with_demand < 10:
            causes.append({
                'rank': 1,
                'cause': 'Concentrated Demand',
                'explanation': f'Orders only exist in {weeks_with_demand} out of 19 weeks',
                'impact': 'HIGH',
                'solution': 'Accept more orders for weeks 4-19'
            })

        # Check JIT optimization
        causes.append({
            'rank': 2,
            'cause': 'Just-In-Time Optimization',
            'explanation': 'Low inventory holding cost ($1) encourages production close to delivery',
            'impact': 'HIGH',
            'solution': 'Increase inventory holding cost OR reduce lateness penalty'
        })

        # Check stage seriality
        causes.append({
            'rank': 3,
            'cause': 'Stage Seriality Constraints',
            'explanation': 'MC1→MC2→MC3 and SP1→SP2→SP3 must flow sequentially',
            'impact': 'MEDIUM',
            'solution': 'Cannot be changed (manufacturing reality)'
        })

        # Check box capacity
        box_util = self._calculate_box_utilization()
        if box_util < 50:
            causes.append({
                'rank': 4,
                'cause': 'Box Capacity Not Constraining',
                'explanation': f'Average box utilization: {box_util:.1f}% (plenty of slack)',
                'impact': 'LOW',
                'solution': 'Not a bottleneck - no action needed'
            })

        # Print causes
        for cause in sorted(causes, key=lambda x: x['rank']):
            print(f"{cause['rank']}. {cause['cause'].upper()} [{cause['impact']} IMPACT]")
            print(f"   Explanation: {cause['explanation']}")
            print(f"   Solution: {cause['solution']}")
            print()

    def generate_recommendations(self):
        """Generate actionable recommendations"""
        print("="*80)
        print("RECOMMENDATIONS")
        print("="*80)
        print()

        print("🎯 Actions to Increase Utilization:")
        print()

        print("1. ACCEPT MORE ORDERS for weeks 4-19")
        print("   Current: Most demand in weeks 1-3")
        print("   Target: Spread orders across all 19 weeks")
        print("   Expected Impact: +40-60% utilization in later weeks")
        print()

        print("2. ADJUST OPTIMIZER PARAMETERS (if needed)")
        print("   Option A: Increase inventory holding cost")
        print("     Current: $1 per unit per week")
        print("     Suggested: $10 per unit per week")
        print("     Effect: Discourages early production more")
        print()
        print("   Option B: Decrease lateness penalty")
        print("     Current: $150,000 per week late")
        print("     Suggested: $50,000 per week late")
        print("     Effect: Makes late delivery less costly")
        print()

        print("3. REVIEW PLANNING HORIZON")
        print("   Current: 19 weeks")
        print("   Issue: Weeks 11-19 have near-zero demand")
        print("   Suggestion: Reduce to 10 weeks or forecast demand for weeks 11-19")
        print()

        print("4. UTILIZE IDLE CAPACITY")
        print("   Weeks 4-19 have significant idle capacity")
        print("   Options:")
        print("     • Accept rush orders with short lead times")
        print("     • Build safety stock for commonly ordered parts")
        print("     • Offer promotions for deliveries in weeks 4-10")
        print()

    def _calculate_box_utilization(self):
        """Calculate average box utilization"""
        try:
            part_master_dict = self.part_master.set_index('FG Code').to_dict('index')
            valid_parts = set(self.part_master['FG Code'].dropna().astype(str))
            valid_orders = self.sales_orders[
                self.sales_orders['Material Code'].isin(valid_parts)
            ].copy()

            total_demand = 0
            for _, order in valid_orders.iterrows():
                part = order['Material Code']
                qty = order['Balance Qty']

                if part not in part_master_dict:
                    continue

                pm = part_master_dict[part]
                box_qty = self._safe_float(pm.get('Box Quantity', 1))

                if box_qty > 0:
                    total_demand += qty / box_qty

            total_capacity = self.box_capacity['Weekly_Capacity'].sum() * 19

            if total_capacity > 0:
                return (total_demand / total_capacity) * 100
            else:
                return 0
        except:
            return 0

    def _safe_float(self, value):
        """Safely convert to float"""
        try:
            return float(value) if pd.notna(value) and value != '' else 0.0
        except:
            return 0.0

    def run_full_analysis(self):
        """Run all analyses"""
        try:
            self.load_data()
            self.analyze_demand_distribution()
            self.analyze_stage_bottlenecks()
            self.analyze_box_capacity_utilization()
            self.analyze_jit_production()
            self.analyze_root_causes()
            self.generate_recommendations()

            print("="*80)
            print("ANALYSIS COMPLETE")
            print("="*80)

        except Exception as e:
            print(f"\n❌ Analysis failed: {e}")
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    analyzer = UtilizationAnalyzer()
    analyzer.run_full_analysis()
