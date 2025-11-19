"""
Order Risk Dashboard
====================
Classifies orders by risk level and identifies at-risk deliveries.
"""

import pandas as pd
from dataclasses import dataclass
from typing import List, Dict
from collections import defaultdict


@dataclass
class OrderRisk:
    """Risk assessment for a single order."""
    order_id: str
    part_code: str
    customer: str
    ordered_qty: int
    delivered_qty: int
    unmet_qty: int
    due_week: int
    current_week: int
    weeks_until_due: int
    risk_level: str  # Critical, High, Medium, Low
    risk_factors: List[str]
    fulfillment_pct: float


@dataclass
class RiskSummary:
    """Summary of order risk analysis."""
    total_orders: int
    critical_orders: int
    high_risk_orders: int
    medium_risk_orders: int
    low_risk_orders: int
    total_at_risk_qty: int
    at_risk_customers: List[str]


class OrderRiskAnalyzer:
    """
    Analyzes orders and classifies them by risk level.

    Usage:
        analyzer = OrderRiskAnalyzer(comprehensive_output_path)
        risks = analyzer.analyze()
        df = analyzer.to_dataframe()
    """

    def __init__(self, comprehensive_output_path: str, current_week: int = 1):
        """
        Args:
            comprehensive_output_path: Path to production_plan_COMPREHENSIVE_test.xlsx
            current_week: Current planning week (default 1)
        """
        self.output_path = comprehensive_output_path
        self.current_week = current_week
        self.data = {}
        self._load_data()

    def _load_data(self):
        """Load relevant sheets from comprehensive output."""
        print("  📊 Loading optimizer results for risk analysis...")

        try:
            xl_file = pd.ExcelFile(self.output_path)

            sheets_to_load = [
                'Order_Fulfillment',
                'Part_Fulfillment',
                'Weekly_Summary',
                'Shipment_Allocation'
            ]

            for sheet in sheets_to_load:
                if sheet in xl_file.sheet_names:
                    self.data[sheet] = pd.read_excel(xl_file, sheet_name=sheet)

            print(f"    ✓ Loaded {len(self.data)} sheets")

        except Exception as e:
            print(f"    ❌ Error loading data: {e}")
            raise

    def analyze(self) -> List[OrderRisk]:
        """Analyze all orders and classify by risk."""
        print("  🔍 Analyzing order risks...")

        order_fulfillment = self.data.get('Order_Fulfillment', pd.DataFrame())

        if order_fulfillment.empty:
            print("    ⚠ No Order_Fulfillment data available")
            return []

        risks = []

        for _, row in order_fulfillment.iterrows():
            order_id = str(row.get('Sales_Order_No', row.get('Order_ID', 'Unknown')))
            part_code = str(row.get('Material_Code', row.get('Part', 'Unknown')))
            customer = str(row.get('Customer', 'Unknown'))

            ordered_qty = int(row.get('Ordered_Qty', row.get('Balance_Qty', 0)) or 0)
            delivered_qty = int(row.get('Delivered_Qty', row.get('Fulfilled_Qty', 0)) or 0)
            unmet_qty = ordered_qty - delivered_qty

            # Parse due week
            due_week = self._parse_week(row.get('Committed_Week', row.get('Due_Week', 1)))

            weeks_until_due = due_week - self.current_week
            fulfillment_pct = (delivered_qty / ordered_qty * 100) if ordered_qty > 0 else 100

            # Determine risk factors and level
            risk_factors = []
            risk_level = 'Low'

            # Factor 1: Fulfillment percentage
            if fulfillment_pct == 0:
                risk_factors.append('No production scheduled')
                risk_level = 'Critical'
            elif fulfillment_pct < 50:
                risk_factors.append(f'Low fulfillment ({fulfillment_pct:.0f}%)')
                risk_level = 'High'
            elif fulfillment_pct < 100:
                risk_factors.append(f'Partial fulfillment ({fulfillment_pct:.0f}%)')
                risk_level = 'Medium'

            # Factor 2: Time until due
            if weeks_until_due <= 0 and unmet_qty > 0:
                risk_factors.append('Already past due')
                risk_level = 'Critical'
            elif weeks_until_due <= 1 and unmet_qty > 0:
                risk_factors.append('Due this/next week')
                if risk_level not in ['Critical']:
                    risk_level = 'High'
            elif weeks_until_due <= 2 and unmet_qty > 0:
                risk_factors.append('Due within 2 weeks')
                if risk_level not in ['Critical', 'High']:
                    risk_level = 'Medium'

            # Factor 3: Quantity size
            if unmet_qty > 100:
                risk_factors.append(f'Large unmet qty ({unmet_qty})')

            # If fully fulfilled, low risk
            if unmet_qty == 0:
                risk_level = 'Low'
                risk_factors = ['Fully fulfilled']

            risks.append(OrderRisk(
                order_id=order_id,
                part_code=part_code,
                customer=customer,
                ordered_qty=ordered_qty,
                delivered_qty=delivered_qty,
                unmet_qty=unmet_qty,
                due_week=due_week,
                current_week=self.current_week,
                weeks_until_due=weeks_until_due,
                risk_level=risk_level,
                risk_factors=risk_factors,
                fulfillment_pct=round(fulfillment_pct, 1)
            ))

        # Sort by risk level and due week
        risk_order = {'Critical': 0, 'High': 1, 'Medium': 2, 'Low': 3}
        risks.sort(key=lambda r: (risk_order.get(r.risk_level, 4), r.due_week))

        print(f"    ✓ Analyzed {len(risks)} orders")

        return risks

    def _parse_week(self, week_value) -> int:
        """Parse week value from various formats."""
        if pd.isna(week_value):
            return 1

        week_str = str(week_value)

        # Handle 'W5' format
        if week_str.startswith('W'):
            try:
                return int(week_str[1:])
            except ValueError:
                return 1

        # Handle numeric
        try:
            return int(float(week_str))
        except (ValueError, TypeError):
            return 1

    def get_summary(self) -> RiskSummary:
        """Get summary statistics of risk analysis."""
        risks = self.analyze()

        critical = sum(1 for r in risks if r.risk_level == 'Critical')
        high = sum(1 for r in risks if r.risk_level == 'High')
        medium = sum(1 for r in risks if r.risk_level == 'Medium')
        low = sum(1 for r in risks if r.risk_level == 'Low')

        at_risk_qty = sum(r.unmet_qty for r in risks if r.risk_level in ['Critical', 'High'])

        at_risk_customers = list(set(
            r.customer for r in risks
            if r.risk_level in ['Critical', 'High']
        ))

        return RiskSummary(
            total_orders=len(risks),
            critical_orders=critical,
            high_risk_orders=high,
            medium_risk_orders=medium,
            low_risk_orders=low,
            total_at_risk_qty=at_risk_qty,
            at_risk_customers=at_risk_customers
        )

    def to_dataframe(self) -> pd.DataFrame:
        """Convert risk analysis to DataFrame for Excel export."""
        risks = self.analyze()

        if not risks:
            return pd.DataFrame({
                'Note': ['No orders found in Order_Fulfillment']
            })

        records = []
        for r in risks:
            records.append({
                'Order_ID': r.order_id,
                'Part': r.part_code,
                'Customer': r.customer,
                'Ordered_Qty': r.ordered_qty,
                'Delivered_Qty': r.delivered_qty,
                'Unmet_Qty': r.unmet_qty,
                'Due_Week': f'W{r.due_week}',
                'Weeks_Until_Due': r.weeks_until_due,
                'Fulfillment_%': r.fulfillment_pct,
                'Risk_Level': r.risk_level,
                'Risk_Factors': '; '.join(r.risk_factors)
            })

        return pd.DataFrame(records)

    def get_risk_by_customer(self) -> pd.DataFrame:
        """Get risk summary by customer."""
        risks = self.analyze()

        customer_data = defaultdict(lambda: {
            'total_orders': 0,
            'critical': 0,
            'high': 0,
            'unmet_qty': 0
        })

        for r in risks:
            customer_data[r.customer]['total_orders'] += 1
            if r.risk_level == 'Critical':
                customer_data[r.customer]['critical'] += 1
            elif r.risk_level == 'High':
                customer_data[r.customer]['high'] += 1
            customer_data[r.customer]['unmet_qty'] += r.unmet_qty

        records = []
        for customer, data in sorted(
            customer_data.items(),
            key=lambda x: x[1]['critical'] + x[1]['high'],
            reverse=True
        ):
            records.append({
                'Customer': customer,
                'Total_Orders': data['total_orders'],
                'Critical_Orders': data['critical'],
                'High_Risk_Orders': data['high'],
                'Total_Unmet_Qty': data['unmet_qty'],
                'Priority': 'High' if data['critical'] > 0 else
                           'Medium' if data['high'] > 0 else 'Normal'
            })

        return pd.DataFrame(records)

    def get_risk_by_week(self) -> pd.DataFrame:
        """Get risk summary by due week."""
        risks = self.analyze()

        week_data = defaultdict(lambda: {
            'orders': 0,
            'critical': 0,
            'high': 0,
            'unmet_qty': 0
        })

        for r in risks:
            week_data[r.due_week]['orders'] += 1
            if r.risk_level == 'Critical':
                week_data[r.due_week]['critical'] += 1
            elif r.risk_level == 'High':
                week_data[r.due_week]['high'] += 1
            week_data[r.due_week]['unmet_qty'] += r.unmet_qty

        records = []
        for week in sorted(week_data.keys()):
            data = week_data[week]
            records.append({
                'Week': f'W{week}',
                'Total_Orders': data['orders'],
                'Critical': data['critical'],
                'High_Risk': data['high'],
                'Unmet_Qty': data['unmet_qty']
            })

        return pd.DataFrame(records)
