"""
Available-to-Promise (ATP) Calculator
=====================================
Calculates available capacity to promise new orders.
"""

import pandas as pd
from dataclasses import dataclass
from typing import Dict, List, Optional
from datetime import datetime, timedelta


@dataclass
class ATPResult:
    """Result of ATP calculation for a potential order."""
    part_code: str
    requested_qty: int
    requested_week: int
    is_feasible: bool
    earliest_delivery_week: int
    capacity_gaps: Dict[str, float]  # resource -> gap in hours/units
    limiting_resource: str
    confidence: str  # High, Medium, Low
    notes: List[str]


@dataclass
class CapacitySlot:
    """Available capacity for a resource in a week."""
    resource: str
    week: int
    total_capacity: float
    used_capacity: float
    available_capacity: float
    unit: str


class ATPCalculator:
    """
    Calculates available capacity to promise new orders.

    Usage:
        calculator = ATPCalculator(comprehensive_output_path)
        result = calculator.check_order('PART-001', qty=100, requested_week=5)
        capacity_df = calculator.get_available_capacity()
    """

    def __init__(self, comprehensive_output_path: str):
        """
        Args:
            comprehensive_output_path: Path to production_plan_COMPREHENSIVE_test.xlsx
        """
        self.output_path = comprehensive_output_path
        self.data = {}
        self._load_data()
        self._calculate_available_capacity()

    def _load_data(self):
        """Load relevant sheets from comprehensive output."""
        print("  📊 Loading optimizer results for ATP calculation...")

        try:
            xl_file = pd.ExcelFile(self.output_path)

            sheets_to_load = [
                'Weekly_Summary',
                'Part_Parameters',
                'Machine_Utilization',
                'Box_Utilization'
            ]

            for sheet in sheets_to_load:
                if sheet in xl_file.sheet_names:
                    self.data[sheet] = pd.read_excel(xl_file, sheet_name=sheet)

            print(f"    ✓ Loaded {len(self.data)} sheets")

        except Exception as e:
            print(f"    ❌ Error loading data: {e}")
            raise

    def _calculate_available_capacity(self):
        """Calculate available capacity by resource and week."""
        self.capacity_slots = []
        weekly_summary = self.data.get('Weekly_Summary', pd.DataFrame())

        if weekly_summary.empty:
            return

        # Define resources and their capacity columns
        resources = [
            ('Big_Line', 'Big_Line_Hours', 'Big_Line_Capacity_Hours', 'hours'),
            ('Small_Line', 'Small_Line_Hours', 'Small_Line_Capacity_Hours', 'hours'),
            ('Casting', 'Casting_Tons', None, 'tons'),
            ('Grinding', 'Grinding_Units', None, 'units'),
            ('MC1', 'MC1_Units', None, 'units'),
            ('MC2', 'MC2_Units', None, 'units'),
            ('MC3', 'MC3_Units', None, 'units'),
            ('SP1', 'SP1_Units', None, 'units'),
            ('SP2', 'SP2_Units', None, 'units'),
            ('SP3', 'SP3_Units', None, 'units'),
        ]

        for _, row in weekly_summary.iterrows():
            week = int(row.get('Week', 0))
            if week == 0:
                continue

            for resource, used_col, cap_col, unit in resources:
                used = float(row.get(used_col, 0) or 0)

                # Get capacity - try explicit column first, then estimate from utilization
                if cap_col and cap_col in row:
                    capacity = float(row.get(cap_col, 0) or 0)
                else:
                    # Estimate from utilization if available
                    util_col = f'{resource}_Util_%'
                    if util_col in row:
                        util = float(row.get(util_col, 0) or 0)
                        capacity = used / (util / 100) if util > 0 else used * 1.5
                    else:
                        # Default estimate
                        capacity = used * 1.2 if used > 0 else 1000

                available = max(0, capacity - used)

                self.capacity_slots.append(CapacitySlot(
                    resource=resource,
                    week=week,
                    total_capacity=round(capacity, 1),
                    used_capacity=round(used, 1),
                    available_capacity=round(available, 1),
                    unit=unit
                ))

    def check_order(self, part_code: str, qty: int, requested_week: int) -> ATPResult:
        """
        Check if a new order can be fulfilled.

        Args:
            part_code: Part/material code
            qty: Requested quantity
            requested_week: Requested delivery week

        Returns:
            ATPResult with feasibility assessment
        """
        print(f"  🔍 Checking ATP for {part_code}: {qty} units by W{requested_week}...")

        # Get part parameters
        part_params = self._get_part_parameters(part_code)

        if not part_params:
            return ATPResult(
                part_code=part_code,
                requested_qty=qty,
                requested_week=requested_week,
                is_feasible=False,
                earliest_delivery_week=0,
                capacity_gaps={},
                limiting_resource='Unknown',
                confidence='Low',
                notes=[f'Part {part_code} not found in Part_Parameters']
            )

        # Calculate required capacity for each stage
        required_capacity = self._calculate_required_capacity(part_params, qty)

        # Check availability for requested week and find earliest possible
        capacity_gaps = {}
        limiting_resource = None
        max_gap = 0

        # Check if requested week is feasible
        is_feasible = True
        notes = []

        for resource, required in required_capacity.items():
            available = self._get_available_capacity(resource, requested_week)

            if available < required:
                gap = required - available
                capacity_gaps[resource] = gap
                is_feasible = False

                if gap > max_gap:
                    max_gap = gap
                    limiting_resource = resource

        # Find earliest delivery week
        earliest_week = self._find_earliest_week(required_capacity, requested_week)

        # Determine confidence
        if is_feasible:
            confidence = 'High'
            notes.append(f'✓ Order can be fulfilled by W{requested_week}')
        elif earliest_week <= requested_week + 2:
            confidence = 'Medium'
            notes.append(f'⚠ Partial capacity available, earliest: W{earliest_week}')
        else:
            confidence = 'Low'
            notes.append(f'❌ Significant capacity constraints, earliest: W{earliest_week}')

        if limiting_resource:
            notes.append(f'Limiting resource: {limiting_resource}')

        return ATPResult(
            part_code=part_code,
            requested_qty=qty,
            requested_week=requested_week,
            is_feasible=is_feasible,
            earliest_delivery_week=earliest_week,
            capacity_gaps=capacity_gaps,
            limiting_resource=limiting_resource or 'None',
            confidence=confidence,
            notes=notes
        )

    def _get_part_parameters(self, part_code: str) -> Optional[Dict]:
        """Get production parameters for a part."""
        params_df = self.data.get('Part_Parameters', pd.DataFrame())

        if params_df.empty:
            return None

        # Find part in parameters
        part_col = None
        for col in ['Part', 'Material_Code', 'FG_Code']:
            if col in params_df.columns:
                part_col = col
                break

        if not part_col:
            return None

        part_row = params_df[params_df[part_col] == part_code]

        if part_row.empty:
            return None

        return part_row.iloc[0].to_dict()

    def _calculate_required_capacity(self, part_params: Dict, qty: int) -> Dict[str, float]:
        """Calculate required capacity for each resource to produce qty units."""
        required = {}

        # Unit weight for casting
        unit_weight = float(part_params.get('Unit_Weight_kg', 0) or 0)
        required['Casting'] = (unit_weight * qty) / 1000  # tons

        # Moulding line
        moulding_line = str(part_params.get('Moulding_Line', '')).upper()
        casting_cycle = float(part_params.get('Casting_Cycle_time_min', 0) or 0)

        if 'BIG' in moulding_line:
            required['Big_Line'] = (casting_cycle * qty) / 60  # hours
        else:
            required['Small_Line'] = (casting_cycle * qty) / 60  # hours

        # Grinding
        required['Grinding'] = qty

        # Machining (simplified - same units through all stages)
        required['MC1'] = qty
        required['MC2'] = qty
        required['MC3'] = qty

        # Painting
        required['SP1'] = qty
        required['SP2'] = qty
        required['SP3'] = qty

        return required

    def _get_available_capacity(self, resource: str, week: int) -> float:
        """Get available capacity for a resource in a specific week."""
        for slot in self.capacity_slots:
            if slot.resource == resource and slot.week == week:
                return slot.available_capacity
        return 0

    def _find_earliest_week(self, required_capacity: Dict[str, float],
                           start_week: int) -> int:
        """Find the earliest week when all capacity is available."""
        max_week = max(slot.week for slot in self.capacity_slots) if self.capacity_slots else start_week

        for week in range(start_week, max_week + 1):
            all_available = True

            for resource, required in required_capacity.items():
                available = self._get_available_capacity(resource, week)
                if available < required:
                    all_available = False
                    break

            if all_available:
                return week

        return max_week + 1  # Beyond planning horizon

    def get_available_capacity(self) -> pd.DataFrame:
        """Get available capacity matrix as DataFrame."""
        if not self.capacity_slots:
            return pd.DataFrame({'Note': ['No capacity data available']})

        records = []
        for slot in self.capacity_slots:
            records.append({
                'Week': f'W{slot.week}',
                'Resource': slot.resource,
                'Total_Capacity': slot.total_capacity,
                'Used_Capacity': slot.used_capacity,
                'Available_Capacity': slot.available_capacity,
                'Utilization_%': round(
                    (slot.used_capacity / slot.total_capacity * 100)
                    if slot.total_capacity > 0 else 0, 1
                ),
                'Unit': slot.unit
            })

        df = pd.DataFrame(records)

        # Pivot for easier viewing
        if not df.empty:
            df = df.sort_values(['Resource', 'Week'])

        return df

    def get_capacity_forecast(self) -> pd.DataFrame:
        """Get capacity forecast summary by week."""
        if not self.capacity_slots:
            return pd.DataFrame()

        # Group by week and summarize
        records = []
        weeks = sorted(set(slot.week for slot in self.capacity_slots))

        for week in weeks:
            week_slots = [s for s in self.capacity_slots if s.week == week]

            total_util = sum(
                s.used_capacity / s.total_capacity * 100
                for s in week_slots if s.total_capacity > 0
            )
            avg_util = total_util / len(week_slots) if week_slots else 0

            constrained = sum(
                1 for s in week_slots
                if s.total_capacity > 0 and
                (s.used_capacity / s.total_capacity * 100) >= 85
            )

            records.append({
                'Week': f'W{week}',
                'Avg_Utilization_%': round(avg_util, 1),
                'Constrained_Resources': constrained,
                'Status': 'Critical' if constrained >= 3 else
                         'Tight' if constrained >= 1 else 'OK'
            })

        return pd.DataFrame(records)

    def check_multiple_orders(self, potential_orders: List[Dict]) -> pd.DataFrame:
        """
        Check feasibility of multiple potential orders.

        Args:
            potential_orders: List of dicts with keys: part_code, qty, requested_week

        Returns:
            DataFrame with feasibility results for each order
        """
        results = []

        for order in potential_orders:
            part_code = order.get('part_code', '')
            qty = int(order.get('qty', 0))
            requested_week = int(order.get('requested_week', 1))

            if not part_code or qty <= 0:
                continue

            result = self.check_order(part_code, qty, requested_week)

            results.append({
                'Part_Code': result.part_code,
                'Requested_Qty': result.requested_qty,
                'Requested_Week': f'W{result.requested_week}',
                'Feasible': 'Yes' if result.is_feasible else 'No',
                'Earliest_Delivery': f'W{result.earliest_delivery_week}',
                'Delay_Weeks': max(0, result.earliest_delivery_week - result.requested_week),
                'Limiting_Resource': result.limiting_resource,
                'Confidence': result.confidence,
                'Notes': '; '.join(result.notes)
            })

        if not results:
            return pd.DataFrame({
                'Note': ['No valid orders to check. Add orders with Part_Code, Qty, and Requested_Week.']
            })

        return pd.DataFrame(results)

    def get_available_parts(self) -> List[str]:
        """Get list of parts that can be produced (from Part_Parameters)."""
        params_df = self.data.get('Part_Parameters', pd.DataFrame())

        if params_df.empty:
            return []

        # Find part column
        for col in ['Part', 'Material_Code', 'FG_Code']:
            if col in params_df.columns:
                return sorted(params_df[col].dropna().unique().tolist())

        return []

    def load_orders_from_file(self, input_path: str = 'ATP_INPUT.xlsx') -> List[Dict]:
        """
        Load potential orders from an Excel input file.

        Expected columns: Part_Code, Qty, Requested_Week

        Args:
            input_path: Path to the input Excel file

        Returns:
            List of order dicts
        """
        import os

        if not os.path.exists(input_path):
            return []

        try:
            df = pd.read_excel(input_path)

            # Normalize column names
            df.columns = [col.strip().replace(' ', '_') for col in df.columns]

            orders = []
            for _, row in df.iterrows():
                # Try different column name variations
                part_code = None
                for col in ['Part_Code', 'Part', 'Material_Code', 'FG_Code']:
                    if col in df.columns and pd.notna(row.get(col)):
                        part_code = str(row[col])
                        break

                qty = 0
                for col in ['Qty', 'Quantity', 'Requested_Qty', 'Order_Qty']:
                    if col in df.columns and pd.notna(row.get(col)):
                        try:
                            qty = int(row[col])
                        except:
                            pass
                        break

                week = 1
                for col in ['Requested_Week', 'Week', 'Delivery_Week']:
                    if col in df.columns and pd.notna(row.get(col)):
                        week_val = row[col]
                        if isinstance(week_val, str) and week_val.startswith('W'):
                            week = int(week_val[1:])
                        else:
                            try:
                                week = int(week_val)
                            except:
                                pass
                        break

                if part_code and qty > 0:
                    orders.append({
                        'part_code': part_code,
                        'qty': qty,
                        'requested_week': week
                    })

            print(f"    ✓ Loaded {len(orders)} orders from {input_path}")
            return orders

        except Exception as e:
            print(f"    ⚠ Could not read {input_path}: {e}")
            return []

    def create_atp_template(self, input_path: str = 'ATP_INPUT.xlsx') -> pd.DataFrame:
        """
        Create ATP results from input file or sample entries.

        If ATP_INPUT.xlsx exists, use those orders.
        Otherwise, create sample entries for demonstration.

        Args:
            input_path: Path to optional input file with orders to check
        """
        # Try to load from input file first
        user_orders = self.load_orders_from_file(input_path)

        if user_orders:
            # User provided orders - check those
            print(f"    📋 Checking {len(user_orders)} orders from input file...")
            return self.check_multiple_orders(user_orders)

        # No input file - create sample entries
        print("    ℹ No ATP_INPUT.xlsx found - using sample orders")
        print("    💡 Create ATP_INPUT.xlsx with columns: Part_Code, Qty, Requested_Week")

        # Get available parts for reference
        available_parts = self.get_available_parts()

        # Get weeks with best capacity
        forecast = self.get_capacity_forecast()
        ok_weeks = []
        if not forecast.empty and 'Status' in forecast.columns:
            ok_weeks = forecast[forecast['Status'] == 'OK']['Week'].tolist()

        # Create sample entries
        sample_orders = []

        # Add 3-5 sample entries using actual parts if available
        sample_parts = available_parts[:5] if available_parts else ['PART-001', 'PART-002', 'PART-003']
        sample_weeks = [5, 6, 7, 8, 9]

        for i, part in enumerate(sample_parts):
            week = sample_weeks[i % len(sample_weeks)]
            sample_orders.append({
                'Part_Code': part,
                'Qty': (i + 1) * 50,
                'Requested_Week': week
            })

        # Check these sample orders
        results_df = self.check_multiple_orders(sample_orders)

        return results_df

    def get_best_weeks_for_capacity(self, num_weeks: int = 5) -> List[int]:
        """Get the weeks with most available capacity."""
        if not self.capacity_slots:
            return list(range(1, num_weeks + 1))

        # Calculate average utilization per week
        week_util = {}
        weeks = sorted(set(slot.week for slot in self.capacity_slots))

        for week in weeks:
            week_slots = [s for s in self.capacity_slots if s.week == week]
            if week_slots:
                total_util = sum(
                    s.used_capacity / s.total_capacity * 100
                    for s in week_slots if s.total_capacity > 0
                )
                avg_util = total_util / len(week_slots)
                week_util[week] = avg_util

        # Sort by utilization (lowest first = most capacity available)
        sorted_weeks = sorted(week_util.keys(), key=lambda w: week_util[w])

        return sorted_weeks[:num_weeks]

    def get_capacity_summary_by_week(self) -> pd.DataFrame:
        """
        Get a pivoted capacity summary showing available capacity per resource per week.
        Useful for quick visual assessment of where capacity exists.
        """
        if not self.capacity_slots:
            return pd.DataFrame()

        # Create pivot table
        records = []
        resources = sorted(set(slot.resource for slot in self.capacity_slots))
        weeks = sorted(set(slot.week for slot in self.capacity_slots))

        for resource in resources:
            row = {'Resource': resource}
            for week in weeks:
                slot = next(
                    (s for s in self.capacity_slots
                     if s.resource == resource and s.week == week),
                    None
                )
                if slot:
                    # Show available capacity with utilization indicator
                    util = (slot.used_capacity / slot.total_capacity * 100) if slot.total_capacity > 0 else 0
                    if util >= 100:
                        row[f'W{week}'] = f'{slot.available_capacity:.0f} (FULL)'
                    elif util >= 85:
                        row[f'W{week}'] = f'{slot.available_capacity:.0f} (TIGHT)'
                    else:
                        row[f'W{week}'] = f'{slot.available_capacity:.0f}'
                else:
                    row[f'W{week}'] = '-'
            records.append(row)

        return pd.DataFrame(records)

    def create_input_template(self, output_path: str = 'ATP_INPUT.xlsx'):
        """
        Create a blank ATP input template file for users to fill in.

        Args:
            output_path: Path to create the template file
        """
        import os

        if os.path.exists(output_path):
            print(f"    ⚠ {output_path} already exists - not overwriting")
            return

        # Get some available parts for examples
        available_parts = self.get_available_parts()
        sample_parts = available_parts[:3] if available_parts else ['PART-001', 'PART-002', 'PART-003']

        # Create template with example rows
        template_data = {
            'Part_Code': sample_parts + [''] * 7,  # 3 examples + 7 blank rows
            'Qty': [100, 200, 150] + [None] * 7,
            'Requested_Week': [5, 6, 7] + [None] * 7
        }

        df = pd.DataFrame(template_data)
        df.to_excel(output_path, index=False)

        print(f"    ✓ Created input template: {output_path}")
        print(f"    📝 Edit this file with your potential orders, then re-run the script")
