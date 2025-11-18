# Decision Support System - Detailed Implementation Plan

## Table of Contents
1. [Architecture Overview](#1-architecture-overview)
2. [Phase 1: Bottleneck Analysis Module](#2-phase-1-bottleneck-analysis-module)
3. [Phase 2: Available-to-Promise (ATP) Calculator](#3-phase-2-available-to-promise-atp-calculator)
4. [Phase 3: Order Priority & Risk Dashboard](#4-phase-3-order-priority--risk-dashboard)
5. [Phase 4: Recommendations Engine](#5-phase-4-recommendations-engine)
6. [Phase 5: What-If Scenario Analyzer](#6-phase-5-what-if-scenario-analyzer)
7. [Output File Structure](#7-output-file-structure)
8. [Integration with Existing Code](#8-integration-with-existing-code)
9. [Database Schema (Future)](#9-database-schema-future)
10. [API Design (Future)](#10-api-design-future)

---

## 1. Architecture Overview

### Current vs Proposed Architecture

```
CURRENT ARCHITECTURE:
┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│ Excel Input │ ──► │ LP Optimizer│ ──► │ Excel Output│
└─────────────┘     └─────────────┘     └─────────────┘

PROPOSED ARCHITECTURE:
┌─────────────┐     ┌─────────────┐     ┌──────────────────┐
│ Excel Input │ ──► │ LP Optimizer│ ──► │ Results Analyzer │
└─────────────┘     └─────────────┘     └────────┬─────────┘
                                                 │
                    ┌────────────────────────────┼────────────────────────────┐
                    │                            │                            │
                    ▼                            ▼                            ▼
          ┌─────────────────┐          ┌─────────────────┐          ┌─────────────────┐
          │ Bottleneck      │          │ ATP Calculator  │          │ Recommendations │
          │ Analyzer        │          │                 │          │ Engine          │
          └────────┬────────┘          └────────┬────────┘          └────────┬────────┘
                   │                            │                            │
                   └────────────────────────────┼────────────────────────────┘
                                                │
                                                ▼
                                    ┌─────────────────────┐
                                    │ Decision Support    │
                                    │ Dashboard (Excel)   │
                                    └─────────────────────┘
```

### New File Structure

```
PP-C/
├── production_plan_test.py                    # Core optimizer (existing)
├── production_plan_executive_test7sheets.py   # Executive reports (existing)
├── decision_support/                          # NEW MODULE
│   ├── __init__.py
│   ├── bottleneck_analyzer.py                 # Phase 1
│   ├── atp_calculator.py                      # Phase 2
│   ├── order_risk_dashboard.py                # Phase 3
│   ├── recommendations_engine.py              # Phase 4
│   ├── scenario_analyzer.py                   # Phase 5
│   └── report_generator.py                    # Consolidated reporting
├── run_decision_support.py                    # Main entry point
└── Master_Data_*.xlsx                         # Input files (existing)
```

---

## 2. Phase 1: Bottleneck Analysis Module

### 2.1 Purpose
Identify which resources are blocking order fulfillment and quantify their impact.

### 2.2 File: `decision_support/bottleneck_analyzer.py`

```python
"""
Bottleneck Analysis Module
==========================
Identifies capacity constraints blocking order fulfillment.
"""

import pandas as pd
from collections import defaultdict
from dataclasses import dataclass
from typing import List, Dict, Tuple

@dataclass
class BottleneckInfo:
    """Information about a single bottleneck."""
    resource_code: str
    resource_name: str
    operation: str  # Casting, Grinding, MC1, SP1, etc.
    week: int
    utilization_pct: float
    capacity_hours: float
    used_hours: float
    overflow_hours: float  # Hours exceeding capacity
    orders_affected: List[str]  # Order IDs delayed by this bottleneck
    parts_affected: List[str]  # Part codes affected
    units_delayed: int

@dataclass
class BottleneckReport:
    """Complete bottleneck analysis report."""
    bottlenecks: List[BottleneckInfo]
    summary_by_resource: Dict[str, float]  # resource -> total overflow hours
    summary_by_week: Dict[int, int]  # week -> number of bottlenecks
    critical_path: List[str]  # Most constrained resources in order
    total_orders_affected: int
    total_units_delayed: int


class BottleneckAnalyzer:
    """
    Analyzes optimizer results to identify bottlenecks.

    Usage:
        analyzer = BottleneckAnalyzer(
            optimizer_results=results_dict,
            model=optimization_model,
            config=config,
            machine_manager=machine_manager,
            orders_meta=orders_metadata
        )
        report = analyzer.analyze()
    """

    def __init__(self, optimizer_results: dict, model, config,
                 machine_manager, orders_meta: List[dict]):
        """
        Args:
            optimizer_results: Dictionary from ComprehensiveResultsAnalyzer
            model: The solved PuLP optimization model
            config: ProductionConfig instance
            machine_manager: MachineResourceManager instance
            orders_meta: List of order metadata dicts with keys:
                - order_id, part, ordered_qty, due_week, delivered_qty
        """
        self.results = optimizer_results
        self.model = model
        self.config = config
        self.machine_manager = machine_manager
        self.orders_meta = orders_meta
        self.weeks = list(range(1, config.PLANNING_WEEKS + 1))

    def analyze(self) -> BottleneckReport:
        """Run complete bottleneck analysis."""
        bottlenecks = []

        # Analyze each resource type
        bottlenecks.extend(self._analyze_casting_bottlenecks())
        bottlenecks.extend(self._analyze_grinding_bottlenecks())
        bottlenecks.extend(self._analyze_machining_bottlenecks())
        bottlenecks.extend(self._analyze_painting_bottlenecks())
        bottlenecks.extend(self._analyze_box_bottlenecks())

        # Sort by severity (overflow hours * units delayed)
        bottlenecks.sort(
            key=lambda b: b.overflow_hours * b.units_delayed,
            reverse=True
        )

        # Build summary
        summary_by_resource = defaultdict(float)
        summary_by_week = defaultdict(int)

        for b in bottlenecks:
            summary_by_resource[b.resource_code] += b.overflow_hours
            summary_by_week[b.week] += 1

        # Critical path = top 5 resources by total overflow
        critical_path = sorted(
            summary_by_resource.keys(),
            key=lambda r: summary_by_resource[r],
            reverse=True
        )[:5]

        # Count affected orders (deduplicated)
        all_affected_orders = set()
        total_units = 0
        for b in bottlenecks:
            all_affected_orders.update(b.orders_affected)
            total_units += b.units_delayed

        return BottleneckReport(
            bottlenecks=bottlenecks,
            summary_by_resource=dict(summary_by_resource),
            summary_by_week=dict(summary_by_week),
            critical_path=critical_path,
            total_orders_affected=len(all_affected_orders),
            total_units_delayed=total_units
        )

    def _analyze_casting_bottlenecks(self) -> List[BottleneckInfo]:
        """Analyze Big Line and Small Line casting bottlenecks."""
        bottlenecks = []
        weekly_summary = self.results.get('weekly_summary', pd.DataFrame())

        if weekly_summary.empty:
            return bottlenecks

        for _, row in weekly_summary.iterrows():
            week = int(row['Week'])

            # Big Line analysis
            big_util = row.get('Big_Line_Util_%', 0)
            if big_util > 95:  # Threshold for bottleneck
                big_hours = row.get('Big_Line_Hours', 0)
                big_cap = self.config.BIG_LINE_HOURS_PER_WEEK
                overflow = max(0, big_hours - big_cap)

                # Find affected orders
                affected = self._find_affected_orders(
                    week, 'casting', 'Big Line'
                )

                if overflow > 0 or big_util >= 100:
                    bottlenecks.append(BottleneckInfo(
                        resource_code='BIG_LINE',
                        resource_name='Big Line Casting',
                        operation='Casting',
                        week=week,
                        utilization_pct=big_util,
                        capacity_hours=big_cap,
                        used_hours=big_hours,
                        overflow_hours=overflow,
                        orders_affected=affected['orders'],
                        parts_affected=affected['parts'],
                        units_delayed=affected['units']
                    ))

            # Small Line analysis
            small_util = row.get('Small_Line_Util_%', 0)
            if small_util > 95:
                small_hours = row.get('Small_Line_Hours', 0)
                small_cap = self.config.SMALL_LINE_HOURS_PER_WEEK
                overflow = max(0, small_hours - small_cap)

                affected = self._find_affected_orders(
                    week, 'casting', 'Small Line'
                )

                if overflow > 0 or small_util >= 100:
                    bottlenecks.append(BottleneckInfo(
                        resource_code='SMALL_LINE',
                        resource_name='Small Line Casting',
                        operation='Casting',
                        week=week,
                        utilization_pct=small_util,
                        capacity_hours=small_cap,
                        used_hours=small_hours,
                        overflow_hours=overflow,
                        orders_affected=affected['orders'],
                        parts_affected=affected['parts'],
                        units_delayed=affected['units']
                    ))

        return bottlenecks

    def _analyze_grinding_bottlenecks(self) -> List[BottleneckInfo]:
        """Analyze grinding resource bottlenecks."""
        bottlenecks = []

        # Get grinding capacity from machine manager
        grinding_cap = self.machine_manager.get_aggregated_capacity('Grinding')
        if grinding_cap == 0:
            return bottlenecks

        # Calculate weekly grinding hours from results
        grinding_plan = self.results.get('grinding_plan', pd.DataFrame())
        if grinding_plan.empty:
            return bottlenecks

        # Implementation: Group by week, sum hours, compare to capacity
        # Similar pattern to casting analysis

        return bottlenecks

    def _analyze_machining_bottlenecks(self) -> List[BottleneckInfo]:
        """Analyze MC1/MC2/MC3 resource bottlenecks."""
        bottlenecks = []

        # Get all unique machining resources
        machining_resources = self.machine_manager.get_resources_by_operation(
            'Machining'
        )

        for resource_code in machining_resources:
            cap = self.machine_manager.get_machine_capacity(resource_code)
            if cap == 0:
                continue

            # Calculate weekly usage from MC1/MC2/MC3 plans
            for week in self.weeks:
                used_hours = self._calculate_resource_usage(
                    resource_code, week, ['mc1', 'mc2', 'mc3']
                )

                util_pct = (used_hours / cap * 100) if cap > 0 else 0

                if util_pct > 95:
                    overflow = max(0, used_hours - cap)
                    affected = self._find_affected_orders(
                        week, 'machining', resource_code
                    )

                    bottlenecks.append(BottleneckInfo(
                        resource_code=resource_code,
                        resource_name=self.machine_manager.get_resource_name(
                            resource_code
                        ),
                        operation='Machining',
                        week=week,
                        utilization_pct=util_pct,
                        capacity_hours=cap,
                        used_hours=used_hours,
                        overflow_hours=overflow,
                        orders_affected=affected['orders'],
                        parts_affected=affected['parts'],
                        units_delayed=affected['units']
                    ))

        return bottlenecks

    def _analyze_painting_bottlenecks(self) -> List[BottleneckInfo]:
        """Analyze SP1/SP2/SP3 resource bottlenecks."""
        # Similar structure to machining analysis
        bottlenecks = []
        # Implementation...
        return bottlenecks

    def _analyze_box_bottlenecks(self) -> List[BottleneckInfo]:
        """Analyze mould box capacity bottlenecks."""
        bottlenecks = []
        # Implementation...
        return bottlenecks

    def _find_affected_orders(self, week: int, stage: str,
                              resource: str) -> Dict:
        """
        Find orders affected by a bottleneck.

        Returns:
            dict with keys: orders (list), parts (list), units (int)
        """
        affected_orders = []
        affected_parts = set()
        affected_units = 0

        for order in self.orders_meta:
            # Check if order uses this resource in this week
            # and is not fully delivered
            if (order['due_week'] >= week and
                order['delivered_qty'] < order['ordered_qty']):

                unmet = order['ordered_qty'] - order['delivered_qty']
                if unmet > 0:
                    affected_orders.append(order['order_id'])
                    affected_parts.add(order['part'])
                    affected_units += unmet

        return {
            'orders': affected_orders,
            'parts': list(affected_parts),
            'units': affected_units
        }

    def _calculate_resource_usage(self, resource_code: str, week: int,
                                  stages: List[str]) -> float:
        """Calculate total hours used by a resource in a week."""
        total_hours = 0.0

        # Sum usage across specified stages
        for stage in stages:
            plan_key = f'{stage}_plan'
            plan = self.results.get(plan_key, pd.DataFrame())

            if plan.empty:
                continue

            # Filter for this week and resource
            # Sum hours based on cycle times and quantities

        return total_hours

    def get_bottleneck_summary_df(self) -> pd.DataFrame:
        """Generate DataFrame summary for Excel export."""
        report = self.analyze()

        rows = []
        for b in report.bottlenecks:
            rows.append({
                'Week': b.week,
                'Resource': b.resource_name,
                'Operation': b.operation,
                'Utilization_%': round(b.utilization_pct, 1),
                'Capacity_Hours': round(b.capacity_hours, 1),
                'Used_Hours': round(b.used_hours, 1),
                'Overflow_Hours': round(b.overflow_hours, 1),
                'Orders_Affected': len(b.orders_affected),
                'Units_Delayed': b.units_delayed,
                'Severity': 'Critical' if b.utilization_pct >= 100 else 'High'
            })

        return pd.DataFrame(rows)

    def get_constraint_shadow_prices(self) -> Dict[str, float]:
        """
        Extract shadow prices from solved LP model.

        Shadow price = marginal value of relaxing a constraint by 1 unit.
        High shadow price = bottleneck (adding capacity here helps most).
        """
        shadow_prices = {}

        for name, constraint in self.model.model.constraints.items():
            # PuLP stores dual values after solving
            if hasattr(constraint, 'pi') and constraint.pi is not None:
                shadow_prices[name] = constraint.pi

        # Sort by absolute value (most impactful)
        sorted_prices = sorted(
            shadow_prices.items(),
            key=lambda x: abs(x[1]),
            reverse=True
        )

        return dict(sorted_prices[:50])  # Top 50 constraints
```

---

## 3. Phase 2: Available-to-Promise (ATP) Calculator

### 3.1 Purpose
Determine if a new order can be fulfilled and find the earliest possible delivery date.

### 3.2 File: `decision_support/atp_calculator.py`

```python
"""
Available-to-Promise (ATP) Calculator
=====================================
Calculates feasibility and delivery dates for new orders.
"""

import pandas as pd
from datetime import datetime, timedelta
from dataclasses import dataclass
from typing import List, Dict, Optional, Tuple
import copy

@dataclass
class ATPResult:
    """Result of ATP calculation."""
    is_feasible: bool
    requested_date: datetime
    earliest_delivery_date: Optional[datetime]
    earliest_delivery_week: Optional[int]
    delay_weeks: int  # 0 if on-time or early
    blocking_resources: List[Dict]  # Resources preventing earlier delivery
    capacity_usage: Dict[str, float]  # Resource -> % capacity used by this order
    confidence: str  # 'High', 'Medium', 'Low'
    notes: List[str]

@dataclass
class NewOrderRequest:
    """New order to check ATP for."""
    part_code: str
    quantity: int
    requested_delivery_date: datetime
    customer_name: Optional[str] = None
    priority: str = 'Normal'  # 'High', 'Normal', 'Low'


class ATPCalculator:
    """
    Calculate Available-to-Promise for new orders.

    Usage:
        calculator = ATPCalculator(
            current_schedule=optimizer_results,
            part_master=part_master_df,
            machine_manager=machine_manager,
            box_manager=box_manager,
            config=config
        )

        # Check single order
        result = calculator.check_order(NewOrderRequest(
            part_code='DGC-001',
            quantity=500,
            requested_delivery_date=datetime(2025, 12, 15)
        ))

        # Find earliest date
        earliest = calculator.find_earliest_delivery(
            part_code='DGC-001',
            quantity=500
        )
    """

    def __init__(self, current_schedule: dict, part_master: pd.DataFrame,
                 machine_manager, box_manager, config):
        self.schedule = current_schedule
        self.part_master = part_master
        self.machine_manager = machine_manager
        self.box_manager = box_manager
        self.config = config
        self.weeks = list(range(1, config.PLANNING_WEEKS + 1))

        # Pre-calculate available capacity per resource per week
        self._calculate_available_capacity()

    def _calculate_available_capacity(self):
        """Calculate remaining capacity for each resource each week."""
        self.available_capacity = defaultdict(lambda: defaultdict(float))

        weekly_summary = self.schedule.get('weekly_summary', pd.DataFrame())

        for week in self.weeks:
            # Casting - Big Line
            big_cap = self.config.BIG_LINE_HOURS_PER_WEEK
            big_used = 0
            if not weekly_summary.empty:
                week_data = weekly_summary[weekly_summary['Week'] == week]
                if len(week_data) > 0:
                    big_used = week_data.iloc[0].get('Big_Line_Hours', 0)
            self.available_capacity['BIG_LINE'][week] = max(0, big_cap - big_used)

            # Casting - Small Line
            small_cap = self.config.SMALL_LINE_HOURS_PER_WEEK
            small_used = 0
            if not weekly_summary.empty and len(week_data) > 0:
                small_used = week_data.iloc[0].get('Small_Line_Hours', 0)
            self.available_capacity['SMALL_LINE'][week] = max(0, small_cap - small_used)

            # Machining resources
            for resource in self.machine_manager.get_all_resources():
                cap = self.machine_manager.get_machine_capacity(resource)
                used = self._get_resource_usage(resource, week)
                self.available_capacity[resource][week] = max(0, cap - used)

            # Box capacities
            for box_size in self.box_manager.get_all_sizes():
                cap = self.box_manager.get_capacity(box_size)
                used = self._get_box_usage(box_size, week)
                self.available_capacity[f'BOX_{box_size}'][week] = max(0, cap - used)

    def check_order(self, order: NewOrderRequest) -> ATPResult:
        """
        Check if a new order can be delivered by requested date.
        """
        # Get part parameters
        part_params = self._get_part_params(order.part_code)
        if not part_params:
            return ATPResult(
                is_feasible=False,
                requested_date=order.requested_delivery_date,
                earliest_delivery_date=None,
                earliest_delivery_week=None,
                delay_weeks=0,
                blocking_resources=[],
                capacity_usage={},
                confidence='Low',
                notes=[f"Part {order.part_code} not found in Part Master"]
            )

        # Calculate required capacity
        required_capacity = self._calculate_required_capacity(
            part_params, order.quantity
        )

        # Convert requested date to week number
        requested_week = self._date_to_week(order.requested_delivery_date)

        # Check capacity availability working backwards from delivery
        feasibility, blocking = self._check_capacity_chain(
            part_params, required_capacity, requested_week
        )

        # Find earliest possible delivery
        earliest_week = self._find_earliest_week(
            part_params, required_capacity
        )
        earliest_date = self._week_to_date(earliest_week)

        delay_weeks = max(0, earliest_week - requested_week)

        # Calculate capacity usage percentages
        capacity_usage = {}
        for resource, hours in required_capacity.items():
            if self.available_capacity[resource][earliest_week] > 0:
                usage_pct = hours / self.available_capacity[resource][earliest_week] * 100
                capacity_usage[resource] = round(usage_pct, 1)

        # Determine confidence
        if feasibility and delay_weeks == 0:
            confidence = 'High'
        elif feasibility and delay_weeks <= 2:
            confidence = 'Medium'
        else:
            confidence = 'Low'

        notes = []
        if delay_weeks > 0:
            notes.append(f"Order will be {delay_weeks} week(s) late")
        if blocking:
            notes.append(f"Blocked by: {', '.join([b['resource'] for b in blocking])}")

        return ATPResult(
            is_feasible=feasibility,
            requested_date=order.requested_delivery_date,
            earliest_delivery_date=earliest_date,
            earliest_delivery_week=earliest_week,
            delay_weeks=delay_weeks,
            blocking_resources=blocking,
            capacity_usage=capacity_usage,
            confidence=confidence,
            notes=notes
        )

    def find_earliest_delivery(self, part_code: str,
                               quantity: int) -> Tuple[datetime, int]:
        """
        Find the earliest possible delivery date for a part/quantity.

        Returns:
            Tuple of (earliest_date, week_number)
        """
        part_params = self._get_part_params(part_code)
        if not part_params:
            return None, None

        required_capacity = self._calculate_required_capacity(
            part_params, quantity
        )

        earliest_week = self._find_earliest_week(part_params, required_capacity)
        earliest_date = self._week_to_date(earliest_week)

        return earliest_date, earliest_week

    def check_multiple_orders(self, orders: List[NewOrderRequest]) -> List[ATPResult]:
        """
        Check ATP for multiple orders, considering cumulative capacity impact.

        Orders are processed in sequence - earlier orders consume capacity
        that affects later orders.
        """
        results = []

        # Make a copy of available capacity to modify
        temp_capacity = copy.deepcopy(self.available_capacity)

        for order in orders:
            result = self._check_order_with_capacity(order, temp_capacity)
            results.append(result)

            # If feasible, consume capacity
            if result.is_feasible and result.earliest_delivery_week:
                self._consume_capacity(
                    temp_capacity, order, result.earliest_delivery_week
                )

        return results

    def _calculate_required_capacity(self, part_params: dict,
                                     quantity: int) -> Dict[str, float]:
        """Calculate capacity required at each resource for this order."""
        required = {}

        # Casting hours
        casting_cycle = part_params.get('casting_cycle', 0)
        moulding_line = part_params.get('moulding_line', '')

        if 'Big Line' in moulding_line:
            required['BIG_LINE'] = (casting_cycle * quantity) / 60.0
        elif 'Small Line' in moulding_line:
            required['SMALL_LINE'] = (casting_cycle * quantity) / 60.0

        # Grinding hours
        grind_cycle = part_params.get('grind_cycle', 0)
        grind_batch = max(1, part_params.get('grind_batch', 1))
        if grind_cycle > 0:
            required['GRINDING'] = (grind_cycle / 60.0) * quantity / grind_batch

        # Machining hours (MC1, MC2, MC3)
        for i, stage in enumerate(['mc1', 'mc2', 'mc3']):
            resources = part_params.get('mach_resources', [])
            cycles = part_params.get('mach_cycles', [])
            batches = part_params.get('mach_batches', [])

            if i < len(resources) and resources[i]:
                cycle = cycles[i] if i < len(cycles) else 0
                batch = max(1, batches[i] if i < len(batches) else 1)
                if cycle > 0:
                    required[resources[i]] = (cycle / 60.0) * quantity / batch

        # Painting hours (SP1, SP2, SP3)
        for i, stage in enumerate(['sp1', 'sp2', 'sp3']):
            resources = part_params.get('paint_resources', [])
            cycles = part_params.get('paint_cycles', [])
            batches = part_params.get('paint_batches', [])

            if i < len(resources) and resources[i]:
                cycle = cycles[i] if i < len(cycles) else 0
                batch = max(1, batches[i] if i < len(batches) else 1)
                if cycle > 0:
                    required[resources[i]] = (cycle / 60.0) * quantity / batch

        # Box capacity (moulds)
        box_size = part_params.get('box_size', '')
        box_qty = max(1, part_params.get('box_quantity', 1))
        if box_size:
            required[f'BOX_{box_size}'] = quantity / box_qty

        return required

    def _check_capacity_chain(self, part_params: dict,
                              required: Dict[str, float],
                              delivery_week: int) -> Tuple[bool, List[Dict]]:
        """
        Check if capacity is available through the production chain.
        """
        blocking = []
        lead_time = max(2, part_params.get('lead_time_weeks', 2))

        # Simple model: everything needs to fit in (delivery_week - lead_time)
        production_week = max(1, delivery_week - lead_time)

        for resource, hours_needed in required.items():
            available = self.available_capacity[resource][production_week]

            if hours_needed > available:
                blocking.append({
                    'resource': resource,
                    'week': production_week,
                    'needed': round(hours_needed, 1),
                    'available': round(available, 1),
                    'shortfall': round(hours_needed - available, 1)
                })

        is_feasible = len(blocking) == 0
        return is_feasible, blocking

    def _find_earliest_week(self, part_params: dict,
                            required: Dict[str, float]) -> int:
        """Find earliest week where all capacity is available."""
        lead_time = max(2, part_params.get('lead_time_weeks', 2))

        for delivery_week in self.weeks:
            production_week = max(1, delivery_week - lead_time)

            all_available = True
            for resource, hours_needed in required.items():
                if hours_needed > self.available_capacity[resource][production_week]:
                    all_available = False
                    break

            if all_available:
                return delivery_week

        # No week found within planning horizon
        return self.config.PLANNING_WEEKS + 1

    def _get_part_params(self, part_code: str) -> Optional[dict]:
        """Get part parameters from part master."""
        row = self.part_master[self.part_master['FG Code'] == part_code]
        if row.empty:
            return None

        row = row.iloc[0]
        # Extract all relevant parameters
        # (Similar to ComprehensiveParameterBuilder)

        return {
            'casting_cycle': row.get('Casting Cycle time (min)', 0),
            'moulding_line': row.get('Moulding Line', ''),
            'grind_cycle': row.get('Grinding Cycle time (min)', 0),
            'grind_batch': row.get('Grinding batch size', 1),
            # ... etc
        }

    def _date_to_week(self, date: datetime) -> int:
        """Convert date to week number."""
        days_diff = (date - self.config.CURRENT_DATE).days
        return max(1, (days_diff // 7) + 1)

    def _week_to_date(self, week: int) -> datetime:
        """Convert week number to date (end of week)."""
        return self.config.CURRENT_DATE + timedelta(weeks=week)

    def get_capacity_forecast_df(self) -> pd.DataFrame:
        """Generate capacity availability forecast as DataFrame."""
        rows = []

        for week in self.weeks:
            row = {'Week': week}
            for resource in self.available_capacity:
                cap = self.available_capacity[resource][week]
                row[f'{resource}_Available'] = round(cap, 1)
            rows.append(row)

        return pd.DataFrame(rows)
```

---

## 4. Phase 3: Order Priority & Risk Dashboard

### 4.1 Purpose
Classify orders by risk level and provide actionable prioritization.

### 4.2 File: `decision_support/order_risk_dashboard.py`

```python
"""
Order Risk Dashboard
====================
Classifies orders by risk level and generates priority actions.
"""

import pandas as pd
from datetime import datetime
from dataclasses import dataclass
from typing import List, Dict
from enum import Enum

class RiskLevel(Enum):
    CRITICAL = 'Critical'  # Will miss delivery, action required NOW
    HIGH = 'High'          # At risk, needs attention this week
    MEDIUM = 'Medium'      # Tight schedule, monitor closely
    LOW = 'Low'            # On track with buffer
    SAFE = 'Safe'          # Comfortable buffer

@dataclass
class OrderRiskInfo:
    """Risk assessment for a single order."""
    order_id: str
    part_code: str
    customer: str
    quantity: int
    due_date: datetime
    due_week: int
    scheduled_delivery_week: int
    delivered_qty: int
    unmet_qty: int
    risk_level: RiskLevel
    days_buffer: int  # Positive = early, negative = late
    blocking_stage: str  # Which stage is delayed
    action_required: str  # Recommended action

@dataclass
class RiskDashboard:
    """Complete risk dashboard."""
    orders: List[OrderRiskInfo]
    summary: Dict[RiskLevel, int]  # Count per risk level
    critical_orders: List[OrderRiskInfo]
    action_items: List[str]


class OrderRiskAnalyzer:
    """
    Analyze orders and classify by risk level.
    """

    def __init__(self, orders_meta: List[dict], schedule: dict,
                 bottlenecks, config):
        self.orders = orders_meta
        self.schedule = schedule
        self.bottlenecks = bottlenecks
        self.config = config

    def analyze(self) -> RiskDashboard:
        """Generate complete risk dashboard."""
        risk_orders = []

        for order in self.orders:
            risk_info = self._assess_order_risk(order)
            risk_orders.append(risk_info)

        # Sort by risk (highest first)
        risk_priority = {
            RiskLevel.CRITICAL: 0,
            RiskLevel.HIGH: 1,
            RiskLevel.MEDIUM: 2,
            RiskLevel.LOW: 3,
            RiskLevel.SAFE: 4
        }
        risk_orders.sort(key=lambda o: (risk_priority[o.risk_level], -o.unmet_qty))

        # Summary counts
        summary = {level: 0 for level in RiskLevel}
        for order in risk_orders:
            summary[order.risk_level] += 1

        # Extract critical orders
        critical = [o for o in risk_orders if o.risk_level == RiskLevel.CRITICAL]

        # Generate action items
        action_items = self._generate_action_items(risk_orders)

        return RiskDashboard(
            orders=risk_orders,
            summary=summary,
            critical_orders=critical,
            action_items=action_items
        )

    def _assess_order_risk(self, order: dict) -> OrderRiskInfo:
        """Assess risk level for a single order."""
        order_id = order.get('order_id', 'Unknown')
        part = order['part']
        quantity = order['ordered_qty']
        delivered = order.get('delivered_qty', 0)
        due_week = order['due_week']

        # Find scheduled delivery week from schedule
        scheduled_week = self._find_scheduled_delivery(part, due_week)

        # Calculate buffer
        buffer_weeks = due_week - scheduled_week
        buffer_days = buffer_weeks * 7

        # Determine risk level
        unmet = quantity - delivered

        if unmet > 0 and scheduled_week > due_week + 2:
            risk = RiskLevel.CRITICAL
            action = f"Expedite {part} - will miss by {scheduled_week - due_week} weeks"
        elif unmet > 0 and scheduled_week > due_week:
            risk = RiskLevel.HIGH
            action = f"Prioritize {part} - at risk of {scheduled_week - due_week} week delay"
        elif unmet > 0 and buffer_weeks <= 1:
            risk = RiskLevel.MEDIUM
            action = f"Monitor {part} - tight schedule, {buffer_days} days buffer"
        elif unmet > 0:
            risk = RiskLevel.LOW
            action = f"On track - {buffer_days} days buffer"
        else:
            risk = RiskLevel.SAFE
            action = "Fully delivered"

        # Find blocking stage
        blocking = self._find_blocking_stage(part, due_week)

        return OrderRiskInfo(
            order_id=order_id,
            part_code=part,
            customer=order.get('customer', 'Unknown'),
            quantity=quantity,
            due_date=self._week_to_date(due_week),
            due_week=due_week,
            scheduled_delivery_week=scheduled_week,
            delivered_qty=delivered,
            unmet_qty=unmet,
            risk_level=risk,
            days_buffer=buffer_days,
            blocking_stage=blocking,
            action_required=action
        )

    def _generate_action_items(self, orders: List[OrderRiskInfo]) -> List[str]:
        """Generate prioritized action items."""
        actions = []

        # Group critical orders by blocking stage
        critical_by_stage = {}
        for order in orders:
            if order.risk_level == RiskLevel.CRITICAL:
                stage = order.blocking_stage
                if stage not in critical_by_stage:
                    critical_by_stage[stage] = []
                critical_by_stage[stage].append(order)

        # Generate actions for each blocking stage
        for stage, stage_orders in critical_by_stage.items():
            total_units = sum(o.unmet_qty for o in stage_orders)
            actions.append(
                f"URGENT: Clear {stage} backlog - {len(stage_orders)} orders, "
                f"{total_units} units affected"
            )

        return actions

    def get_risk_summary_df(self) -> pd.DataFrame:
        """Generate DataFrame for Excel export."""
        dashboard = self.analyze()

        rows = []
        for order in dashboard.orders:
            rows.append({
                'Order_ID': order.order_id,
                'Part': order.part_code,
                'Customer': order.customer,
                'Quantity': order.quantity,
                'Due_Date': order.due_date.strftime('%Y-%m-%d'),
                'Due_Week': order.due_week,
                'Scheduled_Week': order.scheduled_delivery_week,
                'Delivered': order.delivered_qty,
                'Unmet': order.unmet_qty,
                'Risk_Level': order.risk_level.value,
                'Buffer_Days': order.days_buffer,
                'Blocking_Stage': order.blocking_stage,
                'Action': order.action_required
            })

        return pd.DataFrame(rows)
```

---

## 5. Phase 4: Recommendations Engine

### 5.1 Purpose
Generate actionable recommendations to improve fulfillment.

### 5.2 File: `decision_support/recommendations_engine.py`

```python
"""
Recommendations Engine
======================
Generates actionable recommendations to improve fulfillment.
"""

import pandas as pd
from dataclasses import dataclass
from typing import List, Dict, Optional
from enum import Enum

class RecommendationType(Enum):
    OVERTIME = 'Overtime'
    OUTSOURCE = 'Outsource'
    RESCHEDULE = 'Reschedule'
    PRIORITIZE = 'Prioritize'
    ADD_CAPACITY = 'Add Capacity'
    REDUCE_WIP = 'Reduce WIP'

class TimeFrame(Enum):
    IMMEDIATE = 'Immediate'  # This week
    SHORT_TERM = 'Short-term'  # Next 2 weeks
    MEDIUM_TERM = 'Medium-term'  # Next month
    STRATEGIC = 'Strategic'  # Longer term

@dataclass
class Recommendation:
    """Single recommendation."""
    id: str
    type: RecommendationType
    timeframe: TimeFrame
    title: str
    description: str
    resource_affected: str
    impact_orders: int  # Orders that would benefit
    impact_units: int  # Units that would be fulfilled
    estimated_cost: Optional[float]  # Cost to implement
    estimated_hours: Optional[float]  # Hours needed
    priority_score: float  # 0-100, higher = more impactful
    implementation_steps: List[str]


class RecommendationsEngine:
    """
    Generate recommendations based on bottlenecks and risk analysis.
    """

    def __init__(self, bottleneck_report, risk_dashboard, schedule,
                 config, machine_manager):
        self.bottlenecks = bottleneck_report
        self.risk = risk_dashboard
        self.schedule = schedule
        self.config = config
        self.machine_manager = machine_manager

        # Cost parameters (configurable)
        self.overtime_cost_per_hour = 500  # INR
        self.outsource_cost_multiplier = 1.5  # 50% premium

    def generate(self) -> List[Recommendation]:
        """Generate all recommendations."""
        recommendations = []

        # Overtime recommendations
        recommendations.extend(self._generate_overtime_recommendations())

        # Outsourcing recommendations
        recommendations.extend(self._generate_outsource_recommendations())

        # Rescheduling recommendations
        recommendations.extend(self._generate_reschedule_recommendations())

        # Prioritization recommendations
        recommendations.extend(self._generate_priority_recommendations())

        # Strategic recommendations
        recommendations.extend(self._generate_strategic_recommendations())

        # Sort by priority score
        recommendations.sort(key=lambda r: r.priority_score, reverse=True)

        # Assign IDs
        for i, rec in enumerate(recommendations, 1):
            rec.id = f"REC-{i:03d}"

        return recommendations

    def _generate_overtime_recommendations(self) -> List[Recommendation]:
        """Generate overtime recommendations for bottleneck resources."""
        recommendations = []

        for bottleneck in self.bottlenecks.bottlenecks[:5]:  # Top 5 bottlenecks
            if bottleneck.overflow_hours <= 0:
                continue

            # Calculate overtime needed
            overtime_hours = min(
                bottleneck.overflow_hours,
                20  # Cap at 20 hours overtime per week
            )

            # Estimate impact
            if bottleneck.capacity_hours > 0:
                capacity_increase = overtime_hours / bottleneck.capacity_hours
                units_enabled = int(bottleneck.units_delayed * capacity_increase)
            else:
                units_enabled = 0

            cost = overtime_hours * self.overtime_cost_per_hour

            # Priority score based on orders affected and overflow
            priority = min(100, (
                bottleneck.overflow_hours * 5 +
                len(bottleneck.orders_affected) * 2
            ))

            recommendations.append(Recommendation(
                id='',  # Assigned later
                type=RecommendationType.OVERTIME,
                timeframe=TimeFrame.IMMEDIATE,
                title=f"Add {overtime_hours:.0f} hrs overtime to {bottleneck.resource_name}",
                description=(
                    f"Week {bottleneck.week}: {bottleneck.resource_name} is at "
                    f"{bottleneck.utilization_pct:.0f}% utilization with "
                    f"{bottleneck.overflow_hours:.1f} hours overflow. "
                    f"Adding overtime would enable {units_enabled} additional units."
                ),
                resource_affected=bottleneck.resource_code,
                impact_orders=len(bottleneck.orders_affected),
                impact_units=units_enabled,
                estimated_cost=cost,
                estimated_hours=overtime_hours,
                priority_score=priority,
                implementation_steps=[
                    f"1. Approve overtime budget: INR {cost:,.0f}",
                    f"2. Schedule {overtime_hours:.0f} hours for Week {bottleneck.week}",
                    f"3. Notify operators and arrange logistics",
                    f"4. Prioritize orders: {', '.join(bottleneck.orders_affected[:3])}"
                ]
            ))

        return recommendations

    def _generate_outsource_recommendations(self) -> List[Recommendation]:
        """Generate outsourcing recommendations for persistent bottlenecks."""
        recommendations = []

        # Find resources with multi-week bottlenecks
        resource_overflow = {}
        for b in self.bottlenecks.bottlenecks:
            if b.resource_code not in resource_overflow:
                resource_overflow[b.resource_code] = {
                    'total_overflow': 0,
                    'weeks': 0,
                    'units': 0,
                    'orders': set(),
                    'name': b.resource_name
                }
            resource_overflow[b.resource_code]['total_overflow'] += b.overflow_hours
            resource_overflow[b.resource_code]['weeks'] += 1
            resource_overflow[b.resource_code]['units'] += b.units_delayed
            resource_overflow[b.resource_code]['orders'].update(b.orders_affected)

        for resource, data in resource_overflow.items():
            if data['weeks'] >= 2 and data['total_overflow'] > 10:
                # Persistent bottleneck - recommend outsourcing

                priority = min(100, data['total_overflow'] * 3 + len(data['orders']) * 2)

                recommendations.append(Recommendation(
                    id='',
                    type=RecommendationType.OUTSOURCE,
                    timeframe=TimeFrame.SHORT_TERM,
                    title=f"Outsource {data['total_overflow']:.0f} hrs of {data['name']}",
                    description=(
                        f"{data['name']} is bottlenecked for {data['weeks']} weeks "
                        f"with {data['total_overflow']:.0f} total overflow hours. "
                        f"Outsourcing would clear backlog for {data['units']} units."
                    ),
                    resource_affected=resource,
                    impact_orders=len(data['orders']),
                    impact_units=data['units'],
                    estimated_cost=None,  # Requires vendor quotes
                    estimated_hours=data['total_overflow'],
                    priority_score=priority,
                    implementation_steps=[
                        "1. Identify qualified vendors for this operation",
                        "2. Request quotes for the work volume",
                        "3. Arrange material transfer and quality specs",
                        f"4. Target completion by Week {self.bottlenecks.bottlenecks[0].week + 2}"
                    ]
                ))

        return recommendations

    def _generate_strategic_recommendations(self) -> List[Recommendation]:
        """Generate longer-term strategic recommendations."""
        recommendations = []

        # Analyze if adding machines would help
        for resource in self.bottlenecks.critical_path[:3]:
            total_overflow = self.bottlenecks.summary_by_resource.get(resource, 0)

            if total_overflow > 50:  # Significant persistent overflow
                recommendations.append(Recommendation(
                    id='',
                    type=RecommendationType.ADD_CAPACITY,
                    timeframe=TimeFrame.STRATEGIC,
                    title=f"Evaluate additional {resource} capacity",
                    description=(
                        f"{resource} has {total_overflow:.0f} hours of overflow "
                        f"across the planning horizon. Adding capacity would "
                        f"increase throughput by ~{total_overflow/self.config.PLANNING_WEEKS:.0f} "
                        f"hours per week."
                    ),
                    resource_affected=resource,
                    impact_orders=self.bottlenecks.total_orders_affected,
                    impact_units=self.bottlenecks.total_units_delayed,
                    estimated_cost=None,  # Requires capex analysis
                    estimated_hours=total_overflow,
                    priority_score=60,  # Lower priority (strategic)
                    implementation_steps=[
                        "1. Conduct detailed capacity analysis",
                        "2. Evaluate machine options and costs",
                        "3. Calculate ROI based on order pipeline",
                        "4. Prepare capital expenditure proposal"
                    ]
                ))

        return recommendations

    def get_recommendations_df(self) -> pd.DataFrame:
        """Generate DataFrame for Excel export."""
        recommendations = self.generate()

        rows = []
        for rec in recommendations:
            rows.append({
                'ID': rec.id,
                'Type': rec.type.value,
                'Timeframe': rec.timeframe.value,
                'Title': rec.title,
                'Resource': rec.resource_affected,
                'Orders_Impacted': rec.impact_orders,
                'Units_Impacted': rec.impact_units,
                'Est_Cost_INR': rec.estimated_cost,
                'Est_Hours': rec.estimated_hours,
                'Priority_Score': round(rec.priority_score, 1),
                'Description': rec.description
            })

        return pd.DataFrame(rows)
```

---

## 6. Phase 5: What-If Scenario Analyzer

### 6.1 Purpose
Simulate impact of changes before implementing them.

### 6.2 File: `decision_support/scenario_analyzer.py`

```python
"""
What-If Scenario Analyzer
=========================
Simulate impact of capacity changes, new orders, etc.
"""

import pandas as pd
from dataclasses import dataclass
from typing import List, Dict, Optional
import copy

@dataclass
class ScenarioResult:
    """Result of scenario simulation."""
    scenario_name: str
    baseline_fulfillment: float
    new_fulfillment: float
    fulfillment_change: float
    baseline_ontime: float
    new_ontime: float
    ontime_change: float
    orders_improved: int
    units_improved: int
    cost_estimate: Optional[float]
    roi_estimate: Optional[float]  # Orders improved per cost unit
    details: Dict

@dataclass
class Scenario:
    """Scenario definition."""
    name: str
    description: str
    changes: List[Dict]  # List of changes to apply
    # Each change: {'type': 'add_capacity'|'add_shift'|'add_order'|..., ...params}


class ScenarioAnalyzer:
    """
    Analyze what-if scenarios.

    Usage:
        analyzer = ScenarioAnalyzer(
            base_optimizer=optimizer,
            base_results=results,
            config=config
        )

        # Define scenario
        scenario = Scenario(
            name="Add VMC-03",
            description="Add second VMC-03 machine",
            changes=[{
                'type': 'add_capacity',
                'resource': 'VMC-03',
                'additional_hours': 64.8  # 1 machine worth
            }]
        )

        result = analyzer.simulate(scenario)
    """

    def __init__(self, base_optimizer, base_results, config,
                 data_loader, machine_manager):
        self.base_optimizer = base_optimizer
        self.base_results = base_results
        self.config = config
        self.data_loader = data_loader
        self.machine_manager = machine_manager

        # Store baseline metrics
        self.baseline = self._calculate_baseline_metrics()

    def simulate(self, scenario: Scenario) -> ScenarioResult:
        """
        Simulate a scenario by re-running optimizer with modified parameters.

        This is computationally expensive - runs full optimization.
        """
        # Create modified configuration
        modified_config = copy.deepcopy(self.config)
        modified_machine_manager = copy.deepcopy(self.machine_manager)

        # Apply scenario changes
        for change in scenario.changes:
            self._apply_change(change, modified_config, modified_machine_manager)

        # Re-run optimization with modified parameters
        new_results = self._run_modified_optimization(
            modified_config, modified_machine_manager
        )

        # Calculate new metrics
        new_metrics = self._calculate_metrics(new_results)

        # Compare to baseline
        fulfillment_change = new_metrics['fulfillment_pct'] - self.baseline['fulfillment_pct']
        ontime_change = new_metrics['ontime_pct'] - self.baseline['ontime_pct']

        # Estimate cost
        cost = self._estimate_scenario_cost(scenario)

        # Calculate ROI
        orders_improved = new_metrics['fulfilled_orders'] - self.baseline['fulfilled_orders']
        roi = orders_improved / cost if cost and cost > 0 else None

        return ScenarioResult(
            scenario_name=scenario.name,
            baseline_fulfillment=self.baseline['fulfillment_pct'],
            new_fulfillment=new_metrics['fulfillment_pct'],
            fulfillment_change=fulfillment_change,
            baseline_ontime=self.baseline['ontime_pct'],
            new_ontime=new_metrics['ontime_pct'],
            ontime_change=ontime_change,
            orders_improved=orders_improved,
            units_improved=new_metrics.get('units_improved', 0),
            cost_estimate=cost,
            roi_estimate=roi,
            details=new_metrics
        )

    def compare_scenarios(self, scenarios: List[Scenario]) -> pd.DataFrame:
        """Compare multiple scenarios side by side."""
        results = []

        for scenario in scenarios:
            result = self.simulate(scenario)
            results.append({
                'Scenario': result.scenario_name,
                'Fulfillment_%': round(result.new_fulfillment, 1),
                'Change_%': round(result.fulfillment_change, 1),
                'OnTime_%': round(result.new_ontime, 1),
                'OnTime_Change_%': round(result.ontime_change, 1),
                'Orders_Improved': result.orders_improved,
                'Est_Cost': result.cost_estimate,
                'ROI': round(result.roi_estimate, 4) if result.roi_estimate else None
            })

        return pd.DataFrame(results)
```

---

## 7. Output File Structure

### 7.1 New Excel Report: `production_plan_DECISION_SUPPORT.xlsx`

| Sheet Name | Content | Source Module |
|------------|---------|---------------|
| **1_EXECUTIVE_SUMMARY** | One-page overview with KPIs | All modules |
| **2_BOTTLENECK_ANALYSIS** | Resource bottlenecks by week | BottleneckAnalyzer |
| **3_ORDER_RISK_DASHBOARD** | Orders by risk level | OrderRiskAnalyzer |
| **4_CAPACITY_AVAILABILITY** | Available capacity forecast | ATPCalculator |
| **5_RECOMMENDATIONS** | Actionable recommendations | RecommendationsEngine |
| **6_IMPLEMENTATION_PLAN** | Step-by-step actions | RecommendationsEngine |
| **7_ATP_TEMPLATE** | Template for checking new orders | ATPCalculator |
| **8_SCENARIO_COMPARISON** | What-if scenario results | ScenarioAnalyzer |

---

## 8. Integration with Existing Code

### 8.1 Main Entry Point: `run_decision_support.py`

```python
"""
Decision Support System - Main Entry Point
"""

import argparse
from datetime import datetime

# Existing modules
from production_plan_test import (
    ProductionConfig,
    ComprehensiveDataLoader,
    WIPDemandCalculator,
    ComprehensiveParameterBuilder,
    MachineResourceManager,
    BoxCapacityManager,
    ComprehensiveOptimizationModel,
    ComprehensiveResultsAnalyzer
)

# New decision support modules
from decision_support.bottleneck_analyzer import BottleneckAnalyzer
from decision_support.atp_calculator import ATPCalculator, NewOrderRequest
from decision_support.order_risk_dashboard import OrderRiskAnalyzer
from decision_support.recommendations_engine import RecommendationsEngine
from decision_support.report_generator import DecisionSupportReportGenerator


def main():
    parser = argparse.ArgumentParser(description='Production Planning Decision Support')
    parser.add_argument('--input', default='Master_Data_Updated_Nov_Dec.xlsx',
                        help='Input Excel file')
    parser.add_argument('--atp-check', help='Check ATP: "PART,QTY,DATE"')
    args = parser.parse_args()

    print("="*80)
    print("PRODUCTION PLANNING DECISION SUPPORT SYSTEM")
    print("="*80)

    # Step 1: Load data and run optimization (existing code)
    config = ProductionConfig()
    loader = ComprehensiveDataLoader(args.input, config)
    data = loader.load_all_data()

    # ... (existing optimization steps)

    # Step 2: Run decision support analysis
    print("\n" + "="*80)
    print("RUNNING DECISION SUPPORT ANALYSIS")
    print("="*80)

    # 2a. Bottleneck Analysis
    bottleneck_analyzer = BottleneckAnalyzer(...)
    bottleneck_report = bottleneck_analyzer.analyze()

    # 2b. Order Risk Analysis
    risk_analyzer = OrderRiskAnalyzer(...)
    risk_dashboard = risk_analyzer.analyze()

    # 2c. ATP Calculator setup
    atp_calculator = ATPCalculator(...)

    # 2d. Generate Recommendations
    recommendations_engine = RecommendationsEngine(...)
    recommendations = recommendations_engine.generate()

    # Step 3: Generate Decision Support Report
    report_generator = DecisionSupportReportGenerator(...)
    report_generator.generate_report('production_plan_DECISION_SUPPORT.xlsx')

    print("\n✅ Decision support report generated")


if __name__ == '__main__':
    main()
```

### 8.2 Modifications to Existing Code

Add to `MachineResourceManager`:

```python
def get_all_resources(self) -> List[str]:
    """Get list of all resource codes."""
    return list(self.machine_capacity.keys())

def get_resources_by_operation(self, operation: str) -> List[str]:
    """Get resources for a specific operation type."""
    pass

def set_machine_capacity(self, resource: str, capacity: float):
    """Set capacity for a resource (for scenario analysis)."""
    self.machine_capacity[resource] = capacity
```

---

## 9. Database Schema (Future)

For future web-based version:

```sql
CREATE TABLE orders (
    order_id VARCHAR(50) PRIMARY KEY,
    part_code VARCHAR(50) NOT NULL,
    customer_id VARCHAR(50),
    quantity INT NOT NULL,
    due_date DATE NOT NULL,
    status VARCHAR(20) DEFAULT 'Pending',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE production_schedules (
    schedule_id SERIAL PRIMARY KEY,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    planning_start DATE NOT NULL,
    planning_weeks INT NOT NULL,
    fulfillment_pct DECIMAL(5,2),
    ontime_pct DECIMAL(5,2)
);

CREATE TABLE bottlenecks (
    bottleneck_id SERIAL PRIMARY KEY,
    schedule_id INT REFERENCES production_schedules(schedule_id),
    resource_code VARCHAR(50),
    week INT,
    utilization_pct DECIMAL(5,2),
    overflow_hours DECIMAL(10,2),
    orders_affected INT
);

CREATE TABLE atp_checks (
    check_id SERIAL PRIMARY KEY,
    checked_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    part_code VARCHAR(50),
    quantity INT,
    requested_date DATE,
    is_feasible BOOLEAN,
    earliest_date DATE
);
```

---

## 10. API Design (Future)

```python
from fastapi import FastAPI
from pydantic import BaseModel

app = FastAPI(title="Production Planning Decision Support API")

@app.post("/api/v1/atp/check")
async def check_atp(part_code: str, quantity: int, requested_date: date):
    """Check if an order can be delivered by requested date."""
    pass

@app.get("/api/v1/orders/risk")
async def get_order_risks(risk_level: Optional[str] = None):
    """Get orders by risk level."""
    pass

@app.get("/api/v1/bottlenecks")
async def get_bottlenecks(week: Optional[int] = None):
    """Get current bottlenecks."""
    pass

@app.get("/api/v1/recommendations")
async def get_recommendations():
    """Get actionable recommendations."""
    pass

@app.get("/api/v1/dashboard/summary")
async def get_dashboard_summary():
    """Get executive dashboard summary."""
    pass
```

---

## Implementation Timeline

| Phase | Duration | Deliverables |
|-------|----------|--------------|
| **Phase 1: Bottleneck Analysis** | 1-2 weeks | `bottleneck_analyzer.py`, bottleneck report sheet |
| **Phase 2: ATP Calculator** | 2-3 weeks | `atp_calculator.py`, capacity forecast, ATP template |
| **Phase 3: Risk Dashboard** | 1-2 weeks | `order_risk_dashboard.py`, risk classification |
| **Phase 4: Recommendations** | 2-3 weeks | `recommendations_engine.py`, action items |
| **Phase 5: Scenarios** | 2-3 weeks | `scenario_analyzer.py`, comparison tool |
| **Integration & Testing** | 2 weeks | `run_decision_support.py`, full report |

**Total: 10-15 weeks** for complete implementation

---

## Key Success Metrics

1. **Fulfillment Rate Improvement**: Target +10-15%
2. **On-Time Delivery Improvement**: Target +20%
3. **New Order Conversion**: Orders accepted / Orders inquired
4. **Bottleneck Resolution Time**: Days from identification to resolution
5. **Recommendation Adoption Rate**: Recommendations implemented / Total
6. **ATP Accuracy**: Actual delivery date vs. predicted

---

*Last updated: 2025-11-18*
