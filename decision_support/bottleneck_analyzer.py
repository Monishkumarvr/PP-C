"""
Bottleneck Analysis Module
==========================
Identifies capacity constraints blocking order fulfillment.
"""

import pandas as pd
from collections import defaultdict
from dataclasses import dataclass, field
from typing import List, Dict, Optional


@dataclass
class BottleneckInfo:
    """Information about a single bottleneck."""
    resource_code: str
    resource_name: str
    operation: str  # Casting, Grinding, MC1, MC2, MC3, SP1, SP2, SP3
    week: int
    utilization_pct: float
    capacity_value: float  # Hours or units depending on resource
    used_value: float
    overflow_value: float  # Amount exceeding capacity
    capacity_unit: str  # 'hours', 'units', 'tons'
    parts_affected: List[str] = field(default_factory=list)
    severity: str = 'Medium'  # Critical, High, Medium


@dataclass
class BottleneckReport:
    """Complete bottleneck analysis report."""
    bottlenecks: List[BottleneckInfo]
    summary_by_resource: Dict[str, float]  # resource -> total overflow
    summary_by_week: Dict[int, int]  # week -> number of bottlenecks
    critical_path: List[str]  # Most constrained resources in order
    total_bottlenecks: int
    weeks_with_bottlenecks: int


class BottleneckAnalyzer:
    """
    Analyzes optimizer results to identify bottlenecks.

    Reads from the comprehensive output Excel file and identifies
    resources operating at or above capacity.

    Usage:
        analyzer = BottleneckAnalyzer(comprehensive_output_path)
        report = analyzer.analyze()
        df = analyzer.to_dataframe()
    """

    # Utilization thresholds
    CRITICAL_THRESHOLD = 100  # >= 100% is critical
    HIGH_THRESHOLD = 95       # >= 95% is high
    MEDIUM_THRESHOLD = 85     # >= 85% is medium

    def __init__(self, comprehensive_output_path: str):
        """
        Args:
            comprehensive_output_path: Path to production_plan_COMPREHENSIVE_test.xlsx
        """
        self.output_path = comprehensive_output_path
        self.data = {}
        self._load_data()

    def _load_data(self):
        """Load relevant sheets from comprehensive output."""
        print("  📊 Loading optimizer results for bottleneck analysis...")

        try:
            xl_file = pd.ExcelFile(self.output_path)

            # Load key sheets
            sheets_to_load = [
                'Weekly_Summary',
                'Machine_Utilization',
                'Casting',
                'Grinding',
                'Machining_Stage1',
                'Machining_Stage2',
                'Machining_Stage3',
                'Painting_Stage1',
                'Painting_Stage2',
                'Painting_Stage3',
                'Box_Utilization',
                'Order_Fulfillment'
            ]

            for sheet in sheets_to_load:
                if sheet in xl_file.sheet_names:
                    self.data[sheet] = pd.read_excel(xl_file, sheet_name=sheet)

            print(f"    ✓ Loaded {len(self.data)} sheets")

        except Exception as e:
            print(f"    ❌ Error loading data: {e}")
            raise

    def analyze(self) -> BottleneckReport:
        """Run complete bottleneck analysis."""
        print("  🔍 Analyzing bottlenecks...")

        bottlenecks = []

        # Analyze each resource type
        bottlenecks.extend(self._analyze_casting_bottlenecks())
        bottlenecks.extend(self._analyze_stage_bottlenecks('Grinding', 'Grinding'))
        bottlenecks.extend(self._analyze_stage_bottlenecks('MC1', 'Machining_Stage1'))
        bottlenecks.extend(self._analyze_stage_bottlenecks('MC2', 'Machining_Stage2'))
        bottlenecks.extend(self._analyze_stage_bottlenecks('MC3', 'Machining_Stage3'))
        bottlenecks.extend(self._analyze_stage_bottlenecks('SP1', 'Painting_Stage1'))
        bottlenecks.extend(self._analyze_stage_bottlenecks('SP2', 'Painting_Stage2'))
        bottlenecks.extend(self._analyze_stage_bottlenecks('SP3', 'Painting_Stage3'))
        bottlenecks.extend(self._analyze_box_bottlenecks())

        # Sort by severity and utilization
        severity_order = {'Critical': 0, 'High': 1, 'Medium': 2}
        bottlenecks.sort(
            key=lambda b: (severity_order.get(b.severity, 3), -b.utilization_pct)
        )

        # Build summaries
        summary_by_resource = defaultdict(float)
        summary_by_week = defaultdict(int)

        for b in bottlenecks:
            summary_by_resource[b.resource_code] += b.overflow_value
            summary_by_week[b.week] += 1

        # Critical path = top resources by total overflow
        critical_path = sorted(
            summary_by_resource.keys(),
            key=lambda r: summary_by_resource[r],
            reverse=True
        )[:5]

        report = BottleneckReport(
            bottlenecks=bottlenecks,
            summary_by_resource=dict(summary_by_resource),
            summary_by_week=dict(summary_by_week),
            critical_path=critical_path,
            total_bottlenecks=len(bottlenecks),
            weeks_with_bottlenecks=len(summary_by_week)
        )

        print(f"    ✓ Found {len(bottlenecks)} bottlenecks across {len(summary_by_week)} weeks")

        return report

    def _analyze_casting_bottlenecks(self) -> List[BottleneckInfo]:
        """Analyze Big Line and Small Line casting bottlenecks."""
        bottlenecks = []
        weekly_summary = self.data.get('Weekly_Summary', pd.DataFrame())

        if weekly_summary.empty:
            return bottlenecks

        for _, row in weekly_summary.iterrows():
            week = int(row.get('Week', 0))
            if week == 0:
                continue

            # Big Line analysis
            big_util = float(row.get('Big_Line_Util_%', 0) or 0)
            if big_util >= self.MEDIUM_THRESHOLD:
                big_hours = float(row.get('Big_Line_Hours', 0) or 0)
                big_cap = float(row.get('Big_Line_Capacity_Hours', big_hours / (big_util/100)) if big_util > 0 else 0)
                overflow = max(0, big_hours - big_cap)

                severity = self._get_severity(big_util)

                bottlenecks.append(BottleneckInfo(
                    resource_code='BIG_LINE',
                    resource_name='Big Line Casting',
                    operation='Casting',
                    week=week,
                    utilization_pct=round(big_util, 1),
                    capacity_value=round(big_cap, 1),
                    used_value=round(big_hours, 1),
                    overflow_value=round(overflow, 1),
                    capacity_unit='hours',
                    severity=severity
                ))

            # Small Line analysis
            small_util = float(row.get('Small_Line_Util_%', 0) or 0)
            if small_util >= self.MEDIUM_THRESHOLD:
                small_hours = float(row.get('Small_Line_Hours', 0) or 0)
                small_cap = float(row.get('Small_Line_Capacity_Hours', small_hours / (small_util/100)) if small_util > 0 else 0)
                overflow = max(0, small_hours - small_cap)

                severity = self._get_severity(small_util)

                bottlenecks.append(BottleneckInfo(
                    resource_code='SMALL_LINE',
                    resource_name='Small Line Casting',
                    operation='Casting',
                    week=week,
                    utilization_pct=round(small_util, 1),
                    capacity_value=round(small_cap, 1),
                    used_value=round(small_hours, 1),
                    overflow_value=round(overflow, 1),
                    capacity_unit='hours',
                    severity=severity
                ))

        return bottlenecks

    def _analyze_stage_bottlenecks(self, stage_name: str, sheet_name: str) -> List[BottleneckInfo]:
        """Analyze bottlenecks for a production stage using Weekly_Summary."""
        bottlenecks = []
        weekly_summary = self.data.get('Weekly_Summary', pd.DataFrame())

        if weekly_summary.empty:
            return bottlenecks

        # Map stage names to Weekly_Summary columns
        stage_columns = {
            'Grinding': ('Grinding_Units', 'Grinding'),
            'MC1': ('MC1_Units', 'MC1'),
            'MC2': ('MC2_Units', 'MC2'),
            'MC3': ('MC3_Units', 'MC3'),
            'SP1': ('SP1_Units', 'SP1'),
            'SP2': ('SP2_Units', 'SP2'),
            'SP3': ('SP3_Units', 'SP3')
        }

        if stage_name not in stage_columns:
            return bottlenecks

        units_col, display_name = stage_columns[stage_name]

        # Get parts scheduled for this stage to list affected parts
        stage_data = self.data.get(sheet_name, pd.DataFrame())

        for _, row in weekly_summary.iterrows():
            week = int(row.get('Week', 0))
            if week == 0:
                continue

            # Get units for this stage
            units = float(row.get(units_col, 0) or 0)

            # Try to get utilization percentage if available
            util_col = f'{stage_name}_Util_%'
            if util_col in row:
                util_pct = float(row.get(util_col, 0) or 0)
            else:
                # Estimate utilization (assume capacity based on typical values)
                # This is a simplification - actual capacity should come from constraints
                util_pct = 0
                continue  # Skip if we don't have utilization data

            if util_pct >= self.MEDIUM_THRESHOLD:
                # Calculate capacity from utilization
                capacity = units / (util_pct / 100) if util_pct > 0 else 0
                overflow = max(0, units - capacity)

                severity = self._get_severity(util_pct)

                # Find affected parts for this week
                affected_parts = []
                if not stage_data.empty and 'Part' in stage_data.columns:
                    week_col = f'W{week}'
                    if week_col in stage_data.columns:
                        week_parts = stage_data[stage_data[week_col] > 0]['Part'].unique()
                        affected_parts = list(week_parts)

                bottlenecks.append(BottleneckInfo(
                    resource_code=stage_name,
                    resource_name=display_name,
                    operation=self._get_operation_name(stage_name),
                    week=week,
                    utilization_pct=round(util_pct, 1),
                    capacity_value=round(capacity, 0),
                    used_value=round(units, 0),
                    overflow_value=round(overflow, 0),
                    capacity_unit='units',
                    parts_affected=affected_parts,
                    severity=severity
                ))

        return bottlenecks

    def _analyze_box_bottlenecks(self) -> List[BottleneckInfo]:
        """Analyze mould box capacity bottlenecks."""
        bottlenecks = []
        box_util = self.data.get('Box_Utilization', pd.DataFrame())

        if box_util.empty:
            return bottlenecks

        # Box utilization sheet typically has box sizes with weekly utilization
        for _, row in box_util.iterrows():
            box_size = row.get('Box_Size', 'Unknown')

            # Find week columns
            for col in box_util.columns:
                if col.startswith('W') and col[1:].isdigit():
                    week = int(col[1:])
                    util_pct = float(row.get(col, 0) or 0)

                    if util_pct >= self.MEDIUM_THRESHOLD:
                        severity = self._get_severity(util_pct)

                        bottlenecks.append(BottleneckInfo(
                            resource_code=f'BOX_{box_size}',
                            resource_name=f'Mould Box {box_size}',
                            operation='Casting',
                            week=week,
                            utilization_pct=round(util_pct, 1),
                            capacity_value=0,  # Would need actual capacity
                            used_value=0,
                            overflow_value=0,
                            capacity_unit='units',
                            severity=severity
                        ))

        return bottlenecks

    def _get_severity(self, util_pct: float) -> str:
        """Determine severity based on utilization percentage."""
        if util_pct >= self.CRITICAL_THRESHOLD:
            return 'Critical'
        elif util_pct >= self.HIGH_THRESHOLD:
            return 'High'
        else:
            return 'Medium'

    def _get_operation_name(self, stage_name: str) -> str:
        """Get human-readable operation name."""
        operation_map = {
            'Grinding': 'Grinding',
            'MC1': 'Machining Stage 1',
            'MC2': 'Machining Stage 2',
            'MC3': 'Machining Stage 3',
            'SP1': 'Painting Stage 1',
            'SP2': 'Painting Stage 2',
            'SP3': 'Painting Stage 3'
        }
        return operation_map.get(stage_name, stage_name)

    def to_dataframe(self) -> pd.DataFrame:
        """Convert bottleneck report to DataFrame for Excel export."""
        report = self.analyze()

        if not report.bottlenecks:
            return pd.DataFrame({
                'Note': ['No bottlenecks detected (all resources below 85% utilization)']
            })

        records = []
        for b in report.bottlenecks:
            records.append({
                'Week': f'W{b.week}',
                'Resource': b.resource_name,
                'Operation': b.operation,
                'Utilization_%': b.utilization_pct,
                'Capacity': b.capacity_value,
                'Used': b.used_value,
                'Overflow': b.overflow_value,
                'Unit': b.capacity_unit,
                'Severity': b.severity,
                'Parts_Affected': ', '.join(b.parts_affected[:5]) if b.parts_affected else ''
            })

        return pd.DataFrame(records)

    def get_summary_dataframe(self) -> pd.DataFrame:
        """Get summary statistics as DataFrame."""
        report = self.analyze()

        # Resource summary
        resource_records = []
        for resource, overflow in sorted(
            report.summary_by_resource.items(),
            key=lambda x: x[1],
            reverse=True
        ):
            resource_records.append({
                'Resource': resource,
                'Total_Overflow': round(overflow, 1),
                'Is_Critical': resource in report.critical_path
            })

        return pd.DataFrame(resource_records)

    def get_recommendations(self) -> List[str]:
        """Generate recommendations based on bottleneck analysis."""
        report = self.analyze()
        recommendations = []

        if not report.bottlenecks:
            recommendations.append("✓ No significant bottlenecks detected")
            return recommendations

        # Analyze critical path
        for resource in report.critical_path[:3]:
            total_overflow = report.summary_by_resource[resource]

            if 'LINE' in resource:
                recommendations.append(
                    f"⚠ {resource}: Consider adding overtime or additional shifts "
                    f"(overflow: {total_overflow:.1f} hours)"
                )
            elif 'BOX' in resource:
                recommendations.append(
                    f"⚠ {resource}: Consider procuring additional mould boxes "
                    f"or outsourcing casting"
                )
            else:
                recommendations.append(
                    f"⚠ {resource}: Consider capacity expansion or load balancing "
                    f"(overflow: {total_overflow:.1f})"
                )

        # Week-specific recommendations
        critical_weeks = [
            w for w, count in report.summary_by_week.items()
            if count >= 3
        ]
        if critical_weeks:
            weeks_str = ', '.join([f'W{w}' for w in sorted(critical_weeks)])
            recommendations.append(
                f"⚠ Multiple bottlenecks in weeks {weeks_str}: "
                "Consider load leveling or order rescheduling"
            )

        return recommendations
