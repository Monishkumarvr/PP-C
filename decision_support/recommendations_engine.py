"""
Recommendations Engine
======================
Generates actionable recommendations based on analysis results.
"""

import pandas as pd
from dataclasses import dataclass
from typing import List, Dict
from .bottleneck_analyzer import BottleneckAnalyzer, BottleneckReport
from .order_risk_dashboard import OrderRiskAnalyzer, RiskSummary


@dataclass
class Recommendation:
    """A single actionable recommendation."""
    priority: str  # Critical, High, Medium, Low
    category: str  # Capacity, Scheduling, Procurement, Process
    title: str
    description: str
    impact: str
    action_items: List[str]
    affected_resources: List[str]
    estimated_benefit: str


class RecommendationsEngine:
    """
    Generates actionable recommendations based on bottleneck and risk analysis.

    Usage:
        engine = RecommendationsEngine(comprehensive_output_path)
        recommendations = engine.generate()
        df = engine.to_dataframe()
    """

    def __init__(self, comprehensive_output_path: str):
        """
        Args:
            comprehensive_output_path: Path to production_plan_COMPREHENSIVE_test.xlsx
        """
        self.output_path = comprehensive_output_path
        self.bottleneck_analyzer = None
        self.risk_analyzer = None
        self._initialize_analyzers()

    def _initialize_analyzers(self):
        """Initialize the analysis modules."""
        print("  🔧 Initializing recommendation analyzers...")

        self.bottleneck_analyzer = BottleneckAnalyzer(self.output_path)
        self.risk_analyzer = OrderRiskAnalyzer(self.output_path)

        print("    ✓ Analyzers initialized")

    def generate(self) -> List[Recommendation]:
        """Generate recommendations based on analysis."""
        print("  💡 Generating recommendations...")

        recommendations = []

        # Get analysis results
        bottleneck_report = self.bottleneck_analyzer.analyze()
        risk_summary = self.risk_analyzer.get_summary()

        # Generate recommendations based on bottlenecks
        recommendations.extend(
            self._generate_bottleneck_recommendations(bottleneck_report)
        )

        # Generate recommendations based on order risks
        recommendations.extend(
            self._generate_risk_recommendations(risk_summary)
        )

        # Generate general optimization recommendations
        recommendations.extend(
            self._generate_optimization_recommendations(bottleneck_report, risk_summary)
        )

        # Sort by priority
        priority_order = {'Critical': 0, 'High': 1, 'Medium': 2, 'Low': 3}
        recommendations.sort(key=lambda r: priority_order.get(r.priority, 4))

        print(f"    ✓ Generated {len(recommendations)} recommendations")

        return recommendations

    def _generate_bottleneck_recommendations(self, report: BottleneckReport) -> List[Recommendation]:
        """Generate recommendations for bottlenecks."""
        recommendations = []

        if not report.bottlenecks:
            return recommendations

        # Analyze critical path
        for resource in report.critical_path[:3]:
            total_overflow = report.summary_by_resource.get(resource, 0)

            if 'LINE' in resource:
                # Casting line recommendations
                recommendations.append(Recommendation(
                    priority='High',
                    category='Capacity',
                    title=f'Increase {resource.replace("_", " ").title()} Capacity',
                    description=f'{resource} is overloaded by {total_overflow:.1f} hours across the planning horizon.',
                    impact=f'Affects casting schedule and downstream operations',
                    action_items=[
                        'Add overtime shifts (weekend/evening)',
                        'Consider outsourcing overflow castings',
                        'Evaluate equipment utilization improvements',
                        'Review pattern changeover times for optimization'
                    ],
                    affected_resources=[resource],
                    estimated_benefit=f'Recover {total_overflow:.1f} hours capacity'
                ))

            elif 'BOX' in resource:
                # Mould box recommendations
                recommendations.append(Recommendation(
                    priority='High',
                    category='Procurement',
                    title=f'Procure Additional {resource.replace("_", " ")}',
                    description=f'Mould box capacity is constraining production.',
                    impact='Limits casting volume regardless of line capacity',
                    action_items=[
                        'Order additional mould boxes (lead time: 2-4 weeks)',
                        'Repair/maintain existing boxes to extend life',
                        'Optimize box allocation across parts',
                        'Consider box sharing with nearby foundries'
                    ],
                    affected_resources=[resource],
                    estimated_benefit='Increase casting flexibility'
                ))

            elif resource in ['MC1', 'MC2', 'MC3']:
                # Machining recommendations
                recommendations.append(Recommendation(
                    priority='High',
                    category='Capacity',
                    title=f'Address {resource} Machining Bottleneck',
                    description=f'Machining stage {resource} is over capacity.',
                    impact='Delays painting and final delivery',
                    action_items=[
                        'Add overtime for machining operators',
                        'Consider subcontracting machining operations',
                        'Review tool changeover times',
                        'Evaluate CNC program optimization'
                    ],
                    affected_resources=[resource],
                    estimated_benefit=f'Reduce machining backlog'
                ))

            elif resource in ['SP1', 'SP2', 'SP3']:
                # Painting recommendations
                recommendations.append(Recommendation(
                    priority='High',
                    category='Capacity',
                    title=f'Address {resource} Painting Bottleneck',
                    description=f'Painting stage {resource} is over capacity.',
                    impact='Final bottleneck before delivery',
                    action_items=[
                        'Add painting shifts',
                        'Consider outsourcing painting',
                        'Optimize drying rack utilization',
                        'Review paint booth scheduling'
                    ],
                    affected_resources=[resource],
                    estimated_benefit='Faster delivery completion'
                ))

            else:
                # Generic resource
                recommendations.append(Recommendation(
                    priority='Medium',
                    category='Capacity',
                    title=f'Review {resource} Capacity',
                    description=f'{resource} is showing high utilization.',
                    impact='May affect production schedule',
                    action_items=[
                        'Analyze resource utilization patterns',
                        'Consider capacity expansion',
                        'Review maintenance schedules'
                    ],
                    affected_resources=[resource],
                    estimated_benefit='Improved throughput'
                ))

        # Week-specific recommendations
        critical_weeks = [
            w for w, count in report.summary_by_week.items()
            if count >= 3
        ]

        if critical_weeks:
            weeks_str = ', '.join([f'W{w}' for w in sorted(critical_weeks)])
            recommendations.append(Recommendation(
                priority='High',
                category='Scheduling',
                title='Load Leveling Required',
                description=f'Multiple bottlenecks detected in weeks {weeks_str}.',
                impact='Production schedule is unbalanced',
                action_items=[
                    'Pull forward production from constrained weeks',
                    'Negotiate delivery date changes with customers',
                    'Prioritize orders by customer importance',
                    'Consider inventory build for smooth flow'
                ],
                affected_resources=report.critical_path,
                estimated_benefit='Smoother production flow'
            ))

        return recommendations

    def _generate_risk_recommendations(self, summary: RiskSummary) -> List[Recommendation]:
        """Generate recommendations based on order risks."""
        recommendations = []

        if summary.critical_orders > 0:
            recommendations.append(Recommendation(
                priority='Critical',
                category='Scheduling',
                title='Address Critical Orders Immediately',
                description=f'{summary.critical_orders} orders are at critical risk of missing delivery.',
                impact=f'{summary.total_at_risk_qty} units at risk',
                action_items=[
                    'Review critical order list with planning team',
                    'Expedite production for critical items',
                    'Communicate proactively with affected customers',
                    'Identify alternative supply options'
                ],
                affected_resources=[],
                estimated_benefit='Prevent customer escalations'
            ))

        if summary.high_risk_orders > 5:
            recommendations.append(Recommendation(
                priority='High',
                category='Process',
                title='Implement Order Prioritization System',
                description=f'{summary.high_risk_orders} orders classified as high risk.',
                impact='Multiple orders may miss delivery dates',
                action_items=[
                    'Establish daily order review meeting',
                    'Create priority scoring system',
                    'Assign dedicated expediter for high-risk orders',
                    'Set up early warning system for delays'
                ],
                affected_resources=[],
                estimated_benefit='Better delivery performance'
            ))

        if summary.at_risk_customers:
            recommendations.append(Recommendation(
                priority='High',
                category='Process',
                title='Customer Communication Plan',
                description=f'{len(summary.at_risk_customers)} customers have at-risk orders.',
                impact='Customer relationships at stake',
                action_items=[
                    'Prepare status updates for affected customers',
                    'Propose recovery plans with revised dates',
                    'Consider compensation/goodwill gestures',
                    'Schedule follow-up calls'
                ],
                affected_resources=[],
                estimated_benefit='Maintain customer trust'
            ))

        return recommendations

    def _generate_optimization_recommendations(self, bottleneck_report: BottleneckReport,
                                               risk_summary: RiskSummary) -> List[Recommendation]:
        """Generate general optimization recommendations."""
        recommendations = []

        # If everything is running smoothly
        if (bottleneck_report.total_bottlenecks == 0 and
            risk_summary.critical_orders == 0 and
            risk_summary.high_risk_orders == 0):

            recommendations.append(Recommendation(
                priority='Low',
                category='Process',
                title='Production Running Smoothly',
                description='No critical bottlenecks or high-risk orders detected.',
                impact='Positive',
                action_items=[
                    'Continue monitoring key metrics',
                    'Document current best practices',
                    'Consider taking on additional orders'
                ],
                affected_resources=[],
                estimated_benefit='Maintain performance'
            ))

        else:
            # General improvement recommendations
            if bottleneck_report.total_bottlenecks > 5:
                recommendations.append(Recommendation(
                    priority='Medium',
                    category='Process',
                    title='Implement Bottleneck Management System',
                    description='Multiple recurring bottlenecks indicate systemic issues.',
                    impact='Long-term capacity improvement',
                    action_items=[
                        'Create bottleneck tracking dashboard',
                        'Establish weekly capacity review',
                        'Develop capacity investment roadmap',
                        'Consider Theory of Constraints implementation'
                    ],
                    affected_resources=bottleneck_report.critical_path,
                    estimated_benefit='Sustained throughput improvement'
                ))

        return recommendations

    def to_dataframe(self) -> pd.DataFrame:
        """Convert recommendations to DataFrame for Excel export."""
        recommendations = self.generate()

        if not recommendations:
            return pd.DataFrame({
                'Note': ['No specific recommendations at this time']
            })

        records = []
        for i, r in enumerate(recommendations, 1):
            records.append({
                'Priority': r.priority,
                'Category': r.category,
                'Recommendation': r.title,
                'Description': r.description,
                'Impact': r.impact,
                'Action_Items': '\n'.join([f'• {a}' for a in r.action_items]),
                'Affected_Resources': ', '.join(r.affected_resources) if r.affected_resources else '-',
                'Expected_Benefit': r.estimated_benefit
            })

        return pd.DataFrame(records)

    def get_action_plan(self) -> pd.DataFrame:
        """Get a simplified action plan with immediate actions."""
        recommendations = self.generate()

        # Filter to Critical and High priority
        urgent = [r for r in recommendations if r.priority in ['Critical', 'High']]

        if not urgent:
            return pd.DataFrame({
                'Status': ['No urgent actions required']
            })

        records = []
        action_num = 1

        for r in urgent:
            for action in r.action_items[:2]:  # Top 2 actions per recommendation
                records.append({
                    'Action_#': action_num,
                    'Priority': r.priority,
                    'Action': action,
                    'Category': r.category,
                    'Recommendation': r.title
                })
                action_num += 1

        return pd.DataFrame(records)
