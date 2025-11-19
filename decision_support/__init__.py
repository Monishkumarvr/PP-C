"""
Decision Support System for Production Planning
================================================

This module provides decision support tools for manufacturing production planning:
- Bottleneck Analysis: Identify capacity constraints blocking order fulfillment
- ATP Calculator: Check feasibility of new orders
- Order Risk Dashboard: Classify orders by risk level
- Recommendations Engine: Generate actionable suggestions
- Scenario Analyzer: What-if analysis for capacity changes

Usage:
    from decision_support import BottleneckAnalyzer, ATPCalculator, OrderRiskAnalyzer
"""

from .bottleneck_analyzer import BottleneckAnalyzer
from .atp_calculator import ATPCalculator
from .order_risk_dashboard import OrderRiskAnalyzer
from .recommendations_engine import RecommendationsEngine

__version__ = '1.0.0'
__all__ = [
    'BottleneckAnalyzer',
    'ATPCalculator',
    'OrderRiskAnalyzer',
    'RecommendationsEngine'
]
