"""
Daily-Level Production Planning Optimization System
===================================================
Optimizes production schedules at DAILY granularity instead of weekly.

Key Differences from Weekly Version:
- Decision variables: x_casting[(variant, day)] instead of [(variant, week)]
- Constraints: Daily capacity limits instead of weekly
- Lead times: Measured in days (cooling_time, drying_time, etc.)
- Feasibility: Ensures no single day exceeds capacity

Usage:
    python production_plan_daily.py

Generates: production_plan_daily_comprehensive.xlsx
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import pulp
from collections import defaultdict
import holidays

# Reuse existing infrastructure
from production_plan_test import (
    ProductionCalendar,
    ComprehensiveDataLoader,
    WIPDemandCalculator,
    ComprehensiveParameterBuilder,
    MachineResourceManager,
    BoxCapacityManager,
    build_wip_init
)


class DailyProductionConfig:
    """Configuration for daily-level production planning."""
    
    def __init__(self):
        # Planning horizon
        self.CURRENT_DATE = datetime(2025, 10, 1)
        self.PLANNING_BUFFER_DAYS = 14  # Buffer beyond latest order
        self.MAX_PLANNING_DAYS = 210  # ~7 months
        
        # Working schedule  
        self.WORKING_DAYS_PER_WEEK = 6
        self.WEEKLY_OFF_DAY = 6  # Sunday (0=Monday, 6=Sunday)
        
        # Machine configuration
        self.OEE = 0.90
        self.HOURS_PER_SHIFT = 12
        self.SHIFTS_PER_DAY = 2
        
        # Daily capacities (derived from weekly / 6)
        self.BIG_LINE_HOURS_PER_DAY = self.HOURS_PER_SHIFT * self.SHIFTS_PER_DAY * self.OEE
        self.SMALL_LINE_HOURS_PER_DAY = self.HOURS_PER_SHIFT * self.SHIFTS_PER_DAY * self.OEE
        
        # Lead times (in days)
        self.DEFAULT_COOLING_DAYS = 1  # Casting -> Grinding
        self.DEFAULT_GRINDING_TO_MC1_DAYS = 1
        self.DEFAULT_MC_STAGE_DAYS = 1  # Between MC stages
        self.DEFAULT_PAINT_DRYING_DAYS = 1  # Between paint stages
        
        # Penalties
        self.UNMET_DEMAND_PENALTY = 200000
        self.LATENESS_PENALTY_PER_DAY = 5000  # Per day late (was 150k/week)
        self.INVENTORY_HOLDING_COST_PER_DAY = 0.15  # Per unit per day (was 1/week)
        self.SETUP_PENALTY = 5
        
        # Delivery flexibility
        self.DELIVERY_WINDOW_DAYS = 3  # Allow ±3 days from due date
        
        # Vacuum moulding
        self.VACUUM_CAPACITY_PENALTY = 0.70
        
        # Pattern changeover
        self.PATTERN_CHANGE_TIME_MIN = 18
        
        # Dynamic planning horizon (will be calculated)
        self.PLANNING_DAYS = None


print("="*80)
print("DAILY-LEVEL PRODUCTION PLANNING OPTIMIZATION")
print("="*80)
print("\nThis script implements DAILY granularity optimization.")
print("Each day is optimized independently with daily capacity constraints.")
print("\nKey improvements over weekly version:")
print("  ✓ Respects daily capacity limits")
print("  ✓ Accurate lead times (in days, not weeks)")
print("  ✓ Feasible daily schedules")
print("  ✓ Better short-term planning visibility")
print("\n" + "="*80 + "\n")

# Continue in next message due to length...
