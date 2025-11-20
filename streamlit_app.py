"""
Production Planning Optimization System - Streamlit Interface
=============================================================
Interactive web interface for the manufacturing production planning optimizer.

Usage:
    streamlit run streamlit_app.py

Features:
    - Upload Master Data Excel file
    - Configure optimization parameters
    - Run optimization with progress tracking
    - View interactive dashboards and charts
    - Download generated reports
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime, timedelta
import io
import tempfile
import os
import sys
import hashlib
import time
import traceback

# Import optimization modules
from production_plan_test import (
    ProductionConfig,
    ComprehensiveDataLoader,
    WIPDemandCalculator,
    ComprehensiveParameterBuilder,
    MachineResourceManager,
    BoxCapacityManager,
    ComprehensiveOptimizationModel,
    ComprehensiveResultsAnalyzer,
    ShipmentFulfillmentAnalyzer,
    DailyScheduleGenerator,
    build_wip_init
)
import pulp

# Page configuration
st.set_page_config(
    page_title="Production Planning Optimizer",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1F4788;
        text-align: center;
        padding: 1rem 0;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1F4788;
    }
    .success-box {
        padding: 1rem;
        background-color: #d4edda;
        border-radius: 0.5rem;
        border-left: 4px solid #28a745;
    }
    .warning-box {
        padding: 1rem;
        background-color: #fff3cd;
        border-radius: 0.5rem;
        border-left: 4px solid #ffc107;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 10px 20px;
    }
</style>
""", unsafe_allow_html=True)


def create_config_from_inputs(config_inputs):
    """Create ProductionConfig from user inputs."""
    config = ProductionConfig()

    # Update config with user inputs
    config.CURRENT_DATE = config_inputs['current_date']
    config.PLANNING_BUFFER_WEEKS = config_inputs['buffer_weeks']
    config.OEE = config_inputs['oee'] / 100.0
    config.UNMET_DEMAND_PENALTY = config_inputs['unmet_penalty']
    config.LATENESS_PENALTY = config_inputs['lateness_penalty']
    config.INVENTORY_HOLDING_COST = config_inputs['inventory_cost']
    config.PATTERN_CHANGE_TIME_MIN = config_inputs['pattern_change_time']

    return config


def validate_uploaded_file(uploaded_file):
    """Validate the uploaded Excel file has required sheets and columns."""
    validation_results = {
        'valid': True,
        'errors': [],
        'warnings': [],
        'summary': {}
    }

    try:
        # Read Excel file
        excel_file = pd.ExcelFile(uploaded_file)
        available_sheets = excel_file.sheet_names

        # Required sheets
        required_sheets = {
            'Part Master': ['FG Code', 'Standard unit wt.'],
            'Sales Order': ['Material Code', 'Balance Qty'],
            'Machine Constraints': ['Resource Code'],
            'Stage WIP': [],
            'Mould Box Capacity': ['Box_Size']
        }

        # Check for required sheets
        for sheet_name, required_cols in required_sheets.items():
            # Try different naming conventions
            found_sheet = None
            for available in available_sheets:
                if sheet_name.lower().replace(' ', '') in available.lower().replace(' ', ''):
                    found_sheet = available
                    break

            if found_sheet:
                df = pd.read_excel(uploaded_file, sheet_name=found_sheet)
                validation_results['summary'][sheet_name] = {
                    'rows': len(df),
                    'columns': len(df.columns)
                }

                # Check required columns
                for col in required_cols:
                    col_found = False
                    for df_col in df.columns:
                        if col.lower() in str(df_col).lower():
                            col_found = True
                            break
                    if not col_found:
                        validation_results['warnings'].append(
                            f"Column '{col}' not found in {sheet_name}"
                        )
            else:
                if sheet_name in ['Part Master', 'Sales Order']:
                    validation_results['errors'].append(f"Required sheet '{sheet_name}' not found")
                    validation_results['valid'] = False
                else:
                    validation_results['warnings'].append(f"Sheet '{sheet_name}' not found")

        # Load Sales Order for summary
        sales_sheet = None
        for sheet in available_sheets:
            if 'sales' in sheet.lower() or 'order' in sheet.lower():
                sales_sheet = sheet
                break

        if sales_sheet:
            sales_df = pd.read_excel(uploaded_file, sheet_name=sales_sheet)
            validation_results['summary']['total_orders'] = len(sales_df)

            # Find quantity column
            qty_col = None
            for col in sales_df.columns:
                if 'qty' in str(col).lower() or 'quantity' in str(col).lower():
                    qty_col = col
                    break
            if qty_col:
                validation_results['summary']['total_quantity'] = int(sales_df[qty_col].sum())

            # Find date column
            date_col = None
            for col in sales_df.columns:
                if 'date' in str(col).lower() or 'delivery' in str(col).lower():
                    date_col = col
                    break
            if date_col:
                dates = pd.to_datetime(sales_df[date_col], errors='coerce')
                validation_results['summary']['earliest_date'] = dates.min()
                validation_results['summary']['latest_date'] = dates.max()

        # Load Part Master for summary
        part_sheet = None
        for sheet in available_sheets:
            if 'part' in sheet.lower() and 'master' in sheet.lower():
                part_sheet = sheet
                break

        if part_sheet:
            part_df = pd.read_excel(uploaded_file, sheet_name=part_sheet)
            validation_results['summary']['total_parts'] = len(part_df)

    except Exception as e:
        validation_results['valid'] = False
        validation_results['errors'].append(f"Error reading file: {str(e)}")

    return validation_results


def get_file_hash(uploaded_file):
    """Generate a hash of the uploaded file for caching purposes."""
    file_content = uploaded_file.getvalue()
    return hashlib.md5(file_content).hexdigest()


def get_config_hash(config_inputs):
    """Generate a hash of the configuration for caching purposes."""
    config_str = str(sorted(config_inputs.items()))
    return hashlib.md5(config_str.encode()).hexdigest()


class OptimizationError(Exception):
    """Custom exception for optimization errors with detailed messages."""
    def __init__(self, message, stage, details=None):
        self.message = message
        self.stage = stage
        self.details = details or {}
        super().__init__(self.message)


def run_optimization_with_progress(uploaded_file, config, progress_callback=None):
    """Run the optimization with progress tracking and detailed error handling.

    Args:
        uploaded_file: Uploaded Excel file
        config: ProductionConfig object
        progress_callback: Function to call with (progress_percent, status_message)

    Returns:
        Dictionary with all optimization results

    Raises:
        OptimizationError: With detailed error information
    """
    def update_progress(percent, message):
        if progress_callback:
            progress_callback(percent, message)

    # Save uploaded file to temp location
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_path = tmp_file.name

    try:
        # Stage 1: Load data (0-15%)
        update_progress(5, "Loading master data from Excel...")
        try:
            loader = ComprehensiveDataLoader(tmp_path, config)
            data = loader.load_all_data()
        except Exception as e:
            raise OptimizationError(
                f"Failed to load data: {str(e)}",
                "Data Loading",
                {"hint": "Check that all required sheets exist and have correct column names"}
            )

        update_progress(15, f"Loaded {len(data.get('sales_order', []))} orders, {len(data.get('part_master', []))} parts")

        # Stage 2: Calculate demand (15-30%)
        update_progress(20, "Calculating net demand with WIP adjustments...")
        try:
            calculator = WIPDemandCalculator(data['sales_order'], data['stage_wip'], config)
            (net_demand, stage_start_qty, wip_coverage,
             gross_demand, wip_by_part) = calculator.calculate_net_demand_with_stages()
            split_demand, part_week_mapping, variant_windows = calculator.split_demand_by_week(net_demand)
        except KeyError as e:
            raise OptimizationError(
                f"Missing required column: {str(e)}",
                "Demand Calculation",
                {"hint": f"Check that Sales Order and Stage WIP sheets have the required columns"}
            )
        except Exception as e:
            raise OptimizationError(
                f"Failed to calculate demand: {str(e)}",
                "Demand Calculation",
                {"hint": "Verify order quantities and WIP values are numeric"}
            )

        update_progress(30, f"Calculated demand for {len(split_demand)} part-week variants")

        # Stage 3: Build parameters (30-40%)
        update_progress(35, "Building part parameters and routing data...")
        try:
            param_builder = ComprehensiveParameterBuilder(data['part_master'], config)
            params = param_builder.build_parameters()
        except Exception as e:
            raise OptimizationError(
                f"Failed to build parameters: {str(e)}",
                "Parameter Building",
                {"hint": "Check Part Master for missing cycle times or resource codes"}
            )

        update_progress(40, f"Built parameters for {len(params)} parts")

        # Stage 4: Setup resources (40-50%)
        update_progress(45, "Setting up machine and capacity resources...")
        try:
            machine_manager = MachineResourceManager(data['machine_constraints'], config)
            box_manager = BoxCapacityManager(data['box_capacity'], config, machine_manager)
            wip_init = build_wip_init(data['stage_wip'])
        except Exception as e:
            raise OptimizationError(
                f"Failed to setup resources: {str(e)}",
                "Resource Setup",
                {"hint": "Check Machine Constraints and Mould Box Capacity sheets"}
            )

        update_progress(50, "Resource constraints initialized")

        # Stage 5: Build and solve model (50-80%)
        update_progress(55, "Building optimization model...")
        try:
            optimizer = ComprehensiveOptimizationModel(
                split_demand,
                part_week_mapping,
                variant_windows,
                params,
                stage_start_qty,
                machine_manager,
                box_manager,
                config,
                wip_init=wip_init
            )
        except Exception as e:
            raise OptimizationError(
                f"Failed to build model: {str(e)}",
                "Model Building",
                {"hint": "There may be inconsistencies between parts and constraints"}
            )

        update_progress(65, "Solving optimization model (this may take a while)...")
        try:
            status = optimizer.build_and_solve()
        except Exception as e:
            raise OptimizationError(
                f"Solver failed: {str(e)}",
                "Optimization Solving",
                {"hint": "The problem may be infeasible. Try adjusting constraints or capacity."}
            )

        if status != pulp.LpStatusOptimal:
            status_name = pulp.LpStatus.get(status, "Unknown")
            raise OptimizationError(
                f"Optimization did not find optimal solution. Status: {status_name}",
                "Optimization Solving",
                {"status": status_name, "hint": "Try increasing capacity or adjusting penalties"}
            )

        update_progress(80, "Optimization complete, extracting results...")

        # Stage 6: Extract results (80-90%)
        try:
            analyzer = ComprehensiveResultsAnalyzer(
                optimizer,
                split_demand,
                part_week_mapping,
                params,
                machine_manager,
                box_manager,
                config
            )
            results = analyzer.extract_all_results()
        except Exception as e:
            raise OptimizationError(
                f"Failed to extract results: {str(e)}",
                "Results Extraction",
                {"hint": "The model solved but results extraction failed"}
            )

        update_progress(85, "Generating fulfillment reports...")

        # Stage 7: Generate reports (90-100%)
        try:
            fulfillment_analyzer = ShipmentFulfillmentAnalyzer(
                optimizer,
                data['sales_order'],
                split_demand,
                part_week_mapping,
                params,
                config,
                data,
                wip_by_part
            )
            fulfillment_reports = fulfillment_analyzer.generate_all_fulfillment_reports()
        except Exception as e:
            raise OptimizationError(
                f"Failed to generate fulfillment reports: {str(e)}",
                "Fulfillment Analysis",
                {"hint": "Results were extracted but fulfillment analysis failed"}
            )

        update_progress(92, "Generating daily schedules...")

        try:
            daily_generator = DailyScheduleGenerator(
                results['weekly_summary'],
                results,
                config
            )
            daily_schedule = daily_generator.generate_daily_schedule()
            part_daily_schedule = daily_generator.generate_part_level_daily_schedule(data['part_master'])
        except Exception as e:
            raise OptimizationError(
                f"Failed to generate daily schedule: {str(e)}",
                "Schedule Generation",
                {"hint": "Weekly results are available but daily breakdown failed"}
            )

        update_progress(100, "Optimization completed successfully!")

        # Compile all results
        all_results = {
            'status': status,
            'results': results,
            'fulfillment_reports': fulfillment_reports,
            'daily_schedule': daily_schedule,
            'part_daily_schedule': part_daily_schedule,
            'data': data,
            'config': config,
            'split_demand': split_demand,
            'part_week_mapping': part_week_mapping,
            'variant_windows': variant_windows,
            'stage_start_qty': stage_start_qty,
            'wip_init': wip_init,
            'timestamp': datetime.now().isoformat()
        }

        return all_results

    finally:
        # Cleanup temp file
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)


def run_optimization(uploaded_file, config):
    """Legacy wrapper for backward compatibility."""
    return run_optimization_with_progress(uploaded_file, config)


def create_kpi_dashboard(results, fulfillment_reports):
    """Create KPI dashboard with key metrics."""
    weekly_summary = results['weekly_summary']
    order_fulfillment = fulfillment_reports['order_fulfillment']

    # Calculate KPIs
    total_orders = len(order_fulfillment)
    # Fulfilled = On-Time, Late, or Fulfilled (not 'Not Fulfilled' or 'Partial')
    fulfilled_orders = len(order_fulfillment[order_fulfillment['Delivery_Status'].isin(['On-Time', 'Late', 'Fulfilled'])])
    fulfillment_rate = (fulfilled_orders / total_orders * 100) if total_orders > 0 else 0

    on_time_orders = len(order_fulfillment[order_fulfillment['Delivery_Status'] == 'On-Time'])
    on_time_rate = (on_time_orders / total_orders * 100) if total_orders > 0 else 0

    total_demand = order_fulfillment['Ordered_Qty'].sum()
    total_delivered = order_fulfillment['Delivered_Qty'].sum()

    # Average utilization
    if 'Big_Line_Util_%' in weekly_summary.columns:
        avg_casting_util = weekly_summary['Big_Line_Util_%'].mean()
    else:
        avg_casting_util = 0

    # Display KPIs in columns
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric(
            label="Fulfillment Rate",
            value=f"{fulfillment_rate:.1f}%",
            delta=f"{fulfilled_orders}/{total_orders} orders"
        )

    with col2:
        st.metric(
            label="On-Time Delivery",
            value=f"{on_time_rate:.1f}%",
            delta=f"{on_time_orders}/{total_orders} orders"
        )

    with col3:
        st.metric(
            label="Total Demand",
            value=f"{total_demand:,.0f}",
            delta=f"{total_delivered:,.0f} delivered"
        )

    with col4:
        st.metric(
            label="Avg Casting Utilization",
            value=f"{avg_casting_util:.1f}%"
        )


def create_capacity_chart(weekly_summary):
    """Create capacity utilization chart."""
    # Prepare data for plotting
    weeks = weekly_summary['Week'].tolist()

    # Create figure with secondary y-axis
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    # Big Line utilization
    if 'Big_Line_Util_%' in weekly_summary.columns:
        fig.add_trace(
            go.Bar(
                x=weeks,
                y=weekly_summary['Big_Line_Util_%'],
                name='Big Line Util %',
                marker_color='#1F4788'
            ),
            secondary_y=False
        )

    # Small Line utilization
    if 'Small_Line_Util_%' in weekly_summary.columns:
        fig.add_trace(
            go.Bar(
                x=weeks,
                y=weekly_summary['Small_Line_Util_%'],
                name='Small Line Util %',
                marker_color='#366092'
            ),
            secondary_y=False
        )

    # Add 100% capacity line
    fig.add_hline(y=100, line_dash="dash", line_color="red",
                  annotation_text="100% Capacity")

    fig.update_layout(
        title='Weekly Casting Line Utilization',
        xaxis_title='Week',
        yaxis_title='Utilization %',
        barmode='group',
        height=400
    )

    return fig


def create_production_flow_chart(weekly_summary):
    """Create production flow chart showing units through stages."""
    weeks = weekly_summary['Week'].tolist()

    fig = go.Figure()

    stages = [
        ('Casting_Tons', 'Casting (Tons)', '#8B4513'),
        ('Grinding_Units', 'Grinding', '#708090'),
        ('MC1_Units', 'MC1', '#4169E1'),
        ('MC2_Units', 'MC2', '#4682B4'),
        ('MC3_Units', 'MC3', '#5F9EA0'),
        ('SP1_Units', 'Primer', '#32CD32'),
        ('SP2_Units', 'Intermediate', '#228B22'),
        ('SP3_Units', 'Top Coat', '#006400'),
        ('Delivery_Units', 'Delivery', '#FF4500')
    ]

    for col, name, color in stages:
        if col in weekly_summary.columns:
            fig.add_trace(go.Scatter(
                x=weeks,
                y=weekly_summary[col],
                mode='lines+markers',
                name=name,
                line=dict(color=color, width=2)
            ))

    fig.update_layout(
        title='Weekly Production Flow by Stage',
        xaxis_title='Week',
        yaxis_title='Units / Tons',
        height=450,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )

    return fig


def create_fulfillment_chart(order_fulfillment):
    """Create order fulfillment status chart."""
    # Fulfillment status pie chart
    fulfillment_counts = order_fulfillment['Delivery_Status'].value_counts()

    fig = px.pie(
        values=fulfillment_counts.values,
        names=fulfillment_counts.index,
        title='Order Fulfillment Status',
        color_discrete_map={
            'On-Time': '#28a745',
            'Late': '#ffc107',
            'Partial': '#fd7e14',
            'Not Fulfilled': '#dc3545',
            'Fulfilled': '#17a2b8'
        }
    )

    fig.update_layout(height=350)

    return fig


def create_customer_analysis_chart(customer_fulfillment):
    """Create customer fulfillment analysis chart."""
    # Sort by total ordered quantity
    top_customers = customer_fulfillment.nlargest(10, 'Ordered_Qty')

    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=top_customers['Customer'],
        y=top_customers['Ordered_Qty'],
        name='Ordered',
        marker_color='#1F4788'
    ))

    fig.add_trace(go.Bar(
        x=top_customers['Customer'],
        y=top_customers['Delivered_Qty'],
        name='Delivered',
        marker_color='#28a745'
    ))

    fig.update_layout(
        title='Top 10 Customers - Order vs Delivery',
        xaxis_title='Customer',
        yaxis_title='Quantity',
        barmode='group',
        height=400
    )

    return fig


def create_bottleneck_analysis(weekly_summary, config):
    """Analyze bottlenecks and resource constraints."""
    bottlenecks = []

    # Check casting line utilization
    if 'Big_Line_Util_%' in weekly_summary.columns:
        high_util_weeks = weekly_summary[weekly_summary['Big_Line_Util_%'] > 90]
        if len(high_util_weeks) > 0:
            max_util = weekly_summary['Big_Line_Util_%'].max()
            bottlenecks.append({
                'Resource': 'Big Casting Line',
                'Max_Utilization': f"{max_util:.1f}%",
                'Weeks_Over_90%': len(high_util_weeks),
                'Peak_Week': int(weekly_summary.loc[weekly_summary['Big_Line_Util_%'].idxmax(), 'Week']),
                'Severity': 'Critical' if max_util > 100 else 'High' if max_util > 95 else 'Medium'
            })

    if 'Small_Line_Util_%' in weekly_summary.columns:
        high_util_weeks = weekly_summary[weekly_summary['Small_Line_Util_%'] > 90]
        if len(high_util_weeks) > 0:
            max_util = weekly_summary['Small_Line_Util_%'].max()
            bottlenecks.append({
                'Resource': 'Small Casting Line',
                'Max_Utilization': f"{max_util:.1f}%",
                'Weeks_Over_90%': len(high_util_weeks),
                'Peak_Week': int(weekly_summary.loc[weekly_summary['Small_Line_Util_%'].idxmax(), 'Week']),
                'Severity': 'Critical' if max_util > 100 else 'High' if max_util > 95 else 'Medium'
            })

    return pd.DataFrame(bottlenecks) if bottlenecks else pd.DataFrame()


def simulate_capacity_change(weekly_summary, oee_change, overtime_hours):
    """Simulate impact of capacity changes on utilization."""
    simulated = weekly_summary.copy()

    # Calculate new utilization with OEE change
    oee_factor = 1 + (oee_change / 100)

    util_columns = [col for col in simulated.columns if 'Util_%' in col]

    for col in util_columns:
        # Lower utilization with higher OEE (more capacity available)
        simulated[f'{col}_New'] = simulated[col] / oee_factor

        # Further reduce with overtime
        if overtime_hours > 0:
            # Assume base is 6 days * 24 hours = 144 hours/week
            base_hours = 144
            overtime_factor = base_hours / (base_hours + overtime_hours)
            simulated[f'{col}_New'] = simulated[f'{col}_New'] * overtime_factor

    return simulated


def simulate_demand_scaling(weekly_summary, fulfillment_reports, scale_factor):
    """Simulate impact of demand scaling on capacity."""
    # Scale the production quantities
    scaled_summary = weekly_summary.copy()

    quantity_columns = [col for col in scaled_summary.columns
                       if any(x in col for x in ['_Units', '_Tons'])]

    for col in quantity_columns:
        scaled_summary[col] = scaled_summary[col] * scale_factor

    # Recalculate utilization (simplified - assumes linear scaling)
    util_columns = [col for col in scaled_summary.columns if 'Util_%' in col]
    for col in util_columns:
        scaled_summary[col] = scaled_summary[col] * scale_factor

    # Calculate impact on fulfillment
    order_fulfillment = fulfillment_reports['order_fulfillment']
    current_fulfilled = len(order_fulfillment[
        order_fulfillment['Delivery_Status'].isin(['On-Time', 'Late', 'Fulfilled'])
    ])

    # Estimate new fulfillment (simplified model)
    max_util = scaled_summary[[col for col in util_columns if col in scaled_summary.columns]].max().max() if util_columns else 0

    if max_util > 100:
        # Rough estimate: reduce fulfillment proportionally to over-capacity
        estimated_fulfillment = current_fulfilled * (100 / max_util)
    else:
        estimated_fulfillment = current_fulfilled

    return scaled_summary, estimated_fulfillment, max_util


def generate_excel_report(all_results):
    """Generate comprehensive Excel report."""
    output = io.BytesIO()

    results = all_results['results']
    fulfillment_reports = all_results['fulfillment_reports']
    daily_schedule = all_results['daily_schedule']
    part_daily_schedule = all_results['part_daily_schedule']

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Stage plans
        results['casting_plan'].to_excel(writer, sheet_name='Casting', index=False)
        results['grinding_plan'].to_excel(writer, sheet_name='Grinding', index=False)
        results['mc1_plan'].to_excel(writer, sheet_name='Machining_Stage1', index=False)
        results['mc2_plan'].to_excel(writer, sheet_name='Machining_Stage2', index=False)
        results['mc3_plan'].to_excel(writer, sheet_name='Machining_Stage3', index=False)
        results['sp1_plan'].to_excel(writer, sheet_name='Painting_Stage1', index=False)
        results['sp2_plan'].to_excel(writer, sheet_name='Painting_Stage2', index=False)
        results['sp3_plan'].to_excel(writer, sheet_name='Painting_Stage3', index=False)
        results['delivery_plan'].to_excel(writer, sheet_name='Delivery', index=False)

        # Analysis
        results['flow_analysis'].to_excel(writer, sheet_name='Flow_Analysis', index=False)
        results['weekly_summary'].to_excel(writer, sheet_name='Weekly_Summary', index=False)
        results['changeover_analysis'].to_excel(writer, sheet_name='Pattern_Changeovers', index=False)
        results['vacuum_utilization'].to_excel(writer, sheet_name='Vacuum_Utilization', index=False)
        results['wip_consumption'].to_excel(writer, sheet_name='WIP_Consumption', index=False)

        # Daily schedules
        daily_schedule.to_excel(writer, sheet_name='Daily_Schedule', index=False)
        part_daily_schedule.to_excel(writer, sheet_name='Part_Daily_Schedule', index=False)

        # Fulfillment reports
        fulfillment_reports['order_fulfillment'].to_excel(writer, sheet_name='Order_Fulfillment', index=False)
        fulfillment_reports['customer_fulfillment'].to_excel(writer, sheet_name='Customer_Fulfillment', index=False)

    output.seek(0)
    return output


def main():
    """Main Streamlit application."""

    # Header
    st.markdown('<h1 class="main-header">🏭 Production Planning Optimizer</h1>',
                unsafe_allow_html=True)
    st.markdown("---")

    # Sidebar - Configuration
    with st.sidebar:
        st.header("⚙️ Configuration")

        # File upload
        st.subheader("📁 Data Upload")
        uploaded_file = st.file_uploader(
            "Upload Master Data Excel",
            type=['xlsx', 'xls'],
            help="Upload the Master Data file containing Part Master, Sales Orders, etc."
        )

        st.markdown("---")

        # Planning parameters
        st.subheader("📅 Planning Parameters")

        current_date = st.date_input(
            "Planning Start Date",
            value=datetime(2025, 10, 1),
            help="The start date for production planning"
        )

        buffer_weeks = st.slider(
            "Buffer Weeks",
            min_value=1,
            max_value=5,
            value=2,
            help="Buffer beyond latest order date"
        )

        oee = st.slider(
            "OEE %",
            min_value=70,
            max_value=100,
            value=90,
            help="Overall Equipment Effectiveness"
        )

        st.markdown("---")

        # Penalty parameters
        st.subheader("💰 Penalty Parameters")

        unmet_penalty = st.number_input(
            "Unmet Demand Penalty",
            min_value=1000,
            max_value=1000000,
            value=200000,
            step=10000,
            help="Cost of not fulfilling orders"
        )

        lateness_penalty = st.number_input(
            "Lateness Penalty",
            min_value=1000,
            max_value=500000,
            value=150000,
            step=10000,
            help="Cost per week late"
        )

        inventory_cost = st.number_input(
            "Inventory Holding Cost",
            min_value=0,
            max_value=100,
            value=1,
            help="Cost per unit per week"
        )

        pattern_change_time = st.number_input(
            "Pattern Change Time (min)",
            min_value=0,
            max_value=60,
            value=18,
            help="Mould changeover time"
        )

        st.markdown("---")

        # Run optimization button
        run_button = st.button(
            "🚀 Run Optimization",
            type="primary",
            use_container_width=True,
            disabled=uploaded_file is None
        )

    # Main content area
    if uploaded_file is None:
        # Show instructions when no file uploaded
        st.info("👈 Please upload a Master Data Excel file to begin.")

        st.markdown("""
        ### Required Excel Sheets:

        1. **Part Master** - Part specifications (FG Code, cycle times, resources)
        2. **Sales Order** - Customer orders with quantities and delivery dates
        3. **Machine Constraints** - Resource capacities and availability
        4. **Stage WIP** - Work in progress inventory
        5. **Mould Box Capacity** - Casting constraints

        ### How to Use:

        1. Upload your Master Data Excel file
        2. Adjust planning parameters in the sidebar
        3. Click "Run Optimization"
        4. View results in the dashboard tabs
        5. Download the generated report
        """)

        # Show sample data format
        with st.expander("📋 View Sample Data Format"):
            st.markdown("""
            #### Sales Order Sheet Columns:
            - `Material Code` - FG Code
            - `Balance Qty` - Quantity to produce
            - `Comitted Delivery Date` - Due date
            - `Sold_To_Party` - Customer ID

            #### Part Master Sheet Columns:
            - `FG Code`, `CS Code` - Part identifiers
            - `Standard unit wt.` - Unit weight
            - `Box Size`, `Box Quantity` - Packaging
            - `Moulding Line` - Casting line assignment
            - Cycle times for each operation
            """)

    else:
        # File uploaded but not run yet - show validation and preview
        st.subheader("📋 Data Validation & Preview")

        # Validate file
        validation = validate_uploaded_file(uploaded_file)

        # Show validation status
        if validation['valid']:
            st.success("✅ File validation passed!")
        else:
            st.error("❌ File validation failed!")
            for error in validation['errors']:
                st.error(f"• {error}")

        # Show warnings
        if validation['warnings']:
            with st.expander(f"⚠️ Warnings ({len(validation['warnings'])})"):
                for warning in validation['warnings']:
                    st.warning(f"• {warning}")

        # Show data summary
        st.subheader("📊 Data Summary")

        summary = validation['summary']

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric("Total Orders", summary.get('total_orders', 'N/A'))
        with col2:
            st.metric("Total Parts", summary.get('total_parts', 'N/A'))
        with col3:
            st.metric("Total Quantity", f"{summary.get('total_quantity', 0):,}")
        with col4:
            if 'earliest_date' in summary and 'latest_date' in summary:
                earliest = summary['earliest_date']
                latest = summary['latest_date']
                if pd.notna(earliest) and pd.notna(latest):
                    date_range = f"{earliest.strftime('%m/%d')} - {latest.strftime('%m/%d')}"
                else:
                    date_range = "N/A"
            else:
                date_range = "N/A"
            st.metric("Delivery Range", date_range)

        # Show sheet details
        with st.expander("📁 Sheet Details"):
            sheet_data = []
            for sheet_name, details in summary.items():
                if isinstance(details, dict) and 'rows' in details:
                    sheet_data.append({
                        'Sheet': sheet_name,
                        'Rows': details['rows'],
                        'Columns': details['columns']
                    })
            if sheet_data:
                st.dataframe(pd.DataFrame(sheet_data), use_container_width=True)

        st.info("👈 Click 'Run Optimization' in the sidebar to start planning.")

    if run_button and uploaded_file is not None:
        # Run optimization
        config_inputs = {
            'current_date': datetime.combine(current_date, datetime.min.time()),
            'buffer_weeks': buffer_weeks,
            'oee': oee,
            'unmet_penalty': unmet_penalty,
            'lateness_penalty': lateness_penalty,
            'inventory_cost': inventory_cost,
            'pattern_change_time': pattern_change_time
        }

        config = create_config_from_inputs(config_inputs)

        # Check for cached results
        file_hash = get_file_hash(uploaded_file)
        config_hash = get_config_hash(config_inputs)
        cache_key = f"{file_hash}_{config_hash}"

        # Check if we have cached results for this exact configuration
        if (st.session_state.get('cache_key') == cache_key and
            st.session_state.get('optimization_complete', False)):
            st.info("Using cached results. Click 'Run Optimization' again with different settings to re-run.")
        else:
            # Create progress tracking UI
            progress_container = st.container()
            with progress_container:
                progress_bar = st.progress(0)
                status_text = st.empty()
                stage_info = st.empty()

                start_time = time.time()

                def progress_callback(percent, message):
                    """Update progress UI with current status."""
                    progress_bar.progress(percent / 100)
                    elapsed = time.time() - start_time
                    status_text.markdown(f"**{message}**")
                    if percent < 100:
                        stage_info.text(f"Elapsed: {elapsed:.1f}s")
                    else:
                        stage_info.text(f"Total time: {elapsed:.1f}s")

                try:
                    # Run optimization with progress tracking
                    all_results = run_optimization_with_progress(
                        uploaded_file,
                        config,
                        progress_callback
                    )

                    # Store results in session state with cache key
                    st.session_state['optimization_results'] = all_results
                    st.session_state['optimization_complete'] = True
                    st.session_state['cache_key'] = cache_key
                    st.session_state['last_run_time'] = datetime.now().isoformat()

                    # Show success message
                    elapsed = time.time() - start_time
                    st.success(f"Optimization completed successfully in {elapsed:.1f} seconds!")

                    # Clear progress UI and rerun to show results
                    time.sleep(0.5)
                    st.rerun()

                except OptimizationError as e:
                    # Handle custom optimization errors with detailed feedback
                    st.error(f"Optimization failed at stage: **{e.stage}**")
                    st.error(f"{e.message}")

                    if e.details.get('hint'):
                        st.warning(f"**Suggestion:** {e.details['hint']}")

                    with st.expander("Technical Details"):
                        st.code(traceback.format_exc())

                    # Clear incomplete results
                    st.session_state['optimization_complete'] = False
                    st.session_state.pop('optimization_results', None)
                    st.session_state.pop('cache_key', None)

                except Exception as e:
                    # Handle unexpected errors
                    st.error(f"Unexpected error during optimization")
                    st.error(str(e))

                    with st.expander("Error Details"):
                        st.code(traceback.format_exc())

                    st.warning("""
                    **Common issues:**
                    - Missing required columns in Excel sheets
                    - Invalid data types (text where numbers expected)
                    - Missing Part Master entries for ordered parts
                    - Insufficient capacity for demand
                    """)

                    # Clear incomplete results
                    st.session_state['optimization_complete'] = False
                    st.session_state.pop('optimization_results', None)
                    st.session_state.pop('cache_key', None)

    # Display results if available
    if st.session_state.get('optimization_complete', False):
        all_results = st.session_state['optimization_results']
        results = all_results['results']
        fulfillment_reports = all_results['fulfillment_reports']

        # Show run information
        last_run = st.session_state.get('last_run_time', '')
        if last_run:
            try:
                run_dt = datetime.fromisoformat(last_run)
                run_str = run_dt.strftime("%Y-%m-%d %H:%M:%S")
            except:
                run_str = last_run

            col1, col2 = st.columns([3, 1])
            with col1:
                st.caption(f"Last optimization run: {run_str}")
            with col2:
                if st.button("Clear Results", type="secondary", key="clear_results"):
                    st.session_state['optimization_complete'] = False
                    st.session_state.pop('optimization_results', None)
                    st.session_state.pop('cache_key', None)
                    st.session_state.pop('last_run_time', None)
                    st.rerun()

        # Results tabs
        tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
            "📊 Dashboard",
            "📈 Capacity Analysis",
            "📦 Production Schedule",
            "🚚 Delivery Tracking",
            "🔮 What-If Analysis",
            "📥 Download",
            "✏️ Data Editor"
        ])

        with tab1:
            st.header("Executive Dashboard")

            # KPI cards
            create_kpi_dashboard(results, fulfillment_reports)

            st.markdown("---")

            # Charts
            col1, col2 = st.columns(2)

            with col1:
                fig = create_fulfillment_chart(fulfillment_reports['order_fulfillment'])
                st.plotly_chart(fig, use_container_width=True)

            with col2:
                fig = create_customer_analysis_chart(fulfillment_reports['customer_fulfillment'])
                st.plotly_chart(fig, use_container_width=True)

            # Production flow chart
            fig = create_production_flow_chart(results['weekly_summary'])
            st.plotly_chart(fig, use_container_width=True)

        with tab2:
            st.header("Capacity Analysis")

            # Capacity utilization chart
            fig = create_capacity_chart(results['weekly_summary'])
            st.plotly_chart(fig, use_container_width=True)

            # Weekly summary table
            st.subheader("Weekly Summary")
            st.dataframe(
                results['weekly_summary'],
                use_container_width=True,
                height=400
            )

            # Vacuum utilization
            st.subheader("Vacuum Line Utilization")
            st.dataframe(
                results['vacuum_utilization'],
                use_container_width=True
            )

        with tab3:
            st.header("Production Schedule")

            # Stage selection
            stage_options = {
                'Casting': results['casting_plan'],
                'Grinding': results['grinding_plan'],
                'MC1': results['mc1_plan'],
                'MC2': results['mc2_plan'],
                'MC3': results['mc3_plan'],
                'SP1 (Primer)': results['sp1_plan'],
                'SP2 (Intermediate)': results['sp2_plan'],
                'SP3 (Top Coat)': results['sp3_plan'],
                'Delivery': results['delivery_plan']
            }

            selected_stage = st.selectbox(
                "Select Stage",
                options=list(stage_options.keys())
            )

            stage_data = stage_options[selected_stage]

            # Filter options
            col1, col2 = st.columns(2)
            with col1:
                if 'Part' in stage_data.columns:
                    parts = ['All'] + sorted(stage_data['Part'].unique().tolist())
                    selected_part = st.selectbox("Filter by Part", parts)
            with col2:
                if 'Week' in stage_data.columns:
                    weeks = ['All'] + sorted(stage_data['Week'].unique().tolist())
                    selected_week = st.selectbox("Filter by Week", weeks)

            # Apply filters
            filtered_data = stage_data.copy()
            if selected_part != 'All' and 'Part' in filtered_data.columns:
                filtered_data = filtered_data[filtered_data['Part'] == selected_part]
            if selected_week != 'All' and 'Week' in filtered_data.columns:
                filtered_data = filtered_data[filtered_data['Week'] == selected_week]

            st.dataframe(filtered_data, use_container_width=True, height=500)

            # Daily schedule
            st.subheader("Daily Schedule")
            st.dataframe(
                all_results['daily_schedule'],
                use_container_width=True,
                height=300
            )

        with tab4:
            st.header("Delivery Tracking")

            # Order fulfillment
            st.subheader("Order Fulfillment Status")

            order_fulfillment = fulfillment_reports['order_fulfillment']

            # Summary metrics
            col1, col2, col3 = st.columns(3)
            with col1:
                total = len(order_fulfillment)
                fulfilled = len(order_fulfillment[order_fulfillment['Delivery_Status'].isin(['On-Time', 'Late', 'Fulfilled'])])
                st.metric("Fulfilled Orders", f"{fulfilled}/{total}")
            with col2:
                on_time = len(order_fulfillment[order_fulfillment['Delivery_Status'] == 'On-Time'])
                st.metric("On-Time Orders", f"{on_time}/{total}")
            with col3:
                late = len(order_fulfillment[order_fulfillment['Delivery_Status'] == 'Late'])
                st.metric("Late Orders", f"{late}")

            # Filter by status
            available_statuses = order_fulfillment['Delivery_Status'].unique().tolist()
            status_filter = st.multiselect(
                "Filter by Delivery Status",
                options=available_statuses,
                default=available_statuses
            )

            filtered_orders = order_fulfillment[
                order_fulfillment['Delivery_Status'].isin(status_filter)
            ]

            st.dataframe(filtered_orders, use_container_width=True, height=400)

            # Customer fulfillment
            st.subheader("Customer Fulfillment Summary")
            st.dataframe(
                fulfillment_reports['customer_fulfillment'],
                use_container_width=True
            )

        with tab5:
            st.header("What-If Analysis")
            st.markdown("Simulate different scenarios to understand capacity constraints and plan improvements.")

            # Create sub-tabs for different analyses
            whatif_tab1, whatif_tab2, whatif_tab3, whatif_tab4, whatif_tab5 = st.tabs([
                "🔧 Capacity Scenarios",
                "📊 Demand Scaling",
                "⚠️ Bottleneck Analysis",
                "📦 New Order (ATP)",
                "🚨 Rush Orders"
            ])

            with whatif_tab1:
                st.subheader("Capacity Improvement Scenarios")
                st.markdown("Simulate the impact of improving OEE or adding overtime hours.")

                col1, col2 = st.columns(2)

                with col1:
                    oee_change = st.slider(
                        "OEE Improvement (%)",
                        min_value=-10,
                        max_value=20,
                        value=0,
                        step=1,
                        help="Increase OEE to see reduced utilization"
                    )

                with col2:
                    overtime_hours = st.slider(
                        "Weekly Overtime Hours",
                        min_value=0,
                        max_value=48,
                        value=0,
                        step=4,
                        help="Add overtime hours per week"
                    )

                if oee_change != 0 or overtime_hours > 0:
                    simulated = simulate_capacity_change(
                        results['weekly_summary'],
                        oee_change,
                        overtime_hours
                    )

                    # Show comparison
                    st.subheader("Utilization Comparison")

                    # Create comparison chart
                    weeks = simulated['Week'].tolist()

                    fig = go.Figure()

                    if 'Big_Line_Util_%' in simulated.columns:
                        fig.add_trace(go.Bar(
                            x=weeks,
                            y=simulated['Big_Line_Util_%'],
                            name='Current Big Line',
                            marker_color='#1F4788',
                            opacity=0.6
                        ))

                        if 'Big_Line_Util_%_New' in simulated.columns:
                            fig.add_trace(go.Bar(
                                x=weeks,
                                y=simulated['Big_Line_Util_%_New'],
                                name='Improved Big Line',
                                marker_color='#28a745'
                            ))

                    fig.add_hline(y=100, line_dash="dash", line_color="red",
                                  annotation_text="100% Capacity")

                    fig.update_layout(
                        title='Current vs Improved Utilization',
                        xaxis_title='Week',
                        yaxis_title='Utilization %',
                        barmode='group',
                        height=400
                    )

                    st.plotly_chart(fig, use_container_width=True)

                    # Summary metrics
                    col1, col2, col3 = st.columns(3)

                    with col1:
                        if 'Big_Line_Util_%' in simulated.columns:
                            current_max = simulated['Big_Line_Util_%'].max()
                            st.metric("Current Max Util", f"{current_max:.1f}%")

                    with col2:
                        if 'Big_Line_Util_%_New' in simulated.columns:
                            new_max = simulated['Big_Line_Util_%_New'].max()
                            reduction = current_max - new_max
                            st.metric("New Max Util", f"{new_max:.1f}%",
                                     delta=f"-{reduction:.1f}%")

                    with col3:
                        if 'Big_Line_Util_%_New' in simulated.columns:
                            weeks_over = len(simulated[simulated['Big_Line_Util_%_New'] > 100])
                            st.metric("Weeks Over Capacity", weeks_over)

                else:
                    st.info("Adjust the sliders above to simulate capacity improvements.")

            with whatif_tab2:
                st.subheader("Demand Scaling Analysis")
                st.markdown("See how changes in demand affect capacity utilization.")

                scale_percent = st.slider(
                    "Demand Scale (%)",
                    min_value=50,
                    max_value=200,
                    value=100,
                    step=10,
                    help="100% = current demand, 150% = 50% increase"
                )

                scale_factor = scale_percent / 100

                if scale_factor != 1.0:
                    scaled_summary, est_fulfillment, max_util = simulate_demand_scaling(
                        results['weekly_summary'],
                        fulfillment_reports,
                        scale_factor
                    )

                    # Summary metrics
                    col1, col2, col3 = st.columns(3)

                    order_fulfillment = fulfillment_reports['order_fulfillment']
                    current_fulfilled = len(order_fulfillment[
                        order_fulfillment['Delivery_Status'].isin(['On-Time', 'Late', 'Fulfilled'])
                    ])
                    total_orders = len(order_fulfillment)

                    with col1:
                        scaled_orders = int(total_orders * scale_factor)
                        st.metric("Scaled Orders", scaled_orders,
                                 delta=f"{scaled_orders - total_orders:+d}")

                    with col2:
                        st.metric("Peak Utilization", f"{max_util:.1f}%",
                                 delta="Over capacity!" if max_util > 100 else "OK")

                    with col3:
                        est_rate = (est_fulfillment / (total_orders * scale_factor) * 100) if scale_factor > 0 else 0
                        st.metric("Est. Fulfillment Rate", f"{est_rate:.1f}%")

                    # Show scaled utilization chart
                    weeks = scaled_summary['Week'].tolist()

                    fig = go.Figure()

                    util_cols = [col for col in scaled_summary.columns if 'Util_%' in col and '_New' not in col]

                    for col in util_cols[:2]:  # Show first 2 utilization columns
                        fig.add_trace(go.Bar(
                            x=weeks,
                            y=scaled_summary[col],
                            name=col.replace('_', ' ').replace('Util %', 'Utilization')
                        ))

                    fig.add_hline(y=100, line_dash="dash", line_color="red",
                                  annotation_text="100% Capacity")

                    fig.update_layout(
                        title=f'Utilization at {scale_percent}% Demand',
                        xaxis_title='Week',
                        yaxis_title='Utilization %',
                        barmode='group',
                        height=400
                    )

                    st.plotly_chart(fig, use_container_width=True)

                    # Warning if over capacity
                    if max_util > 100:
                        st.warning(f"""
                        **Capacity Exceeded!** At {scale_percent}% demand, peak utilization reaches {max_util:.1f}%.

                        **Recommendations:**
                        - Add overtime hours to increase capacity
                        - Improve OEE through maintenance
                        - Consider outsourcing peak demand
                        - Negotiate delivery dates with customers
                        """)

                else:
                    st.info("Adjust the slider to simulate demand changes.")

            with whatif_tab3:
                st.subheader("Bottleneck Analysis")
                st.markdown("Identify resource constraints limiting production.")

                # Get bottleneck analysis
                bottleneck_df = create_bottleneck_analysis(
                    results['weekly_summary'],
                    all_results['config']
                )

                if not bottleneck_df.empty:
                    # Show severity colors
                    st.dataframe(
                        bottleneck_df,
                        use_container_width=True,
                        column_config={
                            "Severity": st.column_config.TextColumn(
                                "Severity",
                                help="Critical: >100%, High: >95%, Medium: >90%"
                            )
                        }
                    )

                    # Recommendations based on bottlenecks
                    st.subheader("Recommendations")

                    for _, row in bottleneck_df.iterrows():
                        severity = row['Severity']
                        resource = row['Resource']

                        if severity == 'Critical':
                            st.error(f"""
                            **{resource}** - CRITICAL

                            Immediate actions required:
                            - Add overtime shifts in week {row['Peak_Week']}
                            - Consider outsourcing some orders
                            - Review order priorities and reschedule non-critical orders
                            """)
                        elif severity == 'High':
                            st.warning(f"""
                            **{resource}** - HIGH RISK

                            Preventive actions recommended:
                            - Plan overtime for weeks with >95% utilization
                            - Improve OEE through preventive maintenance
                            - Build safety stock in earlier weeks
                            """)
                        else:
                            st.info(f"""
                            **{resource}** - MEDIUM RISK

                            Monitor closely:
                            - Track daily utilization
                            - Prepare contingency plans
                            - Consider minor schedule adjustments
                            """)

                else:
                    st.success("No significant bottlenecks detected! All resources are operating within acceptable limits.")

                # Utilization heatmap
                st.subheader("Weekly Utilization Heatmap")

                weekly_summary = results['weekly_summary']
                util_data = []

                for _, row in weekly_summary.iterrows():
                    week = int(row['Week'])
                    if 'Big_Line_Util_%' in row:
                        util_data.append({
                            'Week': week,
                            'Resource': 'Big Line',
                            'Utilization': row['Big_Line_Util_%']
                        })
                    if 'Small_Line_Util_%' in row:
                        util_data.append({
                            'Week': week,
                            'Resource': 'Small Line',
                            'Utilization': row['Small_Line_Util_%']
                        })

                if util_data:
                    heatmap_df = pd.DataFrame(util_data)
                    pivot_df = heatmap_df.pivot(index='Resource', columns='Week', values='Utilization')

                    fig = px.imshow(
                        pivot_df,
                        labels=dict(x="Week", y="Resource", color="Utilization %"),
                        color_continuous_scale=['green', 'yellow', 'red'],
                        aspect='auto'
                    )

                    fig.update_layout(
                        title='Resource Utilization Heatmap',
                        height=300
                    )

                    st.plotly_chart(fig, use_container_width=True)

            with whatif_tab4:
                st.subheader("New Order Feasibility (ATP)")
                st.markdown("Check if a new order can be fulfilled given current capacity.")

                # Get available parts from the data
                part_master = all_results['data']['part_master']
                available_parts = sorted(part_master['FG Code'].dropna().unique().tolist())

                col1, col2, col3 = st.columns(3)

                with col1:
                    atp_part = st.selectbox(
                        "Part Code",
                        options=available_parts,
                        help="Select the part for the new order"
                    )

                with col2:
                    atp_qty = st.number_input(
                        "Quantity",
                        min_value=1,
                        max_value=10000,
                        value=100,
                        step=10,
                        help="Order quantity"
                    )

                with col3:
                    atp_week = st.number_input(
                        "Delivery Week",
                        min_value=1,
                        max_value=20,
                        value=8,
                        help="Requested delivery week"
                    )

                if st.button("Check Feasibility", type="primary"):
                    # Get current utilization for the requested week
                    weekly_summary = results['weekly_summary']
                    week_data = weekly_summary[weekly_summary['Week'] == atp_week]

                    if len(week_data) > 0:
                        week_row = week_data.iloc[0]

                        # Get part parameters
                        part_info = part_master[part_master['FG Code'] == atp_part]

                        if len(part_info) > 0:
                            part_row = part_info.iloc[0]
                            unit_weight = float(part_row.get('Standard unit wt.', 1) or 1)
                            additional_tons = (atp_qty * unit_weight) / 1000

                            # Check casting capacity
                            current_util = week_row.get('Big_Line_Util_%', 0)

                            # Estimate additional utilization (simplified)
                            # Assume roughly linear relationship
                            current_tons = week_row.get('Casting_Tons', 0)
                            if current_tons > 0:
                                util_per_ton = current_util / current_tons
                                additional_util = additional_tons * util_per_ton
                                new_util = current_util + additional_util
                            else:
                                new_util = current_util + 5  # Default estimate

                            # Determine feasibility
                            if new_util <= 90:
                                feasibility = "Feasible"
                                color = "success"
                                message = f"Order can be fulfilled in Week {atp_week}."
                            elif new_util <= 100:
                                feasibility = "Tight"
                                color = "warning"
                                message = f"Order can be fulfilled but capacity will be tight ({new_util:.1f}%)."
                            else:
                                feasibility = "Not Feasible"
                                color = "error"
                                message = f"Week {atp_week} exceeds capacity ({new_util:.1f}%). Consider alternative weeks."

                            # Display result
                            st.markdown("---")
                            st.subheader("ATP Result")

                            col1, col2, col3, col4 = st.columns(4)
                            with col1:
                                st.metric("Feasibility", feasibility)
                            with col2:
                                st.metric("Additional Load", f"{additional_tons:.2f} tons")
                            with col3:
                                st.metric("Current Util", f"{current_util:.1f}%")
                            with col4:
                                st.metric("Projected Util", f"{new_util:.1f}%")

                            if color == "success":
                                st.success(message)
                            elif color == "warning":
                                st.warning(message)
                            else:
                                st.error(message)

                                # Suggest alternative weeks
                                st.markdown("**Alternative Weeks with Capacity:**")
                                alternatives = weekly_summary[weekly_summary['Big_Line_Util_%'] < 80]
                                if len(alternatives) > 0:
                                    alt_weeks = alternatives['Week'].tolist()[:5]
                                    st.write(f"Weeks with <80% utilization: {', '.join(map(str, alt_weeks))}")
                                else:
                                    st.write("No weeks with significant spare capacity found.")
                        else:
                            st.error(f"Part {atp_part} not found in Part Master.")
                    else:
                        st.warning(f"Week {atp_week} is outside the current planning horizon.")

            with whatif_tab5:
                st.subheader("Rush Order Analysis")
                st.markdown("Analyze late or unfulfilled orders and see what's needed to expedite them.")

                # Get orders that are late or unfulfilled
                order_fulfillment = fulfillment_reports['order_fulfillment']
                problem_orders = order_fulfillment[
                    order_fulfillment['Delivery_Status'].isin(['Late', 'Not Fulfilled', 'Partial'])
                ]

                if len(problem_orders) > 0:
                    # Create selection options
                    order_options = []
                    for _, row in problem_orders.iterrows():
                        order_id = row.get('Sales_Order_No', 'Unknown')
                        part = row.get('Part_Code', row.get('Material_Code', 'Unknown'))
                        qty = row.get('Ordered_Qty', 0)
                        status = row.get('Delivery_Status', 'Unknown')
                        order_options.append(f"{order_id} - {part} ({qty} units) - {status}")

                    selected_order = st.selectbox(
                        "Select Order to Analyze",
                        options=order_options
                    )

                    if selected_order:
                        # Parse selected order
                        order_id = selected_order.split(' - ')[0]
                        order_row = problem_orders[
                            problem_orders['Sales_Order_No'].astype(str) == order_id
                        ]

                        if len(order_row) > 0:
                            order_data = order_row.iloc[0]

                            st.markdown("---")
                            st.subheader("Order Details")

                            col1, col2, col3, col4 = st.columns(4)
                            with col1:
                                st.metric("Order", order_id)
                            with col2:
                                qty = order_data.get('Ordered_Qty', 0)
                                delivered = order_data.get('Delivered_Qty', 0)
                                st.metric("Quantity", f"{delivered}/{qty}")
                            with col3:
                                st.metric("Status", order_data.get('Delivery_Status', 'Unknown'))
                            with col4:
                                days_late = order_data.get('Days_Late', 0)
                                st.metric("Days Late", int(days_late))

                            # Analysis and recommendations
                            st.subheader("Expedite Options")

                            unmet = order_data.get('Unmet_Qty', qty - delivered)
                            part_code = order_data.get('Part_Code', order_data.get('Material_Code', 'Unknown'))

                            # Get part weight
                            part_master = all_results['data']['part_master']
                            part_info = part_master[part_master['FG Code'] == part_code]
                            if len(part_info) > 0:
                                unit_weight = float(part_info.iloc[0].get('Standard unit wt.', 1) or 1)
                            else:
                                unit_weight = 1

                            additional_tons = (unmet * unit_weight) / 1000

                            st.markdown(f"""
                            **To fulfill remaining {int(unmet)} units ({additional_tons:.2f} tons):**

                            1. **Overtime Option**
                               - Add ~{max(4, int(additional_tons * 2))} overtime hours in the next week
                               - Estimated cost: Higher labor costs

                            2. **Outsourcing Option**
                               - Outsource casting of {additional_tons:.2f} tons
                               - Lead time: 1-2 weeks
                               - Estimated cost: Premium pricing

                            3. **Priority Rescheduling**
                               - Reschedule lower-priority orders
                               - Free up {additional_tons:.2f} tons capacity
                               - Impact: Other orders delayed

                            4. **Partial Shipment**
                               - Ship available {int(delivered)} units now
                               - Deliver remaining {int(unmet)} units in Week +1-2
                               - Negotiate with customer
                            """)

                            # Show capacity in upcoming weeks
                            st.subheader("Upcoming Capacity Availability")
                            weekly_summary = results['weekly_summary']

                            # Show next 4 weeks
                            upcoming = weekly_summary.head(4)[['Week', 'Big_Line_Util_%', 'Casting_Tons']]
                            upcoming['Available_Capacity_%'] = 100 - upcoming['Big_Line_Util_%']

                            st.dataframe(upcoming, use_container_width=True)

                else:
                    st.success("No late or unfulfilled orders found! All orders are on track.")

                    # Show orders at risk (near deadline)
                    st.subheader("Orders to Monitor")
                    on_time_orders = order_fulfillment[
                        order_fulfillment['Delivery_Status'] == 'On-Time'
                    ]
                    if len(on_time_orders) > 0:
                        st.info(f"All {len(on_time_orders)} orders are currently on-time. Continue monitoring daily production.")

        with tab6:
            st.header("Download Reports")

            # Generate and download Excel report
            st.subheader("📊 Comprehensive Report")

            excel_data = generate_excel_report(all_results)

            st.download_button(
                label="📥 Download Excel Report",
                data=excel_data,
                file_name=f"production_plan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            st.markdown("""
            The Excel report includes:
            - Stage-wise production plans (Casting, Grinding, Machining, Painting)
            - Weekly summary with utilization metrics
            - Daily schedules
            - Order and customer fulfillment tracking
            - Flow analysis and WIP consumption
            """)

            # Individual sheet downloads
            st.subheader("📁 Individual Sheets")

            col1, col2 = st.columns(2)

            with col1:
                # Weekly summary CSV
                csv_weekly = results['weekly_summary'].to_csv(index=False)
                st.download_button(
                    label="Weekly Summary (CSV)",
                    data=csv_weekly,
                    file_name="weekly_summary.csv",
                    mime="text/csv"
                )

                # Order fulfillment CSV
                csv_orders = fulfillment_reports['order_fulfillment'].to_csv(index=False)
                st.download_button(
                    label="Order Fulfillment (CSV)",
                    data=csv_orders,
                    file_name="order_fulfillment.csv",
                    mime="text/csv"
                )

            with col2:
                # Daily schedule CSV
                csv_daily = all_results['daily_schedule'].to_csv(index=False)
                st.download_button(
                    label="Daily Schedule (CSV)",
                    data=csv_daily,
                    file_name="daily_schedule.csv",
                    mime="text/csv"
                )

                # Customer fulfillment CSV
                csv_customers = fulfillment_reports['customer_fulfillment'].to_csv(index=False)
                st.download_button(
                    label="Customer Fulfillment (CSV)",
                    data=csv_customers,
                    file_name="customer_fulfillment.csv",
                    mime="text/csv"
                )

        with tab7:
            st.header("Data Editor")
            st.markdown("Modify data and re-run optimization to see impact of changes.")

            # Initialize edited data in session state if not present
            if 'edited_orders' not in st.session_state:
                st.session_state['edited_orders'] = None
            if 'edited_wip' not in st.session_state:
                st.session_state['edited_wip'] = None

            # Create sub-tabs for different editors
            editor_tab1, editor_tab2, editor_tab3 = st.tabs([
                "📋 Sales Orders",
                "📦 WIP Inventory",
                "➕ Add New Order"
            ])

            with editor_tab1:
                st.subheader("Edit Sales Orders")
                st.markdown("Modify order quantities or delivery dates.")

                # Get current orders from results
                order_fulfillment = fulfillment_reports['order_fulfillment']

                # Create editable dataframe
                edit_cols = ['Sales_Order_No', 'Part_Code', 'Customer', 'Ordered_Qty',
                            'Delivery_Status', 'Committed_Week']

                # Filter to available columns
                available_cols = [col for col in edit_cols if col in order_fulfillment.columns]

                if available_cols:
                    orders_to_edit = order_fulfillment[available_cols].copy()

                    # Use data editor
                    edited_orders = st.data_editor(
                        orders_to_edit,
                        num_rows="dynamic",
                        use_container_width=True,
                        height=400,
                        column_config={
                            "Ordered_Qty": st.column_config.NumberColumn(
                                "Quantity",
                                min_value=0,
                                step=1
                            ),
                            "Committed_Week": st.column_config.NumberColumn(
                                "Delivery Week",
                                min_value=1,
                                max_value=52
                            )
                        }
                    )

                    # Check for changes
                    if not orders_to_edit.equals(edited_orders):
                        st.session_state['edited_orders'] = edited_orders
                        st.warning("⚠️ You have unsaved changes. Click 'Apply Changes' to update.")

                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button("Apply Changes", type="primary", key="apply_orders"):
                            st.session_state['edited_orders'] = edited_orders
                            st.success("✅ Order changes saved! Re-run optimization to see impact.")

                    with col2:
                        if st.button("Reset to Original", key="reset_orders"):
                            st.session_state['edited_orders'] = None
                            st.rerun()

                else:
                    st.warning("Order data not available for editing.")

            with editor_tab2:
                st.subheader("Edit WIP Inventory")
                st.markdown("Update current work-in-progress quantities at each stage.")

                # Get WIP data
                if 'data' in all_results and 'stage_wip' in all_results['data']:
                    wip_data = all_results['data']['stage_wip'].copy()

                    # Create editable version
                    wip_cols = ['CastingItem', 'FG', 'SP', 'MC', 'GR', 'CS']
                    available_wip_cols = [col for col in wip_cols if col in wip_data.columns]

                    if available_wip_cols:
                        wip_to_edit = wip_data[available_wip_cols].head(50)  # Limit rows for performance

                        edited_wip = st.data_editor(
                            wip_to_edit,
                            use_container_width=True,
                            height=400,
                            column_config={
                                "FG": st.column_config.NumberColumn("Finished Goods", min_value=0),
                                "SP": st.column_config.NumberColumn("Painting WIP", min_value=0),
                                "MC": st.column_config.NumberColumn("Machining WIP", min_value=0),
                                "GR": st.column_config.NumberColumn("Grinding WIP", min_value=0),
                                "CS": st.column_config.NumberColumn("Casting WIP", min_value=0)
                            }
                        )

                        col1, col2 = st.columns(2)
                        with col1:
                            if st.button("Apply WIP Changes", type="primary", key="apply_wip"):
                                st.session_state['edited_wip'] = edited_wip
                                st.success("✅ WIP changes saved! Re-run optimization to see impact.")

                        with col2:
                            if st.button("Reset WIP to Original", key="reset_wip"):
                                st.session_state['edited_wip'] = None
                                st.rerun()

                        # Show WIP summary
                        st.subheader("WIP Summary")
                        col1, col2, col3, col4, col5 = st.columns(5)
                        with col1:
                            st.metric("FG", int(wip_to_edit['FG'].sum()) if 'FG' in wip_to_edit.columns else 0)
                        with col2:
                            st.metric("SP", int(wip_to_edit['SP'].sum()) if 'SP' in wip_to_edit.columns else 0)
                        with col3:
                            st.metric("MC", int(wip_to_edit['MC'].sum()) if 'MC' in wip_to_edit.columns else 0)
                        with col4:
                            st.metric("GR", int(wip_to_edit['GR'].sum()) if 'GR' in wip_to_edit.columns else 0)
                        with col5:
                            st.metric("CS", int(wip_to_edit['CS'].sum()) if 'CS' in wip_to_edit.columns else 0)
                    else:
                        st.warning("WIP columns not found in data.")
                else:
                    st.warning("WIP data not available.")

            with editor_tab3:
                st.subheader("Add New Order")
                st.markdown("Create a new sales order to include in optimization.")

                # Get available parts
                part_master = all_results['data']['part_master']
                available_parts = sorted(part_master['FG Code'].dropna().unique().tolist())

                col1, col2 = st.columns(2)

                with col1:
                    new_part = st.selectbox(
                        "Part Code",
                        options=available_parts,
                        key="new_order_part"
                    )

                    new_qty = st.number_input(
                        "Quantity",
                        min_value=1,
                        max_value=10000,
                        value=100,
                        step=10,
                        key="new_order_qty"
                    )

                with col2:
                    new_customer = st.text_input(
                        "Customer ID",
                        value="NEW_CUSTOMER",
                        key="new_order_customer"
                    )

                    new_week = st.number_input(
                        "Delivery Week",
                        min_value=1,
                        max_value=30,
                        value=10,
                        key="new_order_week"
                    )

                if st.button("Add Order", type="primary", key="add_new_order"):
                    # Store new order
                    if 'new_orders' not in st.session_state:
                        st.session_state['new_orders'] = []

                    new_order = {
                        'Part_Code': new_part,
                        'Quantity': new_qty,
                        'Customer': new_customer,
                        'Delivery_Week': new_week
                    }
                    st.session_state['new_orders'].append(new_order)
                    st.success(f"✅ Added order: {new_qty} units of {new_part} for Week {new_week}")

                # Show pending new orders
                if 'new_orders' in st.session_state and st.session_state['new_orders']:
                    st.subheader("Pending New Orders")
                    new_orders_df = pd.DataFrame(st.session_state['new_orders'])
                    st.dataframe(new_orders_df, use_container_width=True)

                    if st.button("Clear All New Orders", key="clear_new_orders"):
                        st.session_state['new_orders'] = []
                        st.rerun()

            # Re-run optimization section
            st.markdown("---")
            st.subheader("🔄 Re-run Optimization")

            has_changes = (
                st.session_state.get('edited_orders') is not None or
                st.session_state.get('edited_wip') is not None or
                (st.session_state.get('new_orders') and len(st.session_state['new_orders']) > 0)
            )

            if has_changes:
                st.info("You have pending changes. Re-run optimization to apply them.")

                if st.button("🚀 Re-run with Changes", type="primary", use_container_width=True):
                    st.warning("⚠️ Re-running optimization with edited data is not yet implemented. "
                              "Please download the modified data and upload it as a new file.")

                    # Show what would change
                    st.markdown("### Pending Changes Summary")
                    if st.session_state.get('edited_orders') is not None:
                        st.write("• Sales orders modified")
                    if st.session_state.get('edited_wip') is not None:
                        st.write("• WIP inventory modified")
                    if st.session_state.get('new_orders'):
                        st.write(f"• {len(st.session_state['new_orders'])} new orders added")
            else:
                st.info("No changes to apply. Edit data above to make changes.")

    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: gray;'>"
        "Production Planning Optimization System | "
        f"Last updated: {datetime.now().strftime('%Y-%m-%d')}"
        "</div>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
