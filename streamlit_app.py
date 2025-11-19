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


def run_optimization(uploaded_file, config):
    """Run the optimization and return results."""
    # Save uploaded file to temp location
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_path = tmp_file.name

    try:
        # Load data
        loader = ComprehensiveDataLoader(tmp_path, config)
        data = loader.load_all_data()

        # Calculate demand with stage-wise skip logic
        calculator = WIPDemandCalculator(data['sales_order'], data['stage_wip'], config)
        (net_demand, stage_start_qty, wip_coverage,
         gross_demand, wip_by_part) = calculator.calculate_net_demand_with_stages()
        split_demand, part_week_mapping, variant_windows = calculator.split_demand_by_week(net_demand)

        # Build parameters
        param_builder = ComprehensiveParameterBuilder(data['part_master'], config)
        params = param_builder.build_parameters()

        # Setup resources
        machine_manager = MachineResourceManager(data['machine_constraints'], config)
        box_manager = BoxCapacityManager(data['box_capacity'], config, machine_manager)

        # Build WIP init
        wip_init = build_wip_init(data['stage_wip'])

        # Build and solve model
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
        status = optimizer.build_and_solve()

        # Extract results
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

        # Generate fulfillment reports
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

        # Generate daily schedule
        daily_generator = DailyScheduleGenerator(
            results['weekly_summary'],
            results,
            config
        )
        daily_schedule = daily_generator.generate_daily_schedule()
        part_daily_schedule = daily_generator.generate_part_level_daily_schedule(data['part_master'])

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
            'wip_init': wip_init
        }

        return all_results

    finally:
        # Cleanup temp file
        os.unlink(tmp_path)


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

    elif run_button:
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

        with st.spinner("Running optimization... This may take a few minutes."):
            try:
                # Progress indicator
                progress_bar = st.progress(0)
                status_text = st.empty()

                status_text.text("Loading data...")
                progress_bar.progress(10)

                # Run optimization
                all_results = run_optimization(uploaded_file, config)

                progress_bar.progress(100)
                status_text.text("Optimization complete!")

                # Store results in session state
                st.session_state['optimization_results'] = all_results
                st.session_state['optimization_complete'] = True

                st.success("✅ Optimization completed successfully!")
                st.rerun()

            except Exception as e:
                st.error(f"❌ Error during optimization: {str(e)}")
                st.exception(e)

    # Display results if available
    if st.session_state.get('optimization_complete', False):
        all_results = st.session_state['optimization_results']
        results = all_results['results']
        fulfillment_reports = all_results['fulfillment_reports']

        # Results tabs
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "📊 Dashboard",
            "📈 Capacity Analysis",
            "📦 Production Schedule",
            "🚚 Delivery Tracking",
            "🔮 What-If Analysis",
            "📥 Download"
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
            whatif_tab1, whatif_tab2, whatif_tab3 = st.tabs([
                "🔧 Capacity Scenarios",
                "📊 Demand Scaling",
                "⚠️ Bottleneck Analysis"
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
