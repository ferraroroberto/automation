import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from loaders import load_config, load_excel_data  # noqa: F401 — re-exported for callers
from _lib import generate_color_gradient

# Page config is now handled in main.py

def create_performance_chart(df_filtered, df_all, group_col):
    if group_col not in df_filtered.columns or group_col not in df_all.columns:
        return None, None

    # Get total people from ALL data (unfiltered)
    total_stats = df_all.groupby(group_col).agg(
        Total_People=(group_col, 'count')
    ).reset_index()

    # Get contacted/connected from FILTERED data
    filtered_stats = df_filtered.groupby(group_col).agg(
        Contacted=('date_contacted', 'count'),  # Count of rows with contact date in filtered data
        Connected=('date_connected', 'count')
    ).reset_index()

    # Merge the stats
    stats = pd.merge(total_stats, filtered_stats, on=group_col, how='left').fillna(0)

    stats['Rate'] = (stats['Connected'] / stats['Contacted'] * 100).fillna(0).round(1)
    stats['Contacted_Percent'] = (stats['Contacted'] / stats['Total_People'] * 100).fillna(0).round(1)
    # User request: bar length is contacted + connected
    stats['Length'] = stats['Contacted'] + stats['Connected']

    # Sort by Rate ascending (so highest is at top in chart, as plotly builds from bottom)
    stats_chart = stats.sort_values('Rate', ascending=True)

    # Custom gradient from contacted gray (#808080) to connected blue (#0B65C3)
    custom_color_scale = ['#808080', '#0B65C3']

    fig = px.bar(
        stats_chart,
        x='Length',
        y=group_col,
        orientation='h',
        color='Rate',
        color_continuous_scale=custom_color_scale,
        labels={'Length': 'Volume (Contacted + Connected)', 'Rate': 'Success Rate (%)', group_col: group_col.replace('_', ' ').title()}
    )

    fig.update_layout(
        margin=dict(t=10, b=10)
    )

    # Return stats sorted by Contacted descending for tables (highest first)
    return fig, stats.sort_values('Contacted', ascending=False)

def create_contact_chart(df_filtered, df_all, group_col):
    """Create contact chart showing total people vs contacted people."""
    if group_col not in df_filtered.columns or group_col not in df_all.columns:
        return None, None

    # Get total people from ALL data (unfiltered)
    total_stats = df_all.groupby(group_col).agg(
        Total_People=(group_col, 'count')
    ).reset_index()

    # Get contacted from FILTERED data
    filtered_stats = df_filtered.groupby(group_col).agg(
        Contacted=('date_contacted', 'count')  # Count rows where 'date_contacted' is not null in filtered data
    ).reset_index()

    # Merge the stats
    stats = pd.merge(total_stats, filtered_stats, on=group_col, how='left').fillna(0)

    stats['Contact_Rate'] = (stats['Contacted'] / stats['Total_People'] * 100).fillna(0).round(1)

    # Sort by Total_People ascending for chart (highest at top)
    stats_chart = stats.sort_values('Total_People', ascending=True)

    # Create overlay bar chart: grey bars show total people, green bars show contacted subset
    fig = go.Figure()

    # Add total people bars (grey background)
    fig.add_trace(go.Bar(
        x=stats_chart['Total_People'],
        y=stats_chart[group_col],
        orientation='h',
        name='Total People',
        marker_color='#808080',
        showlegend=True
    ))

    # Add contacted people bars (blue overlay)
    fig.add_trace(go.Bar(
        x=stats_chart['Contacted'],
        y=stats_chart[group_col],
        orientation='h',
        name='Contacted',
        marker_color='#0B65C3',
        showlegend=True
    ))

    fig.update_layout(
        barmode='overlay',  # Overlay mode: bars are drawn on top of each other, total width = max bar value
                            # Unlike 'stack' mode which sums bar values, overlay shows subsets within the total
        xaxis_title="Count",
        yaxis_title=group_col.replace('_', ' ').title(),
        legend_title="Legend",
        margin=dict(t=10, b=10)
    )

    # Return stats sorted by Total_People descending for tables
    return fig, stats.sort_values('Total_People', ascending=False)

def main(df_filtered, df_all, filter_params=None):
    """Main dashboard function that orchestrates the Streamlit app."""

    # Extract filter parameters
    if filter_params is None:
        filter_params = {}

    selected_company = filter_params.get('selected_company', 'All')
    selected_search_type = filter_params.get('selected_type', 'All')
    start_date = filter_params.get('start_date')
    end_date = filter_params.get('end_date')
    contacted_filter = filter_params.get('contacted_filter', 'All')
    connected_filter = filter_params.get('connected_filter', 'All')

    # Store filter info for raw data display
    st.session_state['df_filtered'] = df_filtered
    st.session_state['df_all'] = df_all
    st.session_state['selected_company'] = selected_company
    st.session_state['selected_search_type'] = selected_search_type
    st.session_state['filter_params'] = filter_params

    # --- Performance Section ---
    
    total_contacts = df_filtered['date_contacted'].notna().sum()
    connected_count = df_filtered['date_connected'].notna().sum()
    conversion_rate = (connected_count / total_contacts * 100) if total_contacts > 0 else 0
    total_records = len(df_all)
    contacted_percentage = (total_contacts / total_records * 100) if total_records > 0 else 0

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total People", f"{total_records}")
    col2.metric("Contacted", f"{total_contacts}")
    col3.metric("% Contacted", f"{contacted_percentage:.1f}%")
    col4.metric("Connected", f"{connected_count}")
    col5.metric("% Connected", f"{conversion_rate:.1f}%")

    st.markdown("---")

    # --- Visualizations ---
    
    # 1. Contacts & Connections per Day (Time Series) & Response Time
    col_activity_left, col_activity_right = st.columns(2)

    with col_activity_left:
        st.subheader("📅 Activity Over Time")
        if 'date_contacted' in df_filtered.columns:
            # Group by day
            daily_counts = df_filtered.groupby(df_filtered['date_contacted'].dt.date).size().reset_index(name='Contacted')
            daily_connected = df_filtered[df_filtered['date_connected'].notna()].groupby(df_filtered['date_contacted'].dt.date).size().reset_index(name='Connected')
            
            # Merge data for plotting
            daily_stats = pd.merge(daily_counts, daily_connected, on='date_contacted', how='left').fillna(0)
            
            fig_timeline = go.Figure()
            fig_timeline.add_trace(go.Bar(x=daily_stats['date_contacted'], y=daily_stats['Contacted'], name='Contacted', marker_color='#808080'))
            fig_timeline.add_trace(go.Bar(x=daily_stats['date_contacted'], y=daily_stats['Connected'], name='Connected', marker_color='#0B65C3'))
            
            fig_timeline.update_layout(
                barmode='overlay', 
                xaxis_title="Date", 
                yaxis_title="Count",
                margin=dict(t=10, b=10)
            )
            # barmode='overlay': Connected bars overlay on Contacted bars, showing subset relationship
            # Bar width = Contacted (total), green overlay shows Connected (subset)
            st.plotly_chart(fig_timeline, width='stretch')

    with col_activity_right:
        st.subheader("⏱️ Response Time Distribution")
        if 'date_contacted' in df_filtered.columns and 'date_connected' in df_filtered.columns:
            # Calculate days to respond - only for records where date_contacted is not null (contacted records)
            df_resp = df_filtered[df_filtered['date_contacted'].notna()].copy()
            df_resp['days_diff'] = (df_resp['date_connected'] - df_resp['date_contacted']).dt.days
            
            def get_label(x):
                if pd.isna(x):
                    return "Never"
                elif x == 0:
                    return "0 days"
                elif 1 <= x <= 2:
                    return "1-2 days"
                elif 3 <= x <= 5:
                    return "3-5 days"
                else:  # x > 5
                    return "more than 5 days"
            
            df_resp['response_label'] = df_resp['days_diff'].apply(get_label)
            
            # Count for pie chart
            pie_data = df_resp['response_label'].value_counts().reset_index()
            pie_data.columns = ['Label', 'Count']
            
            # Sort: 0 days, 1-2 days, 3-5 days, more than 5 days, Never
            def sort_key(label):
                if label == "0 days":
                    return 0
                elif label == "1-2 days":
                    return 1
                elif label == "3-5 days":
                    return 2
                elif label == "more than 5 days":
                    return 3
                elif label == "Never":
                    return 4
                else:
                    return 5
            
            pie_data['sort_key'] = pie_data['Label'].apply(sort_key)
            pie_data = pie_data.sort_values('sort_key')
            
            # Generate colors: Blue (#0B65C3) -> Gray (#808080)
            # 0 days is Blue, Never is Gray
            colors = generate_color_gradient('#0B65C3', '#808080', len(pie_data))

            fig_pie = px.pie(
                pie_data,
                values='Count',
                names='Label',
                category_orders={'Label': pie_data['Label'].tolist()},
                color_discrete_sequence=colors,
                hole=0.4
            )
            fig_pie.update_traces(sort=False, textinfo='percent+label')
            fig_pie.update_layout(margin=dict(t=10, b=10))
            st.plotly_chart(fig_pie, width='stretch')

    col_left, col_right = st.columns(2)

    with col_left:
        # 2. Performance by Search Type
        st.subheader("🔍 Performance by Search Type")
        if 'search_type' in df_filtered.columns:
            fig_type, type_stats = create_performance_chart(df_filtered, df_all, 'search_type')
            st.plotly_chart(fig_type, width="stretch")

    with col_right:
        # 3. Performance by Company
        st.subheader("🏢 Performance by Company")
        if 'company' in df_filtered.columns:
            fig_company, company_stats = create_performance_chart(df_filtered, df_all, 'company')
            st.plotly_chart(fig_company, width="stretch")

    # Contact Overview Charts

    col_contact_left, col_contact_right = st.columns(2)

    with col_contact_left:
        # 4. Contact Overview by Search Type
        st.subheader("🔍 Contact Overview by Search Type")
        if 'search_type' in df_filtered.columns:
            fig_contact_type, contact_type_stats = create_contact_chart(df_filtered, df_all, 'search_type')
            st.plotly_chart(fig_contact_type, width="stretch")

    with col_contact_right:
        # 5. Contact Overview by Company
        st.subheader("🏢 Contact Overview by Company")
        if 'company' in df_filtered.columns:
            fig_contact_company, contact_company_stats = create_contact_chart(df_filtered, df_all, 'company')
            st.plotly_chart(fig_contact_company, width="stretch")

    # --- Data Tables ---
    st.header("📊 Data Tables")

    # Performance Data Tables
    col_table_left, col_table_right = st.columns(2)

    with col_table_left:
        st.subheader("🔍 by Search Type")
        if 'search_type' in df_filtered.columns:
            st.dataframe(type_stats[['search_type', 'Total_People', 'Contacted', 'Contacted_Percent', 'Connected', 'Rate']].rename(columns={'search_type': 'Search Type', 'Total_People': 'Total People', 'Rate': '% Connected', 'Contacted_Percent': '% Contacted'}), hide_index=True)

    with col_table_right:
        st.subheader("🏢 by Company")
        if 'company' in df_filtered.columns:
            st.dataframe(company_stats[['company', 'Total_People', 'Contacted', 'Contacted_Percent', 'Connected', 'Rate']].rename(columns={'Total_People': 'Total People', 'Rate': '% Connected', 'Contacted_Percent': '% Contacted'}), hide_index=True)

    # --- Raw Data ---
    st.header("📄 Raw Data")

    # Show current filter status
    total_records = len(df_all)

    # Determine which data to show in raw data table
    # If date contacted filter is applied, show only contacted records
    date_filter_applied = False
    if start_date and end_date and 'date_contacted' in df_all.columns:
        min_date_all = df_all['date_contacted'].min()
        max_date_all = df_all['date_contacted'].max()
        if pd.notna(min_date_all) and pd.notna(max_date_all):
            min_date_all = min_date_all.date()
            max_date_all = max_date_all.date()
            if start_date > min_date_all or end_date < max_date_all:
                date_filter_applied = True

    if date_filter_applied:
        # Show metrics for contacted records when date filter is applied
        contacted_records = df_filtered[df_filtered['date_contacted'].notna()]
        filtered_records = len(contacted_records)
        filter_percentage = (filtered_records / total_records * 100) if total_records > 0 else 0
    else:
        # Show metrics for all filtered records for other filters
        filtered_records = len(df_filtered)
        filter_percentage = (filtered_records / total_records * 100) if total_records > 0 else 0

    col_info, col_filters = st.columns([1, 2])

    with col_info:
        st.metric("Total Records", f"{total_records:,}")
        st.metric("Filtered Records", f"{filtered_records:,}")
        if total_records > 0:
            st.metric("Filter Coverage", f"{filter_percentage:.1f}%")

    with col_filters:
        st.subheader("Current Filters Applied:")

        filters_applied = []

        # Show company filter if applied
        if selected_company != 'All':
            st.write(f"**🏢 Company:** {selected_company}")
            filters_applied.append("Company")

        # Show search type filter if applied
        if selected_search_type != 'All':
            st.write(f"**🔍 Search Type:** {selected_search_type}")
            filters_applied.append("Search Type")

        # Show date contacted filter if applied
        if start_date and end_date and 'date_contacted' in df_all.columns:
            # Check if date range is different from min/max to determine if filter is applied
            min_date_all = df_all['date_contacted'].min()
            max_date_all = df_all['date_contacted'].max()
            if pd.notna(min_date_all) and pd.notna(max_date_all):
                min_date_all = min_date_all.date()
                max_date_all = max_date_all.date()
                # Only show as applied if the range is more restrictive than the full data range
                if start_date > min_date_all or end_date < max_date_all:
                    st.write(f"**📅 Date Contacted:** {start_date} to {end_date}")
                    filters_applied.append("Date Contacted")

        # Show contacted status filter if applied
        if contacted_filter != 'All':
            status_text = "Contacted Only" if contacted_filter == "Contacted Only" else "Uncontacted Only"
            st.write(f"**✅ Contacted Status:** {status_text}")
            filters_applied.append("Contacted Status")

        # Show connected status filter if applied
        if connected_filter != 'All':
            status_text = "Connected Only" if connected_filter == "Connected Only" else "Unconnected Only"
            st.write(f"**🔗 Connected Status:** {status_text}")
            filters_applied.append("Connected Status")

        # If no filters are applied
        if not filters_applied:
            st.write("*No filters currently applied - showing all records*")
        else:
            st.write(f"*Showing records that match: {', '.join(filters_applied)} filters*")

    # Determine which data to show in raw data table
    # If date contacted filter is applied, show only contacted records
    date_filter_applied = False
    if start_date and end_date and 'date_contacted' in df_all.columns:
        min_date_all = df_all['date_contacted'].min()
        max_date_all = df_all['date_contacted'].max()
        if pd.notna(min_date_all) and pd.notna(max_date_all):
            min_date_all = min_date_all.date()
            max_date_all = max_date_all.date()
            if start_date > min_date_all or end_date < max_date_all:
                date_filter_applied = True

    if date_filter_applied:
        # Show only contacted records when date filter is applied
        contacted_records = df_filtered[df_filtered['date_contacted'].notna()].copy()
        raw_data_count = len(contacted_records)
        raw_data_title = f"Show {raw_data_count:,} contacted records (filtered by date)"
        display_df = contacted_records
    else:
        # Show all filtered records for other filters
        raw_data_count = filtered_records
        raw_data_title = f"Show {raw_data_count:,} filtered records"
        display_df = df_filtered

    with st.expander(raw_data_title):
        st.dataframe(display_df)

# Dashboard is now called from main.py
