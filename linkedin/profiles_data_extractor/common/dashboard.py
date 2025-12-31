import json
import os
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# Page config is now handled in main.py

def load_config():
    """Load configuration from the JSON file in the same directory."""
    config_path = Path(__file__).parent / "linkedin_profiles_data.json"
    if not config_path.exists():
        st.error(f"Config file not found at {config_path}")
        return None
    
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        st.error(f"Error loading config: {e}")
        return None

def load_data(file_path):
    """Load data from the Excel file."""
    if not os.path.exists(file_path):
        st.error(f"Data file not found at: {file_path}")
        return None
    
    try:
        # Load Excel file
        df = pd.read_excel(file_path)
        
        # Ensure date columns are datetime
        if 'day' in df.columns:
            df['day'] = pd.to_datetime(df['day'], errors='coerce')
        
        if 'date connected' in df.columns:
            df['date connected'] = pd.to_datetime(df['date connected'], errors='coerce')
            
        return df
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return None

def create_performance_chart(df_filtered, df_all, group_col):
    if group_col not in df_filtered.columns or group_col not in df_all.columns:
        return None, None

    # Get total people from ALL data (unfiltered)
    total_stats = df_all.groupby(group_col).agg(
        Total_People=(group_col, 'count')
    ).reset_index()

    # Get contacted/connected from FILTERED data
    filtered_stats = df_filtered.groupby(group_col).agg(
        Contacted=('day', 'count'),  # Count of rows with contact date in filtered data
        Connected=('date connected', 'count')
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
        Contacted=('day', 'count')  # Count rows where 'day' is not null in filtered data
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

def generate_color_gradient(start_hex, end_hex, n):
    """Generate a gradient of n colors between start_hex and end_hex."""
    if n < 1: return []
    if n == 1: return [start_hex]
    
    def hex_to_rgb(h):
        return tuple(int(h.lstrip('#')[i:i+2], 16) for i in (0, 2, 4))
    
    start_rgb = hex_to_rgb(start_hex)
    end_rgb = hex_to_rgb(end_hex)
    
    colors = []
    for i in range(n):
        ratio = i / (n - 1)
        rgb = tuple(int(start_rgb[j] + (end_rgb[j] - start_rgb[j]) * ratio) for j in range(3))
        colors.append('#{:02x}{:02x}{:02x}'.format(*rgb))
        
    return colors

def main(df_filtered, df_all):
    """Main dashboard function that orchestrates the Streamlit app."""
    st.title("📊 LinkedIn Reachout Dashboard")

    # --- Performance Section ---
    
    total_contacts = len(df_filtered)
    connected_count = df_filtered['date connected'].notna().sum()
    conversion_rate = (connected_count / total_contacts * 100) if total_contacts > 0 else 0
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Contacted", f"{total_contacts}")
    col2.metric("Connected", f"{connected_count}")
    col3.metric("Conversion Rate", f"{conversion_rate:.1f}%")

    st.markdown("---")

    # --- Visualizations ---
    
    # 1. Contacts & Connections per Day (Time Series) & Response Time
    col_activity_left, col_activity_right = st.columns(2)

    with col_activity_left:
        st.subheader("📅 Activity Over Time")
        if 'day' in df_filtered.columns:
            # Group by day
            daily_counts = df_filtered.groupby(df_filtered['day'].dt.date).size().reset_index(name='Contacted')
            daily_connected = df_filtered[df_filtered['date connected'].notna()].groupby(df_filtered['day'].dt.date).size().reset_index(name='Connected')
            
            # Merge data for plotting
            daily_stats = pd.merge(daily_counts, daily_connected, on='day', how='left').fillna(0)
            
            fig_timeline = go.Figure()
            fig_timeline.add_trace(go.Bar(x=daily_stats['day'], y=daily_stats['Contacted'], name='Contacted', marker_color='#808080'))
            fig_timeline.add_trace(go.Bar(x=daily_stats['day'], y=daily_stats['Connected'], name='Connected', marker_color='#0B65C3'))
            
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
        if 'day' in df_filtered.columns and 'date connected' in df_filtered.columns:
            # Calculate days to respond
            df_resp = df_filtered.copy()
            df_resp['days_diff'] = (df_resp['date connected'] - df_resp['day']).dt.days
            
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
            st.dataframe(type_stats[['search_type', 'Total_People', 'Contacted', 'Contacted_Percent', 'Connected', 'Rate']].rename(columns={'Total_People': 'Total People', 'Rate': '% Connected', 'Contacted_Percent': '% Contacted'}), hide_index=True)

    with col_table_right:
        st.subheader("🏢 by Company")
        if 'company' in df_filtered.columns:
            st.dataframe(company_stats[['company', 'Total_People', 'Contacted', 'Contacted_Percent', 'Connected', 'Rate']].rename(columns={'Total_People': 'Total People', 'Rate': '% Connected', 'Contacted_Percent': '% Contacted'}), hide_index=True)

    # --- Raw Data ---
    st.header("📄 Raw Data")
    with st.expander("Show detailed records"):
        st.dataframe(df_filtered)

# Dashboard is now called from main.py
