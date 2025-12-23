import json
import os
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# Set page config
st.set_page_config(
    page_title="Reachout Dashboard",
    page_icon="📊",
    layout="wide"
)

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

def create_performance_chart(df, group_col, title):
    if group_col not in df.columns:
        return None, None
        
    stats = df.groupby(group_col).agg(
        Contacted=(group_col, 'count'),
        Connected=('date connected', 'count')
    ).reset_index()
    
    stats['Rate'] = (stats['Connected'] / stats['Contacted'] * 100).fillna(0).round(1)
    # User request: bar length is contacted + connected
    stats['Length'] = stats['Contacted'] + stats['Connected']
    
    # Sort by Rate ascending (so highest is at top in chart, as plotly builds from bottom)
    stats_chart = stats.sort_values('Rate', ascending=True)
    
    # Custom gradient from contacted gray (#808080) to connected green (#00A44E)
    custom_color_scale = ['#808080', '#00A44E']
    
    fig = px.bar(
        stats_chart, 
        x='Length', 
        y=group_col, 
        orientation='h',
        title=title, 
        color='Rate',
        color_continuous_scale=custom_color_scale,
        labels={'Length': 'Volume (Contacted + Connected)', 'Rate': 'Success Rate (%)', group_col: group_col.replace('_', ' ').title()}
    )
    
    # Return stats sorted by Contacted descending for tables (highest first)
    return fig, stats.sort_values('Contacted', ascending=False)

def main():
    st.title("📊 LinkedIn Reachout Dashboard")
    
    # Load config and data
    config = load_config()
    if not config:
        return

    data_path = config.get("destination_file")
    if not data_path:
        st.error("No 'destination_file' specified in config.")
        return

    df = load_data(data_path)
    if df is None:
        return

    # Sidebar Filters
    st.sidebar.header("Filters")
    
    # Date Filter
    if 'day' in df.columns:
        min_date = df['day'].min().date()
        max_date = df['day'].max().date()
        
        date_range = st.sidebar.date_input(
            "Select Date Range",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date
        )
        
        if len(date_range) == 2:
            start_date, end_date = date_range
            mask = (df['day'].dt.date >= start_date) & (df['day'].dt.date <= end_date)
            df_filtered = df.loc[mask]
        else:
            df_filtered = df
    else:
        df_filtered = df

    # Company Filter
    if 'company' in df.columns:
        companies = ['All'] + sorted(df['company'].dropna().unique().tolist())
        selected_company = st.sidebar.selectbox("Select Company", companies)
        if selected_company != 'All':
            df_filtered = df_filtered[df_filtered['company'] == selected_company]

    # Search Type Filter
    if 'search_type' in df.columns:
        search_types = ['All'] + sorted(df_filtered['search_type'].dropna().unique().tolist())
        selected_type = st.sidebar.selectbox("Select Search Type", search_types)
        if selected_type != 'All':
            df_filtered = df_filtered[df_filtered['search_type'] == selected_type]

    # --- Metrics Section ---
    st.header("📈 Key Metrics")
    
    total_contacts = len(df_filtered)
    connected_count = df_filtered['date connected'].notna().sum()
    conversion_rate = (connected_count / total_contacts * 100) if total_contacts > 0 else 0
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Contacted", f"{total_contacts}")
    col2.metric("Connected", f"{connected_count}")
    col3.metric("Conversion Rate", f"{conversion_rate:.1f}%")

    st.markdown("---")

    # --- Visualizations ---
    
    # 1. Contacts & Connections per Day (Time Series)
    st.subheader("📅 Activity Over Time")
    if 'day' in df_filtered.columns:
        # Group by day
        daily_counts = df_filtered.groupby(df_filtered['day'].dt.date).size().reset_index(name='Contacted')
        daily_connected = df_filtered[df_filtered['date connected'].notna()].groupby(df_filtered['day'].dt.date).size().reset_index(name='Connected')
        
        # Merge data for plotting
        daily_stats = pd.merge(daily_counts, daily_connected, on='day', how='left').fillna(0)
        
        fig_timeline = go.Figure()
        fig_timeline.add_trace(go.Bar(x=daily_stats['day'], y=daily_stats['Contacted'], name='Contacted', marker_color='#808080'))
        fig_timeline.add_trace(go.Bar(x=daily_stats['day'], y=daily_stats['Connected'], name='Connected', marker_color='#00A44E'))
        
        fig_timeline.update_layout(barmode='overlay', title="Daily Contacts vs Connections", xaxis_title="Date", yaxis_title="Count")
        st.plotly_chart(fig_timeline, width="stretch")

    col_left, col_right = st.columns(2)

    with col_left:
        # 2. Performance by Search Type
        st.subheader("🔍 Performance by Search Type")
        if 'search_type' in df_filtered.columns:
            fig_type, type_stats = create_performance_chart(df_filtered, 'search_type', "Performance by Search Type")
            st.plotly_chart(fig_type, width="stretch")
            
            st.dataframe(type_stats[['search_type', 'Contacted', 'Connected', 'Rate']], hide_index=True)

    with col_right:
        # 3. Performance by Company
        st.subheader("🏢 Performance by Company")
        if 'company' in df_filtered.columns:
            fig_company, company_stats = create_performance_chart(df_filtered, 'company', "Performance by Company")
            st.plotly_chart(fig_company, width="stretch")
            
            st.dataframe(company_stats[['company', 'Contacted', 'Connected', 'Rate']], hide_index=True)

    # --- Data Table ---
    st.subheader("📄 Raw Data")
    with st.expander("Show detailed records"):
        st.dataframe(df_filtered)

if __name__ == "__main__":
    main()
