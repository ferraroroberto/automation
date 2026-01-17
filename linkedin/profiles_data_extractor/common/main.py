import streamlit as st
from pathlib import Path
import sys
import os
import json
import pandas as pd

# Add the current directory to Python path to import local modules
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

from dashboard import main as dashboard_main
from dataentry import main as dataentry_main
from reachout import main as reachout_main
from extract_data import main as extract_data_main
from history import main as history_main

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

def load_excel_data(file_path):
    """Load data from the Excel file."""
    if not os.path.exists(file_path):
        st.error(f"Data file not found at: {file_path}")
        return None

    try:
        df = pd.read_excel(file_path)

        # Ensure date columns are datetime
        date_columns = ['date_contacted', 'date_connected', 'date_revocation']
        for col in date_columns:
            if col not in df.columns:
                df[col] = pd.NaT  # Create missing column
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')

        return df
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return None

def get_current_theme_mode():
    """Determine if the current theme is light or dark based on config.toml."""
    config_path = Path(__file__).parent / ".streamlit" / "config.toml"
    
    # Default to dark if file doesn't exist or can't be read
    if not config_path.exists():
        return "dark"
        
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            content = f.read()
            # Simple check for background color
            if 'backgroundColor = "#FFFFFF"' in content:
                return "light"
            else:
                return "dark"
    except:
        return "dark"

def toggle_theme():
    """Toggle between light and dark mode by updating config.toml."""
    config_dir = Path(__file__).parent / ".streamlit"
    config_path = config_dir / "config.toml"

    # Ensure directory exists
    config_dir.mkdir(exist_ok=True)

    current_mode = get_current_theme_mode()

    if current_mode == "light":
        # Switch to Dark
        new_config = """[theme]
primaryColor = "#1E88E5"
backgroundColor = "#0E1117"
secondaryBackgroundColor = "#262730"
textColor = "#FAFAFA"
font = "sans serif"
"""
    else:
        # Switch to Light
        new_config = """[theme]
primaryColor = "#1E88E5"
backgroundColor = "#FFFFFF"
secondaryBackgroundColor = "#F0F2F6"
textColor = "#262730"
font = "sans serif"
"""

    with open(config_path, "w", encoding="utf-8") as f:
        f.write(new_config)

    # Trigger a rerun to apply changes
    st.rerun()


def main():
    """Main navigation hub for LinkedIn Reachout Dashboard."""

    # Set page config
    st.set_page_config(
        page_title="LinkedIn Reachout Hub",
        page_icon="🔗",
        layout="wide"
    )

    # Reduce spacing before title
    st.markdown("<style>.block-container { padding-top: 1.5rem; }</style>", unsafe_allow_html=True)

    st.title("🔗 LinkedIn Reachout Hub")

    # Load config and data once for shared use
    config = load_config()
    if not config:
        return

    data_path = config.get("destination_file")
    if not data_path:
        st.error("No 'destination_file' specified in config.")
        return

    df = load_excel_data(data_path)
    if df is None:
        return

    # Shared Sidebar Filters
    st.sidebar.header("App Settings")
    
    # Theme Toggle
    current_mode = get_current_theme_mode()
    if st.sidebar.button(f"Switch to {'Light' if current_mode == 'dark' else 'Dark'} Mode"):
        toggle_theme()
        
    st.sidebar.divider()
    st.sidebar.header("Filters")

    # Note about Data Editor
    st.sidebar.info("ℹ️ Filters apply to Dashboard and Reachout tabs only. Data Editor is not affected by these filters.")

    # Company Filter
    df_filtered = df.copy()
    selected_company = 'All'
    if 'company' in df_filtered.columns:
        companies = ['All'] + sorted(df_filtered['company'].dropna().unique().tolist())
        selected_company = st.sidebar.selectbox("Select Company", companies)
        if selected_company != 'All':
            df_filtered = df_filtered[df_filtered['company'] == selected_company]

    # Search Type Filter
    selected_type = 'All'
    if 'search_type' in df_filtered.columns:
        search_types = ['All'] + sorted(df_filtered['search_type'].dropna().unique().tolist())
        selected_type = st.sidebar.selectbox("Select Search Type", search_types)
        if selected_type != 'All':
            df_filtered = df_filtered[df_filtered['search_type'] == selected_type]

    # Date Contacted Filter
    st.sidebar.subheader("Contact Filters")

    # Date contacted range filter
    if 'date_contacted' in df.columns:
        min_date = df['date_contacted'].min()
        max_date = df['date_contacted'].max()

        if pd.notna(min_date) and pd.notna(max_date):
            # Convert to date objects for streamlit
            min_date = min_date.date()
            max_date = max_date.date()

            col_date1, col_date2 = st.sidebar.columns(2)
            with col_date1:
                start_date = st.sidebar.date_input("Contacted From", value=min_date, min_value=min_date, max_value=max_date)
            with col_date2:
                end_date = st.sidebar.date_input("Contacted To", value=max_date, min_value=min_date, max_value=max_date)

            # Apply date filter
            # Convert date objects to pandas Timestamps for proper comparison
            start_ts = pd.Timestamp(start_date)  # Start of day
            end_ts = pd.Timestamp(end_date) + pd.Timedelta(days=1)  # Start of next day (exclusive)
            df_filtered = df_filtered[
                (df_filtered['date_contacted'].isna()) |
                ((df_filtered['date_contacted'] >= start_ts) & (df_filtered['date_contacted'] < end_ts))
            ]

    # Contacted Yes/No filter
    contacted_filter = st.sidebar.selectbox(
        "Contacted Status",
        ["All", "Contacted Only", "Uncontacted Only"],
        help="Filter by whether profiles have been contacted"
    )

    if contacted_filter == "Contacted Only":
        df_filtered = df_filtered[df_filtered['date_contacted'].notna()]
    elif contacted_filter == "Uncontacted Only":
        df_filtered = df_filtered[df_filtered['date_contacted'].isna()]

    # Connected Yes/No filter
    connected_filter = st.sidebar.selectbox(
        "Connected Status",
        ["All", "Connected Only", "Unconnected Only"],
        help="Filter by whether profiles have connected"
    )

    if connected_filter == "Connected Only":
        df_filtered = df_filtered[df_filtered['date_connected'].notna()]
    elif connected_filter == "Unconnected Only":
        df_filtered = df_filtered[df_filtered['date_connected'].isna()]

    # Create tabs for navigation
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Dashboard", "✏️ Data Entry", "🎯 Reachout Manager", "📥 Extract Data", "📜 History"])

    with tab1:
        # Prepare filter parameters for dashboard
        filter_params = {
            'selected_company': selected_company,
            'selected_type': selected_type,
            'start_date': start_date if 'date_contacted' in df.columns and pd.notna(df['date_contacted'].min()) else None,
            'end_date': end_date if 'date_contacted' in df.columns and pd.notna(df['date_contacted'].max()) else None,
            'contacted_filter': contacted_filter,
            'connected_filter': connected_filter
        }
        dashboard_main(df_filtered, df, filter_params)  # Pass filtered data, full data, and all filter selections

    with tab2:
        dataentry_main()

    with tab3:
        reachout_main(df_filtered, df)  # Pass both filtered and full datasets

    with tab4:
        extract_data_main()

    with tab5:
        history_main()

if __name__ == "__main__":
    main()