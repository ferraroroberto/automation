#!/usr/bin/env python3
"""
Household Inventory & Shopping Helper
Mobile-responsive Streamlit application for managing grocery inventory.

Usage:
    streamlit run app.py

Requirements:
    pip install streamlit pandas openpyxl

Data Structure:
    XLSX file with columns: super, buscador, lugar, comida, cantidad, tenemos, comprar
"""

import json
import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

# Configuration and logging setup
with open("config.json", "r", encoding="utf-8") as f:
    CONFIG = json.load(f)

logging.basicConfig(
    level=getattr(logging, CONFIG["logging"]["level"]),
    format=CONFIG["logging"]["format"]
)
logger = logging.getLogger(__name__)

# Constants from config
COLUMNS = CONFIG["data"]["columns"]
MODES = CONFIG["ui"]["modes"]
UI_LABELS = CONFIG["data"]["ui_labels"]


def load_inventory_data() -> Optional[pd.DataFrame]:
    """Load inventory data from XLSX file."""
    xlsx_path = Path(CONFIG["data"]["xlsx_file"])
    if not xlsx_path.exists():
        st.error(f"❌ Inventory file not found: {xlsx_path}")
        logger.error(f"Inventory file not found: {xlsx_path}")
        return None

    try:
        df = pd.read_excel(xlsx_path, engine="openpyxl")
        required_columns = list(COLUMNS.values())

        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            st.error(f"❌ Missing required columns: {missing_cols}")
            logger.error(f"Missing columns in data: {missing_cols}")
            return None

        # Ensure numeric columns are integers
        df[COLUMNS["cantidad"]] = df[COLUMNS["cantidad"]].astype(int)
        df[COLUMNS["tenemos"]] = df[COLUMNS["tenemos"]].astype(int)

        # Calculate comprar column
        df[COLUMNS["comprar"]] = (df[COLUMNS["cantidad"]] - df[COLUMNS["tenemos"]]).clip(lower=0)

        logger.info(f"✅ Loaded inventory data: {len(df)} items")
        return df

    except Exception as e:
        st.error(f"❌ Error loading inventory data: {e}")
        logger.error(f"Error loading data: {e}")
        return None


def save_inventory_data(df: pd.DataFrame) -> bool:
    """Save inventory data back to XLSX file."""
    try:
        df.to_excel(CONFIG["data"]["xlsx_file"], index=False, engine="openpyxl")
        logger.info("✅ Inventory data saved successfully")
        return True
    except Exception as e:
        st.error(f"❌ Error saving inventory data: {e}")
        logger.error(f"Error saving data: {e}")
        return False


def get_unique_zones(df: pd.DataFrame) -> List[str]:
    """Get unique zones from inventory data."""
    return sorted(df[COLUMNS["lugar"]].unique().tolist())


def get_unique_supermarkets(df: pd.DataFrame) -> List[str]:
    """Get unique supermarkets from inventory data."""
    return sorted(df[COLUMNS["super"]].unique().tolist())


def update_item_quantity(df: pd.DataFrame, item_index: int, delta: int) -> pd.DataFrame:
    """Update the tenemos quantity for an item and recalculate comprar."""
    current_qty = df.at[item_index, COLUMNS["tenemos"]]
    new_qty = max(0, current_qty + delta)  # Prevent negative quantities

    df.at[item_index, COLUMNS["tenemos"]] = new_qty
    df.at[item_index, COLUMNS["comprar"]] = max(0, df.at[item_index, COLUMNS["cantidad"]] - new_qty)

    # Auto-save to Excel after each change
    save_inventory_data(df)

    logger.debug(f"Updated item {item_index}: tenemos={new_qty}, comprar={df.at[item_index, COLUMNS['comprar']]}")
    return df


def update_target_quantity(df: pd.DataFrame, item_index: int, delta: int) -> pd.DataFrame:
    """Update the cantidad (target) quantity for an item and recalculate comprar."""
    current_target = df.at[item_index, COLUMNS["cantidad"]]
    new_target = max(0, current_target + delta)  # Prevent negative quantities

    df.at[item_index, COLUMNS["cantidad"]] = new_target
    df.at[item_index, COLUMNS["comprar"]] = max(0, new_target - df.at[item_index, COLUMNS["tenemos"]])

    # Auto-save to Excel after each change
    save_inventory_data(df)

    logger.debug(f"Updated target for item {item_index}: cantidad={new_target}, comprar={df.at[item_index, COLUMNS['comprar']]}")
    return df


def render_audit_mode(df: pd.DataFrame) -> pd.DataFrame:
    """Render the audit mode interface."""
    st.header(MODES["audit"])

    zones = get_unique_zones(df)
    selected_zone = st.selectbox(
        "🏠 Select Zone",
        zones,
        help="Choose the area of your home to audit"
    )

    # Filter data for selected zone and only show items with target > 0
    zone_data = df[(df[COLUMNS["lugar"]] == selected_zone) & (df[COLUMNS["cantidad"]] > 0)].copy()

    if zone_data.empty:
        st.info(f"No items with target > 0 found in {selected_zone}")
        return df

    st.subheader(f"Items in {selected_zone.title()}")

    # Display items in a single-line mobile layout
    for idx in zone_data.index:
        item_name = zone_data.at[idx, COLUMNS["comida"]]
        current_qty = zone_data.at[idx, COLUMNS["tenemos"]]
        target_qty = zone_data.at[idx, COLUMNS["cantidad"]]
        buy_qty = zone_data.at[idx, COLUMNS["comprar"]]

        # Single line layout: item name | target | current | (+ - buttons) | buy
        with st.container():
            col1, col2, col3, col4, col5, col6 = st.columns([3, 1, 1, 1, 1, 1])

            with col1:
                st.write(f"**{item_name}**")

            with col2:
                st.metric("target", target_qty)

            with col3:
                st.metric("current", current_qty)

            with col4:
                # Large minus button for mobile
                if st.button("➖", key=f"minus_{idx}", help="Decrease quantity"):
                    df = update_item_quantity(df, idx, -1)
                    st.rerun()

            with col5:
                # Large plus button for mobile
                if st.button("➕", key=f"plus_{idx}", help="Increase quantity"):
                    df = update_item_quantity(df, idx, 1)
                    st.rerun()

            with col6:
                st.metric("buy", buy_qty)

        st.divider()

    return df


def render_edit_mode(df: pd.DataFrame) -> pd.DataFrame:
    """Render the edit mode interface for changing target quantities."""
    st.header(MODES["edit"])

    zones = get_unique_zones(df)
    selected_zone = st.selectbox(
        "🏠 Select Zone",
        zones,
        help="Choose the area of your home to edit targets"
    )

    # Filter data for selected zone
    zone_data = df[df[COLUMNS["lugar"]] == selected_zone].copy()

    if zone_data.empty:
        st.info(f"No items found in {selected_zone}")
        return df

    st.subheader(f"Edit Targets in {selected_zone.title()}")

    # Display items in a single-line mobile layout for editing
    for idx in zone_data.index:
        item_name = zone_data.at[idx, COLUMNS["comida"]]
        current_qty = zone_data.at[idx, COLUMNS["tenemos"]]
        target_qty = zone_data.at[idx, COLUMNS["cantidad"]]
        buy_qty = zone_data.at[idx, COLUMNS["comprar"]]

        # Single line layout: item name | target | current | (+ - buttons) | buy
        with st.container():
            col1, col2, col3, col4, col5, col6 = st.columns([3, 1, 1, 1, 1, 1])

            with col1:
                st.write(f"**{item_name}**")

            with col2:
                st.metric("target", target_qty)

            with col3:
                st.metric("current", current_qty)

            with col4:
                # Large minus button for target
                if st.button("➖", key=f"target_minus_{idx}", help="Decrease target quantity"):
                    df = update_target_quantity(df, idx, -1)
                    st.rerun()

            with col5:
                # Large plus button for target
                if st.button("➕", key=f"target_plus_{idx}", help="Increase target quantity"):
                    df = update_target_quantity(df, idx, 1)
                    st.rerun()

            with col6:
                st.metric("buy", buy_qty)

        st.divider()

    return df


def render_shopping_mode(df: pd.DataFrame) -> None:
    """Render the shopping list mode interface."""
    st.header(MODES["shopping"])

    # Initialize bought items tracking if not exists
    if "bought_items" not in st.session_state:
        st.session_state.bought_items = set()

    # Filter items that need to be purchased
    shopping_items = df[df[COLUMNS["comprar"]] > 0].copy()

    if shopping_items.empty:
        st.success("🎉 All items are in stock! No shopping needed.")
        return

    # Group by supermarket
    supermarkets = get_unique_supermarkets(shopping_items)

    total_items = len(shopping_items)
    bought_count = len([idx for idx in shopping_items.index if idx in st.session_state.bought_items])
    st.info(f"🛒 {total_items} items to buy from {len(supermarkets)} supermarket(s) | ✅ {bought_count} bought")

    # Clear all bought items button
    if bought_count > 0:
        col1, col2 = st.columns([4, 1])
        with col2:
            if st.button("🗑️ Clear All", help="Unmark all purchased items"):
                st.session_state.bought_items.clear()
                st.rerun()

    for supermarket in supermarkets:
        supermarket_items = shopping_items[shopping_items[COLUMNS["super"]] == supermarket]

        with st.expander(f"🏪 {supermarket.title()} ({len(supermarket_items)} items)", expanded=True):
            for idx in supermarket_items.index:
                item_name = supermarket_items.at[idx, COLUMNS["comida"]]
                qty_to_buy = supermarket_items.at[idx, COLUMNS["comprar"]]
                buy_url = supermarket_items.at[idx, COLUMNS["buscador"]]
                is_bought = idx in st.session_state.bought_items

                col1, col2, col3, col4 = st.columns([3, 1, 1, 1])

                with col1:
                    if is_bought:
                        st.write(f"~~{item_name}~~")  # Strike through for bought items
                    else:
                        st.write(f"**{item_name}**")

                with col2:
                    if is_bought:
                        st.write(f"✅ {qty_to_buy}")
                    else:
                        st.write(f"buy: {qty_to_buy}")

                with col3:
                    # Buy now link button
                    if is_bought:
                        st.link_button(
                            "🔄 Buy Again",
                            buy_url,
                            help=f"Open {supermarket} product page",
                            use_container_width=True
                        )
                    else:
                        st.link_button(
                            "🛒 Buy Now",
                            buy_url,
                            help=f"Open {supermarket} product page",
                            use_container_width=True
                        )

                with col4:
                    # Mark as bought/not bought button
                    if is_bought:
                        if st.button(
                            "↩️ Unmark",
                            key=f"unmark_{idx}",
                            help="Mark as not bought yet",
                            use_container_width=True
                        ):
                            st.session_state.bought_items.remove(idx)
                            st.rerun()
                    else:
                        if st.button(
                            "✅ Got It",
                            key=f"mark_{idx}",
                            help="Mark as purchased",
                            use_container_width=True,
                            type="secondary"
                        ):
                            st.session_state.bought_items.add(idx)
                            st.rerun()

            st.divider()


def render_export_mode(df: pd.DataFrame) -> None:
    """Render the save/export mode interface."""
    st.header(MODES["export"])

    col1, col2 = st.columns(2)

    with col1:
        if st.button("💾 Save to File", type="primary", use_container_width=True):
            if save_inventory_data(df):
                st.success("✅ Inventory saved successfully!")
            else:
                st.error("❌ Failed to save inventory")

    with col2:
        # Create download button
        csv_data = df.to_csv(index=False)
        st.download_button(
            "📥 Download CSV",
            csv_data,
            "inventory_updated.csv",
            "text/csv",
            use_container_width=True,
            help="Download updated inventory as CSV file"
        )

    # Show summary statistics
    st.subheader("📊 Summary")
    total_items = len(df)
    stocked_items = len(df[df[COLUMNS["comprar"]] == 0])
    shopping_items = len(df[df[COLUMNS["comprar"]] > 0])

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Items", total_items)
    with col2:
        st.metric("Fully Stocked", stocked_items)
    with col3:
        st.metric("Need Shopping", shopping_items)


def main():
    """Main application entry point."""
    # Configure page
    st.set_page_config(**CONFIG["ui"]["page_config"])

    # Title and description
    st.title(CONFIG["app"]["title"])
    st.markdown(CONFIG["app"]["description"])

    # Initialize session state
    if "inventory_data" not in st.session_state:
        st.session_state.inventory_data = load_inventory_data()
        if st.session_state.inventory_data is None:
            st.stop()

    if "current_mode" not in st.session_state:
        st.session_state.current_mode = "audit"

    # Sidebar navigation
    with st.sidebar:
        st.header("📱 Navigation")

        # Mode selection
        mode_options = list(MODES.keys())
        mode_labels = list(MODES.values())

        selected_mode = st.radio(
            "Choose Mode:",
            mode_labels,
            index=mode_options.index(st.session_state.current_mode)
        )

        # Update current mode
        current_mode_key = mode_options[mode_labels.index(selected_mode)]
        if current_mode_key != st.session_state.current_mode:
            st.session_state.current_mode = current_mode_key
            st.rerun()

        # Data status
        st.divider()
        st.subheader("📊 Data Status")
        if st.session_state.inventory_data is not None:
            df = st.session_state.inventory_data
            total_items = len(df)
            shopping_needed = len(df[df[COLUMNS["comprar"]] > 0])
            st.write(f"Total items: {total_items}")
            st.write(f"Need shopping: {shopping_needed}")

    # Main content area
    df = st.session_state.inventory_data

    if st.session_state.current_mode == "audit":
        st.session_state.inventory_data = render_audit_mode(df)
    elif st.session_state.current_mode == "edit":
        st.session_state.inventory_data = render_edit_mode(df)
    elif st.session_state.current_mode == "shopping":
        render_shopping_mode(df)
    elif st.session_state.current_mode == "export":
        render_export_mode(df)

    # Footer
    st.divider()
    st.caption(f"v{CONFIG['app']['version']} - Mobile-optimized for inventory management")


if __name__ == "__main__":
    main()
