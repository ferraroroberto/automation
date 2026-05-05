#!/usr/bin/env python3
"""
Household Inventory & Shopping Helper — entry point.

Mobile-responsive Streamlit app for managing grocery inventory.
Page config, shared session state, sidebar, and mode routing live here;
each mode is implemented in its own module (audit, edit_targets, edit_item,
add_item, shopping, export).

Usage:
    streamlit run app.py
"""

import os

import streamlit as st

import add_item
import audio_audit
import audit
import edit_item
import edit_targets
import export
import shopping
from data import CONFIG, COLUMNS, MODES, get_supermarket_stats, load_inventory_data
from ui_helpers import CSS, copy_to_clipboard, get_local_ip, open_inventory_spreadsheet


def _render_sidebar() -> None:
    """Render the sidebar: utility actions, mode selector, and live stats."""
    with st.sidebar:
        if st.button(
            "📂 Open spreadsheet",
            help="Opens the Excel file in the default app (e.g. Excel). Useful when OneDrive has not refreshed yet.",
            width="stretch",
        ):
            open_inventory_spreadsheet()

        local_url = f"https://{get_local_ip()}:8501"
        if st.button(
            "📋 Copy link",
            help=f"Copies {local_url} to clipboard — paste in Telegram to open on mobile.",
            width="stretch",
        ):
            copy_to_clipboard(local_url)
            st.success("✓ Copied!")

        if st.button("🔴 Close app", help="Stop the Streamlit server.", width="stretch"):
            os._exit(0)

        st.divider()

        mode_options = list(MODES.keys())
        mode_labels = list(MODES.values())

        selected_mode = st.radio(
            "Mode",
            mode_labels,
            index=mode_options.index(st.session_state.current_mode),
            label_visibility="collapsed",
        )

        current_mode_key = mode_options[mode_labels.index(selected_mode)]
        if current_mode_key != st.session_state.current_mode:
            st.session_state.current_mode = current_mode_key
            st.rerun()

        if st.session_state.inventory_data is not None:
            df_stats = st.session_state.inventory_data
            sm_shopping = df_stats[df_stats[COLUMNS["comprar"]] > 0].copy()

            st.divider()
            st.caption(f"{len(df_stats)} total tracked items")

            if not sm_shopping.empty:
                stats_by_supermarket = get_supermarket_stats(sm_shopping, st.session_state.bought_items)
                preferred_order = ["mercadona", "ametller"]
                ordered_supermarkets = [sm for sm in preferred_order if sm in stats_by_supermarket] + [
                    sm for sm in stats_by_supermarket if sm not in preferred_order
                ]

                for sm in ordered_supermarkets:
                    stats = stats_by_supermarket[sm]
                    offset_items = st.session_state.get(f"cart_offset_items_{sm}", 0)
                    offset_units = st.session_state.get(f"cart_offset_units_{sm}", 0)
                    done_u = stats["got_it_unique"] + offset_items
                    total_u = stats["total_unique"]
                    done_q = stats["got_it_quantity"] + offset_units
                    total_q = stats["total_quantity"]
                    st.markdown(f"**{sm.title()}**")
                    pm1, pm2 = st.columns(2)
                    with pm1:
                        st.metric("Items got", f"{done_u}/{total_u}")
                    with pm2:
                        st.metric("Units got", f"{done_q}/{total_q}")
                    oc1, oc2 = st.columns(2)
                    with oc1:
                        st.number_input(
                            "＋items",
                            value=0,
                            min_value=0,
                            step=1,
                            key=f"cart_offset_items_{sm}",
                        )
                    with oc2:
                        st.number_input(
                            "＋units",
                            value=0,
                            min_value=0,
                            step=1,
                            key=f"cart_offset_units_{sm}",
                        )


def _init_session_state() -> None:
    """Initialise session-state keys used across modes."""
    if "inventory_data" not in st.session_state:
        st.session_state.inventory_data = load_inventory_data()
        if st.session_state.inventory_data is None:
            st.stop()

    if "current_mode" not in st.session_state:
        st.session_state.current_mode = "audit"

    if "bought_items" not in st.session_state:
        st.session_state.bought_items = set()

    if "extra_shopping_items" not in st.session_state:
        st.session_state.extra_shopping_items = {}

    if "extra_bought_items" not in st.session_state:
        st.session_state.extra_bought_items = {}

    if "extra_item_counter" not in st.session_state:
        st.session_state.extra_item_counter = 0


def main() -> None:
    """Application entry point."""
    st.set_page_config(**CONFIG["ui"]["page_config"])
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown("### 🛒 Inventory & Shopping Helper")

    _init_session_state()
    _render_sidebar()

    df = st.session_state.inventory_data
    mode = st.session_state.current_mode

    if mode == "audit":
        st.session_state.inventory_data = audit.main(df)
    elif mode == "audio_audit":
        st.session_state.inventory_data = audio_audit.main(df)
    elif mode == "edit":
        st.session_state.inventory_data = edit_targets.main(df)
    elif mode == "edit_item":
        st.session_state.inventory_data = edit_item.main(df)
    elif mode == "add_item":
        st.session_state.inventory_data = add_item.main(df)
    elif mode == "shopping":
        shopping.main(df)
    elif mode == "export":
        export.main(df)


if __name__ == "__main__":
    main()
