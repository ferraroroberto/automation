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
import os
import platform
import subprocess
from pathlib import Path
from typing import Dict, List, Optional

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

_SPREADSHEET_LOCKED_HINT = (
    "The spreadsheet is open in Excel or locked by OneDrive. "
    "Close it in Excel, wait for sync, then try again."
)


def _is_spreadsheet_lock_error(exc: BaseException) -> bool:
    if isinstance(exc, PermissionError):
        return True
    if isinstance(exc, OSError):
        errno = getattr(exc, "errno", None)
        if errno is not None and errno in (13, 11):  # EACCES, EAGAIN (common when file is busy)
            return True
    lowered = str(exc).lower()
    return any(
        phrase in lowered
        for phrase in (
            "permission denied",
            "being used by another process",
            "access is denied",
            "the process cannot access the file",
        )
    )


def open_inventory_spreadsheet() -> None:
    """Open the configured XLSX in the default application (e.g. Excel)."""
    path = Path(CONFIG["data"]["xlsx_file"]).expanduser().resolve()
    if not path.exists():
        st.sidebar.error(f"File not found:\n`{path}`")
        logger.error("Open spreadsheet: file missing at %s", path)
        return
    try:
        if platform.system() == "Windows":
            os.startfile(str(path))  # noqa: S606
        elif platform.system() == "Darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
        st.sidebar.success("Opened. Close the workbook before saving changes from this app.")
    except OSError as e:
        st.sidebar.error(f"Could not open file: {e}")
        logger.error("Open spreadsheet failed: %s", e)


# Constants from config
COLUMNS = CONFIG["data"]["columns"]
MODES = CONFIG["ui"]["modes"]
UI_LABELS = CONFIG["data"]["ui_labels"]

# Injected CSS: tighter padding, compact buttons, reduced gaps
_CSS = """
<style>
/* Shrink top/bottom padding of main area */
.main .block-container {
    padding-top: 0.75rem !important;
    padding-bottom: 0.5rem !important;
}
/* Tighter column cell padding */
[data-testid="column"] {
    padding-left: 0.15rem !important;
    padding-right: 0.15rem !important;
}
/* Smaller buttons */
.stButton > button {
    padding: 0.2rem 0.55rem !important;
    font-size: 0.92rem !important;
    line-height: 1.3 !important;
    min-height: 0 !important;
}
/* Compact link buttons */
.stLinkButton > a {
    padding: 0.2rem 0.55rem !important;
    font-size: 0.85rem !important;
}
/* Compact dividers */
hr { margin: 0.2rem 0 !important; }
/* Compact expander header */
details > summary {
    padding: 0.35rem 0.6rem !important;
    font-size: 0.9rem !important;
}
/* Sidebar compactness */
section[data-testid="stSidebar"] .block-container {
    padding-top: 0.5rem !important;
}
/* Compact radio buttons */
.stRadio > div { gap: 0.1rem !important; }
/* Compact selectbox */
[data-testid="stSelectbox"] { margin-bottom: 0.2rem !important; }
</style>
"""


def load_inventory_data(silent: bool = False) -> Optional[pd.DataFrame]:
    """Load inventory data from XLSX file."""
    xlsx_path = Path(CONFIG["data"]["xlsx_file"])
    if not xlsx_path.exists():
        if not silent:
            st.error(f"❌ Inventory file not found: {xlsx_path}")
        logger.error(f"Inventory file not found: {xlsx_path}")
        return None

    try:
        df = pd.read_excel(xlsx_path, engine="openpyxl")
        required_columns = list(COLUMNS.values())

        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            if not silent:
                st.error(f"❌ Missing required columns: {missing_cols}")
            logger.error(f"Missing columns in data: {missing_cols}")
            return None

        df[COLUMNS["cantidad"]] = df[COLUMNS["cantidad"]].astype(int)
        df[COLUMNS["tenemos"]] = df[COLUMNS["tenemos"]].astype(int)
        df[COLUMNS["comprar"]] = (df[COLUMNS["cantidad"]] - df[COLUMNS["tenemos"]]).clip(lower=0)

        logger.info(f"✅ Loaded inventory data: {len(df)} items")
        return df

    except Exception as e:
        logger.error(f"Error loading data: {e}")
        if not silent:
            if _is_spreadsheet_lock_error(e):
                st.error(f"❌ Could not load inventory. {_SPREADSHEET_LOCKED_HINT}")
            else:
                st.error(f"❌ Error loading inventory data: {e}")
        return None


def save_inventory_data(df: pd.DataFrame) -> bool:
    """Save inventory data back to XLSX file."""
    try:
        df.to_excel(CONFIG["data"]["xlsx_file"], index=False, engine="openpyxl")
        logger.info("✅ Inventory data saved successfully")
        return True
    except Exception as e:
        logger.error(f"Error saving data: {e}")
        if _is_spreadsheet_lock_error(e):
            st.warning(f"Could not save. {_SPREADSHEET_LOCKED_HINT}")
        else:
            st.error(f"❌ Error saving inventory data: {e}")
        return False


def get_unique_zones(df: pd.DataFrame) -> List[str]:
    """Get unique zones from inventory data."""
    return sorted(df[COLUMNS["lugar"]].unique().tolist())


def get_unique_supermarkets(df: pd.DataFrame) -> List[str]:
    """Get unique supermarkets from inventory data."""
    return sorted(df[COLUMNS["super"]].unique().tolist())


def get_supermarket_stats(shopping_items: pd.DataFrame, bought_items: set) -> Dict[str, Dict[str, int]]:
    """Calculate statistics for each supermarket from shopping items."""
    stats = {}
    for supermarket in get_unique_supermarkets(shopping_items):
        sm_items = shopping_items[shopping_items[COLUMNS["super"]] == supermarket]
        bought_in_sm = sm_items[sm_items.index.isin(bought_items)]
        stats[supermarket] = {
            "total_unique": len(sm_items),
            "total_quantity": int(sm_items[COLUMNS["comprar"]].sum()),
            "got_it_unique": len(bought_in_sm),
            "got_it_quantity": int(bought_in_sm[COLUMNS["comprar"]].sum()),
        }
    return stats


def update_item_quantity(df: pd.DataFrame, item_index: int, delta: int) -> pd.DataFrame:
    """Update the tenemos quantity for an item and recalculate comprar."""
    old_tenemos = int(df.at[item_index, COLUMNS["tenemos"]])
    old_comprar = int(df.at[item_index, COLUMNS["comprar"]])
    new_qty = max(0, old_tenemos + delta)
    df.at[item_index, COLUMNS["tenemos"]] = new_qty
    df.at[item_index, COLUMNS["comprar"]] = max(0, df.at[item_index, COLUMNS["cantidad"]] - new_qty)
    if not save_inventory_data(df):
        df.at[item_index, COLUMNS["tenemos"]] = old_tenemos
        df.at[item_index, COLUMNS["comprar"]] = old_comprar
        return df
    logger.debug(f"Updated item {item_index}: tenemos={new_qty}")
    return df


def update_target_quantity(df: pd.DataFrame, item_index: int, delta: int) -> pd.DataFrame:
    """Update the cantidad (target) quantity for an item and recalculate comprar."""
    old_target = int(df.at[item_index, COLUMNS["cantidad"]])
    old_comprar = int(df.at[item_index, COLUMNS["comprar"]])
    new_target = max(0, old_target + delta)
    df.at[item_index, COLUMNS["cantidad"]] = new_target
    df.at[item_index, COLUMNS["comprar"]] = max(0, new_target - df.at[item_index, COLUMNS["tenemos"]])
    if not save_inventory_data(df):
        df.at[item_index, COLUMNS["cantidad"]] = old_target
        df.at[item_index, COLUMNS["comprar"]] = old_comprar
        return df
    logger.debug(f"Updated target for item {item_index}: cantidad={new_target}")
    return df


def _qty_html(current: int, target: int) -> str:
    """Compact inline HTML for current/target display with color coding."""
    color = "#21c354" if current >= target else ("#ffa500" if current > 0 else "#ff4b4b")
    return (
        f"<div style='text-align:center;padding-top:6px;font-size:0.93rem'>"
        f"<span style='color:{color};font-weight:600'>{current}</span>"
        f"<span style='color:#666'>/{target}</span></div>"
    )


def _buy_html(qty: int) -> str:
    """Compact inline HTML for buy quantity display."""
    if qty > 0:
        return (
            f"<div style='text-align:center;padding-top:6px;font-size:0.93rem;"
            f"color:#ff4b4b;font-weight:600'>↓{qty}</div>"
        )
    return "<div style='text-align:center;padding-top:6px;font-size:0.93rem;color:#21c354'>✓</div>"


def render_audit_mode(df: pd.DataFrame) -> pd.DataFrame:
    """Render the audit mode interface."""
    zones = get_unique_zones(df)
    selected_zone = st.selectbox("Zone", zones, label_visibility="collapsed")

    zone_data = df[
        (df[COLUMNS["lugar"]] == selected_zone) & (df[COLUMNS["cantidad"]] > 0)
    ].copy().sort_values(COLUMNS["comida"], key=lambda s: s.str.lower())

    if zone_data.empty:
        st.info(f"No tracked items in {selected_zone}")
        return df

    st.caption(f"{selected_zone.title()} · {len(zone_data)} items")

    # Header row — cols [4,1,1,2,1,1,2] = 12 parts, merged as [4,2,2,2,2]
    _, hv, hd, ht, hb = st.columns([4, 2, 2, 2, 2])
    with hv:
        st.markdown("<div style='text-align:center;font-size:0.72rem;color:#888;padding-bottom:0'>➖ have ➕</div>", unsafe_allow_html=True)
    with hd:
        st.markdown("<div style='text-align:center;font-size:0.72rem;color:#888'>have/tgt</div>", unsafe_allow_html=True)
    with ht:
        st.markdown("<div style='text-align:center;font-size:0.72rem;color:#888;padding-bottom:0'>⊖ target ⊕</div>", unsafe_allow_html=True)
    with hb:
        st.markdown("<div style='text-align:center;font-size:0.72rem;color:#888'>buy</div>", unsafe_allow_html=True)

    for idx in zone_data.index:
        item_name = zone_data.at[idx, COLUMNS["comida"]]
        current_qty = zone_data.at[idx, COLUMNS["tenemos"]]
        target_qty = zone_data.at[idx, COLUMNS["cantidad"]]
        buy_qty = zone_data.at[idx, COLUMNS["comprar"]]

        col1, col2, col3, col4, col5, col6, col7 = st.columns([4, 1, 1, 2, 1, 1, 2])
        with col1:
            st.markdown(f"**{item_name}**")
        with col2:
            if st.button("➖", key=f"minus_{idx}", help="Decrease stock"):
                df = update_item_quantity(df, idx, -1)
                st.rerun()
        with col3:
            if st.button("➕", key=f"plus_{idx}", help="Increase stock"):
                df = update_item_quantity(df, idx, 1)
                st.rerun()
        with col4:
            st.markdown(_qty_html(current_qty, target_qty), unsafe_allow_html=True)
        with col5:
            if st.button("⊖", key=f"target_minus_{idx}", help="Decrease target"):
                df = update_target_quantity(df, idx, -1)
                st.rerun()
        with col6:
            if st.button("⊕", key=f"target_plus_{idx}", help="Increase target"):
                df = update_target_quantity(df, idx, 1)
                st.rerun()
        with col7:
            st.markdown(_buy_html(buy_qty), unsafe_allow_html=True)

    return df


def render_edit_mode(df: pd.DataFrame) -> pd.DataFrame:
    """Render the edit mode interface for changing target quantities."""
    zones = get_unique_zones(df)
    selected_zone = st.selectbox("Zone", zones, label_visibility="collapsed")

    zone_data = df[df[COLUMNS["lugar"]] == selected_zone].copy().sort_values(COLUMNS["comida"], key=lambda s: s.str.lower())

    if zone_data.empty:
        st.info(f"No items found in {selected_zone}")
        return df

    st.caption(f"{selected_zone.title()} · {len(zone_data)} items · have / target · buy")

    for idx in zone_data.index:
        item_name = zone_data.at[idx, COLUMNS["comida"]]
        current_qty = zone_data.at[idx, COLUMNS["tenemos"]]
        target_qty = zone_data.at[idx, COLUMNS["cantidad"]]
        buy_qty = zone_data.at[idx, COLUMNS["comprar"]]

        col1, col2, col3, col4, col5 = st.columns([4, 1, 2, 1, 2])
        with col1:
            st.markdown(f"**{item_name}**")
        with col2:
            if st.button("➖", key=f"target_minus_{idx}"):
                df = update_target_quantity(df, idx, -1)
                st.rerun()
        with col3:
            st.markdown(_qty_html(current_qty, target_qty), unsafe_allow_html=True)
        with col4:
            if st.button("➕", key=f"target_plus_{idx}"):
                df = update_target_quantity(df, idx, 1)
                st.rerun()
        with col5:
            st.markdown(_buy_html(buy_qty), unsafe_allow_html=True)

    return df


def render_shopping_mode(df: pd.DataFrame) -> None:
    """Render the shopping list mode interface."""
    if "bought_items" not in st.session_state:
        st.session_state.bought_items = set()
    if "extra_shopping_items" not in st.session_state:
        st.session_state.extra_shopping_items = {}
    if "extra_bought_items" not in st.session_state:
        st.session_state.extra_bought_items = {}
    if "extra_item_counter" not in st.session_state:
        st.session_state.extra_item_counter = 0

    shopping_items = df[df[COLUMNS["comprar"]] > 0].copy()

    base_supermarkets = get_unique_supermarkets(shopping_items) if not shopping_items.empty else []
    all_supermarkets = sorted(set(base_supermarkets) | set(st.session_state.extra_shopping_items.keys()))

    if shopping_items.empty and not all_supermarkets:
        st.success("🎉 All stocked up — nothing to buy!")
        return

    # Compute totals including extra items
    total_items = len(shopping_items)
    total_qty = int(shopping_items[COLUMNS["comprar"]].sum()) if not shopping_items.empty else 0
    bought_count = len([i for i in shopping_items.index if i in st.session_state.bought_items])
    bought_qty = int(
        shopping_items[shopping_items.index.isin(st.session_state.bought_items)][COLUMNS["comprar"]].sum()
    ) if not shopping_items.empty else 0

    for sm, extras in st.session_state.extra_shopping_items.items():
        total_items += len(extras)
        total_qty += sum(e["qty"] for e in extras)
        extra_bought = st.session_state.extra_bought_items.get(sm, set())
        for e in extras:
            if e["id"] in extra_bought:
                bought_count += 1
                bought_qty += e["qty"]

    # Compact one-line summary with inline clear button
    c1, c2 = st.columns([6, 1])
    with c1:
        progress = f" · ✅ {bought_count}/{total_items} unique · {bought_qty}/{total_qty} units" if bought_count > 0 else ""
        st.caption(f"🛒 {total_items} unique · {total_qty} units · {len(all_supermarkets)} store(s){progress}")
    with c2:
        if bought_count > 0 and st.button("🗑️", help="Unmark all"):
            st.session_state.bought_items.clear()
            st.session_state.extra_bought_items.clear()
            st.rerun()

    supermarket_stats = get_supermarket_stats(shopping_items, st.session_state.bought_items) if not shopping_items.empty else {}

    for supermarket in all_supermarkets:
        stats = supermarket_stats.get(supermarket, {"total_unique": 0, "total_quantity": 0, "got_it_unique": 0, "got_it_quantity": 0})
        sm_items = (
            shopping_items[shopping_items[COLUMNS["super"]] == supermarket]
            .sort_values(COLUMNS["comida"], key=lambda s: s.str.lower())
            if not shopping_items.empty else pd.DataFrame()
        )
        extras = st.session_state.extra_shopping_items.get(supermarket, [])
        extra_bought_set = st.session_state.extra_bought_items.get(supermarket, set())

        total_u = stats["total_unique"] + len(extras)
        total_q = stats["total_quantity"] + sum(e["qty"] for e in extras)
        done_u = stats["got_it_unique"] + len([e for e in extras if e["id"] in extra_bought_set])
        done_q = stats["got_it_quantity"] + sum(e["qty"] for e in extras if e["id"] in extra_bought_set)

        done_txt = f" · ✅ {done_u}/{total_u}" if done_u > 0 else ""
        label = f"🏪 {supermarket.title()} — {total_u} items · {total_q} units{done_txt}"

        with st.expander(label, expanded=True):
            # Inventory items
            for idx in sm_items.index:
                item_name = sm_items.at[idx, COLUMNS["comida"]]
                qty_to_buy = sm_items.at[idx, COLUMNS["comprar"]]
                buy_url = sm_items.at[idx, COLUMNS["buscador"]]
                is_bought = idx in st.session_state.bought_items

                col1, col2, col3 = st.columns([5, 2, 2])

                with col1:
                    if is_bought:
                        st.markdown(f"~~{item_name}~~ · {qty_to_buy}×")
                    else:
                        st.markdown(f"**{item_name}** · {qty_to_buy}×")

                with col2:
                    st.link_button(
                        "🔄 Again" if is_bought else "🛒 Buy",
                        buy_url,
                        use_container_width=True,
                    )

                with col3:
                    if is_bought:
                        if st.button("↩️ Undo", key=f"unmark_{idx}", use_container_width=True):
                            st.session_state.bought_items.remove(idx)
                            st.rerun()
                    else:
                        if st.button(
                            "✅ Got it",
                            key=f"mark_{idx}",
                            use_container_width=True,
                            type="secondary",
                        ):
                            st.session_state.bought_items.add(idx)
                            st.rerun()

            # Extra (ad-hoc) items
            for e in extras:
                eid = e["id"]
                is_extra_bought = eid in extra_bought_set
                col1, col2, col3 = st.columns([5, 2, 2])
                with col1:
                    label_txt = f"~~{e['name']}~~ · {e['qty']}×" if is_extra_bought else f"**{e['name']}** · {e['qty']}×"
                    st.markdown(f"{label_txt} _+_")
                with col2:
                    if st.button("🗑️ Remove", key=f"extra_del_{eid}", use_container_width=True):
                        st.session_state.extra_shopping_items[supermarket] = [
                            x for x in extras if x["id"] != eid
                        ]
                        if supermarket in st.session_state.extra_bought_items:
                            st.session_state.extra_bought_items[supermarket].discard(eid)
                        if not st.session_state.extra_shopping_items[supermarket]:
                            del st.session_state.extra_shopping_items[supermarket]
                        st.rerun()
                with col3:
                    if is_extra_bought:
                        if st.button("↩️ Undo", key=f"extra_unmark_{eid}", use_container_width=True):
                            st.session_state.extra_bought_items[supermarket].discard(eid)
                            st.rerun()
                    else:
                        if st.button("✅ Got it", key=f"extra_mark_{eid}", use_container_width=True, type="secondary"):
                            if supermarket not in st.session_state.extra_bought_items:
                                st.session_state.extra_bought_items[supermarket] = set()
                            st.session_state.extra_bought_items[supermarket].add(eid)
                            st.rerun()

            # Quick-add form
            st.divider()
            with st.form(key=f"qa_form_{supermarket}", clear_on_submit=True):
                qa1, qa2, qa3 = st.columns([5, 1, 2])
                with qa1:
                    new_name = st.text_input("Item", placeholder="Item name…", label_visibility="collapsed")
                with qa2:
                    new_qty = st.number_input("Qty", value=1, min_value=1, step=1, label_visibility="collapsed")
                with qa3:
                    if st.form_submit_button("➕ Add", use_container_width=True):
                        if new_name.strip():
                            item_id = st.session_state.extra_item_counter
                            st.session_state.extra_item_counter += 1
                            if supermarket not in st.session_state.extra_shopping_items:
                                st.session_state.extra_shopping_items[supermarket] = []
                            st.session_state.extra_shopping_items[supermarket].append(
                                {"id": item_id, "name": new_name.strip(), "qty": int(new_qty)}
                            )
                            st.rerun()


def render_export_mode(df: pd.DataFrame) -> None:
    """Render the save/export mode interface."""
    col1, col2 = st.columns(2)

    with col1:
        if st.button("💾 Save to File", type="primary", use_container_width=True):
            if save_inventory_data(df):
                st.success("✅ Saved!")

    with col2:
        st.download_button(
            "📥 Download CSV",
            df.to_csv(index=False),
            "inventory_updated.csv",
            "text/csv",
            use_container_width=True,
        )

    st.subheader("📊 Summary")
    total_items = len(df)
    stocked_items = len(df[df[COLUMNS["comprar"]] == 0])
    shopping_items = len(df[df[COLUMNS["comprar"]] > 0])

    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Total", total_items)
    with c2:
        st.metric("Stocked", stocked_items)
    with c3:
        st.metric("Need Buy", shopping_items)


def render_edit_item_mode(df: pd.DataFrame) -> pd.DataFrame:
    """Render the edit item mode interface for searching and editing individual items."""
    search_term = st.text_input(
        "🔍 Search",
        placeholder="Type item name...",
    ).strip().lower()

    if search_term:
        filtered_df = df[df[COLUMNS["comida"]].str.lower().str.contains(search_term, na=False)].copy()
    else:
        filtered_df = df.copy()
    filtered_df = filtered_df.sort_values(COLUMNS["comida"], key=lambda s: s.str.lower())

    if filtered_df.empty:
        st.info(f"No items found matching '{search_term}'" if search_term else "No items to display.")
        return df

    st.caption(f"{len(filtered_df)} item(s)")

    for idx in filtered_df.index:
        item_name = filtered_df.at[idx, COLUMNS["comida"]]

        with st.expander(f"🔧 {item_name}", expanded=len(filtered_df) == 1):
            with st.form(key=f"edit_form_{idx}"):
                col1, col2 = st.columns(2)

                with col1:
                    current_super = filtered_df.at[idx, COLUMNS["super"]]
                    new_super = st.text_input(
                        "🏪 Supermarket",
                        value=current_super if pd.notna(current_super) else "",
                    )

                    current_lugar = filtered_df.at[idx, COLUMNS["lugar"]]
                    new_lugar = st.text_input(
                        "🏠 Zone",
                        value=current_lugar if pd.notna(current_lugar) else "",
                    )

                    current_comida = filtered_df.at[idx, COLUMNS["comida"]]
                    new_comida = st.text_input(
                        "🥘 Item Name",
                        value=current_comida if pd.notna(current_comida) else "",
                    )

                with col2:
                    current_cantidad = int(filtered_df.at[idx, COLUMNS["cantidad"]])
                    new_cantidad = st.number_input("🎯 Target", value=current_cantidad, min_value=0, step=1)

                    current_tenemos = int(filtered_df.at[idx, COLUMNS["tenemos"]])
                    new_tenemos = st.number_input("📦 Current", value=current_tenemos, min_value=0, step=1)

                    current_buscador = filtered_df.at[idx, COLUMNS["buscador"]]
                    new_buscador = st.text_input(
                        "🔗 URL",
                        value=current_buscador if pd.notna(current_buscador) else "",
                    )

                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    save_clicked = st.form_submit_button("💾 Save", type="primary", use_container_width=True)
                with col_btn2:
                    delete_clicked = st.form_submit_button("🗑️ Delete", type="secondary", use_container_width=True)

                if save_clicked:
                    snap = df.loc[idx].copy()
                    df.at[idx, COLUMNS["super"]] = new_super
                    df.at[idx, COLUMNS["lugar"]] = new_lugar
                    df.at[idx, COLUMNS["comida"]] = new_comida
                    df.at[idx, COLUMNS["cantidad"]] = new_cantidad
                    df.at[idx, COLUMNS["tenemos"]] = new_tenemos
                    df.at[idx, COLUMNS["buscador"]] = new_buscador
                    df.at[idx, COLUMNS["comprar"]] = max(0, new_cantidad - new_tenemos)

                    if save_inventory_data(df):
                        st.success(f"✅ Saved '{new_comida}'")
                        st.rerun()
                    else:
                        df.loc[idx] = snap

                if delete_clicked:
                    backup = df.copy()
                    df = df.drop(idx)
                    if save_inventory_data(df):
                        st.session_state.inventory_data = df
                        st.success(f"✅ Deleted '{item_name}'")
                        st.rerun()
                    else:
                        return backup

    return df


def render_add_item_mode(df: pd.DataFrame) -> pd.DataFrame:
    """Render the add item mode interface for creating new inventory items."""
    existing_supermarkets = get_unique_supermarkets(df)
    existing_zones = get_unique_zones(df)

    with st.form(key="add_item_form"):
        col1, col2 = st.columns(2)

        with col1:
            new_super = st.selectbox("🏪 Supermarket", options=existing_supermarkets)
            new_lugar = st.selectbox("🏠 Zone", options=existing_zones)
            new_comida = st.text_input("🥘 Item Name")

        with col2:
            new_cantidad = st.number_input("🎯 Target", value=0, min_value=0, step=1)
            new_tenemos = st.number_input("📦 Current", value=0, min_value=0, step=1)
            new_buscador = st.text_input("🔗 URL")

        if st.form_submit_button("➕ Add Item", type="primary", use_container_width=True):
            if not new_comida.strip():
                st.error("❌ Item name is required!")
            else:
                new_row = {
                    COLUMNS["super"]: new_super,
                    COLUMNS["lugar"]: new_lugar,
                    COLUMNS["comida"]: new_comida,
                    COLUMNS["cantidad"]: new_cantidad,
                    COLUMNS["tenemos"]: new_tenemos,
                    COLUMNS["buscador"]: new_buscador,
                    COLUMNS["comprar"]: max(0, new_cantidad - new_tenemos),
                }
                df_extended = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

                if save_inventory_data(df_extended):
                    st.session_state.inventory_data = df_extended
                    st.success(f"✅ Added '{new_comida}'")
                    st.rerun()

    return df


def main():
    """Main application entry point."""
    st.set_page_config(**CONFIG["ui"]["page_config"])
    st.markdown(_CSS, unsafe_allow_html=True)

    # Compact header
    st.markdown("### 🛒 Inventory & Shopping Helper")

    # Initialize session state
    if "inventory_data" not in st.session_state:
        st.session_state.inventory_data = load_inventory_data()
        if st.session_state.inventory_data is None:
            st.stop()

    if "current_mode" not in st.session_state:
        st.session_state.current_mode = "audit"

    if "bought_items" not in st.session_state:
        st.session_state.bought_items = set()

    if "extra_shopping_items" not in st.session_state:
        st.session_state.extra_shopping_items = {}  # {sm: [{"id": int, "name": str, "qty": int}]}

    if "extra_bought_items" not in st.session_state:
        st.session_state.extra_bought_items = {}  # {sm: set of item ids}

    if "extra_item_counter" not in st.session_state:
        st.session_state.extra_item_counter = 0

    # Sidebar: navigation + compact stats
    with st.sidebar:
        if st.button(
            "📂 Open spreadsheet",
            help="Opens the Excel file in the default app (e.g. Excel). Useful when OneDrive has not refreshed yet.",
            use_container_width=True,
        ):
            open_inventory_spreadsheet()

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

        # Compact stats
        if st.session_state.inventory_data is not None:
            df_stats = st.session_state.inventory_data
            sm_shopping = df_stats[df_stats[COLUMNS["comprar"]] > 0].copy()
            shopping_needed = len(sm_shopping)
            units_to_buy = int(sm_shopping[COLUMNS["comprar"]].sum()) if not sm_shopping.empty else 0

            st.divider()
            st.caption(f"**{len(df_stats)}** items · **{shopping_needed}** unique / **{units_to_buy}** units to buy")

            if not sm_shopping.empty:
                for sm, stats in get_supermarket_stats(sm_shopping, st.session_state.bought_items).items():
                    offset_items = st.session_state.get(f"cart_offset_items_{sm}", 0)
                    offset_units = st.session_state.get(f"cart_offset_units_{sm}", 0)
                    done_u = stats["got_it_unique"] + offset_items
                    total_u = stats["total_unique"]
                    done_q = stats["got_it_quantity"] + offset_units
                    total_q = stats["total_quantity"]
                    bar = "▓" * min(done_u, total_u) + "░" * max(0, total_u - done_u)
                    st.caption(f"**{sm.title()}** {bar} {done_u}/{total_u} · {done_q}/{total_q} units")
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

    # Main content
    df = st.session_state.inventory_data

    if st.session_state.current_mode == "audit":
        st.session_state.inventory_data = render_audit_mode(df)
    elif st.session_state.current_mode == "edit":
        st.session_state.inventory_data = render_edit_mode(df)
    elif st.session_state.current_mode == "edit_item":
        st.session_state.inventory_data = render_edit_item_mode(df)
    elif st.session_state.current_mode == "add_item":
        st.session_state.inventory_data = render_add_item_mode(df)
    elif st.session_state.current_mode == "shopping":
        render_shopping_mode(df)
    elif st.session_state.current_mode == "export":
        render_export_mode(df)


if __name__ == "__main__":
    main()
