"""
Demo: Data Visualization

Showcases Streamlit's built-in charting and data display features:
- st.dataframe / st.data_editor (editable tables)
- st.metric (KPI cards)
- st.line_chart, st.bar_chart, st.scatter_chart
- st.table (static tables)
- Plotly integration for richer charts
"""

import json
from pathlib import Path

import pandas as pd
import streamlit as st

MOCK_DATA_DIR = Path(__file__).resolve().parent.parent / "data" / "mock_data"


def _load_timeseries() -> pd.DataFrame:
    filepath = MOCK_DATA_DIR / "timeseries.csv"
    return pd.read_csv(filepath, parse_dates=["date"])


def _load_sales() -> pd.DataFrame:
    filepath = MOCK_DATA_DIR / "sales.json"
    with open(filepath, encoding="utf-8") as f:
        data = json.load(f)
    return pd.DataFrame(data)


def render():
    st.header("📊 Data Visualization")
    st.markdown(
        "This page demonstrates tables, editable dataframes, KPI metrics, "
        "and multiple chart types — all powered by built-in Streamlit widgets."
    )

    ts_df = _load_timeseries()
    sales_df = _load_sales()

    # ── KPI Metrics ──────────────────────────────────────────────────────
    st.subheader("KPI Metrics")
    st.markdown(
        "`st.metric` renders a label, a large value, and an optional delta. "
        "Ideal for executive dashboards."
    )

    total_revenue = ts_df["revenue"].sum()
    total_visitors = ts_df["visitors"].sum()
    total_signups = ts_df["signups"].sum()
    avg_conversion = (total_signups / total_visitors * 100) if total_visitors else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Revenue", f"${total_revenue:,.0f}", delta="+12.3%")
    col2.metric("Total Visitors", f"{total_visitors:,}", delta="+5.1%")
    col3.metric("Signups", f"{total_signups:,}", delta="-2.0%", delta_color="inverse")
    col4.metric("Conversion Rate", f"{avg_conversion:.1f}%", delta="+0.8pp")

    st.markdown("---")

    # ── Charts ───────────────────────────────────────────────────────────
    st.subheader("Charts")

    chart_tab1, chart_tab2, chart_tab3 = st.tabs(
        ["📈 Line Chart", "📊 Bar Chart", "🔵 Scatter Chart"]
    )

    with chart_tab1:
        st.markdown("Daily **visitors** and **signups** over time.")
        chart_data = ts_df.set_index("date")[["visitors", "signups"]]
        st.line_chart(chart_data)

    with chart_tab2:
        st.markdown("**Revenue by category** (aggregated from sales data).")
        revenue_by_cat = (
            sales_df.groupby("category")["amount"].sum().reset_index()
        )
        revenue_by_cat.columns = ["Category", "Revenue"]
        st.bar_chart(revenue_by_cat, x="Category", y="Revenue")

    with chart_tab3:
        st.markdown("**Amount vs. Quantity** scatter plot from individual transactions.")
        st.scatter_chart(sales_df, x="amount", y="quantity", color="category")

    st.markdown("---")

    # ── Tables & editable dataframes ─────────────────────────────────────
    st.subheader("Tables & Editable DataFrames")

    table_tab1, table_tab2 = st.tabs(["Interactive DataFrame", "Editable DataFrame"])

    with table_tab1:
        st.markdown(
            "`st.dataframe` renders a scrollable, sortable, and searchable table."
        )
        st.dataframe(
            sales_df.head(50),
            use_container_width=True,
            column_config={
                "amount": st.column_config.NumberColumn("Amount ($)", format="$%.2f"),
                "date": st.column_config.DateColumn("Date"),
            },
        )

    with table_tab2:
        st.markdown(
            "`st.data_editor` lets users **edit cells** directly in the browser. "
            "Changes are returned as a new DataFrame."
        )
        sample = ts_df.head(10).copy()
        edited_df = st.data_editor(
            sample,
            num_rows="dynamic",
            use_container_width=True,
            key="ts_editor",
        )
        st.caption("Edited data (reflected in real-time):")
        st.write(edited_df)
