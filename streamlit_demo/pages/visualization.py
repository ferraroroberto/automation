"""
Demo – Visualization
====================
Demonstrates Streamlit's built-in data display and charting capabilities:
- st.dataframe / st.data_editor (editable tables)
- st.metric for KPI cards
- st.line_chart, st.bar_chart, st.scatter_chart
- Filtering controls that dynamically update the charts

**Why this matters:** Dashboards and data exploration tools are the most
common use case for internal Streamlit apps.
"""

import streamlit as st
import pandas as pd

from data.loader import load_employees, load_sales


def render() -> None:
    st.header("Data Visualization Demo")

    tab_tables, tab_charts, tab_kpis = st.tabs(
        ["Tables & DataFrames", "Charts", "KPI Metrics"]
    )

    # Load data once per render (cached by Streamlit's data loader)
    sales = load_sales()
    employees = load_employees()

    # ==================== Tables ====================
    with tab_tables:
        st.subheader("Read-Only Table")
        st.markdown("Use `st.dataframe` for a scrollable, sortable table.")
        st.dataframe(employees.head(20), use_container_width=True)

        st.markdown("---")
        st.subheader("Editable DataFrame")
        st.markdown(
            "`st.data_editor` lets users edit cells in-place.  "
            "Changes are returned as a new DataFrame."
        )
        edited = st.data_editor(
            employees.head(10),
            num_rows="dynamic",
            use_container_width=True,
            key="viz_editor",
        )
        with st.expander("View edited data as JSON"):
            st.json(edited.to_dict(orient="records"))

    # ==================== Charts ====================
    with tab_charts:
        st.subheader("Interactive Charts")

        # --- Filters ---
        products = ["All"] + sorted(sales["product"].unique().tolist())
        selected_product = st.selectbox("Filter by product", products, key="viz_product")

        filtered = sales if selected_product == "All" else sales[sales["product"] == selected_product]

        # Aggregate by date for time-series
        daily = (
            filtered.assign(date=pd.to_datetime(filtered["date"]))
            .groupby("date")
            .agg(total_revenue=("total", "sum"), units_sold=("quantity", "sum"))
            .sort_index()
        )

        col1, col2 = st.columns(2)
        with col1:
            st.markdown("#### Revenue Over Time (Line)")
            st.line_chart(daily["total_revenue"])
        with col2:
            st.markdown("#### Units Sold Over Time (Bar)")
            st.bar_chart(daily["units_sold"])

        st.markdown("#### Revenue vs. Quantity (Scatter)")
        st.scatter_chart(
            filtered,
            x="quantity",
            y="total",
            color="product",
            size="unit_price",
        )

    # ==================== KPI Metrics ====================
    with tab_kpis:
        st.subheader("KPI Metric Cards")
        st.markdown(
            "`st.metric` displays a big number with an optional delta indicator — "
            "perfect for executive dashboards."
        )

        total_revenue = sales["total"].sum()
        avg_price = sales["unit_price"].mean()
        total_units = sales["quantity"].sum()
        active_employees = employees[employees["active"] == True].shape[0]  # noqa: E712

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Revenue", f"${total_revenue:,.0f}", delta="+12.3%")
        c2.metric("Avg Unit Price", f"${avg_price:,.2f}", delta="-2.1%")
        c3.metric("Total Units Sold", f"{total_units:,}")
        c4.metric("Active Employees", active_employees, delta="+3")


if __name__ == "__main__":
    render()
