"""
Demo: CRUD Operations

Demonstrates a full Create-Read-Update-Delete workflow on an in-memory
dataset backed by st.session_state. Changes persist for the duration
of the Streamlit session.

Key concepts:
- Session-state persistence
- Dynamic form generation
- Conditional rendering based on user actions
"""

import csv
from io import StringIO
from pathlib import Path

import pandas as pd
import streamlit as st

MOCK_DATA_DIR = Path(__file__).resolve().parent.parent / "data" / "mock_data"
DEPARTMENTS = ["Engineering", "Marketing", "Sales", "HR", "Finance", "Operations"]
STATUSES = ["Active", "On Leave", "Resigned"]

_SESSION_KEY = "crud_employees"


def _load_initial_data() -> pd.DataFrame:
    filepath = MOCK_DATA_DIR / "employees.csv"
    return pd.read_csv(filepath)


def _get_data() -> pd.DataFrame:
    """Return the current dataset from session state, loading from disk on first call."""
    if _SESSION_KEY not in st.session_state:
        st.session_state[_SESSION_KEY] = _load_initial_data()
    return st.session_state[_SESSION_KEY]


def _set_data(df: pd.DataFrame):
    st.session_state[_SESSION_KEY] = df.reset_index(drop=True)


def render():
    st.header("🗃️ CRUD Operations")
    st.markdown(
        "Manage an employee dataset with full **Create / Read / Update / Delete** "
        "capabilities. All changes persist in `st.session_state` for the current session."
    )

    df = _get_data()

    # ── Stats bar ────────────────────────────────────────────────────────
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Records", len(df))
    col2.metric("Active", len(df[df["status"] == "Active"]))
    col3.metric("Departments", df["department"].nunique())

    st.markdown("---")

    crud_tab1, crud_tab2, crud_tab3, crud_tab4 = st.tabs(
        ["📋 Browse", "➕ Create", "✏️ Update", "🗑️ Delete"]
    )

    # ── READ ─────────────────────────────────────────────────────────────
    with crud_tab1:
        st.subheader("Browse Employees")
        filter_dept = st.multiselect(
            "Filter by department", DEPARTMENTS, default=[], key="crud_filter_dept"
        )
        filter_status = st.multiselect(
            "Filter by status", STATUSES, default=[], key="crud_filter_status"
        )

        filtered = df.copy()
        if filter_dept:
            filtered = filtered[filtered["department"].isin(filter_dept)]
        if filter_status:
            filtered = filtered[filtered["status"].isin(filter_status)]

        st.dataframe(filtered, use_container_width=True, hide_index=True)
        st.caption(f"Showing {len(filtered)} of {len(df)} records.")

    # ── CREATE ───────────────────────────────────────────────────────────
    with crud_tab2:
        st.subheader("Add New Employee")
        with st.form("crud_create_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                new_name = st.text_input("Name")
                new_dept = st.selectbox("Department", DEPARTMENTS, key="crud_new_dept")
            with col2:
                new_salary = st.number_input(
                    "Salary ($)", min_value=30000, max_value=300000,
                    value=60000, step=5000,
                )
                new_status = st.selectbox("Status", STATUSES, key="crud_new_status")

            if st.form_submit_button("Add Employee"):
                if not new_name.strip():
                    st.error("Name is required.")
                else:
                    new_id = int(df["id"].max()) + 1 if len(df) else 1
                    new_row = pd.DataFrame([{
                        "id": new_id,
                        "name": new_name.strip(),
                        "department": new_dept,
                        "salary": new_salary,
                        "hire_date": pd.Timestamp.now().strftime("%Y-%m-%d"),
                        "status": new_status,
                    }])
                    _set_data(pd.concat([df, new_row], ignore_index=True))
                    st.success(f"Added **{new_name}** (ID {new_id}).")
                    st.rerun()

    # ── UPDATE ───────────────────────────────────────────────────────────
    with crud_tab3:
        st.subheader("Update Employee")
        if df.empty:
            st.info("No records to update.")
        else:
            record_id = st.selectbox(
                "Select employee (by ID – Name)",
                df["id"].tolist(),
                format_func=lambda rid: f"{rid} – {df.loc[df['id'] == rid, 'name'].iloc[0]}",
                key="crud_update_id",
            )
            row = df[df["id"] == record_id].iloc[0]

            with st.form("crud_update_form"):
                col1, col2 = st.columns(2)
                with col1:
                    upd_name = st.text_input("Name", value=row["name"])
                    upd_dept = st.selectbox(
                        "Department", DEPARTMENTS,
                        index=DEPARTMENTS.index(row["department"]),
                        key="crud_upd_dept",
                    )
                with col2:
                    upd_salary = st.number_input(
                        "Salary ($)", min_value=30000, max_value=300000,
                        value=int(row["salary"]), step=5000,
                    )
                    upd_status = st.selectbox(
                        "Status", STATUSES,
                        index=STATUSES.index(row["status"]),
                        key="crud_upd_status",
                    )

                if st.form_submit_button("Save Changes"):
                    idx = df.index[df["id"] == record_id][0]
                    df.at[idx, "name"] = upd_name
                    df.at[idx, "department"] = upd_dept
                    df.at[idx, "salary"] = upd_salary
                    df.at[idx, "status"] = upd_status
                    _set_data(df)
                    st.success(f"Updated record **{record_id}**.")
                    st.rerun()

    # ── DELETE ───────────────────────────────────────────────────────────
    with crud_tab4:
        st.subheader("Delete Employee")
        if df.empty:
            st.info("No records to delete.")
        else:
            del_id = st.selectbox(
                "Select employee to delete",
                df["id"].tolist(),
                format_func=lambda rid: f"{rid} – {df.loc[df['id'] == rid, 'name'].iloc[0]}",
                key="crud_del_id",
            )
            del_row = df[df["id"] == del_id].iloc[0]
            st.warning(
                f"You are about to delete: **{del_row['name']}** "
                f"(Dept: {del_row['department']}, Status: {del_row['status']})"
            )

            if st.button("Confirm Delete", type="primary"):
                updated = df[df["id"] != del_id]
                _set_data(updated)
                st.success(f"Deleted record **{del_id}**.")
                st.rerun()
