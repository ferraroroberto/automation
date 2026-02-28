"""
Demo – CRUD Operations
======================
Shows how to create, read, update, and delete records from an in-memory
dataset, persisting changes across reruns via ``st.session_state``.

Key patterns demonstrated:
- Initialising session state with a DataFrame
- Adding rows through a form
- Inline editing with ``st.data_editor``
- Deleting selected rows
- Resetting to the original dataset

**Why this matters:** Almost every internal tool revolves around a table of
records that users need to manage.
"""

import streamlit as st
import pandas as pd

from data.loader import load_employees

# Session-state key that holds the working copy of the dataset
_STATE_KEY = "crud_employees"


def _init_state() -> None:
    """Load the employees dataset into session state if not present."""
    if _STATE_KEY not in st.session_state:
        st.session_state[_STATE_KEY] = load_employees()


def _next_id() -> int:
    """Return the next available employee ID."""
    df: pd.DataFrame = st.session_state[_STATE_KEY]
    return int(df["id"].max()) + 1 if not df.empty else 1


def render() -> None:
    _init_state()
    st.header("CRUD Operations Demo")
    st.markdown(
        "This page manages an **in-memory employee table**.  Changes persist "
        "for the duration of your session."
    )

    tab_read, tab_create, tab_update, tab_delete = st.tabs(
        ["Read", "Create", "Update", "Delete"]
    )

    df: pd.DataFrame = st.session_state[_STATE_KEY]

    # ==================== READ ====================
    with tab_read:
        st.subheader(f"Employee Table ({len(df)} records)")
        st.dataframe(df, use_container_width=True, hide_index=True)

    # ==================== CREATE ====================
    with tab_create:
        st.subheader("Add a New Employee")
        with st.form("crud_add_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                first = st.text_input("First name")
                last = st.text_input("Last name")
                email = st.text_input("Email")
            with col2:
                dept = st.selectbox(
                    "Department",
                    ["Engineering", "Marketing", "Sales", "HR", "Finance", "Support"],
                    key="crud_dept",
                )
                salary = st.number_input("Salary", min_value=0.0, value=50000.0, step=1000.0)
                hire_date = st.date_input("Hire date", key="crud_hire")

            submitted = st.form_submit_button("Add Employee")

        if submitted:
            if not first or not last:
                st.warning("First name and last name are required.")
            else:
                new_row = {
                    "id": _next_id(),
                    "first_name": first,
                    "last_name": last,
                    "email": email,
                    "department": dept,
                    "salary": salary,
                    "hire_date": str(hire_date),
                    "active": True,
                }
                st.session_state[_STATE_KEY] = pd.concat(
                    [df, pd.DataFrame([new_row])], ignore_index=True
                )
                st.success(f"Employee **{first} {last}** added (ID {new_row['id']}).")
                st.rerun()

    # ==================== UPDATE ====================
    with tab_update:
        st.subheader("Edit Employees Inline")
        st.markdown(
            "Double-click a cell to edit.  Changes are saved automatically "
            "when you click outside the cell."
        )
        edited = st.data_editor(
            df,
            num_rows="fixed",
            use_container_width=True,
            hide_index=True,
            key="crud_editor",
        )
        if st.button("Save Changes", key="crud_save"):
            st.session_state[_STATE_KEY] = edited
            st.success("Changes saved.")
            st.rerun()

    # ==================== DELETE ====================
    with tab_delete:
        st.subheader("Delete Employees")
        if df.empty:
            st.info("No records to delete.")
        else:
            ids_to_delete = st.multiselect(
                "Select employee IDs to delete",
                df["id"].tolist(),
                key="crud_delete_ids",
            )
            if st.button("Delete Selected", type="primary", key="crud_delete_btn"):
                if not ids_to_delete:
                    st.warning("Select at least one ID.")
                else:
                    st.session_state[_STATE_KEY] = df[~df["id"].isin(ids_to_delete)]
                    st.success(f"Deleted {len(ids_to_delete)} record(s).")
                    st.rerun()

    # ==================== Reset ====================
    st.markdown("---")
    if st.button("Reset to Original Data", key="crud_reset"):
        st.session_state[_STATE_KEY] = load_employees()
        st.info("Dataset reset to original mock data.")
        st.rerun()


if __name__ == "__main__":
    render()
