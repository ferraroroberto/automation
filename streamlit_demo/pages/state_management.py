"""
Demo: State Management

Demonstrates how st.session_state works:
- Persisting values across reruns
- Sharing state between different pages / modules
- Callbacks and state mutation
- Inspecting the full session state
"""

import streamlit as st


def render():
    st.header("🧠 State Management")
    st.markdown(
        "Streamlit reruns the entire script on every interaction. "
        "`st.session_state` is a dictionary that persists across reruns, "
        "letting you maintain counters, caches, and shared data."
    )

    counter_tab, shared_tab, inspector_tab = st.tabs(
        ["Counter Demo", "Cross-Page State", "State Inspector"]
    )

    # ── Counter demo ─────────────────────────────────────────────────────
    with counter_tab:
        st.subheader("Counter with Session State")
        st.markdown(
            "Without session state, a counter would reset to 0 on every click. "
            "With `st.session_state`, the value persists."
        )

        if "demo_counter" not in st.session_state:
            st.session_state.demo_counter = 0

        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("➕ Increment"):
                st.session_state.demo_counter += 1
        with col2:
            if st.button("➖ Decrement"):
                st.session_state.demo_counter -= 1
        with col3:
            if st.button("🔄 Reset Counter"):
                st.session_state.demo_counter = 0

        st.metric("Current Count", st.session_state.demo_counter)

        st.markdown("---")

        # Callback pattern
        st.subheader("Callback Pattern")
        st.markdown(
            "You can attach callbacks to widgets via `on_change` / `on_click`. "
            "The callback runs **before** the script reruns."
        )

        def _on_slider_change():
            st.session_state["slider_change_count"] = (
                st.session_state.get("slider_change_count", 0) + 1
            )

        if "callback_slider_val" not in st.session_state:
            st.session_state.callback_slider_val = 50

        st.slider(
            "Move me (callback tracks changes)",
            min_value=0,
            max_value=100,
            key="callback_slider_val",
            on_change=_on_slider_change,
        )
        change_count = st.session_state.get("slider_change_count", 0)
        st.write(f"Slider value: **{st.session_state.callback_slider_val}**")
        st.write(f"Times changed: **{change_count}**")

    # ── Cross-page state ─────────────────────────────────────────────────
    with shared_tab:
        st.subheader("Cross-Page Shared State")
        st.markdown(
            "Any key you set in `st.session_state` is accessible from **every page**. "
            "This is how you share data between modules."
        )

        st.text_input(
            "Set a global message (visible on all pages)",
            key="global_message",
            placeholder="Type something here…",
        )

        if st.session_state.get("global_message"):
            st.info(f"Global message: **{st.session_state.global_message}**")

        st.markdown("---")

        st.markdown("**Shared data from other demos:**")

        # Show data from CRUD demo if available
        if "crud_employees" in st.session_state:
            crud_df = st.session_state["crud_employees"]
            st.write(f"CRUD demo has **{len(crud_df)}** employee records loaded.")
        else:
            st.caption("CRUD demo data not loaded yet — visit the CRUD page first.")

        # Show process runner status if available
        if "process_runner_status" in st.session_state:
            status = st.session_state["process_runner_status"]
            st.write(f"Process Runner last status: **{status}**")
        else:
            st.caption(
                "Process Runner has not been executed yet — visit the Process Runner page."
            )

    # ── State inspector ──────────────────────────────────────────────────
    with inspector_tab:
        st.subheader("Session State Inspector")
        st.markdown(
            "This panel exposes the raw contents of `st.session_state`. "
            "Useful for debugging."
        )

        state_items = dict(st.session_state)
        if not state_items:
            st.info("Session state is empty.")
        else:
            st.write(f"**{len(state_items)}** keys in session state:")

            for key in sorted(state_items.keys()):
                value = state_items[key]
                value_type = type(value).__name__
                value_preview = str(value)[:300]
                with st.expander(f"`{key}` ({value_type})"):
                    st.code(value_preview, language="python")
