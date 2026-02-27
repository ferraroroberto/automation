"""
Demo: Process Runner

Demonstrates how to:
- Trigger a long-running task from a button
- Display a progress bar that updates in real-time
- Stream live log output to the UI
- Show status indicators (spinner, success, error)

The "process" is entirely simulated — no external tools are called.
"""

import random
import time

import streamlit as st

_LOG_KEY = "process_runner_logs"
_STATUS_KEY = "process_runner_status"

SIMULATED_STEPS = [
    ("Initializing environment", 0.4),
    ("Loading configuration", 0.3),
    ("Connecting to data source", 0.6),
    ("Validating schema", 0.5),
    ("Processing batch 1/4", 1.0),
    ("Processing batch 2/4", 1.2),
    ("Processing batch 3/4", 0.9),
    ("Processing batch 4/4", 1.1),
    ("Running post-processing checks", 0.7),
    ("Generating summary report", 0.5),
    ("Cleaning up temporary files", 0.3),
    ("Finalizing", 0.2),
]


def _reset_state():
    st.session_state[_LOG_KEY] = []
    st.session_state[_STATUS_KEY] = "idle"


def render():
    st.header("⚙️ Process Runner")
    st.markdown(
        "Simulate a long-running pipeline with **live progress**, "
        "**streaming logs**, and **status indicators**. "
        "This pattern is common in ETL dashboards and CI/CD monitors."
    )

    if _LOG_KEY not in st.session_state:
        _reset_state()

    st.markdown("---")

    # ── Controls ─────────────────────────────────────────────────────────
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        speed = st.select_slider(
            "Simulation speed",
            options=["Slow", "Normal", "Fast"],
            value="Fast",
        )
    speed_factor = {"Slow": 1.0, "Normal": 0.5, "Fast": 0.15}[speed]

    with col2:
        should_fail = st.checkbox("Simulate failure", value=False)
    with col3:
        fail_at = st.number_input(
            "Fail at step",
            min_value=1,
            max_value=len(SIMULATED_STEPS),
            value=7,
            disabled=not should_fail,
        )

    st.markdown("---")

    # ── Run button ───────────────────────────────────────────────────────
    run_col, reset_col = st.columns([1, 1])
    with run_col:
        run_clicked = st.button("▶️ Run Process", type="primary", use_container_width=True)
    with reset_col:
        if st.button("🔄 Reset", use_container_width=True):
            _reset_state()
            st.rerun()

    # ── Execution ────────────────────────────────────────────────────────
    if run_clicked:
        st.session_state[_LOG_KEY] = []
        st.session_state[_STATUS_KEY] = "running"

        progress_bar = st.progress(0, text="Starting…")
        log_container = st.empty()
        status_container = st.empty()

        logs: list[str] = []
        total = len(SIMULATED_STEPS)

        for i, (step_name, base_duration) in enumerate(SIMULATED_STEPS, start=1):
            # Check for simulated failure
            if should_fail and i == fail_at:
                logs.append(f"❌ [{i}/{total}] FAILED at: {step_name}")
                log_container.code("\n".join(logs), language="log")
                progress_bar.progress(i / total, text=f"Failed at step {i}")
                status_container.error(
                    f"Process failed at step {i}/{total}: {step_name}"
                )
                st.session_state[_STATUS_KEY] = "failed"
                st.session_state[_LOG_KEY] = logs
                return

            duration = base_duration * speed_factor * random.uniform(0.8, 1.2)
            time.sleep(duration)

            logs.append(f"✅ [{i}/{total}] {step_name} ({duration:.2f}s)")
            log_container.code("\n".join(logs), language="log")
            progress_bar.progress(i / total, text=f"Step {i}/{total}: {step_name}")

        st.session_state[_STATUS_KEY] = "success"
        st.session_state[_LOG_KEY] = logs
        progress_bar.progress(1.0, text="Complete!")
        status_container.success(f"Process completed successfully ({total} steps).")
        st.balloons()

    # ── Previous run logs ────────────────────────────────────────────────
    if st.session_state[_LOG_KEY] and not run_clicked:
        st.subheader("Previous Run Logs")
        status = st.session_state[_STATUS_KEY]
        if status == "success":
            st.success("Last run: Completed successfully.")
        elif status == "failed":
            st.error("Last run: Failed.")
        st.code("\n".join(st.session_state[_LOG_KEY]), language="log")
