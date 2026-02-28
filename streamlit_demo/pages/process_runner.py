"""
Demo – Process Runner
=====================
Simulates executing a long-running background task and streaming its
output to the UI in real time:
- A button triggers the process
- A progress bar tracks completion
- Live log lines are appended to a container
- A status indicator shows running / complete / error

**Why this matters:** Internal automation tools often wrap shell scripts or
ETL jobs.  This pattern shows how to keep the user informed while a task runs.
"""

import random
import time

import streamlit as st


def _simulate_task(steps: int, fail_chance: float) -> None:
    """
    Generator-style simulation that yields log lines.
    We write directly to Streamlit containers for a real-time feel.
    """
    log_container = st.empty()  # will hold a growing text block
    progress_bar = st.progress(0)
    status = st.status("Running task...", expanded=True)

    logs: list[str] = []

    for i in range(1, steps + 1):
        # Simulate variable work
        delay = random.uniform(0.1, 0.5)
        time.sleep(delay)

        # Decide if this step "fails"
        if random.random() < fail_chance:
            logs.append(f"[WARN] Step {i}/{steps} – transient warning, retrying...")
        else:
            logs.append(f"[INFO] Step {i}/{steps} – completed in {delay:.2f}s")

        # Update UI
        progress_bar.progress(i / steps)
        log_container.code("\n".join(logs), language="log")
        status.update(label=f"Running task... ({i}/{steps})", state="running")

    # Final status
    status.update(label="Task completed successfully", state="complete", expanded=False)
    logs.append(f"\n[DONE] All {steps} steps finished.")
    log_container.code("\n".join(logs), language="log")


def render() -> None:
    st.header("Process Runner Demo")
    st.markdown(
        "Simulate a long-running process with **live logs**, a **progress bar**, "
        "and a **status indicator**."
    )

    col1, col2 = st.columns(2)
    with col1:
        steps = st.slider("Number of steps", 5, 30, 10, key="proc_steps")
    with col2:
        fail_pct = st.slider(
            "Warning probability (%)", 0, 50, 10, key="proc_fail"
        )

    st.markdown("---")

    if st.button("Run Process", type="primary", key="proc_run"):
        _simulate_task(steps, fail_pct / 100)
        st.balloons()

    st.markdown("---")
    with st.expander("How does this work?"):
        st.markdown(
            """
            1. A `st.progress` bar is created and updated each iteration.
            2. A `st.empty` container is rewritten with the full log buffer
               on every step, giving the illusion of streaming output.
            3. `st.status` wraps the section with a collapsible status spinner
               that switches between *running* and *complete*.
            4. `time.sleep` simulates real work; in production you would call
               a subprocess or async task and poll for output.
            """
        )


if __name__ == "__main__":
    render()
