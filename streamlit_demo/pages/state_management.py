"""
Demo – State Management
=======================
Demonstrates how ``st.session_state`` works and how to share state
across different parts of the application:
- Reading and writing session state values
- A persistent counter that survives reruns
- A shared "notepad" that other modules can read
- Inspecting the full session state dictionary

**Why this matters:** Streamlit reruns the entire script on every
interaction.  Understanding session state is essential for building
apps that remember user actions between reruns.
"""

import streamlit as st


def render() -> None:
    st.header("State Management Demo")
    st.markdown(
        "Streamlit reruns the script from top to bottom on every interaction. "
        "`st.session_state` lets you persist data across those reruns."
    )

    tab_counter, tab_notepad, tab_inspector = st.tabs(
        ["Counter", "Shared Notepad", "State Inspector"]
    )

    # ==================== Counter ====================
    with tab_counter:
        st.subheader("Persistent Counter")
        st.markdown(
            "This counter survives reruns because its value lives in session state."
        )

        # Initialise the counter if it doesn't exist
        if "demo_counter" not in st.session_state:
            st.session_state["demo_counter"] = 0

        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("Increment (+1)", key="state_inc"):
                st.session_state["demo_counter"] += 1
        with col2:
            if st.button("Decrement (-1)", key="state_dec"):
                st.session_state["demo_counter"] -= 1
        with col3:
            if st.button("Reset", key="state_reset"):
                st.session_state["demo_counter"] = 0

        st.metric("Counter Value", st.session_state["demo_counter"])

        with st.expander("Show code"):
            st.code(
                '''
# Initialise once
if "demo_counter" not in st.session_state:
    st.session_state["demo_counter"] = 0

# Mutate on button press
if st.button("Increment"):
    st.session_state["demo_counter"] += 1

st.metric("Counter", st.session_state["demo_counter"])
''',
                language="python",
            )

    # ==================== Shared Notepad ====================
    with tab_notepad:
        st.subheader("Shared Notepad")
        st.markdown(
            "This notepad stores text in session state under the key "
            "`shared_notepad`.  Any module can read this value by accessing "
            "`st.session_state['shared_notepad']`."
        )

        # Initialise
        if "shared_notepad" not in st.session_state:
            st.session_state["shared_notepad"] = ""

        new_text = st.text_area(
            "Write something",
            value=st.session_state["shared_notepad"],
            height=150,
            key="notepad_input",
        )

        if st.button("Save to session state", key="notepad_save"):
            st.session_state["shared_notepad"] = new_text
            st.success("Saved! This value is now readable from any page.")

        st.markdown("---")
        st.markdown("**Current stored value:**")
        st.info(st.session_state.get("shared_notepad", "(empty)") or "(empty)")

    # ==================== State Inspector ====================
    with tab_inspector:
        st.subheader("Session State Inspector")
        st.markdown(
            "Below is the **full contents** of `st.session_state`.  "
            "This is useful for debugging."
        )

        # Filter out internal Streamlit keys (they start with underscore or
        # contain widget IDs that are not human-readable)
        state_dict = {}
        for key, value in sorted(st.session_state.items()):
            # Attempt to make values JSON-serialisable
            try:
                import json
                json.dumps(value)
                state_dict[key] = value
            except (TypeError, ValueError):
                state_dict[key] = str(value)

        st.json(state_dict)


if __name__ == "__main__":
    render()
