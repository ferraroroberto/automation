"""
Demo: Data Input Widgets

Showcases every major input widget Streamlit provides:
- Text input, text area, number input
- Slider, select slider
- Selectbox, multiselect
- Date / time pickers
- Checkbox, radio, toggle
- Color picker
- Forms with submit buttons
"""

import streamlit as st


def render():
    st.header("📝 Data Input Widgets")
    st.markdown(
        "This page demonstrates the breadth of input widgets available in Streamlit. "
        "Each section is self-contained; change any value and see the result instantly."
    )

    # ── Text inputs ──────────────────────────────────────────────────────
    with st.expander("Text Inputs", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("Your name", placeholder="e.g. Jane Doe")
            bio = st.text_area(
                "Short bio",
                height=100,
                placeholder="Tell us about yourself…",
            )
        with col2:
            password = st.text_input("Password (masked)", type="password")
            st.markdown("---")
            st.markdown("**Live preview:**")
            if name:
                st.write(f"Hello, **{name}**!")
            if bio:
                st.info(bio)

    # ── Numeric inputs ───────────────────────────────────────────────────
    with st.expander("Numeric Inputs"):
        col1, col2 = st.columns(2)
        with col1:
            age = st.number_input("Age", min_value=0, max_value=120, value=25, step=1)
            temperature = st.slider(
                "Temperature (°C)", min_value=-20.0, max_value=50.0, value=22.0, step=0.5
            )
        with col2:
            price_range = st.slider(
                "Price range ($)",
                min_value=0,
                max_value=1000,
                value=(100, 500),
            )
            rating = st.select_slider(
                "Rating",
                options=["Poor", "Fair", "Good", "Very Good", "Excellent"],
                value="Good",
            )
        st.write(
            f"Age: **{age}** | Temp: **{temperature}°C** | "
            f"Price: **${price_range[0]}–${price_range[1]}** | Rating: **{rating}**"
        )

    # ── Selection widgets ────────────────────────────────────────────────
    with st.expander("Selection Widgets"):
        col1, col2 = st.columns(2)
        with col1:
            department = st.selectbox(
                "Department",
                ["Engineering", "Marketing", "Sales", "HR", "Finance"],
            )
            skills = st.multiselect(
                "Skills",
                ["Python", "SQL", "JavaScript", "Rust", "Go", "Java", "C++"],
                default=["Python", "SQL"],
            )
        with col2:
            color = st.color_picker("Favorite color", value="#3498db")
            notify = st.checkbox("Enable notifications", value=True)
            priority = st.radio(
                "Priority", ["Low", "Medium", "High"], horizontal=True
            )
        st.write(
            f"Dept: **{department}** | Skills: {skills} | "
            f"Color: {color} | Notify: {notify} | Priority: {priority}"
        )

    # ── Date / time pickers ──────────────────────────────────────────────
    with st.expander("Date & Time Pickers"):
        from datetime import date, time

        col1, col2 = st.columns(2)
        with col1:
            selected_date = st.date_input("Pick a date", value=date.today())
        with col2:
            selected_time = st.time_input("Pick a time", value=time(9, 30))
        st.write(f"Selected: **{selected_date}** at **{selected_time}**")

    # ── Form with submit button ──────────────────────────────────────────
    st.subheader("Form Demo")
    st.markdown(
        "Widgets inside a `st.form` do **not** trigger reruns until the form is submitted. "
        "This is critical for performance when you have many inputs."
    )

    with st.form("registration_form"):
        st.markdown("**New Employee Registration**")
        col1, col2 = st.columns(2)
        with col1:
            form_name = st.text_input("Full name")
            form_dept = st.selectbox(
                "Department",
                ["Engineering", "Marketing", "Sales", "HR", "Finance"],
                key="form_dept",
            )
        with col2:
            form_salary = st.number_input(
                "Starting salary ($)",
                min_value=30000,
                max_value=200000,
                value=60000,
                step=5000,
            )
            form_start = st.date_input("Start date", key="form_start")

        submitted = st.form_submit_button("Submit Registration")

        if submitted:
            if not form_name:
                st.error("Name is required.")
            else:
                st.success(
                    f"Registered **{form_name}** in {form_dept} "
                    f"starting {form_start} at ${form_salary:,.0f}/yr."
                )
