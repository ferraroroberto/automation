"""
Demo – Data Input
=================
Showcases every common Streamlit input widget:
- Text input, text area, number input
- Slider, select box, multi-select
- Date / time pickers
- Checkbox, radio, toggle
- A form with a submit button that groups inputs into a single submission

**Why this matters:** Understanding how Streamlit widgets work (and how they
interact with session state) is the foundation of every app.
"""

import streamlit as st


def render() -> None:
    st.header("Data Input Demo")
    st.markdown(
        "Explore every major input widget Streamlit offers.  "
        "Change values and watch the *Current values* section update in real time."
    )

    # ------------------------------------------------------------------
    # Tab layout – one tab per widget category
    # ------------------------------------------------------------------
    tab_basic, tab_selection, tab_datetime, tab_form = st.tabs(
        ["Basic Inputs", "Selection Widgets", "Date & Time", "Form Submit"]
    )

    # ==================== Basic Inputs ====================
    with tab_basic:
        st.subheader("Text & Number Inputs")

        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("Your name", placeholder="e.g. Alice")
            bio = st.text_area("Short bio", height=100, placeholder="Tell us about yourself...")
        with col2:
            age = st.number_input("Age", min_value=0, max_value=150, value=30, step=1)
            rating = st.slider("Satisfaction (1-10)", 1, 10, 7)

        st.markdown("##### Current values")
        st.json({"name": name, "bio": bio, "age": age, "rating": rating})

    # ==================== Selection Widgets ====================
    with tab_selection:
        st.subheader("Selection Widgets")

        col1, col2 = st.columns(2)
        with col1:
            department = st.selectbox(
                "Department",
                ["Engineering", "Marketing", "Sales", "HR", "Finance"],
            )
            skills = st.multiselect(
                "Skills",
                ["Python", "SQL", "JavaScript", "Go", "Rust", "Excel"],
                default=["Python"],
            )
        with col2:
            level = st.radio("Seniority", ["Junior", "Mid", "Senior", "Lead"])
            agree = st.checkbox("I agree to the terms")

        st.markdown("##### Current values")
        st.json(
            {
                "department": department,
                "skills": skills,
                "level": level,
                "agree": agree,
            }
        )

    # ==================== Date & Time ====================
    with tab_datetime:
        st.subheader("Date & Time Pickers")

        col1, col2 = st.columns(2)
        with col1:
            start_date = st.date_input("Start date")
        with col2:
            start_time = st.time_input("Start time")

        st.markdown("##### Current values")
        st.json({"start_date": str(start_date), "start_time": str(start_time)})

    # ==================== Form Submit ====================
    with tab_form:
        st.subheader("Grouped Form Submission")
        st.markdown(
            "Widgets inside a `st.form` do **not** trigger a rerun until "
            "the submit button is pressed.  This is useful for batch submissions."
        )

        with st.form("demo_form"):
            form_name = st.text_input("Full name")
            form_email = st.text_input("Email")
            form_dept = st.selectbox(
                "Department",
                ["Engineering", "Marketing", "Sales", "HR", "Finance"],
                key="form_dept",
            )
            form_salary = st.number_input(
                "Expected salary", min_value=0, value=50000, step=1000
            )
            submitted = st.form_submit_button("Submit")

        if submitted:
            st.success("Form submitted successfully!")
            st.json(
                {
                    "full_name": form_name,
                    "email": form_email,
                    "department": form_dept,
                    "expected_salary": form_salary,
                }
            )


if __name__ == "__main__":
    render()
