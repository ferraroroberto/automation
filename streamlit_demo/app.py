"""
Streamlit Capabilities Demo — Main Entry Point.

This application serves as an interactive playground showcasing all major
Streamlit features. Each capability is demonstrated in a self-contained
page module, dynamically loaded from the sidebar menu.

Run with:
    streamlit run app.py
"""

import streamlit as st

from menu import PAGE_REGISTRY, load_page

# ── Page config (must be the first Streamlit command) ────────────────────────
st.set_page_config(
    page_title="Streamlit Capabilities Demo",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded",
)


def _render_home():
    """Render the home / welcome page."""
    st.title("🚀 Streamlit Capabilities Demo")
    st.markdown(
        """
        Welcome to the **Streamlit Capabilities Demo**!
        This application is an educational playground that demonstrates
        the most important features you need to build real-world Streamlit apps.

        Use the **sidebar** on the left to navigate between demos.

        ---
        """
    )

    st.subheader("Available Demos")

    cols = st.columns(3)
    for idx, page in enumerate(PAGE_REGISTRY):
        if page.key == "home":
            continue
        with cols[idx % 3]:
            st.markdown(
                f"""
                **{page.icon} {page.title}**

                {page.description}
                """
            )

    st.markdown("---")
    st.caption("Built as an internal training reference • Streamlit Capabilities Demo")


def main():
    # ── Sidebar navigation ───────────────────────────────────────────────
    st.sidebar.title("🧭 Navigation")
    st.sidebar.markdown("---")

    page_options = {page.key: f"{page.icon} {page.title}" for page in PAGE_REGISTRY}
    selected_key = st.sidebar.radio(
        "Go to",
        options=list(page_options.keys()),
        format_func=lambda k: page_options[k],
        label_visibility="collapsed",
    )

    st.sidebar.markdown("---")
    st.sidebar.caption("Streamlit Capabilities Demo v1.0")

    # ── Render selected page ─────────────────────────────────────────────
    selected_page = next(p for p in PAGE_REGISTRY if p.key == selected_key)

    if selected_page.key == "home":
        _render_home()
    else:
        render_fn = load_page(selected_page)
        render_fn()


if __name__ == "__main__":
    main()
