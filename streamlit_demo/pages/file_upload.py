"""
Demo – File Handling
====================
Demonstrates uploading and downloading files with Streamlit:
- Upload CSV, TXT, or JSON files and preview their content
- Download a generated CSV or JSON file

**Why this matters:** Many internal tools need to accept user-provided data
files and produce downloadable reports or exports.
"""

import io
import json

import pandas as pd
import streamlit as st

from data.loader import load_employees


def render() -> None:
    st.header("File Handling Demo")

    tab_upload, tab_download = st.tabs(["Upload Files", "Download Files"])

    # ==================== UPLOAD ====================
    with tab_upload:
        st.subheader("Upload a File")
        st.markdown("Supported formats: **CSV**, **TXT**, **JSON**.")

        uploaded = st.file_uploader(
            "Choose a file",
            type=["csv", "txt", "json"],
            key="file_uploader",
        )

        if uploaded is not None:
            st.success(f"Uploaded: **{uploaded.name}** ({uploaded.size:,} bytes)")

            ext = uploaded.name.rsplit(".", 1)[-1].lower()

            if ext == "csv":
                df = pd.read_csv(uploaded)
                st.markdown(f"**{len(df)} rows x {len(df.columns)} columns**")
                st.dataframe(df, use_container_width=True)

            elif ext == "json":
                data = json.load(uploaded)
                if isinstance(data, list):
                    df = pd.DataFrame(data)
                    st.markdown(f"**{len(df)} records**")
                    st.dataframe(df, use_container_width=True)
                else:
                    st.json(data)

            elif ext == "txt":
                text = uploaded.read().decode("utf-8", errors="replace")
                st.text_area("File contents", text, height=300)

    # ==================== DOWNLOAD ====================
    with tab_download:
        st.subheader("Download Generated Files")
        st.markdown(
            "Click a button below to download the mock employees dataset "
            "in different formats."
        )

        df = load_employees()

        col1, col2, col3 = st.columns(3)

        # CSV download
        with col1:
            csv_bytes = df.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="Download as CSV",
                data=csv_bytes,
                file_name="employees.csv",
                mime="text/csv",
                key="dl_csv",
            )

        # JSON download
        with col2:
            json_bytes = df.to_json(orient="records", indent=2).encode("utf-8")
            st.download_button(
                label="Download as JSON",
                data=json_bytes,
                file_name="employees.json",
                mime="application/json",
                key="dl_json",
            )

        # Excel-style download (as CSV with .xls extension for simplicity,
        # or use openpyxl if available — here we keep it dependency-free)
        with col3:
            buffer = io.BytesIO()
            df.to_csv(buffer, index=False)
            st.download_button(
                label="Download as TXT",
                data=buffer.getvalue(),
                file_name="employees.txt",
                mime="text/plain",
                key="dl_txt",
            )

        st.markdown("---")
        st.dataframe(df.head(10), use_container_width=True)
        st.caption("Preview of the dataset available for download.")


if __name__ == "__main__":
    render()
