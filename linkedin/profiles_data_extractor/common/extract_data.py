import streamlit as st
import subprocess
import os
from pathlib import Path
import time

def main():
    """Main function for the LinkedIn Data Extraction tab."""

    st.header("📥 LinkedIn Data Extraction")

    # Chrome Debug Section
    st.subheader("🌐 1. Start Chrome Debug Mode")

    st.markdown("""
    **First, start Chrome with remote debugging enabled for data extraction.**

    This launches Chrome on port 9222 with debugging enabled, allowing the extraction tool to access tab data.
    """)

    col1, col2 = st.columns(2)

    with col1:
        if st.button("🚀 Start Chrome Debug", type="primary", use_container_width=True):
            with st.spinner("Starting Chrome with debugging..."):
                try:
                    # Start Chrome directly with debugging enabled
                    chrome_exe = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
                    user_profile = os.environ.get('USERPROFILE', '')
                    debug_profile_dir = os.path.join(user_profile, r'AppData\Local\Google\Chrome\User Data\DebugProfile')

                    # Create debug profile directory if it doesn't exist
                    os.makedirs(debug_profile_dir, exist_ok=True)

                    # Launch Chrome with remote debugging
                    chrome_cmd = [
                        chrome_exe,
                        '--remote-debugging-port=9222',
                        '--remote-allow-origins=*',
                        f'--user-data-dir={debug_profile_dir}'
                    ]

                    # Use subprocess.Popen to start Chrome asynchronously
                    process = subprocess.Popen(
                        chrome_cmd,
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                        creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                    )

                    # Give Chrome a moment to start
                    time.sleep(2)

                    # Check if Chrome is running by testing the debug port
                    try:
                        import requests
                        response = requests.get("http://localhost:9222/json/version", timeout=3)
                        if response.status_code == 200:
                            st.success("✅ Chrome started successfully with remote debugging enabled!")
                            st.info("🔍 Chrome should now be open. Navigate to LinkedIn and open the profiles you want to extract.")
                        else:
                            st.warning("⚠️ Chrome may have started but debug connection check failed.")
                    except:
                        st.warning("⚠️ Chrome startup initiated. Please check if Chrome opened successfully.")

                except Exception as e:
                    st.error(f"❌ Error starting Chrome debug: {e}")

    with col2:
        if st.button("🔍 Check Chrome Status", use_container_width=True):
            try:
                # Try to connect to Chrome debug port
                import requests
                response = requests.get("http://localhost:9222/json/version", timeout=2)

                if response.status_code == 200:
                    st.success("✅ Chrome is running with remote debugging enabled!")
                    data = response.json()
                    st.info(f"**Chrome Version:** {data.get('Browser', 'Unknown')}")
                else:
                    st.warning("⚠️ Chrome debug connection check failed")

            except Exception:
                st.error("❌ Chrome with remote debugging is not running on port 9222")
                st.info("💡 Use the 'Start Chrome Debug' button to launch Chrome with debugging enabled.")

    st.divider()

    # Orchestrator Section
    st.subheader("🔧 2. Run Data Extraction Orchestrator")

    st.markdown("""
    **After Chrome is running in debug mode, extract LinkedIn profile data.**

    This tool extracts profile data from your open Chrome tabs and merges new profiles into your Excel file.

    **Requirements:**
    - Chrome must be running with remote debugging enabled (see step 1 above)
    - LinkedIn profile tabs should be open with the data you want to extract
    """)

    if st.button("🚀 Run DevTools Orchestrator", type="primary", use_container_width=True):
        with st.spinner("Launching DevTools Orchestrator..."):
            try:
                # Path to the orchestrator batch file - use absolute path from project root
                current_file = Path(__file__).resolve()
                project_root = current_file.parent.parent.parent.parent  # Go up to automation/automation/
                orchestrator_bat_path = project_root / "linkedin" / "profiles_data_extractor" / "devtools" / "linkedin_profiles_data_orchestrator_devtools.bat"

                if not orchestrator_bat_path.exists():
                    st.error(f"Orchestrator batch file not found: {orchestrator_bat_path}")
                    return

                # Launch the orchestrator in a new command prompt window
                # This allows user interaction with the console prompts
                subprocess.Popen(
                    [str(orchestrator_bat_path)],
                    cwd=str(orchestrator_bat_path.parent),
                    creationflags=subprocess.CREATE_NEW_CONSOLE if os.name == 'nt' else 0
                )

                st.success("✅ DevTools Orchestrator launched!")
                st.info("🔍 A new command prompt window has opened. Follow the prompts in that window to run the extraction.")

            except Exception as e:
                st.error(f"❌ Error launching orchestrator: {e}")