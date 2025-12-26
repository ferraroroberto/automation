"""
LinkedIn Profile Search Checker Module - Chrome DevTools Version

Uses Chrome DevTools Protocol to directly query DOM elements for profile names
and pagination information, eliminating the need for text copy-paste operations.
"""

import json
import logging
import os
import re
import time
import requests
import websocket
from typing import Optional, Dict, List, Tuple, Set, Any
import pandas as pd
import win32gui
import win32con
from pynput import keyboard

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class ChromeDevToolsClient:
    """Chrome DevTools Protocol client for DOM manipulation and data extraction."""

    def __init__(self, debug_port: int = 9222):
        """Initialize Chrome DevTools client.

        Args:
            debug_port: Chrome remote debugging port (default: 9222)
        """
        self.debug_port = debug_port
        self.ws_url = None
        self.ws = None
        self.command_id = 1

    def connect_to_chrome(self) -> bool:
        """Connect to Chrome's debugging interface.

        Returns:
            True if connection successful, False otherwise
        """
        try:
            # Get list of available tabs
            response = requests.get(f"http://localhost:{self.debug_port}/json", timeout=5)
            response.raise_for_status()
            tabs = response.json()

            # Find a LinkedIn tab
            linkedin_tab = None
            for tab in tabs:
                if 'linkedin.com' in tab.get('url', '').lower():
                    linkedin_tab = tab
                    break

            if not linkedin_tab:
                # If no LinkedIn tab found, use the first tab
                if tabs:
                    linkedin_tab = tabs[0]
                    logger.warning("⚠️  No LinkedIn tab found, using first available tab")
                else:
                    logger.error("❌ No Chrome tabs found. Make sure Chrome is running with --remote-debugging-port=9222 --remote-allow-origins=*")
                    return False

            self.ws_url = linkedin_tab['webSocketDebuggerUrl']
            logger.info(f"✅ Connected to Chrome tab: {linkedin_tab.get('title', 'Unknown')}")

            # Connect to WebSocket
            self.ws = websocket.create_connection(self.ws_url)
            logger.info("✅ WebSocket connection established")
            return True

        except requests.exceptions.RequestException as e:
            logger.error(f"❌ Failed to connect to Chrome debugging port: {e}")
            logger.error("   Make sure Chrome is running with: chrome.exe --remote-debugging-port=9222 --remote-allow-origins=*")
            return False
        except Exception as e:
            logger.error(f"❌ Failed to establish WebSocket connection: {e}")
            return False

    def send_command(self, method: str, params: Dict = None) -> Dict:
        """Send a command to Chrome DevTools.

        Args:
            method: DevTools method name
            params: Method parameters

        Returns:
            Response from DevTools
        """
        if not self.ws:
            raise Exception("WebSocket not connected")

        command = {
            "id": self.command_id,
            "method": method,
            "params": params or {}
        }

        self.ws.send(json.dumps(command))
        self.command_id += 1

        # Receive response
        response = json.loads(self.ws.recv())
        if 'error' in response:
            raise Exception(f"DevTools error: {response['error']}")

        return response.get('result', {})

    def evaluate_javascript(self, script: str) -> Any:
        """Execute JavaScript in the page context.

        Args:
            script: JavaScript code to execute

        Returns:
            Result of JavaScript execution
        """
        result = self.send_command("Runtime.evaluate", {
            "expression": script,
            "returnByValue": True
        })
        return result.get('result', {}).get('value')

    def close(self):
        """Close the WebSocket connection."""
        if self.ws:
            self.ws.close()
            self.ws = None
            logger.info("🔌 WebSocket connection closed")


class LinkedInProfileSearchCheckerDevTools:
    """Checks LinkedIn search results using Chrome DevTools Protocol."""

    def __init__(self, config_path: str, debug_port: int = 9222):
        """Initialize search checker with Chrome DevTools.

        Args:
            config_path: Path to the JSON configuration file.
            debug_port: Chrome remote debugging port.
        """
        self.keyboard_controller = keyboard.Controller()
        self.config = self.load_config(config_path)
        self.chrome = ChromeDevToolsClient(debug_port)

    def find_chrome_window(self) -> Optional[int]:
        """Find Chrome window handle using win32gui.

        Returns:
            Window handle (HWND) if found, None otherwise.
        """
        chrome_windows = []

        def enum_handler(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                if "Google Chrome" in window_title:
                    windows.append(hwnd)

        try:
            win32gui.EnumWindows(enum_handler, chrome_windows)
            # Return the first Chrome window found
            return chrome_windows[0] if chrome_windows else None
        except Exception as e:
            logger.error(f"❌ Error finding Chrome window: {e}")
            return None

    def activate_chrome(self) -> bool:
        """Find and activate Chrome window.

        Returns:
            True if Chrome was found and activated, False otherwise.
        """
        hwnd = self.find_chrome_window()
        if not hwnd:
            logger.error("❌ Chrome window not found. Please open Chrome first.")
            return False

        try:
            # Restore window if minimized
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

            # Bring window to foreground
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.8)  # Wait for focus to settle
            logger.info("✅ Chrome window activated")
            return True
        except Exception as e:
            logger.error(f"❌ Failed to activate Chrome window: {e}")
            return False

    def load_config(self, config_path: str) -> Dict:
        """Load configuration from JSON file.

        Args:
            config_path: Path to the configuration file.

        Returns:
            Configuration dictionary.
        """
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            logger.info(f"✅ Configuration loaded from {config_path}")
            return config
        except Exception as e:
            logger.error(f"❌ Failed to load configuration: {e}")
            raise

    def extract_profile_names_from_dom(self) -> List[str]:
        """Extract profile names directly from DOM using DevTools.

        Returns:
            List of profile names found on current page.
        """
        # JavaScript to extract profile names from LinkedIn search results
        # This targets the specific DOM structure for LinkedIn profile cards
        js_script = """
        (function() {
            const names = [];

            // Primary selector based on actual LinkedIn DOM structure
            // Look for profile names in the specific span structure: span[dir="ltr"] span[aria-hidden="true"]
            const profileElements = document.querySelectorAll('span[dir="ltr"] span[aria-hidden="true"]');

            for (const element of profileElements) {
                const name = element.textContent?.trim();
                if (name && name.length > 2 && name.split(' ').length >= 2) {
                    // Clean up the name (remove extra whitespace, normalize)
                    const cleanName = name.replace(/\s+/g, ' ').trim();
                    if (!names.includes(cleanName)) {
                        names.push(cleanName);
                    }
                }
            }

            // Fallback selectors if primary doesn't work
            if (names.length === 0) {
                const fallbackSelectors = [
                    'a[data-test-id="profile-result-card"] span[dir="ltr"]',
                    '.entity-result__title-text a span[aria-hidden="true"]',
                    '.entity-result__title span[dir="ltr"]'
                ];

                for (const selector of fallbackSelectors) {
                    const elements = document.querySelectorAll(selector);
                    for (const element of elements) {
                        const name = element.textContent?.trim();
                        if (name && name.length > 2 && name.split(' ').length >= 2) {
                            const cleanName = name.replace(/\s+/g, ' ').trim();
                            if (!names.includes(cleanName)) {
                                names.push(cleanName);
                            }
                        }
                    }
                    if (names.length > 0) break;
                }
            }

            return names;
        })();
        """

        try:
            names = self.chrome.evaluate_javascript(js_script)
            if names and isinstance(names, list):
                logger.info(f"✅ Extracted {len(names)} profile names from DOM")
                return names
            else:
                logger.warning("⚠️  No profile names found in DOM")
                return []
        except Exception as e:
            logger.error(f"❌ Failed to extract profile names: {e}")
            return []

    def extract_page_number_from_dom(self) -> Optional[int]:
        """Extract current page number directly from DOM using DevTools.

        Returns:
            Current page number if found, None otherwise.
        """
        # JavaScript to extract pagination information
        js_script = """
        (function() {
            // Primary selector: look for button with aria-current="true" (current page)
            const currentPageButton = document.querySelector('button[aria-current="true"]');
            if (currentPageButton) {
                // The page number should be in a span inside the button
                const pageSpan = currentPageButton.querySelector('span');
                if (pageSpan) {
                    const pageNum = pageSpan.textContent?.trim();
                    if (pageNum && /^\d+$/.test(pageNum)) {
                        return parseInt(pageNum);
                    }
                }
                // Fallback: check the button text directly
                const buttonText = currentPageButton.textContent?.trim();
                if (buttonText && /^\d+$/.test(buttonText)) {
                    return parseInt(buttonText);
                }
            }

            // Fallback selectors for pagination text like "Page 1 of 5"
            const paginationSelectors = [
                '.artdeco-pagination__page-state',
                '.search-results__pagination-page-state',
                '[data-test-pagination-page-current]',
                '.pagination__page-state'
            ];

            for (const selector of paginationSelectors) {
                const element = document.querySelector(selector);
                if (element) {
                    const text = element.textContent || element.innerText;
                    // Match patterns like "Page 3 of 10" or "3 of 10"
                    const match = text.match(/(?:Page\s+)?(\d+)\s+of\s+\d+/i);
                    if (match) {
                        return parseInt(match[1]);
                    }
                }
            }

            // Alternative: look for active pagination button (older selectors)
            const activePageSelectors = [
                '.artdeco-pagination__indicator--active button',
                '.pagination__indicator--active',
                '[aria-current="page"]'
            ];

            for (const selector of activePageSelectors) {
                const element = document.querySelector(selector);
                if (element) {
                    const pageNum = element.textContent?.trim();
                    if (pageNum && /^\d+$/.test(pageNum)) {
                        return parseInt(pageNum);
                    }
                }
            }

            return null;
        })();
        """

        try:
            page_number = self.chrome.evaluate_javascript(js_script)
            if page_number and isinstance(page_number, int):
                logger.info(f"✅ Extracted page number: {page_number}")
                return page_number
            else:
                logger.warning("⚠️  Could not extract page number from DOM")
                return None
        except Exception as e:
            logger.error(f"❌ Failed to extract page number: {e}")
            return None

    def switch_to_next_tab(self) -> None:
        """Switch to the next tab using keyboard simulation (Ctrl+Tab)."""

        logger.info("🔄 Switching to next tab...")
        self.keyboard_controller.press(keyboard.Key.ctrl)
        self.keyboard_controller.press(keyboard.Key.tab)
        self.keyboard_controller.release(keyboard.Key.tab)
        self.keyboard_controller.release(keyboard.Key.ctrl)
        time.sleep(1.5)  # Wait for tab switch to complete

    def check_search_tabs_devtools(self, max_tabs: int = 20) -> Dict[int, List[str]]:
        """Check all search result tabs using DevTools and extract names with page numbers.

        Args:
            max_tabs: Maximum number of tabs to process.

        Returns:
            Dictionary mapping page numbers to lists of names found on that page.
        """
        logger.info("🚀 Starting DevTools-based search tab checking")

        # Wait for user to switch to Chrome
        logger.info("⏳ Waiting 5 seconds for you to switch to Chrome window...")
        time.sleep(5)

        # Activate Chrome window
        if not self.activate_chrome():
            logger.error("❌ Could not activate Chrome. Please make sure Chrome is open.")
            return {}

        # Connect to Chrome DevTools
        if not self.chrome.connect_to_chrome():
            logger.error("❌ Could not connect to Chrome DevTools")
            return {}

        try:
            page_names = {}  # page_number -> list of names
            processed_pages = set()
            tab_count = 0

            while tab_count < max_tabs:
                logger.info(f"📄 Processing tab {tab_count + 1} (max: {max_tabs})")

                # Extract data from current tab
                page_number = self.extract_page_number_from_dom()
                names = self.extract_profile_names_from_dom()

                if page_number and page_number not in processed_pages:
                    processed_pages.add(page_number)
                    page_names[page_number] = names
                    logger.info(f"✅ Processed page {page_number} with {len(names)} profiles")
                elif page_number in processed_pages:
                    logger.info(f"⏭️  Page {page_number} already processed - stopping")
                    break
                else:
                    logger.warning(f"⚠️  Could not extract page number from tab {tab_count + 1}")

                tab_count += 1

                # Switch to next tab if not the last iteration
                if tab_count < max_tabs:
                    self.switch_to_next_tab()
                    # Reconnect to the new active tab
                    self.chrome.close()  # Close current connection
                    time.sleep(0.5)  # Brief pause for tab switch to settle
                    if not self.chrome.connect_to_chrome():  # Reconnect to new active tab
                        logger.warning("⚠️  Could not reconnect to new tab - stopping")
                        break

            logger.info(f"✅ Completed DevTools checking: found {len(page_names)} pages with profiles")
            return page_names

        finally:
            self.chrome.close()

    def load_existing_contacts(self) -> Set[str]:
        """Load existing contact names from the destination Excel file.

        Returns:
            Set of existing contact names (normalized to lowercase).
        """
        destination_file = self.config.get('destination_file')
        if not destination_file:
            logger.error("❌ destination_file not specified in configuration")
            return set()

        if not os.path.exists(destination_file):
            logger.info(f"📄 Destination file doesn't exist: {destination_file}")
            return set()

        try:
            df = pd.read_excel(destination_file, engine='openpyxl')
            logger.info(f"📖 Read existing Excel file with {len(df)} rows")

            existing_names = set()
            if 'name' in df.columns:
                existing_names = set(
                    name.lower().strip()
                    for name in df['name'].dropna()
                    if isinstance(name, str) and name.strip()
                )

            logger.info(f"✅ Loaded {len(existing_names)} existing contact names")
            return existing_names

        except Exception as e:
            logger.error(f"❌ Failed to load existing contacts: {e}")
            return set()

    def find_missing_profiles(self, page_names: Dict[int, List[str]]) -> Dict[int, List[str]]:
        """Find profiles that don't exist in the contacts file.

        Args:
            page_names: Dictionary mapping page numbers to lists of names.

        Returns:
            Dictionary mapping page numbers to lists of missing names.
        """
        # Load existing contacts
        existing_contacts = self.load_existing_contacts()

        missing_profiles = {}

        for page_num, names in page_names.items():
            missing_names = []
            for name in names:
                normalized_name = name.lower().strip()
                if normalized_name not in existing_contacts:
                    missing_names.append(name)

            if missing_names:
                missing_profiles[page_num] = missing_names

        logger.info(f"✅ Found {sum(len(names) for names in missing_profiles.values())} missing profiles")
        return missing_profiles

    def display_missing_profiles(self, missing_profiles: Dict[int, List[str]]) -> None:
        """Display missing profiles in the requested format.

        Args:
            missing_profiles: Dictionary mapping page numbers to lists of missing names.
        """
        print("\n" + "="*80)
        print("MISSING PROFILES TO OPEN (DevTools Version)")
        print("="*80)

        if not missing_profiles:
            print("🎉 No missing profiles found! All search results are already in contacts.")
            return

        total_missing = 0
        for page_num in sorted(missing_profiles.keys()):
            names = missing_profiles[page_num]
            total_missing += len(names)

            for name in names:
                # Handle Unicode encoding issues on Windows console
                try:
                    print(f"page {page_num}: {name},")
                except UnicodeEncodeError:
                    # Fallback: encode to ASCII, replacing problematic characters
                    safe_name = name.encode('ascii', 'replace').decode('ascii')
                    print(f"page {page_num}: {safe_name},")

        print("="*80)
        print(f"Total missing profiles: {total_missing}")
        print("You need to open these profiles and extract their data.")
        print("="*80 + "\n")

    def run_check(self) -> None:
        """Run the complete search check process using DevTools."""
        logger.info("🎯 Starting LinkedIn Profile Search Checker (DevTools Version)")
        logger.info("=" * 80)

        try:
            # Step 1: Check search tabs using DevTools
            logger.info("📋 Step 1: Extracting names and page numbers from search tabs...")
            max_tabs = self.config.get('max_tabs', 20)
            page_names = self.check_search_tabs_devtools(max_tabs)

            if not page_names:
                logger.warning("⚠️  No search data extracted")
                return

            # Step 2: Find missing profiles
            logger.info("📋 Step 2: Comparing with existing contacts...")
            missing_profiles = self.find_missing_profiles(page_names)

            # Step 3: Display results
            logger.info("📋 Step 3: Displaying missing profiles...")
            self.display_missing_profiles(missing_profiles)

            logger.info("✅ DevTools search check completed successfully")

        except Exception as e:
            logger.error(f"❌ DevTools search check failed: {e}")
            raise


def main() -> None:
    """Main execution function."""
    # Load configuration from the common directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    common_dir = os.path.join(os.path.dirname(script_dir), "common")
    config_path = os.path.join(common_dir, "linkedin_profiles_data.json")

    if not os.path.exists(config_path):
        logger.error(f"❌ Configuration file not found: {config_path}")
        logger.error(f"   Current working directory: {os.getcwd()}")
        logger.error(f"   Script directory: {script_dir}")
        return

    # Initialize and run DevTools checker
    checker = LinkedInProfileSearchCheckerDevTools(config_path)
    checker.run_check()


if __name__ == "__main__":
    main()
