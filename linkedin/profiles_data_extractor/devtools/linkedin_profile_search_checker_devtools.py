"""
LinkedIn Profile Search Checker Module - Chrome DevTools Version

Uses Chrome DevTools Protocol to directly query DOM elements for profile names
and pagination information, eliminating the need for text copy-paste operations.
"""

import json
import os
import time
from typing import Optional, Dict, List, Set
import pandas as pd
from pynput import keyboard

from _chrome_client import ChromeDevToolsClient, setup_logging, activate_chrome

logger = setup_logging(__name__)


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
        if not activate_chrome():
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
