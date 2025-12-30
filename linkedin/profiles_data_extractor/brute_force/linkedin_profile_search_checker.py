"""
LinkedIn Profile Search Checker Module

Extracts profile names and page numbers from LinkedIn search result tabs in Chrome,
then compares with existing contacts to show which profiles need to be opened.
"""

import hashlib
import json
import logging
import os
import re
import time
from datetime import datetime
from typing import Optional, Dict, List, Tuple, Set

import pandas as pd
import pyperclip
import win32gui
import win32con
from pynput import keyboard

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class LinkedInProfileSearchChecker:
    """Checks LinkedIn search results for profiles not yet in contacts."""

    def __init__(self, config_path: str):
        """Initialize search checker with configuration.

        Args:
            config_path: Path to the JSON configuration file.
        """
        self.controller = keyboard.Controller()
        self.config = self.load_config(config_path)

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
            logger.info("")
            logger.info("✅ Chrome window activated")
            logger.info("")
            return True
        except Exception as e:
            logger.error(f"❌ Failed to activate Chrome window: {e}")
            return False

    def copy_all_content(self) -> str:
        """Copy all content from the current page using CTRL+A and CTRL+C."""
        try:
            time.sleep(0.5)

            # Move focus from address bar to page content
            self.controller.press(keyboard.Key.tab)
            self.controller.release(keyboard.Key.tab)
            time.sleep(0.3)

            # CTRL+A to select all
            self.controller.press(keyboard.Key.ctrl)
            self.controller.press('a')
            self.controller.release('a')
            self.controller.release(keyboard.Key.ctrl)
            time.sleep(0.5)

            # CTRL+C to copy
            self.controller.press(keyboard.Key.ctrl)
            self.controller.press('c')
            self.controller.release('c')
            self.controller.release(keyboard.Key.ctrl)
            time.sleep(0.5)

            content = pyperclip.paste()
            if content:
                logger.info(f"✅ Content copied ({len(content)} characters)")
            else:
                logger.warning("⚠️  Clipboard is empty after copying content")

            return content if content else ""
        except Exception as e:
            logger.error(f"❌ Failed to copy content: {e}")
            return ""

    def extract_names_from_tab(self, content: str) -> List[str]:
        """Extract profile names from tab content using 'NameView Name's profile' pattern.

        Args:
            content: The copied content from the Chrome tab.

        Returns:
            List of extracted names.
        """
        names = []

        if not content:
            return names

        logger.info(f"🔍 Analyzing content ({len(content)} chars)...")

        # Debug: Look for any "View" patterns in the content
        view_matches = re.findall(r'View[^\']*', content, re.IGNORECASE)
        if view_matches:
            logger.info(f"✅ Found 'View' patterns")
        else:
            logger.warning("⚠️  No 'View' patterns found in content")

        # Debug: Look for profile patterns
        profile_matches = re.findall(r'[^\']*profile', content, re.IGNORECASE)
        if not profile_matches:
            logger.warning("⚠️  No 'profile' patterns found in content")

        # Pattern: "View [NAME]'s profile" (with space after View)
        # Simply capture whatever is between "View " and "'s profile"
        # Allow straight or curly apostrophes and keep the name capture non-greedy
        pattern = r"View\s+(.+?)[’']s\s+profile"
        matches = re.findall(pattern, content, re.IGNORECASE)

        for match in matches:
            name = match.strip()
            # Skip single-word names and ensure minimum length
            if name and len(name.split()) >= 2:
                names.append(name)
                logger.info(f"✅ Found profile: {name}")

        # Remove duplicates while preserving order
        seen = set()
        unique_names = []
        for name in names:
            normalized = name.lower().strip()
            if normalized not in seen:
                seen.add(normalized)
                unique_names.append(name)

        logger.info(f"📊 Total unique names extracted: {len(unique_names)}")
        return unique_names

    def extract_page_number(self, content: str) -> Optional[int]:
        """Extract current page number from 'Currently on the page X of Y search result pages.' pattern.

        Args:
            content: The copied content from the Chrome tab.

        Returns:
            Current page number if found, None otherwise.
        """
        if not content:
            return None

        # Pattern: "Currently on the page X of Y search result pages."
        pattern = r'Currently\s+on\s+the\s+page\s+(\d+)\s+of\s+\d+\s+search\s+result\s+pages'
        match = re.search(pattern, content, re.IGNORECASE)

        if match:
            page_num = int(match.group(1))
            logger.info(f"✅ Found page number: {page_num}")
            return page_num

        logger.warning("⚠️  Page number pattern not found in content")
        return None

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

    def switch_to_next_tab(self) -> None:
        """Switch to the next tab using CTRL+Tab."""

        logger.info("")
        logger.info("🔄 Switching to next tab...")
        logger.info("")
        self.controller.press(keyboard.Key.ctrl)
        self.controller.press(keyboard.Key.tab)
        self.controller.release(keyboard.Key.tab)
        self.controller.release(keyboard.Key.ctrl)
        time.sleep(1.5)  # Wait for tab switch to complete

    def check_search_tabs(self, max_tabs: int = 20) -> Dict[int, List[str]]:
        """Check all search result tabs and extract names with their page numbers.

        Args:
            max_tabs: Maximum number of tabs to process.

        Returns:
            Dictionary mapping page numbers to lists of names found on that page.
        """
        logger.info("🚀 Starting search tab checking")

        # Wait for user to switch to Chrome
        logger.info("⏳ Waiting 5 seconds for you to switch to Chrome window...")
        time.sleep(5)

        # Activate Chrome
        if not self.activate_chrome():
            logger.error("❌ Could not activate Chrome. Please make sure Chrome is open.")
            return {}

        page_names = {}  # page_number -> list of names
        processed_pages = set()
        content_hashes = set()  # Track content to detect cycling
        previous_content_hash = None
        consecutive_no_progress = 0
        MIN_CONTENT_LENGTH = 100  # Minimum content length to consider valid

        for tab_num in range(max_tabs):
            logger.info(f"📄 Processing tab {tab_num + 1} (max: {max_tabs})")

            # Copy content from current tab
            content = self.copy_all_content()

            if not content:
                logger.warning(f"⚠️  No content copied from tab {tab_num + 1}")
                consecutive_no_progress += 1
                if consecutive_no_progress >= 2:
                    logger.info("⏹️  No content found for 2 consecutive tabs - stopping")
                    break
                if tab_num < max_tabs - 1:
                    self.switch_to_next_tab()
                continue

            # Check if content is too short (likely not a LinkedIn search page)
            if len(content) < MIN_CONTENT_LENGTH:
                logger.warning(f"⚠️  Content too short ({len(content)} chars) - likely not a LinkedIn search page")
                consecutive_no_progress += 1
                if consecutive_no_progress >= 2:
                    logger.info("⏹️  Non-search pages detected - stopping")
                    break
                if tab_num < max_tabs - 1:
                    self.switch_to_next_tab()
                continue

            # Create content hash to detect cycling
            content_hash = hashlib.md5(content.encode('utf-8')).hexdigest()

            # Check if we've seen this exact content before (cycling back)
            if content_hash in content_hashes:
                logger.info("⏹️  Detected tab cycling (same content as previous tab) - stopping")
                break

            # Check if content is identical to previous tab (also indicates cycling)
            if content_hash == previous_content_hash:
                logger.info("⏹️  Same content as previous tab - stopping")
                break

            content_hashes.add(content_hash)
            previous_content_hash = content_hash
            consecutive_no_progress = 0  # Reset counter on successful content

            # Extract page number and names
            page_number = self.extract_page_number(content)
            names = self.extract_names_from_tab(content)

            if page_number and page_number not in processed_pages:
                processed_pages.add(page_number)
                page_names[page_number] = names
                logger.info(f"✅ Processed page {page_number} with {len(names)} profiles")
            elif page_number in processed_pages:
                logger.info(f"⏭️  Page {page_number} already processed - stopping")
                break
            else:
                logger.warning(f"⚠️  Could not extract page number from tab {tab_num + 1}")

            # Switch to next tab
            if tab_num < max_tabs - 1:  # Don't switch on last iteration
                self.switch_to_next_tab()

        logger.info(f"✅ Completed checking: found {len(page_names)} pages with profiles")
        return page_names

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
        print("MISSING PROFILES TO OPEN")
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
        """Run the complete search check process."""
        logger.info("🎯 Starting LinkedIn Profile Search Checker")
        logger.info("=" * 80)

        try:
            # Step 1: Check search tabs
            logger.info("📋 Step 1: Extracting names and page numbers from search tabs...")
            max_tabs = self.config.get('max_tabs', 20)
            page_names = self.check_search_tabs(max_tabs)

            if not page_names:
                logger.warning("⚠️  No search data extracted")
                return

            # Step 2: Find missing profiles
            logger.info("📋 Step 2: Comparing with existing contacts...")
            missing_profiles = self.find_missing_profiles(page_names)

            # Step 3: Display results
            logger.info("📋 Step 3: Displaying missing profiles...")
            self.display_missing_profiles(missing_profiles)

            logger.info("✅ Search check completed successfully")

        except Exception as e:
            logger.error(f"❌ Search check failed: {e}")
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

    # Initialize and run checker
    checker = LinkedInProfileSearchChecker(config_path)
    checker.run_check()


if __name__ == "__main__":
    main()
