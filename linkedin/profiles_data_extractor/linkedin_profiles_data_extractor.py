"""
LinkedIn Profiles Data Extractor Module

Extracts profile data from multiple Chrome tabs automatically.
Cycles through all open tabs, extracts profile information, and saves to Excel.
"""

import logging
import re
import time
from datetime import datetime
from typing import Optional, Dict, List, Tuple
import json
import os

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


class LinkedInProfileExtractor:
    """Extract LinkedIn profile data from Chrome tabs."""

    def __init__(self):
        self.controller = keyboard.Controller()

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

    def copy_url(self, silent: bool = False) -> str:
        """Copy the URL from Chrome address bar using keyboard shortcuts.

        Args:
            silent: If True, don't log the URL copy operation.
        """
        try:
            time.sleep(0.3)

            # CTRL+L to focus address bar
            self.controller.press(keyboard.Key.ctrl)
            self.controller.press('l')
            self.controller.release('l')
            self.controller.release(keyboard.Key.ctrl)
            time.sleep(0.5)

            # CTRL+C to copy URL
            self.controller.press(keyboard.Key.ctrl)
            self.controller.press('c')
            self.controller.release('c')
            self.controller.release(keyboard.Key.ctrl)
            time.sleep(0.5)

            url = pyperclip.paste()
            if url and not silent:
                logger.info(f"✅ URL copied: {url}")
            elif not url:
                logger.warning("⚠️  Clipboard is empty after copying URL")

            # ESC twice to deselect text in address bar (important for tab switching)
            self.controller.press(keyboard.Key.esc)
            self.controller.release(keyboard.Key.esc)
            time.sleep(0.1)
            self.controller.press(keyboard.Key.esc)
            self.controller.release(keyboard.Key.esc)
            time.sleep(0.2)

            return url if url else ""
        except Exception as e:
            logger.error(f"❌ Failed to copy URL: {e}")
            return ""

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

    def extract_follow_date(self, text: str) -> Optional[str]:
        """Extract the date from text following 'following you since'."""
        if not text:
            return None

        # Extract month abbreviation and 4-digit year after "since"
        pattern = r'following you since\s+(\w+)\s+(\d{4})'
        match = re.search(pattern, text, re.IGNORECASE)

        if match:
            month = match.group(1)
            year = match.group(2)
            date_text = f"{month} {year}"
            logger.info(f"✅ Found date text: {date_text}")
            return date_text

        logger.warning("⚠️  'following you since' pattern not found in text")
        return None

    def convert_date_format(self, date_text: str) -> Optional[str]:
        """Convert date from 'Oct 2023' format to Excel-compatible format."""
        if not date_text:
            return None

        try:
            # Parse the date using datetime - this will handle various formats
            parsed_date = datetime.strptime(date_text.strip(), '%b %Y')
            # Format as MM/YYYY for Excel compatibility (e.g., "10/2023")
            formatted_date = parsed_date.strftime('%m/%Y')
            logger.info(f"✅ Converted date: {date_text} -> {formatted_date}")
            return formatted_date
        except ValueError as e:
            logger.warning(f"⚠️  Could not parse date '{date_text}': {e}")
            return None

    def extract_name(self, text: str) -> Optional[str]:
        """Extract the person's name appearing after 'For Business', 'Learning', or 'Background Image'."""
        if not text:
            return None

        # Find the position of the first marker that appears near the beginning of profile content
        # Usually appears after navigation elements like "Home", "My Network", etc.
        markers = ['For Business', 'Learning', 'Background Image']

        # First, try to find markers that appear after common navigation elements
        # Look for the first marker that appears after navigation keywords
        nav_keywords = ['Home', 'My Network', 'Jobs', 'Messaging', 'Notifications']
        best_pos = -1

        for marker in markers:
            marker_pos = text.upper().find(marker.upper())
            if marker_pos >= 0:
                # Check if this marker appears after navigation elements
                # Find the last navigation keyword before this marker
                last_nav_pos = -1
                for nav in nav_keywords:
                    nav_pos = text.upper().rfind(nav.upper(), 0, marker_pos)
                    if nav_pos > last_nav_pos:
                        last_nav_pos = nav_pos

                # If we found navigation before this marker, and this marker is closer to the start than previous best
                if last_nav_pos >= 0 and (best_pos == -1 or marker_pos < best_pos):
                    best_pos = marker_pos

        # If we didn't find a marker after navigation, fall back to finding the first marker
        if best_pos == -1:
            for marker in markers:
                pos = text.upper().find(marker.upper())
                if pos >= 0 and (best_pos == -1 or pos < best_pos):
                    best_pos = pos

        if best_pos >= 0:
            # Extract everything after the marker
            marker_line = text[best_pos:].split('\n')[0]
            section_text = text[best_pos + len(marker_line):].strip()
            # Stop at *** if it exists
            star_pos = section_text.find('***')
            if star_pos >= 0:
                section_text = section_text[:star_pos].strip()

            # Split into lines and find potential names
            lines = [line.strip() for line in section_text.split('\n') if line.strip() and not line.strip().startswith('*')]

            # Look for names that appear multiple times (allowing for whitespace differences)
            for line in lines:
                # Check if line is name-like: starts with capital, reasonable length, not a marker
                if (2 < len(line) <= 50 and line[0].isupper() and
                    line not in ['For Business', 'Learning', 'Background Image']):

                    # Count occurrences allowing for slight whitespace differences
                    # Remove extra whitespace and compare normalized versions
                    normalized_line = ' '.join(line.split())
                    normalized_section = ' '.join(section_text.split())

                    # Count how many times this name appears (normalized)
                    count = normalized_section.lower().count(normalized_line.lower())
                    if count >= 2:
                        logger.info(f"✅ Found name: {line}")
                        return line

            # Fallback: if no repeated names found, take the first valid name-like line
            # This handles cases where the name appears only once but is clearly a name
            for line in lines:
                if (2 < len(line) <= 50 and line[0].isupper() and
                    line not in ['For Business', 'Learning', 'Background Image'] and
                    not any(char.isdigit() for char in line)):  # Names don't contain digits
                    logger.info(f"✅ Found name (single occurrence): {line}")
                    return line

        logger.warning("⚠️  Name pattern not found in text")
        return None

    def extract_job_title(self, text: str, name: Optional[str] = None) -> Optional[str]:
        """Extract the job title appearing after connection info (usually after 'degree connection')."""
        if not text:
            return None

        # Primary method: find text after connection info (most reliable)
        # Look for "degree connection" followed by the job title on the next line
        pattern = r'\d+(?:st|nd|rd|th)?\s*degree\s*connection[^\n\r]*?\s*\n\s*([^\n\r]+)'
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            job_title = match.group(1).strip()
            # Clean up any trailing connection info that might have been captured
            job_title = re.sub(r'\s+\d+(?:st|nd|rd|th)?\s*degree\s*connection.*$', '', job_title, flags=re.IGNORECASE)
            # Filter out pronouns and other non-job-title content
            if job_title and not re.match(r'^(he/him|she/her|they/them|he|she|they)$', job_title, re.IGNORECASE):
                logger.info(f"✅ Found job title: {job_title[:100]}...")
                return job_title

        # If name is provided, search after it (fallback)
        if name:
            # Find text after name (second occurrence), before connection info
            # The name appears twice, so we want the text after the second occurrence
            name_escaped = re.escape(name)
            pattern = rf'{name_escaped}.*?{name_escaped}\s+([^\n\r]+?)(?:\s+\d+(?:st|nd|rd|th)\s+degree\s+connection|\d+degree\s+connection|$)'
            match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE | re.DOTALL)
            if match:
                job_title = match.group(1).strip()
                # Clean up if it contains connection info
                job_title = re.sub(r'\s+\d+(?:st|nd|rd|th)?\s*degree\s*connection.*$', '', job_title, flags=re.IGNORECASE)
                # Filter out pronouns
                if job_title and not re.match(r'^(he/him|she/her|they/them|he|she|they)$', job_title, re.IGNORECASE):
                    logger.info(f"✅ Found job title: {job_title[:100]}...")
                    return job_title

        # Fallback: find text after markers, before connection info
        pattern = r'(?:For Business|Learning|Background Image)\s+[A-Z][a-zA-Z\s\.]+\s+[A-Z][a-zA-Z\s\.]+\s+([^\n\r]+?)(?:\s+\d+(?:st|nd|rd|th)?\s*degree\s*connection|$)'
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE | re.DOTALL)
        if match:
            job_title = match.group(1).strip()
            job_title = re.sub(r'\s+\d+(?:st|nd|rd|th)?\s*degree\s*connection.*$', '', job_title, flags=re.IGNORECASE)
            # Filter out pronouns
            if job_title and not re.match(r'^(he/him|she/her|they/them|he|she|they)$', job_title, re.IGNORECASE):
                logger.info(f"✅ Found job title: {job_title[:100]}...")
                return job_title

        logger.warning("⚠️  Job title pattern not found in text")
        return None

    def extract_company(self, text: str, job_title: Optional[str] = None) -> Optional[str]:
        """Extract the company name appearing after the job title with empty lines separating them."""
        if not text:
            return None

        if job_title:
            # Find text after job title, separated by newlines (job title line, empty line, company line)
            # Escape job title but allow flexible matching
            job_title_escaped = re.escape(job_title)
            pattern = rf'{job_title_escaped}\s*\n\s*\n\s*([^\n\r]+)'
            match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE | re.DOTALL)
            if match:
                company = match.group(1).strip()
                # Clean up trailing whitespace and location markers
                company = re.sub(r'\s+(?:Metropolitan Area|Area|City|County|State|Country|contact info|Contact info|CONTACT INFO).*$', '', company, flags=re.IGNORECASE)
                # Limit to reasonable company name length (usually under 100 chars)
                if company and len(company) < 200 and not re.search(r'\d+(?:st|nd|rd|th)\s*degree', company, re.IGNORECASE):
                    logger.info(f"✅ Found company: {company}")
                    return company

        # Fallback: find text after connection info, before location markers
        pattern = r'\d+(?:st|nd|rd|th)?\s*degree\s*connection[^\n\r]*?\s*\n\s*\n\s*([^\n\r]+)'
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE | re.DOTALL)
        if match:
            company = match.group(1).strip()
            company = re.sub(r'\s+(?:Metropolitan Area|Area|City|County|State|Country|contact info|Contact info|CONTACT INFO).*$', '', company, flags=re.IGNORECASE)
            if company and len(company) < 200 and not re.search(r'\d+(?:st|nd|rd|th)\s*degree', company, re.IGNORECASE):
                logger.info(f"✅ Found company: {company}")
                return company

        logger.warning("⚠️  Company pattern not found in text")
        return None

    def extract_location(self, text: str, company: Optional[str] = None) -> Optional[str]:
        """Extract the location appearing before 'contact info' on the same line."""
        if not text:
            return None

        # Simple approach: find any line containing "Contact info" and extract everything before it on that line
        lines = text.split('\n')
        for line in lines:
            line = line.strip()
            if 'contact info' in line.lower():
                # Find the position of "contact info" (case insensitive)
                contact_pos = line.lower().find('contact info')
                if contact_pos > 0:
                    location = line[:contact_pos].strip()
                    if location and len(location) < 150:
                        logger.info(f"✅ Found location: {location}")
                        return location

        logger.warning("⚠️  Location pattern not found in text")
        return None

    def extract_profile_data(self) -> Dict[str, str]:
        """Extract all profile data from the current tab.

        Returns:
            Dictionary containing extracted profile data.
        """
        # Copy all page content first (before getting URL, which focuses address bar)
        logger.info("📋 Copying all page content...")
        content = self.copy_all_content()

        if not content:
            logger.error("❌ Failed to copy content.")
            return {}

        # Get URL after copying content (address bar focus won't interfere now)
        logger.info("📋 Copying URL from address bar...")
        url = self.copy_url()

        # Extract all fields
        logger.info("🔍 Searching for profile information...")
        name = self.extract_name(content)
        job_title = self.extract_job_title(content, name)
        company = self.extract_company(content, job_title)
        location = self.extract_location(content, company)
        date_text = self.extract_follow_date(content)

        # Convert date format if found
        if date_text:
            formatted_date = self.convert_date_format(date_text)
            follows_from = formatted_date if formatted_date else date_text
        else:
            follows_from = ""

        return {
            'name': name or "",
            'url': url,
            'job_title': job_title or "",
            'follows_from': follows_from,
            'company': company or "",
            'location': location or ""
        }

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

    def extract_all_tabs(self, max_tabs: int = 50) -> List[Dict[str, str]]:
        """Extract profile data from all Chrome tabs.

        Args:
            max_tabs: Maximum number of tabs to process.

        Returns:
            List of dictionaries containing profile data from each tab.
        """
        logger.info("🚀 Starting extraction from all Chrome tabs")

        # Wait for user to switch to Chrome
        logger.info("⏳ Waiting 5 seconds for you to switch to Chrome window...")
        time.sleep(5)

        # Activate Chrome
        if not self.activate_chrome():
            logger.error("❌ Could not activate Chrome. Please make sure Chrome is open.")
            return []

        all_profiles = []
        processed_urls = set()  # Track processed URLs to avoid duplicates

        for tab_num in range(max_tabs):
            logger.info(f"📄 Processing tab {tab_num + 1}/{max_tabs}")

            # Extract data from current tab
            profile_data = self.extract_profile_data()

            # Check if we got valid data and haven't processed this URL before
            if profile_data.get('url') and profile_data['url'] not in processed_urls:
                processed_urls.add(profile_data['url'])

                # Only save profiles with complete data (all required fields present)
                required_fields = ['name', 'job_title', 'company', 'location']
                has_all_fields = all(
                    profile_data.get(field) and profile_data[field].strip()
                    for field in required_fields
                )

                if has_all_fields:
                    all_profiles.append(profile_data)
                    logger.info(f"✅ Extracted complete data from tab {tab_num + 1}: {profile_data.get('name', 'Unknown')}")
                else:
                    missing_fields = [
                        field for field in required_fields
                        if not profile_data.get(field) or not profile_data[field].strip()
                    ]
                    logger.warning(f"⚠️  Incomplete data from tab {tab_num + 1} - missing: {', '.join(missing_fields)}")
            elif profile_data.get('url') in processed_urls:
                logger.info(f"⏭️  Skipping duplicate URL: {profile_data.get('url')}")
            else:
                logger.warning(f"⚠️  No valid data extracted from tab {tab_num + 1}")

            # Switch to next tab
            if tab_num < max_tabs - 1:  # Don't switch on last iteration
                self.switch_to_next_tab()

                # Check if we're back to the first tab (Chrome cycles through tabs)
                # We'll stop if we detect we're looping back
                if tab_num > 0 and len(all_profiles) > 0:
                    current_url = self.copy_url(silent=True)
                    if current_url == all_profiles[0]['url']:
                        logger.info("🔄 Detected tab cycle - stopping extraction")
                        break

        logger.info(f"✅ Completed extraction: {len(all_profiles)} profiles found")
        return all_profiles


def save_to_excel(profiles_data: List[Dict[str, str]], output_path: str) -> str:
    """Save profile data to Excel file with timestamp.

    Args:
        profiles_data: List of profile dictionaries.
        output_path: Directory path for output file.

    Returns:
        Path to the created Excel file.
    """
    if not profiles_data:
        logger.warning("⚠️  No data to save")
        return ""

    # Create DataFrame
    df = pd.DataFrame(profiles_data)

    # Generate timestamp filename
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    filename = f"contacts_{timestamp}.xlsx"
    filepath = os.path.join(output_path, filename)

    # Ensure output directory exists
    os.makedirs(output_path, exist_ok=True)

    # Save to Excel
    df.to_excel(filepath, index=False, engine='openpyxl')
    logger.info(f"💾 Saved {len(profiles_data)} profiles to: {filepath}")

    return filepath


def main() -> None:
    """Main execution function."""
    logger.info("🚀 LinkedIn Profiles Data Extractor Module started")

    # Initialize extractor
    extractor = LinkedInProfileExtractor()

    # Extract data from all tabs
    profiles_data = extractor.extract_all_tabs()

    if profiles_data:
        # Save to Excel
        output_path = "E:\\onedrive\\Documentos\\Roberto\\areas\\work\\contacts reachout"
        excel_file = save_to_excel(profiles_data, output_path)

        # Display summary
        print("\n" + "="*80)
        print("EXTRACTION COMPLETE")
        print("="*80)
        print(f"Total profiles extracted: {len(profiles_data)}")
        print(f"Excel file saved to: {excel_file}")
        print("\nProfile Summary:")
        for i, profile in enumerate(profiles_data, 1):
            print(f"{i}. {profile.get('name', 'Unknown')} - {profile.get('job_title', 'Unknown')} at {profile.get('company', 'Unknown')}")
        print("="*80 + "\n")
    else:
        logger.error("❌ No profile data was extracted")


if __name__ == "__main__":
    main()
