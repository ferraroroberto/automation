"""
LinkedIn Profiles Data Extractor Module - Chrome DevTools Version

Extracts profile data from multiple Chrome tabs using DevTools Protocol.
Cycles through all open tabs, extracts profile information, and saves to Excel.
"""

import hashlib
import json
import logging
import os
import re
import time
import requests
import websocket
from datetime import datetime
from typing import Optional, Dict, List, Tuple

import pandas as pd
import win32gui
import win32con
from pynput import keyboard
from pathlib import Path

# Configure logging to file only (INFO level) and console (WARNING+ only)
log_file = Path(__file__).parent.parent / "logging.log"
log_file.parent.mkdir(parents=True, exist_ok=True)

# Create file handler for detailed logs
file_handler = logging.FileHandler(log_file, encoding='utf-8')
file_handler.setLevel(logging.INFO)
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

# Create console handler for warnings and errors only
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.WARNING)
console_handler.setFormatter(logging.Formatter('%(levelname)s - %(message)s'))

# Configure logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
logger.addHandler(file_handler)
logger.addHandler(console_handler)


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

    def _is_valid_linkedin_search_tab(self, tab: Dict) -> bool:
        """Check if a tab is a valid LinkedIn search page (not a tracking/analytics page).

        Args:
            tab: Tab dictionary from Chrome DevTools

        Returns:
            True if this is a valid LinkedIn search tab
        """
        url = tab.get('url', '').lower()
        tab_type = tab.get('type', '')

        # Must be a 'page' type (not background_page, service_worker, etc.)
        if tab_type != 'page':
            return False

        # Must be a LinkedIn URL
        if 'linkedin.com' not in url:
            return False

        # Exclude tracking/analytics pages
        excluded_patterns = [
            'merchantpool',
            'beacon',
            'tracking',
            'analytics',
            'px.ads',
            'platform.linkedin.com',
        ]
        for pattern in excluded_patterns:
            if pattern in url:
                return False

        # Prefer search results pages
        return True

    def get_linkedin_search_tabs(self) -> List[Dict]:
        """Get all valid LinkedIn search tabs.

        Returns:
            List of tab dictionaries for LinkedIn search pages
        """
        try:
            response = requests.get(f"http://localhost:{self.debug_port}/json", timeout=5)
            response.raise_for_status()
            tabs = response.json()

            # Filter to only valid LinkedIn search tabs
            valid_tabs = [tab for tab in tabs if self._is_valid_linkedin_search_tab(tab)]
            return valid_tabs

        except Exception as e:
            logger.error(f"❌ Failed to get Chrome tabs: {e}")
            return []

    def connect_to_tab(self, tab: Dict) -> bool:
        """Connect to a specific Chrome tab.

        Args:
            tab: Tab dictionary from Chrome DevTools

        Returns:
            True if connection successful, False otherwise
        """
        try:
            self.ws_url = tab['webSocketDebuggerUrl']
            self.ws = websocket.create_connection(self.ws_url)
            logger.info(f"✅ Connected to: {tab.get('title', 'Unknown')[:60]}")
            return True
        except Exception as e:
            logger.error(f"❌ Failed to connect to tab: {e}")
            return False

    def connect_to_chrome(self) -> bool:
        """Connect to Chrome's debugging interface (first valid LinkedIn tab).

        Returns:
            True if connection successful, False otherwise
        """
        try:
            # Get list of available tabs
            response = requests.get(f"http://localhost:{self.debug_port}/json", timeout=5)
            response.raise_for_status()
            tabs = response.json()

            # Find a valid LinkedIn search tab
            linkedin_tab = None
            for tab in tabs:
                if self._is_valid_linkedin_search_tab(tab):
                    linkedin_tab = tab
                    break

            if not linkedin_tab:
                # If no valid LinkedIn tab found, use the first page tab
                page_tabs = [t for t in tabs if t.get('type') == 'page']
                if page_tabs:
                    linkedin_tab = page_tabs[0]
                    logger.warning("⚠️  No LinkedIn search tab found, using first available page")
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
            logger.error("   Make sure Chrome is running with --remote-debugging-port=9222 --remote-allow-origins=*")
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

    def evaluate_javascript(self, script: str) -> any:
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

    def get_current_url(self) -> str:
        """Get the current page URL.

        Returns:
            Current page URL
        """
        result = self.evaluate_javascript("window.location.href")
        return result or ""

    def close(self):
        """Close the WebSocket connection."""
        if self.ws:
            self.ws.close()
            self.ws = None
            logger.info("🔌 WebSocket connection closed")


class LinkedInProfileExtractorDevTools:
    """Extract LinkedIn profile data using Chrome DevTools Protocol."""

    def __init__(self, debug_port: int = 9222):
        """Initialize DevTools extractor.

        Args:
            debug_port: Chrome remote debugging port
        """
        self.keyboard_controller = keyboard.Controller()
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

    def connect_to_chrome(self) -> bool:
        """Connect to Chrome DevTools.

        Returns:
            True if connection successful, False otherwise
        """
        return self.chrome.connect_to_chrome()

    def extract_profile_data(self) -> Dict[str, str]:
        """Extract profile data from the current LinkedIn page using DevTools.

        Works for both search results (first profile) and profile pages.

        Returns:
            Dictionary containing profile data.
        """
        profile_js = """
        (function() {
            const profile = {};

            // Try search results first
            const searchResult = document.querySelector('div[data-view-name="people-search-result"]');
            if (searchResult) {
                // Extract name from the title link
                const nameLink = searchResult.querySelector('a[data-view-name="search-result-lockup-title"]');
                if (nameLink) {
                    profile.name = nameLink.textContent?.trim();
                }

                // Extract URL
                const profileLink = searchResult.querySelector('a[href*="/in/"]');
                if (profileLink) {
                    let href = profileLink.getAttribute('href');
                    if (href) {
                        if (href.startsWith('/in/')) {
                            href = 'https://www.linkedin.com' + href;
                        }
                        profile.url_profile = href.split('?')[0];
                    }
                }

                // Extract job title and location from paragraphs
                const paragraphs = searchResult.querySelectorAll('p');
                let jobTitleFound = false;

                for (const p of paragraphs) {
                    const text = p.textContent?.trim();
                    if (!text) continue;
                    if (profile.name && text.includes(profile.name)) continue;
                    if (text.match(/^[•\\s]*[123](?:st|nd|rd)$/)) continue;
                    if (text.includes('followers')) continue;
                    if (text.includes('mutual connection')) continue;

                    if (!jobTitleFound) {
                        profile.job_title = text;
                        jobTitleFound = true;
                        continue;
                    }

                    profile.location = text;
                    break;
                }
            } else {
                // Fallback: Profile page selectors
                const nameEl = document.querySelector('h1.text-heading-xlarge') ||
                               document.querySelector('h1[data-test-id="hero__page__title"]');
                if (nameEl) profile.name = nameEl.textContent?.trim();

                const jobEl = document.querySelector('.pv-text-details__left-panel .text-body-medium');
                if (jobEl) profile.job_title = jobEl.textContent?.trim();

                const locEl = document.querySelector('.pv-text-details__left-panel .text-body-small');
                if (locEl) profile.location = locEl.textContent?.trim();

                profile.url_profile = window.location.href.split('?')[0];
            }

            // Extract company from search filter pill
            const companySelectors = [
                'button[id="searchFilter_currentCompany"]',
                '.artdeco-pill[aria-label*="Current company"]'
            ];

            for (const selector of companySelectors) {
                const element = document.querySelector(selector);
                if (element) {
                    const text = element.textContent?.trim();
                    if (text) {
                        profile.company = text.replace(/\\s+\\d+$/, '');
                        break;
                    }
                }
            }

            profile.follows_from = null;

            return profile;
        })();
        """

        result = self.chrome.evaluate_javascript(profile_js) or {}

        return {
            'name': result.get('name') or "",
            'url_profile': result.get('url_profile') or "",
            'job_title': result.get('job_title') or "",
            'follows_from': result.get('follows_from'),
            'company': result.get('company') or "",
            'location': result.get('location') or ""
        }

    def extract_search_results_profiles(self) -> List[Dict[str, str]]:
        """Extract multiple profiles from a LinkedIn search results page using DevTools.

        Returns:
            List of dictionaries containing profile data from all search results.
        """

        profiles_js = """
        (function() {
            const profiles = [];

            // Find all search result containers using the new LinkedIn structure
            const searchResults = document.querySelectorAll('div[data-view-name="people-search-result"]');

            for (let i = 0; i < searchResults.length; i++) {
                const result = searchResults[i];
                try {
                    const profile = {};

                    // Extract name from the title link
                    const nameLink = result.querySelector('a[data-view-name="search-result-lockup-title"]');
                    if (nameLink) {
                        profile.name = nameLink.textContent?.trim();
                    }

                    // Extract URL from LinkedIn profile link
                    const profileLink = result.querySelector('a[href*="/in/"]');
                    if (profileLink) {
                        let href = profileLink.getAttribute('href');
                        if (href) {
                            // Make sure it's a full URL
                            if (href.startsWith('/in/')) {
                                href = 'https://www.linkedin.com' + href;
                            }
                            // Clean URL by removing query parameters
                            profile.url_profile = href.split('?')[0];
                        }
                    }

                    // Find all <p> tags that end content blocks (followed by </p>)
                    // The structure after the name is: job_title in first p, location in second p
                    const paragraphs = result.querySelectorAll('p');
                    let jobTitleFound = false;
                    let locationFound = false;

                    for (const p of paragraphs) {
                        const text = p.textContent?.trim();
                        if (!text) continue;

                        // Skip if it contains the name (already extracted)
                        if (profile.name && text.includes(profile.name)) continue;

                        // Skip connection degree indicators and social proof
                        if (text.match(/^[•\\s]*[123](?:st|nd|rd)$/)) continue;
                        if (text.includes('followers')) continue;
                        if (text.includes('mutual connection')) continue;

                        // First non-name paragraph is job title
                        if (!jobTitleFound) {
                            profile.job_title = text;
                            jobTitleFound = true;
                            continue;
                        }

                        // Second is location
                        if (!locationFound) {
                            profile.location = text;
                            locationFound = true;
                            break;
                        }
                    }

                    // Extract company from search filter pill
                    const companySelectors = [
                        'button[id="searchFilter_currentCompany"]',
                        '.artdeco-pill[aria-label*="Current company"]',
                        '.search-reusables__filter-pill-button[aria-label*="company"]'
                    ];

                    for (const selector of companySelectors) {
                        const element = document.querySelector(selector);
                        if (element) {
                            const text = element.textContent?.trim();
                            if (text) {
                                profile.company = text.replace(/\\s+\\d+$/, '');
                                break;
                            }
                        }
                    }

                    // follows_from is always null
                    profile.follows_from = null;

                    // Only add profile if we have at least a name or URL
                    if (profile.name || profile.url_profile) {
                        profiles.push(profile);
                    }

                } catch (error) {
                    console.log('Error extracting profile:', error);
                }
            }

            return profiles;
        })();
        """

        try:
            profiles_data = self.chrome.evaluate_javascript(profiles_js)
            if profiles_data and isinstance(profiles_data, list):
                # Convert any None values to empty strings for consistency
                for profile in profiles_data:
                    for key in profile:
                        if profile[key] is None:
                            profile[key] = ""
                logger.info(f"📊 Extracted {len(profiles_data)} profiles from search results")
                return profiles_data
            else:
                logger.warning("⚠️  No profiles found or invalid response from search results extraction")
                return []
        except Exception as e:
            logger.error(f"❌ Failed to extract profiles from search results: {e}")
            return []

    def extract_all_tabs_devtools(self, max_tabs: int = 50, exit_check=None) -> Tuple[List[Dict[str, str]], int]:
        """Extract profile data from all Chrome tabs using DevTools.

        Iterates directly through valid LinkedIn search tabs via DevTools API,
        avoiding hidden tracking/analytics pages.

        Args:
            max_tabs: Maximum number of tabs to process.
            exit_check: Optional callable that returns True if extraction should exit.

        Returns:
            Tuple of (list of profile dictionaries, actual tabs processed)
        """
        logger.info("🚀 Starting DevTools extraction from all Chrome tabs")

        # Get all valid LinkedIn search tabs directly from DevTools
        valid_tabs = self.chrome.get_linkedin_search_tabs()
        if not valid_tabs:
            logger.error("❌ No valid LinkedIn search tabs found")
            return [], 0

        logger.info(f"📑 Found {len(valid_tabs)} valid LinkedIn search tabs")

        # Limit to max_tabs
        tabs_to_process = valid_tabs[:max_tabs]

        all_profiles = []
        content_hashes = set()  # Track content to detect duplicates
        actual_tabs_processed = 0

        try:
            for tab_num, tab in enumerate(tabs_to_process):
                # Check if exit was requested
                if exit_check and exit_check():
                    logger.info("⏹️  Extraction stopped by user (exit requested)")
                    break

                actual_tabs_processed = tab_num + 1
                logger.info(f"📄 Processing tab {actual_tabs_processed}/{len(tabs_to_process)}")

                # Close previous connection if any
                self.chrome.close()

                # Connect directly to this tab
                if not self.chrome.connect_to_tab(tab):
                    logger.warning(f"⚠️  Could not connect to tab {tab_num + 1} - skipping")
                    continue

                # Get content hash to detect duplicates
                content_hash_js = """
                (function() {
                    const results = document.querySelectorAll('div[data-view-name="people-search-result"]');
                    let content = '';
                    for (const result of results) {
                        content += result.textContent || '';
                    }
                    let hash = 0;
                    for (let i = 0; i < content.length; i++) {
                        const char = content.charCodeAt(i);
                        hash = ((hash << 5) - hash) + char;
                        hash = hash & hash;
                    }
                    return Math.abs(hash).toString();
                })();
                """
                content_hash = self.chrome.evaluate_javascript(content_hash_js)

                if content_hash in content_hashes:
                    logger.info(f"🔄 Duplicate content detected - skipping tab")
                    continue

                # Extract profiles from this tab
                profiles_data = self.extract_search_results_profiles()

                if not profiles_data:
                    logger.warning(f"⚠️  No profiles extracted from tab {tab_num + 1}")
                    continue

                # Track content hash
                content_hashes.add(content_hash)

                # Add all extracted profiles
                for profile_data in profiles_data:
                    all_profiles.append(profile_data)
                    logger.info(f"✅ Extracted: {profile_data.get('name', 'Unknown')} - {profile_data.get('job_title', 'Unknown')}")

                logger.info(f"📊 Extracted {len(profiles_data)} profiles from tab {actual_tabs_processed}")

            logger.info(f"✅ DevTools extraction completed: {len(all_profiles)} profiles extracted from {actual_tabs_processed} tabs")
            return all_profiles, actual_tabs_processed

        finally:
            self.chrome.close()

    def save_to_excel(profiles: List[Dict[str, str]], output_file: str) -> None:
        """Save profiles to Excel file.

        Args:
            profiles: List of profile dictionaries.
            output_file: Path to output Excel file.
        """
        if not profiles:
            logger.warning("⚠️  No profiles to save")
            return

        try:
            df = pd.DataFrame(profiles)

            # Convert empty strings to NaN for better Excel handling
            df = df.replace('', pd.NA)

            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Profiles', index=False)

            logger.info(f"✅ Saved {len(profiles)} profiles to {output_file}")

        except Exception as e:
            logger.error(f"❌ Failed to save to Excel: {e}")
            raise


def main() -> None:
    """Main execution function for testing."""
    extractor = LinkedInProfileExtractorDevTools()

    if not extractor.connect_to_chrome():
        logger.error("❌ Could not connect to Chrome DevTools")
        return

    try:
        # Test single profile extraction
        profile = extractor.extract_profile_data()
        logger.info(f"📊 Extracted profile: {profile}")

    finally:
        extractor.chrome.close()


if __name__ == "__main__":
    main()
