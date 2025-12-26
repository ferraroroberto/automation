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
        """Extract profile data from the current LinkedIn profile page using DevTools.

        Returns:
            Dictionary containing profile data.
        """
        # Extract name
        name_js = """
        (function() {
            // Primary selector for profile name
            const nameSelectors = [
                'h1[data-test-id="hero__page__title"]',
                '.pv-text-details__left-panel h1',
                'h1.text-heading-xlarge'
            ];

            for (const selector of nameSelectors) {
                const element = document.querySelector(selector);
                if (element) {
                    return element.textContent?.trim();
                }
            }

            // Fallback: look for name in the specific span structure from search results
            const profileSpans = document.querySelectorAll('span[dir="ltr"] span[aria-hidden="true"]');
            for (const span of profileSpans) {
                const name = span.textContent?.trim();
                if (name && name.split(' ').length >= 2) {
                    return name;
                }
            }

            return null;
        })();
        """
        name = self.chrome.evaluate_javascript(name_js)

        # Extract URL from profile link
        url_js = """
        (function() {
            // Get the profile URL from the link's href attribute
            const profileLinks = document.querySelectorAll('a[href*="linkedin.com/in/"]');
            for (const link of profileLinks) {
                const href = link.getAttribute('href');
                if (href && href.includes('linkedin.com/in/')) {
                    // Clean up the URL - remove query parameters after profile name if needed
                    // But keep the miniProfileUrn parameter as it contains the profile ID
                    return href;
                }
            }

            return null;
        })();
        """
        url = self.chrome.evaluate_javascript(url_js)

        # Extract job title
        job_title_js = """
        (function() {
            // Strategy 1: Search Results List (Robust)
            // Iterate over each profile card to find the first valid job title
            const searchResults = document.querySelectorAll('div[data-view-name="search-entity-result-universal-template"]');
            for (const result of searchResults) {
                // Orient based on entity-result__insights
                const insights = result.querySelector('.entity-result__insights');
                if (insights) {
                    // The profile info is in the previous sibling container
                    const profileContainer = insights.previousElementSibling;
                    if (profileContainer) {
                        // Look for the job title div.
                        // Based on analysis:
                        // - Job Title: t-14 t-black t-normal
                        
                        // Iterate over direct children of the profile container
                        for (const child of profileContainer.children) {
                            if (child.classList.contains('t-14') && 
                                child.classList.contains('t-black') && 
                                child.classList.contains('t-normal')) {
                                
                                const text = child.textContent?.trim();
                                if (text) return text;
                            }
                        }
                    }
                }
            }

            // Fallback to profile page selectors
            const profileSelectors = [
                '.pv-text-details__left-panel .text-body-medium',
                '.pv-text-details__left-panel div[data-test-id="profile-card__primary-headline"]'
            ];

            for (const selector of profileSelectors) {
                const element = document.querySelector(selector);
                if (element) {
                    return element.textContent?.trim();
                }
            }

            return null;
        })();
        """
        job_title = self.chrome.evaluate_javascript(job_title_js)

        # Extract location
        location_js = """
        (function() {
            // Strategy 1: Search Results List (Robust)
            // Iterate over each profile card to find the first valid location
            // We use the data-view-name to identify the profile card container
            const searchResults = document.querySelectorAll('div[data-view-name="search-entity-result-universal-template"]');
            for (const result of searchResults) {
                // Orient based on entity-result__insights as requested
                const insights = result.querySelector('.entity-result__insights');
                if (insights) {
                    // The profile info is in the previous sibling container
                    const profileContainer = insights.previousElementSibling;
                    if (profileContainer) {
                        // Look for the location div.
                        // Based on analysis:
                        // - Job Title: t-14 t-black t-normal
                        // - Location: t-14 t-normal (and usually NOT t-black)
                        
                        // Iterate over direct children of the profile container
                        for (const child of profileContainer.children) {
                            // Check for t-14 and t-normal
                            if (child.classList.contains('t-14') && child.classList.contains('t-normal')) {
                                // Exclude Job Title (which has t-black)
                                if (!child.classList.contains('t-black')) {
                                    const text = child.textContent?.trim();
                                    if (text) return text;
                                }
                            }
                        }
                    }
                }
            }

            // Fallback to profile page selectors
            const profileSelectors = [
                '.pv-text-details__left-panel .text-body-small.inline.t-black--light.break-words',
                '.pv-text-details__left-panel span[data-test-id="profile-card__location"]',
                '.VlSKoXSzfRCLBjvFdhXHtrnCCaNHklv span.text-body-small'
            ];

            for (const selector of profileSelectors) {
                const element = document.querySelector(selector);
                if (element) {
                    return element.textContent?.trim();
                }
            }

            return null;
        })();
        """
        location = self.chrome.evaluate_javascript(location_js)

        # Extract company from search filter pill (as shown in user's example)
        company_js = """
        (function() {
            const companySelectors = [
                'button[id="searchFilter_currentCompany"]',
                '.artdeco-pill[aria-label*="Current company"]',
                '.search-reusables__filter-pill-button[aria-label*="company"]'
            ];

            for (const selector of companySelectors) {
                const element = document.querySelector(selector);
                if (element) {
                    // Look for the company name, excluding the count
                    const text = element.textContent?.trim();
                    if (text) {
                        // Remove the count (e.g., "HP 1" -> "HP")
                        return text.replace(/\s+\d+$/, '');
                    }
                }
            }

            return null;
        })();
        """
        company = self.chrome.evaluate_javascript(company_js)

        # follows_from is always null as requested
        follows_from = None

        # Clean URL by removing query parameters after ?
        if url:
            url = url.split('?')[0]

        return {
            'name': name or "",
            'url': url,
            'job_title': job_title or "",
            'follows_from': follows_from,
            'company': company or "",
            'location': location or ""
        }

    def extract_search_results_profiles(self) -> List[Dict[str, str]]:
        """Extract multiple profiles from a LinkedIn search results page using DevTools.

        Returns:
            List of dictionaries containing profile data from all search results.
        """

        profiles_js = """
        (function() {
            const profiles = [];

            // Find all search result containers
            const searchResults = document.querySelectorAll('div[data-view-name="search-entity-result-universal-template"]');

            for (let i = 0; i < searchResults.length; i++) {
                const result = searchResults[i];
                try {
                    const profile = {};

                    // Extract name - look for spans with aria-hidden="true" within this result
                    const nameSpans = result.querySelectorAll('span[aria-hidden="true"]');
                    for (const span of nameSpans) {
                        const name = span.textContent?.trim();
                        if (name && name.split(' ').length >= 2) {
                            profile.name = name;
                            break;
                        }
                    }

                    // Extract URL from LinkedIn profile links within this result
                    const profileLinks = result.querySelectorAll('a[href*="linkedin.com/in/"]');
                    for (const link of profileLinks) {
                        const href = link.getAttribute('href');
                        if (href && href.includes('linkedin.com/in/')) {
                            // Clean URL by removing query parameters after ?
                            profile.url = href.split('?')[0];
                            break;
                        }
                    }

                    // Extract job title and company - look for text that contains job info
                    // Try to find text that looks like job titles (contains common job keywords or "at")
                    const allTextElements = result.querySelectorAll('span, div, p');
                    let jobInfo = '';

                    for (const element of allTextElements) {
                        const text = element.textContent?.trim() || '';
                        // Look for patterns that suggest job information
                        if ((text.includes(' at ') || text.includes(' · ') ||
                             /\b(manager|director|engineer|developer|specialist|analyst|consultant|lead|senior|vp|chief)\b/i.test(text)) &&
                            text.length > 5 && text.length < 100) {
                            jobInfo = text;
                            break;
                        }
                    }

                    if (jobInfo) {
                        // Parse job title and company
                        const atIndex = jobInfo.indexOf(' at ');
                        const dotIndex = jobInfo.indexOf(' · ');

                        if (atIndex !== -1) {
                            profile.job_title = jobInfo.substring(0, atIndex).trim();
                            profile.company = jobInfo.substring(atIndex + 4).trim();
                        } else if (dotIndex !== -1) {
                            profile.job_title = jobInfo.substring(0, dotIndex).trim();
                            profile.company = jobInfo.substring(dotIndex + 3).trim();
                        } else {
                            // If no separator found, use heuristics
                            const words = jobInfo.split(' ');
                            if (words.length > 3) {
                                profile.job_title = words.slice(0, -1).join(' ');
                                profile.company = words[words.length - 1];
                            } else {
                                profile.job_title = jobInfo;
                                profile.company = '';
                            }
                        }
                    }

                    // Extract location - look for location-like text
                    let locationInfo = '';
                    for (const element of allTextElements) {
                        const text = element.textContent?.trim() || '';
                        // Location patterns: contains geographic indicators or is short and not job-related
                        if ((text.includes(',') || /\b(area|region|city|province|state|country)\b/i.test(text) ||
                             /^[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*$/.test(text)) &&
                            text.length > 2 && text.length < 50 && !text.includes(' at ') && !text.includes(' · ')) {
                            locationInfo = text;
                            break;
                        }
                    }

                    if (locationInfo) {
                        profile.location = locationInfo;
                    }

                    // follows_from is always null
                    profile.follows_from = null;

                    // Only add profile if we have at least a name or URL
                    if (profile.name || profile.url) {
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

    def switch_to_next_tab_devtools(self) -> None:
        """Switch to the next tab using keyboard simulation (Ctrl+Tab)."""

        logger.info("🔄 Switching to next tab...")
        self.keyboard_controller.press(keyboard.Key.ctrl)
        self.keyboard_controller.press(keyboard.Key.tab)
        self.keyboard_controller.release(keyboard.Key.tab)
        self.keyboard_controller.release(keyboard.Key.ctrl)
        time.sleep(1.5)  # Wait for tab switch to complete

    def extract_all_tabs_devtools(self, max_tabs: int = 50, exit_check=None) -> Tuple[List[Dict[str, str]], int]:
        """Extract profile data from all Chrome tabs using DevTools.

        Args:
            max_tabs: Maximum number of tabs to process.
            exit_check: Optional callable that returns True if extraction should exit.

        Returns:
            Tuple of (list of profile dictionaries, actual tabs processed)
        """
        logger.info("🚀 Starting DevTools extraction from all Chrome tabs")

        # Wait for user to switch to Chrome
        logger.info("⏳ Waiting 5 seconds for you to switch to Chrome window...")
        time.sleep(5)

        # Activate Chrome window
        if not self.activate_chrome():
            logger.error("❌ Could not activate Chrome. Please make sure Chrome is open.")
            return [], 0

        # Connect to Chrome DevTools
        if not self.connect_to_chrome():
            logger.error("❌ Could not connect to Chrome DevTools")
            return [], 0

        try:
            all_profiles = []
            processed_urls = set()  # Track processed URLs to avoid duplicates
            content_hashes = set()  # Track content to detect cycling
            previous_content_hash = None
            consecutive_no_progress = 0
            actual_tabs_processed = 0

            for tab_num in range(max_tabs):
                # Check if exit was requested
                if exit_check and exit_check():
                    logger.info("⏹️  Extraction stopped by user (exit requested)")
                    break

                actual_tabs_processed = tab_num + 1
                logger.info(f"📄 Processing tab {actual_tabs_processed} (max: {max_tabs})")

                # Get current URL first
                current_url = self.chrome.get_current_url()

                if not current_url:
                    logger.warning(f"⚠️  Could not get URL from tab {tab_num + 1}")
                    consecutive_no_progress += 1
                    if consecutive_no_progress >= 2:
                        logger.info("⏹️  No URL found for 2 consecutive tabs - stopping")
                        break
                    if tab_num < max_tabs - 1:
                        self.switch_to_next_tab_devtools()
                        # Reconnect to the new active tab
                        self.chrome.close()  # Close current connection
                        time.sleep(0.5)  # Brief pause for tab switch to settle
                        if not self.chrome.connect_to_chrome():  # Reconnect to new active tab
                            logger.warning("⚠️  Could not reconnect to new tab - stopping")
                            break
                    continue

                # Check if this is a LinkedIn page (search results or profile)
                if 'linkedin.com' not in current_url:
                    logger.warning(f"⚠️  Not a LinkedIn page: {current_url}")
                    consecutive_no_progress += 1
                    if consecutive_no_progress >= 2:
                        logger.info("⏹️  Non-LinkedIn pages detected - stopping")
                        break
                    if tab_num < max_tabs - 1:
                        self.switch_to_next_tab_devtools()
                        # Reconnect to the new active tab
                        self.chrome.close()  # Close current connection
                        time.sleep(0.5)  # Brief pause for tab switch to settle
                        if not self.chrome.connect_to_chrome():  # Reconnect to new active tab
                            logger.warning("⚠️  Could not reconnect to new tab - stopping")
                            break
                    continue

                # Check for duplicate URLs (but allow processing same URL if content is different)
                # We'll use content hash to detect actual duplicates
                content_hash_js = """
                (function() {
                    // Get a hash of the search results content
                    const results = document.querySelectorAll('div[data-view-name="search-entity-result-universal-template"]');
                    let content = '';
                    for (const result of results) {
                        content += result.textContent || '';
                    }
                    // Simple hash function
                    let hash = 0;
                    for (let i = 0; i < content.length; i++) {
                        const char = content.charCodeAt(i);
                        hash = ((hash << 5) - hash) + char;
                        hash = hash & hash; // Convert to 32-bit integer
                    }
                    return Math.abs(hash).toString();
                })();
                """
                content_hash = self.chrome.evaluate_javascript(content_hash_js)

                if content_hash in content_hashes:
                    logger.info(f"🔄 Content already processed (hash: {content_hash}) - cycled back to start, stopping extraction")
                    break

                # Extract multiple profiles from search results
                profiles_data = self.extract_search_results_profiles()

                # Validate that we got some data
                if not profiles_data:
                    logger.warning(f"⚠️  No profiles extracted from tab {tab_num + 1}")
                    consecutive_no_progress += 1
                    if consecutive_no_progress >= 2:
                        logger.info("⏹️  No profiles found for 2 consecutive tabs - stopping")
                        break
                else:
                    consecutive_no_progress = 0
                    processed_urls.add(current_url)
                    content_hashes.add(content_hash)

                    # Add all extracted profiles
                    for profile_data in profiles_data:
                        all_profiles.append(profile_data)
                        logger.info(f"✅ Extracted: {profile_data.get('name', 'Unknown')} - {profile_data.get('job_title', 'Unknown')}")

                    logger.info(f"📊 Extracted {len(profiles_data)} profiles from tab {actual_tabs_processed}")

                # Switch to next tab
                if tab_num < max_tabs - 1:
                    self.switch_to_next_tab_devtools()
                    # Reconnect to the new active tab
                    self.chrome.close()  # Close current connection
                    time.sleep(0.5)  # Brief pause for tab switch to settle
                    if not self.chrome.connect_to_chrome():  # Reconnect to new active tab
                        logger.warning("⚠️  Could not reconnect to new tab - stopping")
                        break

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
