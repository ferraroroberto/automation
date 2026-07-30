"""
Shared Chrome DevTools Protocol client + window-activation helpers for the
DevTools-based LinkedIn scripts in this directory.

All three devtools scripts (``linkedin_profile_search_checker_devtools``,
``linkedin_profiles_data_extractor_devtools``,
``linkedin_profiles_data_orchestrator_devtools``) import ``setup_logging``,
``ChromeDevToolsClient``, ``find_chrome_window``, and ``activate_chrome`` from
here so the connection/window-focus plumbing and logging bootstrap each live
in exactly one place instead of being copy-pasted per script.
"""

import json
import logging
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests
import websocket
import win32con
import win32gui


def setup_logging(name: str) -> logging.Logger:
    """Configure a logger that writes INFO+ to the package's shared logging.log
    and WARNING+ to the console, matching the convention used across this
    package's other modules.
    """
    log_file = Path(__file__).parent.parent / "logging.log"
    log_file.parent.mkdir(parents=True, exist_ok=True)

    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.WARNING)
    console_handler.setFormatter(logging.Formatter('%(levelname)s - %(message)s'))

    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    return logger


logger = setup_logging(__name__)


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


def find_chrome_window() -> Optional[int]:
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


def activate_chrome() -> bool:
    """Find and activate Chrome window.

    Returns:
        True if Chrome was found and activated, False otherwise.
    """
    hwnd = find_chrome_window()
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
