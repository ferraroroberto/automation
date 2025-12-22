"""
Chrome Follow Date Extractor

Extracts the "following you since" date and URL from the active Chrome tab.
Automatically finds Chrome, activates it, and extracts the content and URL.
"""

import logging
import re
import time
from typing import Optional

import pyperclip
import win32gui
import win32con
from pynput import keyboard

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


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


def copy_url() -> str:
    """Copy the URL from Chrome address bar using keyboard shortcuts."""
    try:
        controller = keyboard.Controller()
        time.sleep(0.3)
        
        # CTRL+L to focus address bar
        controller.press(keyboard.Key.ctrl)
        controller.press('l')
        controller.release('l')
        controller.release(keyboard.Key.ctrl)
        time.sleep(0.5)
        
        # CTRL+C to copy URL
        controller.press(keyboard.Key.ctrl)
        controller.press('c')
        controller.release('c')
        controller.release(keyboard.Key.ctrl)
        time.sleep(0.5)
        
        url = pyperclip.paste()
        if url:
            logger.info(f"✅ URL copied: {url}")
        else:
            logger.warning("⚠️  Clipboard is empty after copying URL")
        
        return url if url else ""
    except Exception as e:
        logger.error(f"❌ Failed to copy URL: {e}")
        return ""


def copy_all_content() -> str:
    """Copy all content from the current page using CTRL+A and CTRL+C."""
    try:
        controller = keyboard.Controller()
        time.sleep(0.5)
        
        # Move focus from address bar to page content
        controller.press(keyboard.Key.tab)
        controller.release(keyboard.Key.tab)
        time.sleep(0.3)
        
        # CTRL+A to select all
        controller.press(keyboard.Key.ctrl)
        controller.press('a')
        controller.release('a')
        controller.release(keyboard.Key.ctrl)
        time.sleep(0.5)
        
        # CTRL+C to copy
        controller.press(keyboard.Key.ctrl)
        controller.press('c')
        controller.release('c')
        controller.release(keyboard.Key.ctrl)
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


def extract_follow_date(text: str) -> Optional[str]:
    """Extract the date from text following 'following you since'."""
    if not text:
        return None
    
    pattern = r'following you since\s+([^\n\r]+)'
    match = re.search(pattern, text, re.IGNORECASE)
    
    if match:
        date_text = match.group(1).strip()
        logger.info(f"✅ Found date text: {date_text}")
        return date_text
    
    logger.warning("⚠️  'following you since' pattern not found in text")
    return None


def main() -> None:
    """Main execution function."""
    logger.info("🚀 Chrome Follow Date Extractor started")
    
    # Automatically find and activate Chrome window
    logger.info("🔍 Looking for Chrome window...")
    if not activate_chrome():
        logger.error("❌ Could not activate Chrome. Please make sure Chrome is open.")
        return
    
    # Copy all page content first (before getting URL, which focuses address bar)
    logger.info("📋 Copying all page content...")
    content = copy_all_content()
    
    if not content:
        logger.error("❌ Failed to copy content. Please try again.")
        return
    
    # Get URL after copying content (address bar focus won't interfere now)
    logger.info("📋 Copying URL from address bar...")
    url = copy_url()
    
    if not url:
        logger.warning("⚠️  Failed to get URL, but continuing with content extraction...")
        url = ""
    
    # Extract date
    logger.info("🔍 Searching for 'following you since' date...")
    date_text = extract_follow_date(content)
    
    # Display results
    print("\n" + "="*60)
    print("📊 EXTRACTION RESULTS")
    print("="*60)
    print(f"🔗 URL: {url}")
    if date_text:
        print(f"📅 Date: {date_text}")
    else:
        print("📅 Date: Not found")
    print("="*60 + "\n")


if __name__ == "__main__":
    main()

