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


def extract_name(text: str) -> Optional[str]:
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


def extract_job_title(text: str, name: Optional[str] = None) -> Optional[str]:
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


def extract_company(text: str, job_title: Optional[str] = None) -> Optional[str]:
    """Extract the company name appearing just after the job title."""
    if not text:
        return None
    
    if job_title:
        # Find text after job title, before location or contact info
        # Escape job title but allow flexible matching
        job_title_escaped = re.escape(job_title)
        pattern = rf'{job_title_escaped}\s+([A-Z][a-zA-Z\s&,\.\-]{2,150}?)(?:\s+(?:Metropolitan Area|Area|City|County|State|Country|contact info|Contact info|CONTACT INFO)|$)'
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
    pattern = r'\d+(?:st|nd|rd|th)?\s*degree\s*connection[^\n\r]*?\s+([A-Z][a-zA-Z\s&,\.\-]{2,150}?)(?:\s+(?:Metropolitan Area|Area|City|County|State|Country|contact info|Contact info|CONTACT INFO)|$)'
    match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE | re.DOTALL)
    if match:
        company = match.group(1).strip()
        company = re.sub(r'\s+(?:Metropolitan Area|Area|City|County|State|Country|contact info|Contact info|CONTACT INFO).*$', '', company, flags=re.IGNORECASE)
        if company and len(company) < 200 and not re.search(r'\d+(?:st|nd|rd|th)\s*degree', company, re.IGNORECASE):
            logger.info(f"✅ Found company: {company}")
            return company
    
    logger.warning("⚠️  Company pattern not found in text")
    return None


def extract_location(text: str, company: Optional[str] = None) -> Optional[str]:
    """Extract the location appearing after the company and before 'contact info'."""
    if not text:
        return None
    
    if company:
        # Find text after company, before contact info
        company_escaped = re.escape(company)
        pattern = rf'{company_escaped}\s+([A-Z][a-zA-Z\s,\.]+?)(?:\s+contact info|Contact info|CONTACT INFO|$)'
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE | re.DOTALL)
        if match:
            location = match.group(1).strip()
            # Clean up if it contains contact info
            location = re.sub(r'\s+contact info.*$', '', location, flags=re.IGNORECASE)
            if location and len(location) < 150:  # Reasonable location length
                logger.info(f"✅ Found location: {location}")
                return location
    
    # Fallback: find location pattern before contact info (look for common location patterns)
    pattern = r'([A-Z][a-zA-Z\s,\.]+?(?:Metropolitan Area|Area|City|County|State|Country|Region))(?:\s+contact info|Contact info|CONTACT INFO|$)'
    match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE | re.DOTALL)
    if match:
        location = match.group(1).strip()
        location = re.sub(r'\s+contact info.*$', '', location, flags=re.IGNORECASE)
        if location and len(location) < 150:
            logger.info(f"✅ Found location: {location}")
            return location
    
    logger.warning("⚠️  Location pattern not found in text")
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
    
    # Extract all fields
    logger.info("🔍 Searching for profile information...")
    name = extract_name(content)
    job_title = extract_job_title(content, name)
    company = extract_company(content, job_title)
    location = extract_location(content, company)
    date_text = extract_follow_date(content)
    
    # Display results
    print("\n" + "="*60)
    print("EXTRACTION RESULTS")
    print("="*60)
    print(f"URL: {url}")
    if name:
        print(f"Name: {name}")
    else:
        print("Name: Not found")
    if job_title:
        print(f"Job Title: {job_title}")
    else:
        print("Job Title: Not found")
    if company:
        print(f"Company: {company}")
    else:
        print("Company: Not found")
    if location:
        print(f"Location: {location}")
    else:
        print("Location: Not found")
    if date_text:
        print(f"Date: {date_text}")
    else:
        print("Date: Not found")
    print("="*60 + "\n")


if __name__ == "__main__":
    main()

