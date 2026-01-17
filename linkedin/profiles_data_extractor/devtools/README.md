# DevTools Version

Chrome DevTools Protocol-based LinkedIn profile data extraction for reliable, structured data collection.

## 📋 Summary

The DevTools extraction method leverages Chrome's debugging protocol to directly access and extract LinkedIn profile data from the DOM. This approach provides the most reliable and maintainable solution, eliminating the need for clipboard operations and keyboard simulation while offering direct structured data access.

## Overview

Uses Chrome's DevTools Protocol for reliable DOM access instead of keyboard simulation and text parsing. More robust and maintainable than clipboard-based methods. Now with auto-confirmation for unattended operation.

## Key Features

- **Direct DOM Access**: JavaScript selectors query LinkedIn's HTML structure
- **Smart Tab Filtering**: Automatically skips hidden tracking pages (merchantpool, analytics)
- **Command-Line Interface**: Run with parameters instead of interactive prompts
- **Direct API Iteration**: Connects to tabs via DevTools API (no keyboard simulation)
- **Deduplication**: Automatically skips profiles already in your Excel file
- **Updated Selectors**: Works with LinkedIn's 2026 HTML structure (`people-search-result`)
- **Real-time Extraction**: Structured data extraction with WebSocket communication

## Prerequisites

- Chrome with remote debugging enabled (port 9222)
- `websocket-client` package

## Quick Start

1. **Start Chrome in Debug Mode**:
   ```bash
   start_chrome_debug.bat
   ```
   This launches Chrome with remote debugging on port 9222.

2. **Open LinkedIn Search Pages**:
   - Navigate to LinkedIn and perform your searches
   - Open multiple search result pages in separate Chrome tabs
   - The extractor will automatically process all valid LinkedIn search tabs

3. **Run the Orchestrator** (Recommended):
   ```bash
   linkedin_profiles_data_orchestrator_devtools.bat
   ```
   - Processes all tabs automatically
   - Merges data into Excel with deduplication
   - Applies formatting and hyperlinks

   **Command-Line Options:**
   ```bash
   python linkedin_profiles_data_orchestrator_devtools.py --run              # Run full extraction
   python linkedin_profiles_data_orchestrator_devtools.py --run --save-format # Save format and run
   python linkedin_profiles_data_orchestrator_devtools.py --test               # Quick test
   ```

   **OR** Run the standalone extractor:
   ```bash
   linkedin_profiles_data_extractor_devtools.bat
   ```

## How It Works

1. **Tab Discovery**: Connects to Chrome DevTools API and retrieves all open tabs
2. **Smart Filtering**: Filters out non-LinkedIn pages and tracking/analytics pages
3. **Data Extraction**: For each valid tab, extracts profile data using updated selectors:
   - `data-view-name="people-search-result"` for profile containers
   - `data-view-name="search-result-lockup-title"` for names
   - Sequential `<p>` tags for job title and location
4. **Deduplication**: Compares with existing Excel data (by name, case-insensitive)
5. **Excel Merge**: Appends only new profiles to your Excel file
6. **Formatting**: Applies saved formatting and converts URLs to hyperlinks

## Files

- `linkedin_profiles_data_extractor_devtools.py` - Core DevTools data extractor with updated selectors
- `linkedin_profiles_data_orchestrator_devtools.py` - Orchestrates extraction workflow with auto-confirmation
- `linkedin_profile_search_checker_devtools.py` - Search result validation
- `start_chrome_debug.bat` - Chrome debugging launcher

## Recent Updates (January 2026)

### LinkedIn HTML Structure Changes
- Updated selectors to match LinkedIn's new HTML structure
- Changed from `search-entity-result-universal-template` to `people-search-result`
- Updated name extraction to use `search-result-lockup-title` attribute
- Simplified job title and location extraction using sequential `<p>` tags

### Improved Reliability
- **Smart Tab Filtering**: Automatically excludes hidden tracking pages:
  - `merchantpool*.linkedin.com` (analytics)
  - `beacon*.linkedin.com` (tracking)
  - `platform.linkedin.com` (SDK)
  - Other non-page resources (service workers, background pages)
- **Direct API Access**: Iterates through tabs via DevTools API instead of keyboard shortcuts
- **No Manual Switching**: Removed 5-second wait and Chrome activation (connects directly to tabs)

### Automation Features
- **Command-Line Interface**: No interactive prompts - use `--run` and `--save-format` flags
- **Streamlit Integration**: Buttons in the Streamlit app for "Run extractor and merger" and "Save JSON format and run"
- **Simplified Workflow**: Removed timing inputs and interactive loops for cleaner automation

## Troubleshooting

### No Profiles Extracted
- **Check HTML Structure**: LinkedIn may have updated their HTML. Inspect the page and verify:
  - Profile containers use `data-view-name="people-search-result"`
  - Name links use `data-view-name="search-result-lockup-title"`
- **Verify Search Results**: Make sure you're on a LinkedIn search results page, not a profile page
- **Check Console**: Look for JavaScript errors in the extraction logs

### Hidden Tabs Processed
- The extractor now automatically filters out tracking pages
- Only processes tabs with `type: 'page'` and valid LinkedIn URLs
- Excludes: merchantpool, beacon, tracking, analytics, platform.linkedin.com

### Excel File Locked
- Close Excel before running the extractor
- The script will automatically retry (up to 30 times with 1 second delay) if the file is locked

### Chrome Connection Failed
- Ensure Chrome is running with `--remote-debugging-port=9222`
- Check that no other application is using port 9222
- Try closing and restarting Chrome in debug mode
