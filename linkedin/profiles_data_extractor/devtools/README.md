# DevTools Version

Chrome DevTools Protocol-based LinkedIn profile data extraction.

## Overview

Uses Chrome's DevTools Protocol for reliable DOM access instead of keyboard simulation and text parsing. More robust and maintainable than clipboard-based methods.

## Key Features

- Direct DOM element querying with JavaScript selectors
- No clipboard operations required
- Real-time structured data extraction
- WebSocket communication with Chrome

## Prerequisites

- Chrome with remote debugging enabled (port 9222)
- `websocket-client` package

## Setup

1. **Start Chrome Debug Mode**:
   ```bash
   start_chrome_debug.bat
   ```

2. **Run Extractor**:
   ```bash
   linkedin_profiles_data_extractor_devtools.bat
   ```

3. **Run Orchestrator**:
   ```bash
   linkedin_profiles_data_orchestrator_devtools.bat
   ```

## Files

- `linkedin_profiles_data_extractor_devtools.py` - DevTools data extractor
- `linkedin_profiles_data_orchestrator_devtools.py` - Orchestrates extraction workflow
- `linkedin_profile_search_checker_devtools.py` - Search result validation
- `start_chrome_debug.bat` - Chrome debugging launcher

See `README_DevTools.md` for detailed usage instructions.
