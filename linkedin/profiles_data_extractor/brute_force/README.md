# Brute Force Version

Main copy-paste workflow for LinkedIn profile data extraction.

## Overview

Uses keyboard simulation and clipboard operations to extract profile data from LinkedIn search results. Simple but requires careful window focus management.

## Key Features

- Keyboard shortcuts (Ctrl+A, Ctrl+C) for content copying
- Regex pattern matching for data extraction
- Real-time Excel file merging
- Keyboard interrupt handling (press 'x' to stop)

## Usage

### Main Orchestrator
```bash
linkedin_profiles_data.bat
```
Runs continuous extraction loop with automatic Excel merging.

### Search Checker
```bash
linkedin_profile_search_checker.bat
```
Validates search results against existing contacts database.

## Files

- `linkedin_profiles_data_extractor.py` - Core extraction logic
- `linkedin_profiles_data_orchestrator.py` - Workflow orchestration
- `linkedin_profile_search_checker.py` - Search result validation

## Workflow

1. Open LinkedIn search results in browser
2. Run orchestrator - it will cycle through tabs automatically
3. Press 'x' key to stop extraction
4. Data is merged into configured Excel file

## Notes

- Requires browser window focus for keyboard simulation
- May need timing adjustments for different systems
- Less reliable than DevTools version but simpler setup
