# LinkedIn Profiles Data Extractor

Automated LinkedIn profile data extraction with multiple implementation approaches.

## Project Structure

### 📁 `common/`
Shared utilities, dashboard, and configuration files used by all extraction methods.

- **Dashboard**: Streamlit visualization (`dashboard.bat`)
- **Excel Tools**: Formatting and hyperlink management
- **Configuration**: Extraction settings and paths

### 📁 `devtools/`
Chrome DevTools Protocol-based extraction (recommended).

- **Pros**: Reliable, structured data access, no clipboard operations
- **Cons**: Requires Chrome debugging setup
- **Usage**: `start_chrome_debug.bat` → extraction scripts

### 📁 `brute_force/`
Keyboard simulation and clipboard-based extraction.

- **Pros**: Simple setup, works with any browser
- **Cons**: Requires window focus, timing-dependent
- **Usage**: Direct batch file execution

## Quick Start

### DevTools Version (Recommended)
```bash
cd devtools
start_chrome_debug.bat
linkedin_profiles_data_orchestrator_devtools.bat
```

### Brute Force Version
```bash
cd brute_force
linkedin_profiles_data.bat
```

### Dashboard
```bash
cd common
dashboard.bat
```

## Configuration

Edit `common/linkedin_profiles_data.json` to configure:
- Output Excel file paths
- Tab management settings
- Wait times and delays

## Dependencies

- Python packages: pandas, pynput, requests, websocket-client
- Chrome browser (for DevTools version)
- Excel for data output

## Architecture Notes

- Relative imports maintain modularity across versions
- Shared `excel_format_manager` handles consistent formatting
- Each version can run independently
- Configuration centralized in `common/` folder
