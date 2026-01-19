# LinkedIn Reachout Hub - Common Components

Shared utilities, intelligent data entry interface, and visualization dashboard for LinkedIn profile data extraction and management workflows.

## 📋 Summary

The common components provide the core infrastructure and user interface for the LinkedIn Profiles Data Extractor. This includes a powerful Streamlit dashboard for data visualization, an intelligent data entry system with fuzzy search capabilities, Excel formatting utilities, and centralized configuration management. These shared components ensure consistency across all extraction methods while providing a user-friendly interface for data management and analysis.

## 📁 Files Overview

| File | Purpose |
|------|---------|
| `dashboard.py` | Streamlit dashboard for data visualization and analytics |
| `dataentry.py` | Interactive data entry interface with intelligent search |
| `reachout.py` | Reachout manager for uncontacted, revoked, and discarded contacts |
| `history_manager.py` | History tracking and audit logging for all record changes |
| `excel_format_manager.py` | Excel formatting, hyperlink conversion, and styling utilities |
| `linkedin_profiles_data.json` | Configuration file with extraction settings and paths |
| `excel_format_spec.json` | Excel formatting specifications and styling rules |
| `main.py` | Main application launcher with tabbed interface |

## 🚀 Usage

### Main Application
```bash
main.bat
```
Launches the complete LinkedIn Reachout Hub with both dashboard and data entry tabs.

### Dashboard Only
```bash
dashboard.bat
```
Launches Streamlit dashboard for data visualization and analytics.

### Test Excel Formatting
```bash
test_excel_format.bat
```
Tests Excel formatting and hyperlink functionality.

## 🔍 Data Entry Features

The data entry interface (`dataentry.py`) and reachout manager (`reachout.py`) provide powerful search and editing capabilities:

### Intelligent Search
- **Multi-word search**: Search for multiple terms simultaneously (e.g., `"ana izq"`)
- **Fuzzy matching**: Handles typos and similar words using difflib
- **Partial matching**: Finds substrings within names
- **Prefix matching**: Matches name beginnings (e.g., `"ana"` finds `"ana maría"`)
- **Relevance ranking**: Results sorted by match quality and coverage

### Auto-Selection
- Automatically selects the first search result for faster workflow
- Eliminates manual clicking when editing records

### Data Management
- Edit profile names, connection dates, response status, and chat URLs
- Manage date fields: Contacted, Connected, Revocation, and Discarded dates
- Clear checkboxes for each date field for quick reset
- Discard profiles to hide them from uncontacted view
- Filter by Uncontacted, Revoked, or Discarded contacts
- Automatic Excel formatting after saves (hyperlinks, styling)
- Real-time validation and error handling
- History tracking for all record changes (create, update, delete, discard)

### Search Examples
```
"ana izq"     → Finds "Ana Izquierdo", "Ana María Izquierdo"
"maría"       → Finds all names containing "María"
"izq"         → Finds "Izquierdo" variations using fuzzy matching
"ana maría"   → Finds names containing both "ana" and "maría"
```

## ⚙️ Configuration

Edit `linkedin_profiles_data.json` to configure:
- **Output paths**: Excel file locations and directories
- **Extraction settings**: Data collection parameters
- **Tab management**: Browser automation settings
- **Destination file**: Primary data storage location

Edit `excel_format_spec.json` to customize:
- **Cell formatting**: Colors, fonts, borders
- **Column styling**: Widths, alignments, number formats
- **Hyperlink formatting**: URL column detection and styling

## 🏗️ Architecture

### Search Algorithm
The fuzzy search implements a multi-strategy approach:

1. **Exact Matching** (100 pts): `"ana"` exactly matches `"ana"`
2. **Partial Matching** (80 pts): `"izq"` found in `"izquierdo"`
3. **Prefix Matching** (70 pts): `"ana"` matches start of `"ana maría"`
4. **Fuzzy Matching** (60+ pts): Similarity-based matching for typos

Final scores include bonuses for:
- Matching multiple search words (+10 pts per additional word)
- Having exact matches (+20 pts bonus)

### Data Flow
1. Load configuration and Excel data
2. User performs intelligent search
3. Auto-select first result or manual selection
4. Edit record data with validation
5. Save to Excel with automatic formatting
6. Update session state and refresh UI

## 📊 Data Format

Expected Excel columns:
- `name`: Full profile name (string)
- `date_contacted`: Date when contact was made (datetime)
- `date_connected`: Connection date (datetime)
- `date_revocation`: Date when contact was revoked (datetime)
- `date_discarded`: Date when profile was discarded (datetime)
- `ind_answered`: Response status (0/1 integer)
- `url_chat`: LinkedIn chat URL (string)
- `company`: Company name (string)
- `job_title`: Job title/position (string)
- `location`: Location/city (string)
- `reach_out_type`: Type of reachout made (string)

## 🔧 Dependencies

- `pandas`: Data manipulation and Excel handling
- `streamlit`: Web interface and interactive components
- `openpyxl`: Excel file operations
- `difflib`: Fuzzy string matching (built-in)
- `re`: Regular expressions for text processing (built-in)

## 🐛 Troubleshooting

### Common Issues
- **Config file not found**: Ensure `linkedin_profiles_data.json` exists in the same directory
- **Excel file locked**: Close Excel application before saving
- **No search results**: Check column names match expected format (`name`, etc.)
- **Formatting fails**: Verify `excel_format_spec.json` is valid JSON

### Debug Mode
Enable debug logging by checking the console output when running the application.
