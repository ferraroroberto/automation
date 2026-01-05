# LinkedIn Profiles Data Extractor

Automated LinkedIn profile data extraction with multiple implementation approaches for streamlined lead generation and networking workflows.

## 📋 Project Summary

This project provides automated tools to extract LinkedIn profile data from search results, offering both reliability and flexibility through multiple extraction methods. Whether you need robust DevTools-based extraction or simple clipboard-based workflows, this suite provides everything needed for efficient LinkedIn data collection and management.

**Key Features:**
- DevTools-based extraction for reliability
- Intelligent data entry and search interface
- Excel formatting and hyperlink management
- Real-time dashboard for data visualization
- Cross-platform compatibility

## 🏗️ Project Structure

### 📁 [`common/`](./common/)
Shared utilities, dashboard, and configuration files used by all extraction methods.

- **Dashboard**: Streamlit visualization for data analytics and management
- **Data Entry**: Intelligent search and editing interface with fuzzy matching
- **Excel Tools**: Formatting, hyperlink conversion, and styling utilities
- **Configuration**: Centralized settings and extraction parameters

### 📁 [`devtools/`](./devtools/)
Chrome DevTools Protocol-based extraction (recommended for reliability).

- **Pros**: Direct DOM access, no clipboard operations, highly reliable
- **Cons**: Requires Chrome debugging setup
- **Best for**: Production use, large-scale extraction

## 🚀 Quick Start

### Prerequisites
- Python 3.8+
- Chrome browser (for DevTools version)
- Microsoft Excel (for data output)
- Required Python packages: `pandas`, `pynput`, `requests`, `websocket-client`, `streamlit`, `openpyxl`

### DevTools Version (Recommended)
```bash
cd devtools
start_chrome_debug.bat
linkedin_profiles_data_orchestrator_devtools.bat
```

### Dashboard & Data Management
```bash
cd common
dashboard.bat
```

## ⚙️ Configuration

Edit `common/linkedin_profiles_data.json` to configure:
- **Output Paths**: Excel file locations and directories
- **Extraction Settings**: Data collection parameters and selectors
- **Tab Management**: Browser automation settings and timing
- **Destination Files**: Primary data storage locations

Edit `common/excel_format_spec.json` to customize:
- **Cell Formatting**: Colors, fonts, borders, and styling
- **Column Layout**: Widths, alignments, and number formats
- **Hyperlink Detection**: URL column identification and formatting

## 📊 Data Format

The system expects and produces Excel files with these standard columns:
- `name`: Full profile name (string)
- `date connected`: Connection date (datetime)
- `answered`: Response status (0/1 integer)
- `chat_url`: LinkedIn chat URL (string, auto-converted to hyperlinks)

## 🔧 Dependencies

### Required Python Packages
```bash
pip install pandas pynput requests websocket-client streamlit openpyxl
```

### System Requirements
- **Chrome Browser**: Required for DevTools version (remote debugging enabled)
- **Excel**: Microsoft Excel or compatible spreadsheet application
- **Python**: Version 3.8 or higher

## 🏛️ Architecture Notes

- **Modular Design**: Each extraction method can run independently
- **Shared Components**: Common utilities maintain consistency across versions
- **Configuration Centralization**: All settings managed through `common/` folder
- **Relative Imports**: Maintain modularity and portability
- **Error Handling**: Comprehensive logging and user feedback
- **Cross-Platform**: Works on Windows, macOS, and Linux

## 🔍 Workflow Overview

1. **Setup**: Configure extraction parameters in `common/linkedin_profiles_data.json`
2. **Extract**: Run DevTools extraction scripts on LinkedIn search results
3. **Manage**: Use dashboard for data visualization and editing
4. **Format**: Automatic Excel formatting with hyperlinks and styling

## 🐛 Troubleshooting

### Common Issues
- **Chrome DevTools Connection**: Ensure Chrome is running with `--remote-debugging-port=9222`
- **Excel File Locked**: Close Excel application before running extraction
- **Configuration Errors**: Verify JSON files are valid and paths exist
- **Browser Focus**: For brute force method, keep browser window active
- **Limits**: see this conversation for a review on total limits of reachout connections > https://gemini.google.com/app/11b97204c7f26bd7

### Debug Mode
Enable detailed logging by checking console output when running applications. Test scripts provide step-by-step validation of functionality.

## 📈 Performance Tips

- **DevTools Method**: Most reliable for large-scale extraction
- **Dashboard**: Use for data review and manual corrections
- **Batch Processing**: Configure appropriate delays for system performance

---

*Built for efficient LinkedIn networking and lead generation workflows.*
