# Tests Directory

Comprehensive testing suite for Excel formatting and data management functionality in the LinkedIn Profiles Data Extractor project.

## 📋 Summary

The tests directory provides automated validation and debugging tools for the Excel format management system. These tests ensure that Excel formatting, hyperlink conversion, and file operations work correctly across different scenarios and edge cases.

**Key Testing Areas:**
- Excel format extraction and preservation
- File copying and format application
- URL hyperlink conversion
- Configuration loading and validation
- Error handling and user feedback

## 📁 Files Overview

| File | Purpose |
|------|---------|
| `test_excel_format.py` | Main test script for Excel format manager functionality |
| `test_excel_format.bat` | Windows batch launcher for the test script |

## 🧪 Test Coverage

### Part 1: Format Extraction
- **Purpose**: Extract and save Excel formatting to JSON specification
- **Function**: `save_excel_format_to_json()`
- **Validates**: Cell styles, column widths, number formats, borders
- **Output**: `excel_format_spec.json` in common directory

### Part 2: Format Application
- **Purpose**: Copy Excel file and apply saved formatting
- **Functions**: `apply_format_from_json()`, DataFrame export
- **Validates**: Format preservation, file operations, styling accuracy
- **Output**: Test file with `_test` suffix (e.g., `data_test.xlsx`)

### Part 3: Hyperlink Conversion
- **Purpose**: Convert URL columns to clickable Excel hyperlinks
- **Function**: `convert_url_columns_to_hyperlinks()`
- **Validates**: URL detection, hyperlink creation, formatting
- **Features**: Automatic URL column identification, hyperlink styling

## 🚀 Usage

### Run All Tests
```bash
test_excel_format.bat
```

This launches the comprehensive test suite that runs through all three parts sequentially with user interaction points.

### Manual Execution
```bash
python test_excel_format.py
```

## 🔍 Test Workflow

1. **Initialization**: Load configuration from `common/linkedin_profiles_data.json`
2. **Part 1 - Format Extraction**:
   - Read source Excel file formatting
   - Save format specification to JSON
   - Validate successful extraction
3. **Part 2 - Format Application**:
   - Create test copy of Excel file (raw format)
   - Apply formatting from JSON specification
   - Validate formatting accuracy
4. **Part 3 - Hyperlink Conversion**:
   - Identify URL columns in test file
   - Convert plain URLs to Excel hyperlinks
   - Apply hyperlink styling

## ⚙️ Configuration

Tests use the main project configuration file:
- **Config Path**: `../common/linkedin_profiles_data.json`
- **Required Setting**: `destination_file` (path to source Excel file)
- **Format Spec**: `../common/excel_format_spec.json` (auto-generated)

## 📊 Test Output

### Console Logging
- Detailed progress information for each test part
- Success/failure indicators with emojis
- File paths and operation status
- Error messages with troubleshooting hints

### File Outputs
- **Format JSON**: `excel_format_spec.json` (format specification)
- **Test Excel**: `{source}_test.xlsx` (formatted test file)
- **Hyperlinked Excel**: Updated test file with clickable URLs

## 🔧 Dependencies

### Required Imports
```python
import pandas as pd
from excel_format_manager import (
    save_excel_format_to_json,
    convert_url_columns_to_hyperlinks,
    apply_format_from_json
)
```

### System Requirements
- **Python**: 3.8+ with pandas and openpyxl
- **Source File**: Valid Excel file at configured `destination_file` path
- **Permissions**: Write access to output directories

## 🐛 Troubleshooting

### Common Issues
- **Config Not Found**: Ensure `common/linkedin_profiles_data.json` exists and is valid
- **Excel File Missing**: Verify the `destination_file` path points to existing Excel file
- **Permission Denied**: Close Excel application and ensure write permissions
- **Import Errors**: Install required packages and ensure proper Python path

### Debug Features
- **Step-by-Step Execution**: Tests pause between parts for manual verification
- **Detailed Logging**: Comprehensive console output for issue diagnosis
- **File Validation**: Checks file existence and accessibility before operations

## 🎯 Test Objectives

- **Reliability**: Ensure Excel formatting operations work consistently
- **Data Integrity**: Verify no data loss during format operations
- **User Experience**: Provide clear feedback and error handling
- **Regression Prevention**: Catch formatting issues before production use

## 📈 Best Practices

- **Run Before Production**: Execute tests before large-scale data extraction
- **Verify Output**: Manually check test Excel files for formatting accuracy
- **Update Format Spec**: Regenerate `excel_format_spec.json` when changing Excel templates
- **Monitor Logs**: Review console output for warnings or errors

---

*Essential testing tools for maintaining Excel formatting reliability in LinkedIn data extraction workflows.*