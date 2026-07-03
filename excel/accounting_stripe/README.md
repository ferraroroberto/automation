# CSV Structure Comparison and Processing Tool

A comprehensive automation tool for comparing CSV file structures and processing financial data from Stripe payments.

## 🚀 Overview

This tool performs two main functions:
1. **Structure Comparison**: Analyzes and reports differences between two CSV files (column names, order, and data types)
2. **Data Processing**: Filters, cleans, and exports financial data to Excel format with intelligent date-based filtering

Perfect for financial data reconciliation, automated reporting, and data quality assurance workflows.

## ✨ Features

### 🔍 Structure Analysis
- **Column Comparison**: Detects added/removed columns and order changes
- **Type Analysis**: Identifies data type changes between files
- **Detailed Reporting**: Clear, actionable reports of structural differences

### 🔄 Data Processing
- **Column Filtering**: Keeps only specified columns for focused analysis
- **Intelligent Date Filtering**: Automatic last closed quarter detection
- **Excel Export**: Clean Excel output with proper formatting
- **Safety First**: Confirmation prompt before processing to prevent accidents

### 🛡️ Safety & Reliability
- **User Confirmation**: Interactive prompts prevent accidental data processing
- **Error Handling**: Graceful handling of missing files, columns, and data issues
- **Logging**: Comprehensive logging with emoji indicators for easy reading
- **Configuration-Driven**: All settings managed through JSON config file

## 📋 Requirements

- **Python**: 3.8+ with pandas and openpyxl
- **OS**: Windows 10+ (PowerShell/Command Prompt)
- **Dependencies**: Listed in project root `requirements.txt`

## 🛠️ Installation

1. **Clone the repository** (if not already done)
2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
3. **Navigate to the tool directory**:
   ```bash
   cd excel/accounting_stripe
   ```

## ⚙️ Configuration

Edit `csv_compare_config.json` to customize behavior:

```json
{
  "default_directory": "E:\\onedrive\\Documentos\\Roberto\\areas\\finance\\stripe",
  "file_new": "unified_payments_all.csv",
  "file_old": "unified_payments_all_old.csv",
  "columns_to_keep": [
    "id",
    "Created date (UTC)",
    "Currency",
    "Converted Amount",
    "Converted Amount Refunded",
    "Description",
    "Fee"
  ],
  "date_column": "Created date (UTC)",
  "filter_by_quarter": true,
  "output_file": "filtered_payments.xlsx"
}
```

### Configuration Options

| Option | Description | Default |
|--------|-------------|---------|
| `default_directory` | Directory containing CSV files | Current directory |
| `file_new` | Primary CSV file to process | `unified_payments_all.csv` |
| `file_old` | Reference CSV for structure comparison | `unified_payments_all_old.csv` |
| `columns_to_keep` | Columns to retain during filtering | See example above |
| `date_column` | Column containing dates for filtering | `Created date (UTC)` |
| `filter_by_quarter` | Enable automatic quarter filtering | `true` |
| `output_file` | Excel output filename | `filtered_payments.xlsx` |

## 🎯 Usage

### Quick Start (Interactive Mode)

```bash
# Double-click launcher.bat or run:
python csv_structure_compare.py
```

### Command Line Options

```bash
# Use default directory
python csv_structure_compare.py

# Specify custom directory
python csv_structure_compare.py "C:\path\to\your\data"

# Custom date range (overrides quarter filtering)
python csv_structure_compare.py --start-date 2024-01-01 --end-date 2024-03-31

# Combined options
python csv_structure_compare.py "C:\data" --start-date 2024-07-01 --end-date 2024-09-30
```

### Date Filtering Logic

- **Default**: Automatically selects last closed quarter
- **Example**: Running on January 11, 2026 → filters Q4 2025 (Oct 1 - Dec 31, 2025)
- **Custom**: Override with `--start-date` and `--end-date` parameters
- **Format**: Use `YYYY-MM-DD` format for custom dates

## 📊 Output & Reports

### Structure Comparison Report

```
============================================================
CSV STRUCTURE COMPARISON REPORT
============================================================
Old file: unified_payments_all_old.csv
New file: unified_payments_all.csv

[WARNING] Structural changes detected:

• Added Columns:
   - New Column Name

• Removed Columns:
   - Old Column Name

• Order Changed:
   - Column moved from position X to Y

• Type Changed:
   - column_name: old_type -> new_type
```

### Processing Summary

```
📊 Loaded 1,234 rows, 16 columns
📊 Keeping 7 columns: ['id', 'Created date (UTC)', 'Currency', ...]
📅 Last closed quarter: Q4 2025 (2025-10-01 to 2025-12-31)
📅 Filtered 1,234 rows to 567 rows (date range: 2025-10-01 to 2025-12-31)
✅ Processed 567 rows, 7 columns
💾 Saved 567 rows to filtered_payments.xlsx
```

## 🔄 Workflow Example

1. **Preparation**: Place CSV files in the configured directory
2. **Comparison**: Tool analyzes structural differences
3. **Review**: User reviews the comparison report
4. **Confirmation**: Interactive prompt asks to proceed with processing
5. **Processing**:
   - Filters to specified columns only
   - Applies date range filtering (last quarter by default)
   - Exports clean data to Excel format
6. **Completion**: Ready-to-use Excel file for financial analysis

## 🛡️ Safety Features

- **Confirmation Prompts**: Must explicitly confirm before data processing
- **Non-Destructive**: Original CSV files are never modified
- **Clear Logging**: All actions logged with timestamps and status indicators
- **Error Recovery**: Graceful handling of missing files, invalid data, etc.

## 🐛 Troubleshooting

### Common Issues

**"Directory not found"**
- Verify the path in `csv_compare_config.json`
- Use absolute paths for reliability

**"Column not found" warnings**
- Check column names in your CSV files
- Update `columns_to_keep` in config if needed

**"No data to save"**
- Verify date column exists and contains valid dates
- Check if date range filters out all data
- Disable filtering temporarily with `--start-date` and `--end-date`

**Unicode encoding errors**
- Ensure CSV files are UTF-8 encoded
- Check for special characters in file paths

### Debug Mode

Enable detailed logging by modifying the logging level in the script:
```python
logging.basicConfig(level=logging.DEBUG, ...)
```

## 📁 File Structure

```
excel/accounting_stripe/
├── csv_structure_compare.py    # Main processing script
├── csv_compare_config.json     # Configuration settings
├── launcher.bat               # Windows batch launcher
└── README.md                  # This documentation
```

## 🔧 Technical Details

### Dependencies
- **pandas**: Data manipulation and CSV processing
- **openpyxl**: Excel file generation
- **pathlib**: Modern path handling
- **datetime**: Date calculations and filtering

### Date Quarter Calculation

```python
# Current date: January 11, 2026
current_quarter = 1  # Q1
# Last closed quarter = Q4 of previous year
# Result: Q4 2025 (2025-10-01 to 2025-12-31)
```

### Performance Notes
- **Memory Usage**: Loads entire CSV into memory (suitable for <1M rows)
- **Processing Speed**: Typically <30 seconds for 10K rows
- **File Sizes**: Excel output ~2-5x larger than filtered CSV

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

This project follows the automation monorepo license terms.

## 🆘 Support

For issues or questions:
1. Check the troubleshooting section above
2. Review the logs for error details
3. Verify configuration settings
4. Test with sample data first