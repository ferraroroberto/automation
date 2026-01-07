# Household Inventory & Shopping Helper

Mobile-responsive Streamlit application for managing household grocery inventory with intelligent shopping list generation and automated purchase tracking.

## 📋 Project Summary

This project provides a comprehensive solution for household inventory management, offering both reliability and flexibility through multiple operational modes. Whether you need to audit your current stock, edit target quantities, generate shopping lists, or export data, this application provides everything needed for efficient grocery management and shopping workflows.

**Key Features:**
- Mobile-optimized interface with touch-friendly controls
- Room-by-room inventory auditing with auto-save
- Intelligent shopping list generation grouped by supermarket
- Real-time purchase tracking with visual feedback
- Excel-based data storage with automatic calculations
- Cross-platform compatibility (works on any device with a browser)

## 🏗️ Project Structure

### 📁 [`.`](./)
Main application files and configuration.

- **app.py**: Main Streamlit application with all operational modes
- **config.json**: Application configuration and UI settings
- **laucher.bat**: Windows batch file for easy app launching
- **.streamlit/config.toml**: Streamlit theme customization
- **README.md**: This documentation file

## 🚀 Quick Start

### Prerequisites
- Python 3.8+
- Microsoft Excel (for data storage)
- Required Python packages: `streamlit`, `pandas`, `openpyxl`

### Method 1: Using the Batch File (Recommended)
Simply double-click `laucher.bat` in the grocery folder to launch the app.

### Method 2: Manual Launch
```bash
# Install dependencies (if not already installed)
pip install streamlit pandas openpyxl

# Navigate to the grocery folder
cd E:\automation\automation\system\grocery

# Run the app
streamlit run app.py
```

## ⚙️ Configuration

Edit `config.json` to customize:
- **UI Settings**: Page configuration, mode labels, and layout options
- **Data Paths**: Excel file location and column mappings
- **Logging**: Log level and format configuration
- **Display Options**: UI labels and styling preferences

## 📊 Data Format

The system expects and produces Excel files with these standard columns:

| Column | Description | Required | Excel Column |
|--------|-------------|----------|--------------|
| `super` | Supermarket name (e.g., "mercadona", "ametller") | Yes | A |
| `buscador` | Product URL for online shopping | Optional | B |
| `lugar` | Location in house (e.g., "fridge", "pantry", "garage") | Yes | C |
| `comida` | Item name | Yes | D |
| `cantidad` | Target quantity (how many you want to maintain) | Yes | E |
| `tenemos` | Current quantity (how many you actually have) | Yes | F |
| `comprar` | Auto-calculated: max(0, cantidad - tenemos) | Auto | G |

## 🔧 Dependencies

### Required Python Packages
```bash
pip install streamlit pandas openpyxl
```

### System Requirements
- **Python**: Version 3.8 or higher
- **Excel**: Microsoft Excel or compatible spreadsheet application
- **Browser**: Any modern web browser (Chrome, Firefox, Safari, Edge)

## 🏛️ Architecture Notes

- **Streamlit Framework**: Web-based interface for cross-platform compatibility
- **Excel Integration**: Direct read/write operations with auto-save functionality
- **Mobile-First Design**: Touch-optimized controls and responsive layouts
- **Session Management**: Persistent state for shopping progress tracking
- **Error Handling**: Comprehensive validation and user feedback
- **Configuration-Driven**: Centralized settings for easy customization

## 🔍 Workflow Overview

1. **Setup**: Configure your Excel file with inventory data and place it in the expected location
2. **Edit Targets**: Set desired quantities for items you want to track
3. **Audit Inventory**: Walk through your home and update current stock levels
4. **Generate Shopping List**: View items grouped by supermarket with purchase tracking
5. **Shop**: Use direct links to product pages and mark items as purchased
6. **Export**: Save changes and export data as needed

## 📱 Features

### 🔍 Audit Inventory Mode
- Walk through your home room-by-room
- Update current stock levels with touch-friendly +/- buttons
- Shows items with target > 0 only
- **Auto-saves** every change to Excel

### ✏️ Edit Targets Mode
- Set target quantities for items you want to track
- Adjust desired inventory levels
- Shows all items (not just those with targets)
- **Auto-saves** every change to Excel

### 🛒 Shopping List Mode
- View items that need to be purchased (comprar > 0)
- Items grouped by supermarket for efficient shopping
- **Purchase tracking**: Mark items as bought with visual feedback
- Direct links to product pages
- Progress counter shows completed purchases

### 💾 Save/Export Mode
- Manual save to Excel file
- Download updated inventory as CSV
- Summary statistics and reports

## 🖥️ Usage Guide

### 1. Edit Targets Mode (First Time Setup)
- Set target quantities for items you want to track
- Use +/- buttons to adjust desired inventory levels
- Only items with targets > 0 will appear in audit mode

### 2. Audit Inventory Mode (Regular Use)
- Select a zone (fridge, pantry, etc.)
- Count your actual inventory
- Use +/- buttons to update current quantities
- Changes save automatically

### 3. Shopping List Mode (Shopping Time)
- Items are grouped by supermarket
- Click "🛒 Buy Now" to open product pages
- Click "✅ Got It" to mark items as purchased
- Strikethrough text shows completed purchases
- "🗑️ Clear All" resets all purchase marks

### 4. Save/Export Mode
- "💾 Save to File" - manual save to Excel
- "📥 Download CSV" - export current data

## 🎨 UI Labels

The app uses these consistent labels throughout:
- **target**: Target quantity (from `cantidad` column)
- **current**: Current quantity (from `tenemos` column)
- **buy**: Amount to purchase (from `comprar` column)

## 📱 Mobile Optimization

- **Single-line layouts** - everything visible without scrolling
- **Large touch buttons** - easy to tap on mobile devices
- **Visual feedback** - clear indication of purchased items
- **Responsive design** - works on phones and tablets
- **Auto-save** - no need to manually save changes

## 🔧 Auto-Save Feature

Every time you update any quantity (target or current), the app automatically saves changes to your Excel file. No manual saving required!

## 📈 Progress Tracking

- Real-time counters show shopping progress
- Visual indicators for completed purchases
- Strikethrough formatting for bought items
- Summary statistics in export mode

## 🐛 Troubleshooting

### Common Issues
- **Excel File Not Found**: Ensure your Excel file is in the correct location (see File Location section)
- **Permission Errors**: Close Excel application before running the app
- **Configuration Errors**: Verify `config.json` is valid JSON and paths exist
- **Browser Issues**: Clear browser cache if interface appears broken

### Debug Mode
Enable detailed logging by checking the console output when running the application. Error messages provide specific details about any issues encountered.

## 📈 Performance Tips

- **Auto-Save**: Changes are saved immediately - no need to worry about losing data
- **Mobile Use**: Best experience on mobile devices with touch controls
- **Batch Updates**: Use the shopping mode for efficient purchase tracking
- **Regular Audits**: Frequent inventory checks keep your data accurate

---

*Built for efficient household inventory management and grocery shopping workflows.*