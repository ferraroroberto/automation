# Household Inventory & Shopping Helper

Mobile-responsive Streamlit application for managing household grocery inventory with intelligent shopping list generation.

## 🚀 Quick Start

### Method 1: Using the Batch File (Recommended)
Simply double-click `dashboard.bat` in the grocery folder to launch the app.

### Method 2: Manual Launch
```bash
# Install dependencies (if not already installed)
pip install streamlit pandas openpyxl

# Navigate to the grocery folder
cd E:\automation\automation\system\grocery

# Run the app
streamlit run app.py
```

## 📁 File Location

Your Excel inventory file should be located at:
```
C:\Users\rober\Downloads\list.xlsx
```

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

## 📊 Data Structure

Your `list.xlsx` file must contain these columns:

| Column | Description | Excel Column |
|--------|-------------|--------------|
| `super` | Supermarket name (e.g., "mercadona", "ametller") | A |
| `buscador` | Product URL for online shopping | B |
| `lugar` | Location in house (e.g., "fridge", "pantry", "garage") | C |
| `comida` | Item name | D |
| `cantidad` | Target quantity (how many you want to maintain) | E |
| `tenemos` | Current quantity (how many you actually have) | F |
| `comprar` | Auto-calculated: max(0, cantidad - tenemos) | G |

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