# Household Inventory & Shopping Helper

Mobile-responsive Streamlit application for managing household grocery inventory with intelligent shopping list generation and real-time purchase tracking.

## 📋 Project Summary

Comprehensive household inventory management across multiple operational modes. Audit current stock room-by-room, edit target quantities, track shopping in real time, and add on-the-fly items directly to the shopping list.

**Key Features:**
- Mobile access over local Wi-Fi — use the **Copy link** button in the sidebar to get the URL and open it on your phone
- Room-by-room inventory auditing with auto-save (best done from mobile)
- Shopping list grouped by supermarket with per-store progress bars (best done from desktop)
- Cart offset counters to account for items already in the cart
- Quick-add items (name + quantity) to any supermarket's shopping list
- Excel-based data storage with automatic calculations
- Cross-platform compatibility (works on any device with a browser)

## 🏗️ Project Structure

- **app.py** — Entry point: page config, session state, sidebar, mode routing
- **data.py** — Config, XLSX load/save, supermarket stats, quantity mutators
- **ui_helpers.py** — CSS, inline HTML formatters, sidebar utility actions
- **audit.py / edit_targets.py / edit_item.py / add_item.py / shopping.py / export.py** — One file per mode, each exposing `main(df)`
- **config.json** — Application configuration and UI settings
- **launcher.bat** — Windows batch file for easy app launching
- **.streamlit/config.toml** — Streamlit theme customization

## 🚀 Quick Start

### Prerequisites
- Python 3.8+
- Required Python packages: `streamlit`, `pandas`, `openpyxl`

### Method 1: Using the Batch File (Recommended)
Double-click `launcher.bat` in the grocery folder.

### Method 2: Manual Launch
```bash
pip install streamlit pandas openpyxl
cd E:\automation\automation\system\grocery
streamlit run app.py
```

### Mobile Access (same Wi-Fi network)
The app binds to all network interfaces automatically. To open it on your phone:
1. Launch `launcher.bat` on the PC as usual
2. Click **📋 Copy link** in the sidebar — this copies `http://<local-ip>:8501` to the clipboard
3. Paste the URL into Telegram (or any messaging app) and open it on your phone

> **Firewall:** if the phone cannot connect on first use, run this once in PowerShell (admin):
> ```powershell
> New-NetFirewallRule -DisplayName "Streamlit Grocery" -Direction Inbound -Protocol TCP -LocalPort 8501 -Action Allow
> ```

> **Audit mode on mobile:** rotate your phone to **landscape** for the best layout — the row-per-item grid fits without horizontal scrolling.

## ⚙️ Configuration

Edit `config.json` to customize:
- **Data Paths** — Excel file location and column mappings
- **UI Settings** — Page config, mode labels, layout
- **Logging** — Log level and format

## 📊 Data Format

Excel file columns:

| Column | Description | Notes |
|--------|-------------|-------|
| `super` | Supermarket name (e.g., `mercadona`, `ametller`) | Required |
| `buscador` | Product URL for online shopping | Optional |
| `lugar` | Zone in the house (e.g., `fridge`, `pantry`) | Required |
| `comida` | Item name | Required |
| `cantidad` | Target quantity to maintain | Required |
| `tenemos` | Current quantity on hand | Required |
| `comprar` | Auto-calculated: `max(0, cantidad − tenemos)` | Auto |

## 📱 Modes

### 🔍 Audit Inventory
Walk through each zone of the house, update current stock levels with ±1 buttons. Auto-saves every change to Excel.
Best done from mobile — rotate to **landscape** for optimal layout.

### ✏️ Edit Targets
Set or adjust target quantities per item. Auto-saves every change.

### 🔧 Edit Item
Search for any item and edit all its fields (name, supermarket, zone, URL, quantities) or delete it.

### ➕ Add Item
Add new items to the inventory via a form.

### 🛒 Shopping List
View items that need to be purchased, grouped by supermarket.

**Cart offset counters (sidebar):**
Each supermarket shows an editable `＋items` and `＋units` counter below its progress bar. Use these when items were already placed in the physical cart before opening the app — the bar and totals update immediately to reflect the combined count.

**Quick-add items:**
At the bottom of each supermarket's expander, a small inline form lets you add ad-hoc items (name + quantity). These are session-only and support the full `✅ Got it` / `↩️ Undo` / `🗑️ Remove` workflow. Works for both Ametller and Mercadona (and any other supermarket in the list).

### 💾 Save / Export
Manual save to Excel or download as CSV, plus summary statistics.

## 🖥️ Typical Workflow

1. **Edit Targets** — set desired quantities for tracked items
2. **Audit Inventory** — walk through zones and update current stock
3. **Shopping List** — check what to buy, mark as bought while shopping
4. Use **cart offset counters** if items were already in the cart
5. Use **quick-add** for anything not in the system

## 🐛 Troubleshooting

| Issue | Fix |
|-------|-----|
| Excel file not found | Verify the path in `config.json` |
| Permission error on save | Close Excel before running the app |
| Interface appears broken | Clear browser cache |
| Config errors | Validate `config.json` is well-formed JSON |

---

*Built for efficient household inventory management and grocery shopping.*
