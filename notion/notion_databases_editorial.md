# Notion Databases Editorial

Copies each cleaned database export into a shared editorial-calendar workbook, one sheet per database, refreshing that sheet's contents on every run. This is a purely local, file-to-file copy — it does **not** talk to the Notion API.

> **Not the same script as `notion_databases_add_editorial.py`.** `notion_databases_add_editorial.py` connects to Notion directly and creates missing calendar-day *pages* inside a Notion database. `notion_databases_editorial.py` (this script) never touches the Notion API at all — it copies data between two local Excel workbooks: the per-database "clean" export produced by `notion_databases_clean.py`, and a single multi-sheet editorial-calendar spreadsheet.

## 🚀 Overview

The script reads the master database list, filters it to the rows flagged `ind_editorial = 1`, and for each of those databases copies the cell values from its clean Excel export (`output_path_clean`) into a sheet named `editorial_name` inside a shared editorial workbook. If that sheet already exists it is cleared first (every cell set to `None`) before the copy, so re-running the script always leaves the sheet matching the latest clean export.

## 📋 Usage

### Basic Usage

```bash
python notion_databases_editorial.py
```

There are no command-line arguments — everything comes from the params file described below and the `excel_path` workbook's own `ind_editorial`, `editorial_name`, and `output_path_clean` columns.

## 🔧 Configuration

The script has no JSON config file. It loads a legacy `key = value` text file via `utils.read_params_from_txt_file`, pointed to by the `NOTION_PARAMS_FILE` environment variable:

```
excel_path = E:\automation\notion-automation-files\databases.xlsx
editorial_excel_path = E:\automation\notion-automation-files\editorial-calendar.xlsx
```

### Configuration Fields

- **excel_path**: Path to the master database-list Excel file. Must contain `ind_editorial`, `name`, `editorial_name`, and `output_path_clean` columns.
- **editorial_excel_path**: Path to the shared multi-sheet editorial-calendar workbook that each database's data is copied into.

## 🔍 How It Works

### Processing Flow

1. **Load params**: Reads `excel_path` and `editorial_excel_path` from the params file.
2. **Select rows**: `read_database_editorial()` loads `excel_path` and keeps only rows where `ind_editorial == 1`.
3. **Per database** (`editorial_databases()`):
   - Opens the source workbook at `output_path_clean` (its active sheet) and the destination workbook at `editorial_excel_path`.
   - If a sheet named `editorial_name` already exists in the destination workbook, clears every cell in it; otherwise creates a new sheet with that name.
   - Copies every cell's value from the source sheet to the destination sheet, coordinate by coordinate.
   - Saves the destination workbook.
4. **Logging**: Reports rows copied per database and any per-database errors (caught and logged, not fatal — processing continues with the next database).

### Key Functions

- `read_database_editorial()`: Filters the master list to editorial-flagged databases.
- `editorial_databases()`: Performs the clear-then-copy for each flagged database.

## ⚠️ Important Notes

1. **Destructive refresh**: An existing destination sheet is fully cleared before the copy — there is no merge or append mode.
2. **Values only**: Only cell values are copied; formatting, formulas, and column widths are not preserved.
3. **Per-database errors don't stop the run**: A failure loading or saving one database's workbook is logged and the loop continues to the next database.
4. **Expects `notion_databases_clean.py` to have already run**: `output_path_clean` is the clean-export output path produced by that script.

## 📚 Dependencies

- `pandas`: Reads the master database list.
- `openpyxl`: Reads/writes both workbooks and clears/copies cells.
- `utils.read_params_from_txt_file`, `utils.DEFAULT_PARAMS_FILE`: shared params-file loader.
