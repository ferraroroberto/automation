# Notion Databases Clean

Applies per-column cleaning rules from a metadata spreadsheet to each raw database export, dropping, extracting, renaming, and reordering columns, and writes a `_clean.xlsx` copy of each database.

## 🚀 Overview

For every database flagged `ind_clean = 1` in the master list, the script reads its raw exported Excel file and, for each column that has a matching row in the metadata sheet, either extracts a value out of a JSON string, extracts a substring via a regex built from the metadata's `extract`/`between` values, or leaves the column untouched. Metadata also controls which columns are kept, dropped, or kept in both raw and cleaned form, and what final name and column order to apply. The result is written to `dump_path` as `database-dump-{name}_clean.xlsx`, optionally with the previous file's column widths reapplied.

## 📋 Usage

### Basic Usage

```bash
python notion_databases_clean.py
```

There are no command-line arguments — everything comes from the params file described below plus the `excel_path` and `metadata_path` workbooks' own columns. Verbosity is controlled by the `verbose` key in the params file, not a flag.

## 🔧 Configuration

The script has no JSON config file. It loads a legacy `key = value` text file via `utils.read_params_from_txt_file`, pointed to by the `NOTION_PARAMS_FILE` environment variable:

```
excel_path = E:\automation\notion-automation-files\databases.xlsx
dump_path = E:\automation\notion-automation-files\dumps
metadata_path = E:\automation\notion-automation-files\clean-metadata.xlsx
verbose = False
```

### Configuration Fields

- **excel_path**: Path to the master database-list Excel file. Must contain `ind_clean`, `name`, `output_path`, and `ind_format` columns.
- **dump_path**: Directory the cleaned `database-dump-{name}_clean.xlsx` files are written to.
- **metadata_path**: Path to the per-column cleaning-rules Excel file (see columns below). Read as-is — the literal string `"True"` in `verbose` enables debug-level logging, anything else is treated as off.
- **verbose**: Must be the literal string `"True"` to enable per-row debug logging; any other value is treated as off.

### Metadata columns (`metadata_path`)

Each row describes how to clean one `(database, column)` pair:

| Column | Meaning |
|---|---|
| `database` | Database `name` this rule applies to |
| `column` | Source column name in the raw export |
| `clean` | `1` to apply regex extraction, otherwise leave the value as-is |
| `json` | `1` to treat the value as a JSON string and extract a nested field |
| `extract` | Literal text preceding the value to extract (used to build the regex, or as the JSON extraction key path prefix) |
| `between` | Delimiter surrounding the extracted value, or the literal string `"it's a number"` to extract a numeric value instead |
| `dictionary1` / `dictionary2` | Nested dictionary keys used when `json = 1`, e.g. `json_data[0][dictionary1][dictionary2]` |
| `ind_keep` | `0` = drop the column and its cleaned copy; `1` = keep only the cleaned copy, renamed to `column_name_final`; `2` = keep both the original and cleaned columns |
| `column_name_final` | Final column name when `ind_keep = 1` |
| `order_final` | Sort key controlling final column order (only columns with a non-null value are reordered) |

## 🔍 How It Works

### Processing Flow

1. **Load params and metadata**: Reads the params file, then loads `metadata_path` into a DataFrame.
2. **Select databases**: `read_database_list()` filters `excel_path` to rows where `ind_clean == 1`.
3. **Per database** (`process_databases()`):
   - Reads the raw export at `output_path`.
   - For each column, looks up the matching `(database, column)` row in the metadata. Columns with no metadata match are skipped (logged as a warning) and dropped from consideration for the output.
   - Applies `clean_data()` per row to build a `{col}_clean` column, then replaces literal `"[]"` values with `NaN`.
   - Uses `ind_keep` to decide whether to drop, keep-renamed, or keep-both for each column, and accumulates `order_final` for later reordering.
   - Drops the marked columns, filters to the columns to keep, and reorders by `order_final` if any were set.
   - Writes the result to `dump_path/database-dump-{name}_clean.xlsx`. If `ind_format == 1`, reads column widths from any existing file at that path first (via `utils.get_column_widths`) and reapplies them after writing (via `utils.apply_column_widths`).

### `clean_data()` extraction logic

- If `json == 1`: parses the cell with `ast.literal_eval` and returns `json_data[0][dictionary1][dictionary2]` if present, otherwise falls back to the original value.
- Else if `clean == 1` and `between == "it's a number"`: extracts the first number following `extract` in the string and returns it as a `float`.
- Else if `clean == 1`: extracts the substring between two occurrences of `between` that follow `extract`, via a regex built from `extract` and `between`.
- Otherwise: returns the value unchanged.

## ⚠️ Important Notes

1. **Metadata-driven, not schema-driven**: A column with no matching metadata row is dropped from the cleaned output and only logged as a warning — it is not an error.
2. **Overwrites output**: Re-running the script overwrites the `_clean.xlsx` file for each processed database.
3. **`verbose` must be the string `"True"`**, not a boolean, because it comes from a plain-text params file.
4. **Column-width preservation is best-effort**: If the previous `_clean.xlsx` can't be read (missing, corrupt, wrong format), a warning is logged and widths are simply not reapplied.

## 📚 Dependencies

- `pandas`: Reads the raw exports and metadata, builds the cleaned DataFrame.
- `openpyxl` (via `pandas`/`utils`): Excel I/O and column-width handling.
- `utils.read_params_from_txt_file`, `utils.get_column_widths`, `utils.apply_column_widths`, `utils.DEFAULT_PARAMS_FILE`: shared helpers.
