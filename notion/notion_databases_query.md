# Notion Databases Query

Lists every database visible to a Notion integration and exports the name, ID, and a direct workspace URL for each one to an Excel file. Useful as the discovery step before other `notion/` scripts that need a specific `database_id`.

## 🚀 Overview

The script authenticates against the Notion API, calls the search endpoint filtered to `object = database`, and writes one row per database (`name`, `id`, `url`) to an Excel workbook. Database IDs are normalized by stripping the `-` characters Notion includes in its UUIDs, and the export URL is built as `{workspace_url}{id}?v=b` so each row can be opened directly in the browser.

## 📋 Usage

### Basic Usage

```bash
python notion_databases_query.py
```

There are no command-line arguments — the script takes its inputs entirely from the params file described below and runs straight through, then exits.

## 🔧 Configuration

The script has no JSON config file. It loads a legacy `key = value` text file via `utils.read_params_from_txt_file`, pointed to by the `NOTION_PARAMS_FILE` environment variable:

```
api_token = secret_xxx
db_excel_path = E:\automation\notion-automation-files\databases.xlsx
workspace_url = https://www.notion.so/yourworkspace/
```

### Configuration Fields

- **api_token**: Notion integration token used to authenticate `notion_client.Client`.
- **db_excel_path**: Output path for the Excel file listing every discovered database.
- **workspace_url**: Base workspace URL prefixed to each database ID to build a clickable link.

## 🔍 How It Works

### Processing Flow

1. **Load params**: Reads `api_token`, `db_excel_path`, and `workspace_url` from the params file at `NOTION_PARAMS_FILE`.
2. **Authenticate**: Creates a `notion_client.Client` with `api_token`.
3. **Search databases**: `get_database_list()` calls `notion.search(filter={"property": "object", "value": "database"})` and collects `results`.
4. **Build rows**: For each database, strips hyphens from the ID, uses the title's `plain_text` (or `"Untitled"` if the database has no title), and composes `url = f"{workspace_url}{clean_id}?v=b"`.
5. **Export**: `save_to_excel()` writes the `name` / `id` / `url` columns to `db_excel_path` via `pandas.DataFrame.to_excel` (engine `openpyxl`).

### Key Functions

- `get_database_list()`: Queries Notion and returns the list of `{name, id, url}` dicts.
- `save_to_excel()`: Writes that list to the configured Excel path.

## ⚠️ Important Notes

1. **Integration access**: Only databases the integration has been explicitly shared with are returned by `notion.search`.
2. **No filtering options**: The script always lists every visible database; there is no `--days` or name-filter flag.
3. **Overwrites output**: Re-running the script overwrites `db_excel_path` in place.

## 📚 Dependencies

- `notion-client`: Notion API SDK.
- `pandas` / `openpyxl`: Excel export.
- `utils.read_params_from_txt_file`, `utils.DEFAULT_PARAMS_FILE`: shared params-file loader.
