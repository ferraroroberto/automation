# Notion Excel Sync

Pushes rows from an Excel workbook into a Notion database as new pages, with an optional step to archive every existing page in that database first. Intended as a one-way, Excel-to-Notion bulk import.

## 🚀 Overview

The script reads a three-sheet workbook (`data`, `metadata`, `db`), looks up the target database's ID and an archive flag from the `db` sheet, then (if the archive flag is set) marks every existing page's `archived` checkbox property as `True`. It then walks every row of the `data` sheet and, using the `metadata` sheet to map Excel columns to Notion property names/types, creates a new Notion page per row.

## 📋 Usage

### Basic Usage

```bash
python notion_excel_sync.py
```

There are no command-line arguments. The script prompts once for confirmation before making any changes:

```
Press Enter to continue...
```

Everything else — the workbook to sync, the API token, and the verbosity flag — comes from the params file described below.

## 🔧 Configuration

The script has no JSON config file. It loads a legacy `key = value` text file via `utils.read_params_from_txt_file`, pointed to by the `NOTION_PARAMS_FILE` environment variable:

```
sync_path = E:\automation\notion-automation-files\excel-sync.xlsx
api_token = secret_xxx
verbose = True
```

### Configuration Fields

- **sync_path**: Path to the Excel workbook to sync. Must contain three sheets: `data`, `metadata`, `db`.
- **api_token**: Notion integration token used to authenticate `notion_client.Client`.
- **verbose**: When the literal string `"True"` (compared with `== True`, so any other value is treated as off), logs the full `page_data` payload for every page before creating it.

### Workbook sheets (`sync_path`)

- **`db`** (single row read): `id` — the target Notion database ID; `clean` — truthy to archive all existing pages in the database before importing.
- **`metadata`**: One row per Excel column being synced, with `excelColumn`, `notionColumn`, `type`, and `keep` columns. `type` must be one of `title`, `date`, `number`, `url`, `relation`, `rich_text`, `checkbox`. `keep` (falsy) skips the column entirely.
- **`data`**: The rows to import. Each non-null cell in a column present (and kept) in `metadata` is mapped into the new page's `properties` under the matching `notionColumn`.

## 🔍 How It Works

### Processing Flow

1. **Load params and workbook**: Reads `sync_path`, `api_token`, `verbose`, then loads the `data`, `metadata`, and `db` sheets.
2. **Resolve target database**: Reads `notion_db_id` and `clean_db` from the first row of `db`; builds `metadata_dict` keyed by `excelColumn`.
3. **Confirm database**: Calls `client.databases.retrieve()` and `client.databases.query()` to log the database name and current row count, then blocks on `input("Press Enter to continue...")`.
4. **Optional archive**: If `clean_db` is truthy, queries all pages in the database and calls `client.pages.update()` on each, setting a checkbox property literally named `archived` to `True`.
5. **Create pages**: For each row in `data`, builds a `properties` dict from every column that has `keep` set and a non-null value in `metadata`, mapping by `type` (see column-type mapping below), then calls `client.pages.create()`.

### Column-type mapping

| `type` | Property payload |
|---|---|
| `title` | `{"title": [{"text": {"content": str(value)}}]}` |
| `date` | `{"date": {"start": value.strftime('%Y-%m-%d')}}` (value must already be a datetime) |
| `number` | `{"number": value}` |
| `url` | `{"url": str(value)}` |
| `relation` | `{"relation": [{"id": id} for id in str(value).split(',')]}` (comma-separated page IDs) |
| `rich_text` | `{"rich_text": [{"text": {"content": str(value)}}]}` |
| `checkbox` | `{"checkbox": bool(value)}` |

## ⚠️ Important Notes

1. **Destructive when `clean_db` is set**: Every page in the target database is marked `archived = True` before the import runs. This relies on an actual checkbox property named `archived` existing on the database — it does not call Notion's native page-archive API.
2. **Interactive by design**: The script always pauses for `Enter` after showing the current database name and row count, as a manual sanity check before any writes.
3. **`date` values must be real datetimes**: Cells that pandas didn't parse as dates will raise on `.strftime()`.
4. **Per-page errors are logged, not fatal**: Both the archive loop and the create loop catch `notion_client.APIResponseError` per call and continue.
5. **No dry-run mode**: Unlike `normalize_url.py`, there is no `--dry-run` equivalent — review `sync_path` and the `db`/`metadata` sheets before running.

## 📚 Dependencies

- `notion-client`: Notion API SDK (`Client`, `APIResponseError`).
- `pandas`: Reads the `data`, `metadata`, and `db` sheets.
- `utils.read_params_from_txt_file`, `utils.DEFAULT_PARAMS_FILE`: shared params-file loader.
