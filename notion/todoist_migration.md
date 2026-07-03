# Todoist Migration

Converts a folder of Todoist CSV project exports into a single consolidated Excel workbook of tasks and notes, ready for import elsewhere (e.g. into a Notion database via another script). This script makes no Notion API calls itself — it is a local CSV-to-Excel transform.

## 🚀 Overview

The script walks a source folder recursively for `.csv` files (one file per Todoist project export, named `ProjectName [projectid].csv`), parses each row's `TYPE` column (`task` or `note`, or blank to append to the previous row's comment), assigns a zero-padded `TASK_####`/`NOTE_####` ID to each, resolves whether each row's `DATE` is an exact date (also handling day/month-only dates like `31 Aug` by assuming the current year) and whether that date is in the past, and appends the result to a single DataFrame. Task rows also have a `TITLE`/`URL` pair split out of Todoist's markdown-style `[title](url)` content. The combined DataFrame is written to `Todoist_Migration.xlsx` in the source folder.

## 📋 Usage

### Basic Usage

```bash
python todoist_migration.py
```

There are no command-line arguments. The script always prompts once at startup:

```
Do you want verbose processing? (Y/N):
```

Answering `Y` enables debug-level logging (per-task/per-note detail); anything else runs with progress logged only every 100 rows per file. If the output file `Todoist_Migration.xlsx` is already open and can't be overwritten, the script prompts again:

```
Would you like to retry saving? (Y/N):
```

## 🔧 Configuration

There is no config file (JSON or params `.txt`) for this script. The only external input is an environment variable:

```
TODOIST_SOURCE_FOLDER=E:\automation\notion-automation-files\todoist-export
```

### Configuration Fields

- **TODOIST_SOURCE_FOLDER**: Root folder to walk recursively for `.csv` files. Read once at import time via `os.getenv("TODOIST_SOURCE_FOLDER", "")` — set it before running the script, there is no `--source` flag.

### Expected CSV shape

Each CSV is a standard Todoist project export, with the project name and ID encoded in the filename as `ProjectName [projectid].csv`, and columns including `TYPE`, `CONTENT`, `DESCRIPTION`, `PRIORITY`, `INDENT`, `AUTHOR`, `RESPONSIBLE`, `DATE`, `DATE_LANG`, `TIMEZONE`, `DURATION`, `DURATION_UNIT`.

## 🔍 How It Works

### Processing Flow

1. **Prompt for verbosity**: Reads `Y`/`N` from stdin.
2. **Walk source folder**: `main()` walks `TODOIST_SOURCE_FOLDER` recursively for `*.csv` files.
3. **Per file** (`process_file()`): Reads the CSV, skips it if empty, and processes each row by `TYPE`:
   - **Blank `TYPE`**: Appends the row's `CONTENT` to the previous appended row's `COMMENT`.
   - **`task`**: Assigns a new `TASK_####` ID, splits `TITLE`/`URL` out of `CONTENT` via `extract_title_url()`, resolves `IND_DATE`/`EXACT_DATE`/`IND_PAST` via `is_exact_date()` / `determine_past_future()`, and appends a full task row.
   - **`note`**: Assigns a new `NOTE_####` ID, links it to the most recent task via `TASK_ID = last_task_id`, resolves the same date fields, and appends a note row.
4. **Concatenate**: All rows from all files are combined into one DataFrame.
5. **Save**: `save_output()` writes the DataFrame to `TODOIST_SOURCE_FOLDER\Todoist_Migration.xlsx`, retrying on `PermissionError` (e.g. the file is open in Excel) if the user answers `Y` to the retry prompt.

### Date handling (`is_exact_date()`)

1. Tries `pandas.to_datetime(date_str)` directly.
2. If that fails, retries with the current year appended (handles Todoist's `31 Aug`-style recurring-looking dates that are actually missing a year).
3. If both fail, returns `(0, "", None)` — treated as not an exact date.

`determine_past_future()` then compares the resolved date to today's date, returning `1` (past), `0` (today or future), or `-1` (no valid date).

## ⚠️ Important Notes

1. **No Notion API calls**: Despite living in the `notion/` folder, this script only produces a local Excel file — a separate step is needed to actually import it into Notion.
2. **Filename format matters**: `project_name`/`project_id` are parsed from the filename as `name.split(' [')[0]` / `name.split('[')[-1].split(']')[0]` — files not matching `ProjectName [id].csv` will produce an empty or wrong `PROJECT_ID`.
3. **Interactive by design**: Both the initial verbosity prompt and the save-retry prompt block on stdin — this script is not meant for unattended/scheduled runs.
4. **Overwrites output**: Re-running the script overwrites `Todoist_Migration.xlsx` in the source folder.

## 📚 Dependencies

- `pandas`: CSV parsing, date parsing, and the final Excel export.
- `numpy`: Imported for numeric handling in the data pipeline.
