# build_newsletter.py

Builds **Substack-ready HTML** from a Notion **newsletter** record and related **articles**: fetches pages, groups articles by topic, writes `build_newsletter.html` next to the script, opens it in the default browser, then prompts for a **must-read** order and copies a **single line of three titles** to the **clipboard**.

## Requirements

- **Python**: Repo virtual environment (`.venv`) with `requests`, `python-dotenv` (see root `requirements.txt`).
- **Notion**: Integration token in **`.env`** as `NOTION_API_TOKEN` (the script also reads `notion_api_key` from JSON if present; env wins for the token in practice).
- **Config**: `notion/build_newsletter.json` — `articles_database_id`, `newsletter_database_id`.

## Notion shape

**Newsletter database**

| Property | Type  | Use |
|----------|-------|-----|
| `number` | Title | Newsletter id, e.g. `N057` (script accepts `057` and normalizes to `N057` for the query) |

**Articles database**

| Property | Type        | Use |
|----------|-------------|-----|
| `article` | Title      | Article name |
| `link`    | URL        | Destination URL |
| `topic`   | Select     | One of: `personal development`, `innovation`, `leadership and management` |
| `star`    | Checkbox   | Sorting (starred first within a topic) |
| `niche`   | Multi-select | Secondary sort |
| `news`    | Relation   | Links to the newsletter page |

## How to run

From the repo root (with `.venv` activated):

```bash
python notion/build_newsletter.py --newsletter 057 --config notion/build_newsletter.json
```

Or from `notion/`:

```bash
python build_newsletter.py --newsletter 057
```

**Batch file**: `notion/build_newsletter.bat` — prompts for the number, then runs the script.

### Newsletter number

- **Preferred**: three digits, e.g. `057`, `001`.
- **Still accepted**: `N057` (same as after normalization).

Omit `--newsletter` to be prompted interactively.

### Command-line options

| Option | Default | Description |
|--------|---------|-------------|
| `--newsletter` | *(prompt)* | 3-digit id or `N` + 3 digits |
| `--config` | `build_newsletter.json` | Path to JSON config (resolved next to the script if missing) |
| `--debug` | off | Debug logging and full traceback on unexpected errors |

## After the HTML step

1. Logs the **top article** (first after internal sort) for each topic, in this order: **1** Personal development, **2** Innovation, **3** Leadership and management.
2. Asks: **Which is the "must read"? (1/2/3)**  
   The three titles are concatenated in this order (period + space between, final period):
   - **1** → 1, 2, 3  
   - **2** → 2, 1, 3  
   - **3** → 3, 1, 2  
3. Copies that line to the **system clipboard** (on **Windows**, via `clip.exe` with **UTF-8**).

If any topic has **no** articles, the must-read step is **skipped** (with a warning). EOF / Ctrl+C on the prompt skips the clipboard step without failing the run.

## Outputs

- **`notion/build_newsletter.html`** — overwritten each run; opened in the browser automatically.

## Related

- Root **`README.md`** → Notion Integration, usage example.
- **`build_newsletter.json`** — database IDs for your workspace.
