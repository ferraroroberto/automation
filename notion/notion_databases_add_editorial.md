# notion_databases_add_editorial.py

Fills a Notion database with **one row per calendar day** for a date range you choose. Days that already exist (same `date` property in range) are skipped; missing days are created.

## Requirements

- **Python**: Use this repo’s virtual environment (`.venv`), which should include `notion-client` (2.x).
- **Notion integration**: A [Notion integration](https://www.notion.so/my-integrations) with the token stored in your params file, and **access to the target database** (share the database with the integration).
- **Params file**: Same pattern as other `notion/` scripts — a text file readable by `utils.read_params_from_txt_file` with an `api_token` key (see your existing `notion-params.txt` layout).

## Notion database shape

The script expects these **property names** (types matter):

| Property | Type        | Content |
|----------|-------------|---------|
| `day`    | Title       | Date as `YYYYMMDD` text |
| `date`   | Date        | ISO date `YYYY-MM-DD` |
| `DoW`    | Rich text   | Weekday abbreviation, lowercase (e.g. `mon`) |

If your properties differ, update `get_existing_dates` / `add_date_record` in the `.py` file to match.

## How to run

From the `notion` folder:

```bat
..\.venv\Scripts\python notion_databases_add_editorial.py
```

Or double-click / run:

```bat
notion_databases_add_editorial.bat
```

You will be prompted for **start** and **end** dates as `YYYYMMDD` (e.g. `20260401` … `20260430`).

## Configuration (in the script)

In `if __name__ == "__main__"`:

- **`params_file_path`** — Path to your Notion API token file (default in the repo points to a machine-specific path; change it if needed).
- **`database_id`** — 32-character Notion database ID (no hyphens), or the same ID with hyphens. You can copy it from the database URL: the segment before `?` after the last `/`, e.g.  
  `https://www.notion.so/workspace/<database_id>?v=...`

## Behaviour

1. **Retrieve** the database with `databases.retrieve`, then read the first **`data_sources`** entry (required by **notion-client 2.x**; the old `databases.query` helper was removed from the SDK).
2. **Query** that data source with a **date filter** limited to your requested range (fewer rows than scanning the whole DB).
3. **Create** pages for dates in range that are not already present, using `pages.create` with `parent: { database_id }`.

The client uses a **180s HTTP timeout** to reduce spurious timeouts on slow responses.

## Troubleshooting

| Symptom | What to check |
|--------|----------------|
| Timeout errors | Network; whether the integration can access the DB; try again; optionally raise `timeout_ms` in `init_notion_client`. |
| No data sources / empty `data_sources` | Database must be API-visible; integration must be connected; Notion API version must support data sources. |
| Property errors on create | Property names and types in Notion must match `day`, `date`, `DoW` as above. |
| Wrong or empty results when querying | Confirm the date property is named exactly `date` (used in filters and reads). |

## Related

- Original discussion context: [ChatGPT thread](https://chatgpt.com/c/6724bbf2-a618-8009-a4e3-a1f6e6828ae0) (referenced in the script header).
- Other Notion tooling in this repo: root `README.md` → **Notion Integration** section.
