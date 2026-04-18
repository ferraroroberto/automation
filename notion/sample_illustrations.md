# sample_illustrations

Curates a browseable "inspiration samples" folder from the Notion illustrations archive: **one `.png` per visual type** (the most recently created one), organized both flat and nested, plus an XLS index with full Notion-ID traceability.

## Why

The `archived/` folder holds 1000+ illustrations. Many visual types have dozens of variants (barchart alone = 88). Browsing the full archive for inspiration is noisy. This script produces a curated subset where each visual type appears exactly once.

## Files

- `sample_illustrations.py` — the script
- `sample_illustrations.json` — config (DB IDs, source + dest folders, xlsx path)

## What it reads from Notion

Three databases, queried with standard cursor pagination:

| DB | ID | Fields consumed |
|---|---|---|
| illustrations | `f7008956ade24505bdc2c7faa9ed7902` | `illustration` (title), `theme` (multi_select), `Tags` (select), `Created` (created_time) |
| visual types | `81a8464d50af40aab88df43048495844` | `visualtype` (title), `concept` (relation), `illustrations` (relation) |
| concepts | `569fc5035d6d434796a2e039c66829cd` | `concept` (title), `type` (multi_select) |

## Sample-selection logic

For each visual type:
1. Filter its `illustrations` relation to IDs that exist in the illustrations map.
2. Sort by `Created` descending.
3. Pick the first → that's the sample.
4. Visual types with zero linked illustrations are skipped.

## Output layout

Destination: `.../inspiration samples/`

```
inspiration samples/
├── all/                                                      (flat — 1 file per visual type)
│   ├── chart - barchart - barchart - impossible things 10 years of practice.png
│   ├── metaphor - air balloon - air balloon - sacks of sand harmful habits.png
│   └── …
├── chart/                                                    (nested — concept_type)
│   ├── barchart/                                             (concept)
│   │   └── barchart - impossible things 10 years of practice.png
│   ├── area chart/
│   └── …
├── metaphor/
│   ├── ad-hoc/
│   └── …
├── untyped/                                                  (concepts with no type multi_select)
└── samples_index.xlsx
```

### Filename rules

- **Flat** (`all/`): `"{concept_type} - {concept} - {visualtype} - {topic}.png"` (big-picture → detail).
- **Nested** (`{concept_type}/{concept}/`): `"{visualtype} - {topic}.png"`.
- `concept_type` — first value of `concepts.type` (multi_select). Multiple values joined with `+`, empty → `"untyped"`.
- `concept` — `concepts.concept` title.
- `visualtype` — `visual_types.visualtype` title.
- `topic` — illustration title with the `"{visualtype} - "` prefix stripped. Falls back to the full title if no prefix match.
- All Windows-illegal chars (`<>:"/\|?*`) are stripped from each segment.

### Rebuild behavior

Every run wipes the destination folder's contents (folder itself preserved for iCloud sync) and rebuilds from scratch. No partial-update mode.

## XLS index (`samples_index.xlsx`)

One row per sampled visual type. Text fields are paired with their Notion IDs so rows can be traced back to source pages for debugging.

| Column | Notes |
|---|---|
| `concept_type` | e.g. `"chart"`, `"metaphor"`, `"untyped"` |
| `concept`, `concept_id` | From the concepts DB |
| `visual_type`, `visual_type_id` | From the visual types DB |
| `illustration_title`, `illustration_id` | From the illustrations DB |
| `topic` | Derived (title minus visualtype prefix) |
| `theme` | multi_select joined with `;` |
| `tags` | single select value |
| `created` | ISO timestamp |
| `source_path` | Expected `.png` path in `archived/` |
| `dest_path_flat` | Path written under `all/` |
| `dest_path_nested` | Path written under `{concept_type}/{concept}/` |
| `n_total_in_visual_type` | Count of illustrations linked to this visual type |
| `missing_source_file` | `True` when no `.png` was found at `source_path` (row is still written, no file copied) |

## Usage

```bash
# Default run (wipes + rebuilds)
python sample_illustrations.py

# Dry run — prints planned sample count + first 5 targets, no filesystem writes
python sample_illustrations.py --dry-run

# Custom config path
python sample_illustrations.py --config path/to/config.json
```

Auth: reads `NOTION_API_TOKEN` from `E:\automation\automation\.env` (same env as the other Notion scripts in this folder).

## Typical output

```
Fetched 23 pages from db 569fc503…
Fetched 510 pages from db 81a8464d…
Fetched 2700 pages from db f7008956…
Planned samples: 502 | Visual types with no illustrations: 8
48 sample(s) have no matching .png in source
Copied 454 illustration(s) to flat + nested layouts
Wrote XLS: …\inspiration samples\samples_index.xlsx (502 rows)
```

- **Planned samples** — visual types that will appear as xlsx rows.
- **Visual types with no illustrations** — skipped entirely.
- **Missing source** — sample picked in Notion but no matching `.png` in `archived/`. Row is still in the xlsx with `missing_source_file=True`; useful for debugging stale/renamed files.
- **Copied** — actual files written (planned − missing).

## Patterns reused from the rest of the project

- Env loading + `${NOTION_API_TOKEN}` placeholder resolution — same approach as [normalize_names.py](./normalize_names.py).
- Paginated DB query (`start_cursor` / `has_more`) — same as [normalize_names.py:114-162](./normalize_names.py#L114-L162).
- Property extraction (title / relation / multi_select / select) — same shapes as [articles_sync/notion_articles_sync.py](./articles_sync/notion_articles_sync.py) `extract_property`.
- XLS writing via `pd.to_excel(..., engine='openpyxl')` — consistent with the other Notion-to-Excel scripts in the folder.
