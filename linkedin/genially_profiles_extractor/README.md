# Genially Profiles Extractor

Extracts alumni profile cards (name, role, company, email, LinkedIn URL, photo) from a public Genially presentation, downloads the photos locally, and renders a self-contained, searchable HTML directory.

## How it works

Genially's view-API (`https://view.genially.com/api/view/<id>`) returns the full deck as JSON. Each profile card on a slide is built from independent widgets:

- a **Text** block — the name / role / company caption,
- an **Image** — the round portrait,
- two **Svg** icons whose `interactivities` map to:
  - an `htmlTooltip` action holding the email,
  - an `openLink` action holding the LinkedIn URL.

The script associates them per-card by clustering all widgets of each type into rows by y-coordinate, sorting each row by x, and pairing them by index. This is robust against the email/LinkedIn icons sitting ~300 px to the right of their caption (where naive 2D nearest-neighbour matching would leak across cards).

## Setup

1. Copy `config.example.json` to `config.json` and fill in the values:

   ```json
   {
     "url": "https://view.genially.com/<your-id>",
     "output_dir": "E:\\path\\to\\output",
     "output_stem": "genially_profiles",
     "html_title": "Alumnxs",
     "download_photos": true,
     "write_html": true,
     "save_raw": false,
     "verbose": false
   }
   ```

   `config.json` is gitignored (contains personal data); `config.example.json` is checked in.

2. Dependencies: `requests` (already in the project `.venv`).

## Run

```powershell
& .\.venv\Scripts\python.exe linkedin\genially_profiles_extractor\genially_profiles_extractor.py
```

Optional: `--config <path>` to point at a different config file.

## Output

Written to `output_dir`:

| File | Content |
|------|---------|
| `<stem>.json` | Profiles as JSON (one object per person). |
| `<stem>.csv` | Same data as CSV, UTF-8 BOM (Excel-friendly). |
| `<stem>.html` | Self-contained, searchable directory page. Open in any browser. |
| `photos/` | Downloaded portraits, named after the email local-part. |
| `<stem>_raw.json` | Raw Genially API JSON (only when `save_raw: true`). |

The HTML page inlines all data and references photos via relative paths, so it works offline as long as `photos/` sits next to the file.

## Configuration keys

| Key | Type | Default | Meaning |
|-----|------|---------|---------|
| `url` | string | — | Genially view URL (or bare 24-char id). Required. |
| `output_dir` | string | — | Where outputs are written. Required. |
| `output_stem` | string | `genially_profiles` | Filename stem for `.json` / `.csv` / `.html`. |
| `html_title` | string | `Alumnxs` | Heading shown at the top of the HTML page. |
| `download_photos` | bool | `true` | Download portraits into `photos/`. Cached: existing files are reused. |
| `write_html` | bool | `true` | Generate the searchable HTML page. |
| `save_raw` | bool | `false` | Also persist the raw Genially API JSON. |
| `verbose` | bool | `false` | Enable DEBUG-level logging. |

## Files in this folder

- `genially_profiles_extractor.py` — the extractor script.
- `config.example.json` — config template, committed.
- `config.json` — your real config, gitignored.
- `README.md` — this file.
