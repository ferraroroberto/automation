# Layout

A multi-tool `pipeline-cli` repo: each top-level folder is a tool domain holding independent scripts, not one application. The structure ratchet (`scripts/structure_ratchet.py`, configured in `.fleet.toml`) holds this layout: a root folder must be in the `pipeline-cli` preset or in `extra_roots`, and the folders below are exactly that list.

## Root

| Path | Holds |
|---|---|
| `browser_stealth.py` | Shared Playwright launch kwargs and init script every browser-driving tool imports |
| `no_window.py` | Shared `NO_WINDOW` subprocess flag for Windows spawns |
| `requirements.txt` | The one dependency list for every tool |
| `scripts/` | `verify-before-ship.ps1` (the gate) and the vendored `structure_ratchet.py`, `instruction_ratchet.py`, `doc_refs_ratchet.py`, `dead_code_ratchet.py` and `static_ratchet.py` |
| `docs/` | Durable reference: this file and `architecture.mmd` |
| `tests/` | Every test, under `tests/<domain>/` (see below) |

## Tool domains

Each of these is declared in `extra_roots` because the repo is a grab-bag of tools, so no single `src/` fits. A new tool goes into the domain whose subject it shares; a new domain is declared in `extra_roots` and gets a line here.

| Folder | Holds |
|---|---|
| `audio/` | Audio extraction, normalisation, trimming, OGG→MP3 conversion, transcript collation |
| `excel/` | Excel automation |
| `excel/accounting_stripe/` | Stripe accounting CSV structure comparison |
| `git-housekeeping/` | Git author-email policy and history-rewrite recipes (docs only) |
| `google/` | Gmail, Drive and Google Photos automation, with the shared `_auth.py` |
| `html/` | Standalone HTML utilities (countdown timer, settings demo) |
| `html/assets/` | Images the HTML pages use |
| `image/` | Image formatting, resizing, collage, PDF/HEIC conversion, photo archive, illustration distribution |
| `image/pdf_to_jpg/` | Standalone PDF→JPG converter with its own requirements |
| `linkedin/` | LinkedIn tooling |
| `linkedin/check_ip/` | Reverse-image profile search (SerpAPI + Imgur) |
| `linkedin/genially_profiles_extractor/` | Profile extraction from Genially pages |
| `linkedin/open_profiles/` | Opens a list of profiles in the browser |
| `linkedin/profiles_data_extractor/` | Profile data extraction; its Streamlit surface is `common/` |
| `linkedin/profiles_data_extractor/common/` | The dashboard, data-entry and reach-out pages and their shared helpers |
| `linkedin/profiles_data_extractor/devtools/` | Browser-devtools orchestrator and profile search checker |
| `notion/` | Notion API integrations, with the shared `utils.py` |
| `notion/articles_sync/` | Incremental articles database sync |
| `parking/` | Parking reservation poller: a deterministic poller that books free slots (a package; its `config.json` and `.env` live inside it) |
| `smart_life/` | Smart Life / IoT device automation |
| `system/` | Windows utilities: Outlook web mail, Word→Markdown, venv manager, mouse mover, markdown preview |
| `system/_lib/` | Helpers shared by the system and video tools (console decoding, single-instance lock) |
| `system/apple/` | vCard unification and conversion for Apple Contacts |
| `system/copilot_remap/` | Copilot key remap with autostart scripts |
| `system/foldersearcher/` | Folder index, CLI and tray search |
| `system/quickdeck/` | Quick-action deck, its config GUI and actions |
| `system/textexpander/` | Text expander with tray, monitor and GUI |
| `system/treesize/` | Folder size analyzer GUI |
| `system/wifi/` | Wi-Fi network scanner, saved-profile connect and `.bat` generator |
| `text/` | Text processing and PDF→text conversion |
| `video/` | Screen recording, trimming, re-encoding, concatenation, download, GPU recovery |
| `video/video_extensions_vlc/` | VLC file-extension association scripts |

## Tests

Tests never sit beside the code. `tests/<domain>/` mirrors the tool path it covers: `tests/system/wifi/` tests `system/wifi/`, `tests/parking/` tests `parking/`. Each test file puts its tool folder on `sys.path` from its own location, so it runs from any working directory, and the gate runs each folder as its own suite (`scripts/verify-before-ship.ps1`).

A tests folder holds no `__init__.py`: `unittest discover -s tests/<domain>` and pytest put that folder on the path, and tools import each other as top-level packages (`from system import outlook_web`), so a `tests` package would add a second import root for nothing.

## Baselined offenders

The ratchet records these at their size on adoption, so they may not grow. New code goes in a new module beside them, never into them.

- Files over 1,000 lines: `notion/articles_sync/notion_articles_sync.py`, `system/apple/unify_vcards.py`, `system/wifi/network_scanner.py`.
- Functions over 100 lines: 27, listed as `fn:` entries in `.structure-ratchet-baseline.json`. The longest are `main` in `linkedin/profiles_data_extractor/common/reachout.py`, `dataentry.py` and `dashboard.py`.
