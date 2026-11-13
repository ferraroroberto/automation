# Layout

A multi-tool `pipeline-cli` repo: each folder under `tools/` is a tool domain holding independent scripts, not one application. The structure ratchet (`scripts/structure_ratchet.py`, configured in `.fleet.toml`) holds this layout: a root folder must be in the `pipeline-cli` preset, and `tools/` is the preset's home for one package per tool.

## Root

| Path | Holds |
|---|---|
| `browser_stealth.py` | Shared Playwright launch kwargs and init script every browser-driving tool imports |
| `no_window.py` | Shared `NO_WINDOW` subprocess flag for Windows spawns |
| `requirements.txt` | The one dependency list for every tool |
| `scripts/` | `verify-before-ship.ps1` (the gate) and the vendored `structure_ratchet.py`, `instruction_ratchet.py`, `doc_refs_ratchet.py`, `dead_code_ratchet.py`, `static_ratchet.py`, `docs_ratchet.py`, `comment_ratchet.py`, `conventions_ratchet.py` and its sibling `conventions_readers.py` |
| `docs/` | Durable reference: this file, `architecture.mmd`, `parking-outcome-state.md` and `git-author-email.md` (the git author-email policy and history-rewrite recipes) |
| `tests/` | Every test, under `tests/<domain>/` (see below) |
| `tools/` | One folder per tool domain, the multi-tool home of the `pipeline-cli` shape (see below) |

## Tool domains

Each of these sits under the preset's `tools/` root because the repo is a grab-bag of tools, so no single `src/` fits. A new tool goes into the domain whose subject it shares; a new domain is a new `tools/<domain>/` folder and gets a line here. Scripts run as `python tools/<domain>/<script>.py` or `python -m tools.<domain>.<module>` from the repo root, and tools import each other as `tools.<domain>` (`from tools.system._lib import runtime_files`).

| Folder | Holds |
|---|---|
| `tools/audio/` | Audio extraction, normalisation, trimming, OGG→MP3 conversion, transcript collation |
| `tools/excel/` | Excel automation |
| `tools/excel/accounting_stripe/` | Stripe accounting CSV structure comparison |
| `tools/google/` | Gmail, Drive and Google Photos automation, with the shared `_auth.py` |
| `tools/html/` | Standalone HTML utilities (countdown timer, settings demo) |
| `tools/html/assets/` | Images the HTML pages use |
| `tools/image/` | Image formatting, resizing, collage, PDF/HEIC conversion, photo archive, illustration distribution |
| `tools/image/pdf_to_jpg/` | Standalone PDF→JPG converter with its own requirements |
| `tools/linkedin/` | LinkedIn tooling |
| `tools/linkedin/check_ip/` | Reverse-image profile search (SerpAPI + Imgur) |
| `tools/linkedin/genially_profiles_extractor/` | Profile extraction from Genially pages |
| `tools/linkedin/open_profiles/` | Opens a list of profiles in the browser |
| `tools/linkedin/profiles_data_extractor/` | Profile data extraction; its Streamlit surface is `common/` |
| `tools/linkedin/profiles_data_extractor/common/` | The dashboard, data-entry and reach-out pages and their shared helpers |
| `tools/linkedin/profiles_data_extractor/devtools/` | Browser-devtools orchestrator and profile search checker |
| `tools/notion/` | Notion API integrations, with the shared `utils.py` |
| `tools/notion/articles_sync/` | Incremental articles database sync |
| `tools/parking/` | Parking reservation poller: a deterministic poller that books free slots (a package; its `config.json` and `.env` live inside it) |
| `tools/smart_life/` | Smart Life / IoT device automation |
| `tools/system/` | Windows utilities: Outlook web mail, Word→Markdown, venv manager, mouse mover, markdown preview |
| `tools/system/_lib/` | Helpers shared across the tool domains: console decoding, single-instance lock, `pooled_http.py` (the vendored keep-alive `requests` session; every outbound HTTP call goes through its `pooled_request`) and `runtime_files.py` (writers for operator-named paths that raise when a path is inside the checkout, plus `runtime_data_dir()`, the machine-local folder for logs and scratch output) and `notifier.py` (`FleetNotifier`, the Telegram sender through the fleet `notify_send.py`, used by `tools/parking/` and the Photos weekly job) |
| `tools/system/apple/` | vCard unification and conversion for Apple Contacts |
| `tools/system/copilot_remap/` | Copilot key remap with autostart scripts |
| `tools/system/foldersearcher/` | Folder index, CLI and tray search (retired; folder-explorer replaces it, task-os vendors the core) |
| `tools/system/quickdeck/` | Quick-action deck, its config GUI and actions |
| `tools/system/snap/` | Hotkey region screenshot saved as a timestamped JPEG (`snap_core.py` holds the testable logic, `snap_tool.py` the hotkey, overlay and tray) |
| `tools/system/textexpander/` | Text expander with tray, monitor and GUI |
| `tools/system/treesize/` | Folder size analyzer GUI |
| `tools/system/wifi/` | Wi-Fi network scanner, saved-profile connect and `.bat` generator |
| `tools/text/` | Text processing and PDF→text conversion |
| `tools/video/` | Screen recording, trimming, re-encoding, concatenation, download, GPU recovery |
| `tools/video/video_extensions_vlc/` | VLC file-extension association scripts |

## Tests

Tests never sit beside the code. `tests/<domain>/` mirrors the tool path it covers: `tests/system/wifi/` tests `tools/system/wifi/`, `tests/parking/` tests `tools/parking/`. Each test file puts its tool folder on `sys.path` from its own location, so it runs from any working directory, and the gate runs each folder as its own suite (`scripts/verify-before-ship.ps1`).

A tests folder holds no `__init__.py`: `unittest discover -s tests/<domain>` and pytest put that folder on the path, and tools import each other as `tools.<domain>` packages (`from tools.system import outlook_web`), so a `tests` package would add a second import root for nothing.

## Baselined offenders

None: the structure baseline is empty. `tools/notion/articles_sync/notion_articles_sync.py` stays under the file limit by keeping the property mapping (`property_mapping.py`) and change detection (`change_detection.py`) in mixin modules beside it; its tests still build `NotionArticlesSync` directly with `__new__` and set the state by hand. New code goes in a module of its own, not back into a file near the limit.
