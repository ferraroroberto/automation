# Folder Searcher

A Tkinter tray app that indexes folder *names* across one or more root folders and searches them instantly. Faster and more focused than Windows search, because it never looks at files — only folders — and it searches a pre-built index rather than the disk.

## Features

- **System Tray App**: Runs in the system tray; click the icon (or pick **Open**) to show the window, close the window to minimize back to tray, **Quit** from the tray menu. Single-instance — launching it twice just notifies and exits.
- **Multi-Root Indexing**: Scan any number of root folders into a single index.
- **Instant Search on Open**: The saved index loads at startup, so searching works immediately without rescanning.
- **Multi-Term Search**: Space- or semicolon-separated terms, case-insensitive, AND-ed against the full path.
- **Email-Branch Pruning**: Collapses an item's contact and sibling branches into a single owning-item result.
- **Path-Depth Display**: Trims repetitive leading path components from the result list without affecting what gets opened.
- **Windows Explorer Integration**: Double-click a result to open that folder.
- **Persistent Data**: Config and index survive between sessions.

## Requirements

- Python 3.7 or higher (the core uses dataclasses)
- Windows (Explorer integration and the single-instance mutex are Windows-specific)
- Tkinter (usually included with Python)
- pystray (for the system tray icon)
- Pillow (for drawing the tray icon)

Install from the main project requirements:

```bash
pip install -r ../../requirements.txt
```

## Usage

### Launching

Run `system\foldersearcher.bat` — the launcher one level up, alongside the other `system/` launchers, and the only one. It starts `foldersearcher.py` silently with the repo `.venv`'s `pythonw.exe`, so no console window appears. The app starts in the tray — click the tray icon, or choose **Open**, to show the window.

### First-time setup

1. Open the window and switch to the **Folders** tab.
2. Click **Add Root** and pick a folder. Repeat for each root you want indexed.
3. Click **Scan All Roots** to build `folder_structure.txt`.

### Searching

The **Search / Access** tab is selected when the window opens and already has the saved index loaded — no scan needed.

1. Type one or more terms in **Search word(s)**. Terms are split on whitespace and `;`, lowercased, and AND-ed: a folder matches only when its full path contains every term.
2. Press Enter or click **Search**.
3. Double-click a result to open it in Windows Explorer.

### Closing vs quitting

- The window's **Close** (X) button hides the app back into the tray; the index and last search stay loaded.
- To fully exit, right-click the tray icon and choose **Quit**.

### Headless scan (CLI)

`foldersearcher_cli.py` rebuilds `folder_structure.txt` without the tray app, from the roots in `foldersearcher.json`. From the repo root:

```powershell
& .\.venv\Scripts\python.exe system\foldersearcher\foldersearcher_cli.py scan
```

It logs the roots, folder count and duration to stdout. Unlike **Scan All Roots** it is built for unattended runs:

- **A missing root aborts the run** and leaves the existing index byte-identical — the core would otherwise skip that root and silently shrink the index.
- **The index is written atomically**: to a temp file in the same folder, then swapped in with `os.replace`, so the tray never reads a half-written file.
- It imports no GUI module, takes no single-instance mutex, and never writes `foldersearcher.json` back (a legacy `root_folder` config is migrated in memory only), so it runs fine while the tray app is open.

| Exit code | Meaning |
| --- | --- |
| `0` | Every configured root existed and the index was replaced. |
| `1` | The index could not be written, or an unexpected crash. |
| `2` | Usage error (unknown subcommand). |
| `3` | No roots configured — index left untouched. |
| `4` | One or more configured roots are missing — index left untouched. |

The tray app picks up a rebuilt index without a restart: whenever the window opens or a search runs, it reloads `folder_structure.txt` if the file's modification time changed since it was loaded.

### Nightly job

`run-scan-nightly.bat` runs the CLI scan in the foreground with the repo `.venv`'s `python.exe` (`PYTHONUTF8=1`) and exits with its exit code. It is registered in app-launcher as the machine-local job `foldersearcher-scan-nightly` (schedule `none`, `session_less`, `alert_on_failure`), chained from both `on_success` and `on_failure` of `email-archiver-scan-nightly` — so the nightly chain is backup → email-archiver scan → folder scan, and folders created during the day are in the next morning's index.

## Configuration

`foldersearcher.json`, alongside the script:

| Key | Default | Meaning |
| --- | --- | --- |
| `root_paths` | `[]` | List of root folders to index. |
| `structure_file` | `<script dir>/folder_structure.txt` | Where the index is written. |
| `skip_depth` | `4` | Leading path components hidden in the result list. |
| `prune_email_branches` | `true` | Collapse item subtrees to the owning item. |

**Save Config** on the Folders tab persists `skip_depth` and `prune_email_branches`. Adding or removing a root saves immediately.

### `skip_depth`

Display only. With `skip_depth: 4`, a result at `E:/onedrive/Documentos/clientes/acme` is shown as `acme` — the first four components are hidden to keep the list readable. The full absolute path is preserved internally and is what double-click opens, so changing this never affects which folder you land in.

### `prune_email_branches`

In trees where each item folder holds a contact subfolder named after an email address, plus sibling branches such as `muestras`, `temporal`, or `inventario`, one conceptual hit would otherwise fill the list with a row per branch.

With pruning on, a folder that has a **direct email-named child** becomes a leaf result: any match at or below it collapses onto that folder, and results are de-duplicated. Nested items collapse upward to the shallowest owning item. Matches with no owning ancestor are returned unchanged. A folder counts as email-named when it matches `^[^@\s\\/]+@[^@\s\\/]+\.[A-Za-z]{2,}$` — so `info@acme.com` qualifies, while names like `@vscode` or `cl@ve` do not.

Set it to `false` (or untick the Folders-tab checkbox) to get every matching branch individually.

## File structure

```
foldersearcher.py                # Tray + Tkinter UI only
foldersearcher_core.py           # All logic: scan, persist, search, prune (no GUI imports)
foldersearcher_cli.py            # Headless `scan` subcommand (atomic write, strict roots)
run-scan-nightly.bat             # Foreground launcher for the nightly app-launcher job
test_foldersearcher_core.py      # Focused tests for the core
test_foldersearcher_cli.py       # CLI: rebuild, missing/no roots, atomic write
test_foldersearcher_tray_reload.py  # Tray reloads an index rewritten on disk
foldersearcher.json              # Configuration
folder_structure.txt             # Index (auto-generated)
foldersearcher.md                # This documentation
```

`foldersearcher_core.py` imports no `tkinter`, `pystray`, or `win32` module, so it runs and tests headlessly.

## Index format

New scans write explicit root sections:

```
Root: E:/onedrive/Documentos
E:/onedrive/Documentos/Ana
  E:/onedrive/Documentos/Ana/archive

Root: D:/work
...
```

A legacy headerless `folder_structure.txt` (root-relative paths, pre-multi-root) still loads and is resolved against the first configured root. Likewise, a legacy `foldersearcher.json` carrying a single `root_folder` string migrates to a one-element `root_paths` on load. Both legacy formats are read-only compatibility paths — the next scan or **Save Config** rewrites them in the current format.

## How it works

1. **Scanning** walks every configured root with `os.walk` and records each folder and its direct children as absolute paths. A root that does not exist is skipped with a warning rather than aborting the whole scan.
2. **Storage** writes that index to `folder_structure.txt` in root sections.
3. **Searching** matches terms against the absolute path, then optionally prunes to owning items. Pruning reads only the index, never the disk, so search stays fast and testable.
4. **Opening** hands the stored absolute path to `os.startfile`, falling back to `explorer`.

## Testing

From the repo root:

```powershell
& .\.venv\Scripts\python.exe -m unittest discover -s system/foldersearcher -p "test_*.py"
```

Covers root normalization, legacy config migration, sectioned round-trip, legacy index loading, search matching, `skip_depth` display, and pruning with the flag on and off; the headless CLI's rebuild, missing-root and no-root refusals (old index byte-identical), and atomic write; and the tray's reload of an index rewritten on disk.

## Troubleshooting

| Symptom | Cause / fix |
| --- | --- |
| "No index found" on open | Add roots on the **Folders** tab, then **Scan All Roots**. |
| "No index loaded" when searching | Same — the index file is missing or empty. |
| A root's folders are missing after a scan | The root was unreachable at scan time; the log records `Skipping missing root: …`. Reconnect the drive and rescan. |
| Nightly job `foldersearcher-scan-nightly` failed with exit 4 | A configured root was unreachable; the run's `output.log` names it. The previous index was kept. |
| Results show too many rows per item | `prune_email_branches` is off, or the item has no direct email-named child. |
| Result rows are unreadably long | Raise `skip_depth` on the Folders tab and **Save Config**. |
| "Folder does not exist" on double-click | The index is stale — rescan. |
| Second launch does nothing | Single-instance by design; the existing instance notifies from the tray. |

## Logging

Startup, config load/save, scan progress, search result counts, and errors are logged to the console. The headless CLI logs to stdout, which the nightly job captures in its `output.log`. Raise verbosity by changing the `logging.basicConfig` level in `foldersearcher.py`.

## License

MIT, as part of this repository.
