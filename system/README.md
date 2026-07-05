# System Scripts

Index of the general-purpose Windows utility scripts in `system/` that don't yet have their own doc. Each is a standalone script — run it directly with `python system/<script>.py [args]` from the repo root, or via its launcher `.bat` where one exists.

## 🎨 background.py

Sets the Windows desktop background color and matching taskbar/accent theme via the registry and `ctypes` calls into `user32` — no wallpaper image, just a flat `COLOR_DESKTOP` fill plus light/dark app theme and accent color.

- **How to run:** `python system/background.py <theme>` where `<theme>` is a single positional argument, either `black` or `"light grey"` (quote it — it contains a space). No argument defaults to `"light grey"`. Any other value raises `ValueError: Theme must be either 'black' or 'light grey'`.
- **Launchers:** `background_black.bat` and `background_grey.bat` call `python background.py black` / `python background.py "light grey"` respectively — double-click either, no args needed.
- **Requirements:** Windows only, standard library only (`ctypes`, `winreg`).

## ⌨️ keycaster.py

A small PySimpleGUI app that "types" pasted text into the last active Google Chrome window, character by character, using the `keyboard` library — useful for pasting into inputs that block real clipboard paste. Newlines in the text are sent as `shift+enter` (soft newline) rather than `enter`.

- **How to run:** `python system/keycaster.py` — no command-line arguments; everything is driven from the GUI (a multiline text box, a typing-delay-in-ms field defaulting to `1`, and Start/Stop buttons). Start activates the most recently active non-minimized Chrome window, waits 0.2s, then types.
- **Launcher:** `keycaster.bat` activates a dedicated virtualenv hardcoded at `E:\onedrive\Documentos\Roberto\projects\automation\notion-automation\local\.venv_keycaster`, runs `python keycaster.py`, then pauses so errors stay visible — that venv path is a leftover from an older repo location and may need updating before the launcher works.
- **Requirements:** `PySimpleGUI`, `pygetwindow`, `pyautogui`, `keyboard`. Windows only (targets a Chrome window by title).

## 🔐 unzip_with_password.py

Extracts a password-protected (or plain) `.zip`, `.7z`, or `.rar` archive to a new folder named after the archive (same parent directory, name minus extension). Tries Python libraries first (`zipfile`, then `py7zr`, then `rarfile` depending on extension), and falls back to the `7z` CLI if those aren't installed or fail.

- **How to run:** `python system/unzip_with_password.py [--archive-path PATH] [--password PASSWORD] [--verbose|-v] [--check-deps]`
  - `--archive-path`: path to the archive. If omitted, a tkinter file-picker dialog opens.
  - `--password`: archive password. If omitted, prompts interactively via `input()` (press Enter for no password).
  - `--verbose` / `-v`: enable debug-level logging.
  - `--check-deps`: just report which optional libraries (`py7zr`, `rarfile`) are installed, then exit — does not extract anything.
- **7-Zip fallback:** looks for `7z` on `PATH`, then `C:\Program Files\7-Zip\7z.exe`, then the `(x86)` variant.
- **Requirements:** `py7zr` and `rarfile` are optional (only needed for `.7z`/`.rar` via the Python path — `rarfile` additionally needs WinRAR/UnRAR installed); the 7-Zip CLI fallback needs 7-Zip installed for any format.

## 🏷️ rename_files.py

Walks a directory tree and strips a trailing version tag (`v01`, `V02`, … any `[vV]` + 2 digits right before the extension) from every filename that has one, e.g. `report_v02.docx` → `report.docx`. If the de-versioned name already exists, the file is skipped (logged) rather than overwritten.

- **How to run:** `python system/rename_files.py` — no command-line flags. The script prompts for the directory path with a plain `input('Enter the directory path: ')` call, and that call runs at module level (not inside a `main()`/`__main__` guard), so it fires as soon as the file is executed or imported.
- **Requirements:** standard library only (`os`, `re`, `logging`).

## 📁 copy_git_project.py

Copies a git project directory to a new location while honoring the source's `.gitignore` (via `pathspec`) and always skipping any `.git` directory, so the result is a clean copy without version-control internals or ignored build artifacts.

- **How to run:** `python system/copy_git_project.py [source] [destination]` — both are optional positional arguments; any omitted one falls back to an interactive prompt (`input()`, with surrounding quotes stripped). Aborts with an error if source and destination resolve to the same path, if the source has no `.gitignore`, or if the destination already exists and is non-empty.
- **Requirements:** `pathspec` (the script exits immediately with an install hint if it's missing).

## 📝 word_to_markdown.py

Converts a `.docx` Word document to Markdown, preserving heading levels (`Heading 1`–`6` and `Title` styles), list-style paragraphs (rendered as `- item`), plain paragraphs, and tables (rendered as Markdown pipe tables).

- **How to run:** `python system/word_to_markdown.py [file_path] [--test] [--gui] [--debug]`
  - `file_path`: optional positional path to a single `.docx` file; converts it and writes a sibling file with the same name and a `.md` extension.
  - `--test`: finds every `.docx` under the project root (detected by walking up from the current directory to the first ancestor containing `.git`), copies each into `system/test-word-to-markdown/`, converts them all, and writes the `.md` output alongside the copies in that same test folder.
  - `--gui`: launches a small Tkinter window with "Select File" and "Run Test Conversion" buttons (requires Tkinter).
  - `--debug`: enables debug-level logging (verbose per-paragraph style tracing).
  - With no arguments at all: launches the Tkinter GUI if available, otherwise prints `--help`.
- **Requirements:** `python-docx` (hard dependency — the script exits with an install hint if missing); `tkinter` is optional and only needed for `--gui` / the no-argument default.

## 📊 list_files_to_xls.py

Lists the files directly inside a chosen folder (top level only, not recursive) and writes them to an Excel workbook named `lista.xlsx` inside that same folder, with columns for filename, extension, human-readable size, and last-modified timestamp.

- **How to run:** `python system/list_files_to_xls.py` — no command-line arguments. It always opens a Tkinter folder-picker dialog first; there is no way to pass the folder as an argument. On completion it shows a message box confirming the save path and file count.
- **Requirements:** `openpyxl`, `tkinter`.

## 🌡️ hwinfo_restart.py

Restarts HWiNFO64 (a system monitoring app that stops reporting after a continuous 12-hour run) — terminates it if running, then relaunches it from its install path. Intended to be run on a recurring schedule (e.g. every 8 hours) by an external scheduler; the script itself does a single restart cycle per invocation with no internal loop or sleep.

- **How to run:** `python system/hwinfo_restart.py` — no command-line arguments. Looks for the install path at `C:\Program Files\HWiNFO64\HWiNFO64.exe`, falling back to the `(x86)` variant; raises `FileNotFoundError` if neither exists.
- **Launcher:** `hwinfo_restart.bat` calls `python hwinfo_restart.py` — no args needed.
- **Requirements:** `psutil`. Windows only.

## Already documented

These `system/` tools have their own dedicated doc file — see there instead of here:

- [`base64_encode_decode.md`](base64_encode_decode.md)
- [`markdown_preview.md`](markdown_preview.md)
- [`mouse_mover.md`](mouse_mover.md)
- [`venv_manager.md`](venv_manager.md)
- [`open_file_latency_diag.md`](open_file_latency_diag.md)
