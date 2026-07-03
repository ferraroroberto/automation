# Text Tools

## 🚀 Overview

Two standalone scripts for text-file processing: redacting sensitive data out of arbitrary text, and extracting text from PDF files. Neither uses `argparse` — both take plain positional arguments via `sys.argv`, with a GUI or interactive fallback when no arguments are given.

## 🧹 clean_sensitive_data.py

Redacts common categories of sensitive data from text using regex substitution: file paths, email addresses, IP addresses, URLs, secrets/keys, known internal service names, long numeric sequences, and GUIDs/UUIDs. Each category is matched with its own regex and replaced with a `[REDACTED_*]` placeholder; the secrets pattern is a special case that keeps the captured label (e.g. `api_key=`) and replaces only the value with `[REDACTED_SECRET]`.

The script picks a mode automatically based on how it's invoked and whether a display is available:

- **CLI file mode** — triggered whenever any command-line argument is present, regardless of platform/display:
  ```bash
  python text/clean_sensitive_data.py <input_file> [output_file]
  ```
  Reads `<input_file>` (UTF-8), cleans it, and either writes the result to `output_file` (if given) or prints it to stdout. A missing input file or any other read/write error is logged rather than raised.

- **GUI mode** — used with no arguments when a display is available (always true on Windows, since `TKINTER_AVAILABLE` + `sys.platform == 'win32'` is enough to satisfy `has_display()`):
  ```bash
  python text/clean_sensitive_data.py
  ```
  Opens a Tkinter window with side-by-side input/output text panes, Paste/Clear buttons, a "Clean Text" button, a "Copy to Clipboard" button, and one checkbox per redaction category (File Paths, Email Addresses, IP Addresses, URLs, Secrets/Keys, Service Names, Numeric Sequences, GUIDs/UUIDs — all on by default). On startup it also auto-pastes clipboard contents into the input pane if the clipboard has text and the input is empty.

- **CLI interactive mode** — used with no arguments when no display is available (non-Windows, headless):
  Prompts for multi-line text at `>`, with `config` to toggle the same eight categories, `help` for command list, and `quit` to exit.

### Configuration — `cleaning_patterns.json`

Regex patterns are loaded from `cleaning_patterns.json` in the same directory as the script. If the file doesn't exist, the script writes it out with built-in defaults on first run, so it's self-bootstrapping. If it exists but fails to load, the script falls back to the same in-code defaults and logs a warning. The file is a flat JSON object with one regex string per category:

| Key | Matches |
|---|---|
| `paths` | Windows/Unix-style file paths (optional drive letter + separator + non-whitespace run) |
| `emails` | `user@domain.tld`-shaped strings |
| `ips` | dotted-quad IPv4 addresses |
| `urls` | `http://` / `https://` URLs |
| `secrets` | `api_key`/`token`/`secret`/`password`/`pw`/`pwd`/`auth` (case-insensitive) followed by a `=`/`:`/quote-delimited value |
| `services` | a fixed list of internal service names (`internal_db`, `prod_db`, `auth_service`, `user_service`, `admin_service`, `payment_service`) |
| `numbers` | runs of 5+ digits |
| `guids` | standard 8-4-4-4-12 hex GUID/UUID format |

Each category can be toggled independently — the toggle state lives on the `DataCleanerBase` instance (`clean_paths`, `clean_emails`, `clean_ips`, `clean_urls`, `clean_secrets`, `clean_services`, `clean_numbers`, `clean_guids`, all `True` by default) and is exposed as checkboxes in the GUI or numbered toggles in the CLI's `config` command.

### Requirements

- `pyperclip` — clipboard paste/copy (GUI mode only; CLI file mode doesn't import it at call time but the module-level import is skipped gracefully if `tkinter`/`pyperclip` aren't available, which also forces CLI interactive mode)
- `tkinter` (stdlib) — GUI window and its widgets

## 📄 convert_pdf_to_txt.py

Extracts all text from a PDF (page by page, via `PyPDF2.PdfReader`) and writes it to a `.txt` file with the same base name, next to the source PDF. Logs the input path, any extraction/save errors, and the final word count.

```bash
python text/convert_pdf_to_txt.py [path/to/file.pdf]
```

- With a path argument: uses it directly. The script errors out (via `logger.error`, no exception raised) if the path doesn't exist or doesn't end in `.pdf`.
- With no argument: opens a Tkinter "Select a PDF file" dialog filtered to `*.pdf`. If nothing is selected, it errors out the same way.

Output is always `<input-basename>.txt` written next to the input file, UTF-8 encoded, overwriting any existing file of that name.

### Requirements

- `PyPDF2` — PDF parsing and text extraction
- `tkinter` (stdlib) — file-picker dialog when no path is passed on the command line

### Packaging

`convert_pdf_to_txt.spec` is a PyInstaller spec for building this script into a standalone windowed executable (`console=False`, UPX-compressed, no extra data/binaries bundled).
