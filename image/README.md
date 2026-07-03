# Image Scripts Index

## 🚀 Overview

Index for the `image/` scripts that don't have their own dedicated doc yet: an Affinity-file/PNG pair checker, an illustration-subscriber distributor, an image-to-PDF collator, a HEIC→JPG converter, an Imgur rate-limit checker, and a folder file-lister. Other scripts in this folder (e.g. `photos_archive.py`) already have their own doc — see the pointer list at the bottom.

## 🖼️ illustrations_check.py

Scans a folder for `*.afdesign` and `*.png` files and reports whether every Affinity Designer file has a matching PNG (and vice versa), plus any other "spare" files in the folder.

### Usage

```bash
# Uses source_folder from Illustrations_check.json
python image/illustrations_check.py

# Or pass the folder explicitly (overrides the config)
python image/illustrations_check.py "C:/path/to/folder"
```

If no folder is given on the command line and the config has no `source_folder`, the script prompts interactively (`Enter source folder path:`).

There is also a `illustrations_check.bat` launcher that reads `VENV_FOLDER` from `E:\automation\automation\.env` (falling back to `E:\automation\automation\.venv`), `cd`s into the `image` folder, and runs `python.exe illustrations_check.py %*` — any arguments passed to the `.bat` are forwarded to the script.

### Configuration — `Illustrations_check.json`

Single-key config file, read from the same directory as the script:

- **source_folder**: folder path to scan for `.afdesign`/`.png` pairs. Overridden by a `source_folder` CLI argument if one is given.

### Output

Prints counts of `.afdesign` and `.png` files, matched pairs, orphan `.afdesign` (no matching `.png`), orphan `.png` (no matching `.afdesign`), and any other files in the folder that are neither extension. Exits with code 1 if the folder is missing or not a directory, otherwise 0.

## 📬 illustrations_subscribers.py

Distributes archived illustrations to subscriber folders based on two Excel metadata files, and can optionally embed a hidden ownership message into each copied PNG using LSB steganography.

### Usage

```bash
python image/illustrations_subscribers.py [options]
```

### Command-line arguments

- `--config PATH` (default: `illustrations_subscribers.json`, resolved next to the script if relative): path to the JSON config file
- `--subscribers PATH`: overrides `metadata_subscribers` from the config
- `--illustrations PATH`: overrides `metadata_illustrations` from the config
- `--source PATH`: overrides `metadata_source_folder` from the config
- `--archive PATH`: overrides `metadata_archive_folder` from the config
- `--encode {y,n}`: whether to encode the ownership message into copied files; if omitted, the script prompts interactively
- `--log-level {DEBUG,INFO,WARNING,ERROR,CRITICAL}` (default: `INFO`)

### Configuration — `illustrations_subscribers.json`

- **metadata_subscribers**: path to the subscribers Excel file (read with `pandas.read_excel`)
- **metadata_archive_folder**: root folder where per-subscriber subfolders are created/populated
- **metadata_illustrations**: path to the illustrations metadata Excel file (`database-dump-illustrations_clean.xlsx`)
- **metadata_source_folder**: folder containing the source `.png` illustrations
- **log_level**: logging level string

Any of the four path keys and `log_level` can be overridden by the matching CLI flag; paths are normalized with `os.path.normpath` after merging.

### What it does

1. Loads the subscribers and illustrations Excel files into DataFrames (retries with a prompt on `FileNotFoundError`/`PermissionError`/other load errors).
2. Filters subscribers to rows whose `sub type` is `Wanted illustrations - active` or `Wanted illustrations - gift` **and** whose `update` column equals today's date.
3. For each matching subscriber, creates `metadata_archive_folder/<name>/` if missing; if it already has subfolders, interactively asks (`y`/`n`) whether to delete them first.
4. Filters illustrations to rows where `FK_TAGS == 'archived'` and the sum of `N_TIMES_IG + N_TIMES_LI + N_TIMES_TW + N_TIMES_CMT + N_TIMES_PST` is greater than 0.
5. Copies `<DE_ILLUSTRATION>.png` from the source folder into every matched subscriber's folder; for illustrations whose exact filename isn't found, retries with a `<DE_ILLUSTRATION>*.png` glob and copies any matches found.
6. Logs a final report of files processed/copied/skipped.
7. Optionally (via `--encode` or an interactive Y/N prompt) LSB-encodes a fixed ownership message into every `.png` in each subscriber's folder: `"copy for {name}. For personal use only, according to guidelines at https://robertoferraro.art/. No copy or further distribution is allowed"`. A second prompt (or `log_each_file` input) controls whether each encoded file is logged individually.

### Requirements

`pandas` (with an Excel engine such as `openpyxl`), `Pillow`.

## 📄 image_to_pdf_collate.py

Tkinter GUI app: pick a folder and collate every image inside it into a single PDF (one page per image), saved as `<folder_name>.pdf` inside that same folder.

### Usage

```bash
python image/image_to_pdf_collate.py
```

There are no command-line arguments — it opens a small window titled "Image → PDF Collate" with a single "Choose Folder & Create PDF" button. The `image_to_pdf_collate.bat` launcher activates the repo's `.venv` (`E:\automation\automation\.venv\Scripts\activate.bat`) and runs the script with no arguments.

### Behavior

- Collects files directly inside the chosen folder (non-recursive) whose extension is one of `.jpg .jpeg .png .bmp .gif .tiff .tif .webp .heic .heif`, sorted by filename.
- Converts `RGBA`/`LA`/`P` images onto a white background and anything else to `RGB` before adding it as a PDF page.
- Shows an error dialog if no images are found or if PDF creation raises an exception; shows a success dialog with the output path otherwise.

### Requirements

`Pillow`, `tkinter` (stdlib).

## 🔄 heic_to_jpg_converter.py

Tkinter GUI app: pick a folder and convert every `.heic`/`.heif` file in it (top level only, not recursive) to a `.jpg` in the same folder, same base filename.

### Usage

```bash
python image/heic_to_jpg_converter.py
```

No command-line arguments — it opens a folder-picker dialog ("Select folder containing HEIC images") and then processes that folder.

### Behavior

- Registers `pillow-heif`'s HEIF opener with PIL at import time; if `pillow-heif` isn't installed, it logs a warning and continues (conversions will then fail per-file with a clear error).
- Validates each `.heic`/`.heif` file first with `Image.verify()`, skipping files that don't validate.
- Skips a file if a same-named `.jpg` already exists (never overwrites).
- Converts to `RGB` if the mode isn't already `RGB`/`L`, then saves with `quality=95, optimize=True, progressive=True`.
- Logs a final summary (converted count, error count) and shows a completion message box.

### Requirements

`Pillow`, `pillow-heif`, `tkinter` (stdlib).

## 📊 imgur_check_rate.py

Prints the current Imgur API rate-limit status for the configured account.

### Usage

```bash
python image/imgur_check_rate.py
```

No command-line arguments.

### Requirements

- Environment variable `IMGUR_ACCESS_TOKEN` (loaded via `python-dotenv`'s `load_dotenv()`), used as a `Bearer` token against `https://api.imgur.com/3/credits`.
- `requests`, `python-dotenv`.

### Output

On success, logs remaining per-hour client requests (`ClientRemaining`/`ClientLimit`), remaining per-day user uploads (`UserRemaining`/`UserLimit`), and the reset countdown (`UserReset` seconds, converted to minutes). Logs an error with the HTTP status and response body on failure, or an error if `IMGUR_ACCESS_TOKEN` isn't set.

## 🗂️ files_list.py

Tkinter GUI app: pick a folder and recursively list every file in it (name + modification date), saving the result as JSON back into that same folder.

### Usage

```bash
python image/files_list.py
```

No command-line arguments — opens a folder-picker dialog ("Select folder to list files").

### Behavior

- Walks the selected folder recursively (`Path.rglob('*')`), collecting `{"filename": ..., "modified_date": ...}` for every file (`modified_date` is `datetime.fromtimestamp(stat.st_mtime).isoformat()`).
- Sorts the results by `modified_date` descending (newest first).
- Writes the list to `files_list_<folder_name>_<YYYYmmdd_HHMMSS>.json` inside the scanned folder (`json.dump(..., indent=2, ensure_ascii=False)`).
- Shows a message box with "No Files Found" if the folder is empty, or a success message box with the output filename and file count otherwise.

### Requirements

`tkinter` (stdlib).

## 📚 Already documented elsewhere

- `photos_archive.py` — see `photos_archive.md` in this same folder.
