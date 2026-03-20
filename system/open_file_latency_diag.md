# Windows file-open latency diagnostic

Small utility script: [`open_file_latency_diag.py`](open_file_latency_diag.py).

## Why this exists

On Windows, attaching or opening a file through a GUI **Open** dialog can *feel* slower than “just reading the file” from a terminal. Common hypotheses include **shell extensions** (thumbnails, overlays, property handlers), **security software**, **sync clients** (e.g. cloud placeholders), or **storage** (slow volume, sleep, network-backed paths).

This script gives **two timed measurements** for the **same path** so you can compare **raw I/O** vs **Shell-oriented work** in one run.

## What it measures

| Phase | What happens |
|--------|----------------|
| **Raw read** | Reads the **entire file** into memory (`Path.read_bytes()`). Dominated by **bytes moved** and **cache state** (cold vs warm). |
| **Shell simulation** | Uses **pywin32** / Shell APIs: create a shell item for the path, display names, attributes, **icon + overlay** resolution (`SHGetFileInfo`), and a full **property store** walk (`SHGetPropertyStoreFromParsingName` + `GetValue` per key). |

It does **not** show a real `IFileOpenDialog` window. It still exercises much of the Shell stack that runs when Explorer or a picker **touches** a file (metadata, handlers, overlays).

## Requirements

- **Windows**
- **Python 3** with **`pywin32`** (see project `requirements.txt`)
- **`tkinter`** (usually included with Windows Python builds) — used for an optional file picker when you run without `--file`

## How to run

From the repo root, using the project virtualenv:

```text
.venv\Scripts\python.exe system\open_file_latency_diag.py
```

- Omit `--file` → a file picker opens.
- Or pass a path explicitly:

```text
.venv\Scripts\python.exe system\open_file_latency_diag.py --file "D:\path\to\file.pptx"
```

Optional: more runs for mean/stdev:

```text
.venv\Scripts\python.exe system\open_file_latency_diag.py --iterations 5
```

## Reading the results (important)

### Apples vs oranges on **large** files

The **raw** test always reads **the whole file**. For a **very large** document (e.g. a big Office archive), the **first** run can take hundreds of milliseconds or seconds depending on **disk**, **filters** (sync/AV), and whether data is already **cached**. **Later runs** are often much faster because the OS **page cache** is warm.

The **Shell** phase mostly does **metadata and handler work**; it does **not** necessarily read every byte of a huge file. So for large targets you may see **raw mean ≫ shell mean** — that does **not** disprove extension or AV issues; it often means **“full read cost”** dominated the comparison.

Use a **small** test file on the same volume if you want the ratio to reflect **Shell overhead** more clearly.

### When **shell ≫ raw** on a **small** file

If the file is small but **Shell** is still **many times** slower than **raw**, that pattern is more consistent with **extra work in the Shell layer** (handlers, overlays, property enumeration, or filters on those APIs). It still isn’t a full reproduction of a browser’s picker (the browser may add its own delay).

### PowerShell caveat (historical)

Measuring “read speed” with `Get-Content` on **binary** Office files is **misleading**: line-oriented text processing over a large binary blob is not the same as a **raw byte read**. For a fair CLI baseline on Windows, prefer something like `[System.IO.File]::ReadAllBytes($path)` or this script’s raw phase.

## If the delay is only in the browser

If the script looks fine but the browser still pauses:

- Try another browser or an incognito profile (extensions).
- Capture a short trace with **Process Monitor** filtered to the process and path.
- Consider **ShellExView** (disable non-Microsoft shell extensions temporarily, restart Explorer) as a structured isolation step — after backing up / noting what you change.

## License / scope

This is a **diagnostic aid**, not a guarantee of root cause. Interpret numbers in context of **file size**, **cache**, **volume type**, and **sync/AV** stack.
