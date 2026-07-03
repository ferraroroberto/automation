# LinkedIn Profile Opener

## 🚀 Overview

Opens a batch of LinkedIn profile/company activity pages from an Excel export, one browser tab per row, so a backlog of contacts can be skimmed for recent posts without manually pasting each URL. A Tkinter dialog lets you filter the rows (alert flag, star flag, circle, topic) and choose how many tabs to open at a time before pausing for a "Continue" click.

## 📋 Usage

### Basic Usage

```bash
python linkedin_open.py
python linkedin_open.py --config custom_config.json
python linkedin_open.py --verbose --chunk-size 5
```

Launching the script opens the "LinkedIn Profile Opener - Filter Options" window; there is no headless/CLI-only path — filter selection and processing both happen inside that dialog.

### Command Line Arguments

- `--config` (default: `linkedin_open.json` in the script's own directory): path to the JSON configuration file
- `--verbose`, `-v`: enable debug-level logging
- `--chunk-size`: number of profiles to open at once, overrides `default_chunk_size` from the config

### Launcher

`launcher.bat` reads `VENV_FOLDER` out of `E:\automation\automation\.env` (falling back to `E:\automation\automation\.venv` if that line is missing), `cd`s into `E:\automation\automation\linkedin\open_profiles`, and runs:

```
"%VENV_DIR%\Scripts\python.exe" linkedin_open.py
```

i.e. it always uses the default config path (`linkedin_open.json` next to the script) with no extra flags. On a non-zero exit code it prints an error and pauses so the console stays open to read the logs.

## 🔧 Configuration

The tool uses a JSON configuration file (`linkedin_open.json`, gitignored — see `linkedin_open.example.json` for the schema) with the following structure:

```json
{
    "file_path": "path/to/your/database-dump-connections.xlsx",
    "verbose": false,
    "default_chunk_size": 10,
    "filter_options": {
        "alert_choices": ["True", "False", "Any"],
        "star_choices": ["True", "False", "Any"],
        "circle_choices": ["Any", "Top5", "Key50", "Vital100", "None"],
        "topic_choices": ["all", "innovation", "personal development", "leadership and management", "LinkedIn", "visual illustration"]
    }
}
```

### Configuration Fields

- **file_path**: path to the Excel workbook (read with `pandas.read_excel`) that holds the contact/profile data
- **verbose**: when `true`, logs each rewritten LinkedIn URL at debug level
- **default_chunk_size**: how many profiles the dialog's "Profiles to Open at Once" field defaults to; overridable per run via `--chunk-size`
- **filter_options.alert_choices**: values populated into the "Alert Filter (IND_ALERT)" dropdown (default selection is `True`)
- **filter_options.star_choices**: values populated into the "Star Filter (IND_STAR)" dropdown (default selection is `False`)
- **filter_options.circle_choices**: values populated into the "Circle Filter" dropdown (default selection is `Any`)
- **filter_options.topic_choices**: values populated into the "Topic Filter (FK_TOPIC)" dropdown (default selection is `all`)

If `filter_options` (or one of its sub-keys) is missing, the script falls back to the same value lists shown above, hardcoded in `load_defaults()`.

## 📊 Expected Excel Columns

The workbook loaded from `file_path` must contain:

- `URL_LINKEDIN`: the LinkedIn profile or company URL to open
- `IND_ALERT`: boolean-like column matched against the Alert filter
- `IND_STAR`: boolean-like column matched against the Star filter
- `FK_CIRCLE`: column matched against the Circle filter (`None` in the dropdown matches rows where this is null)
- `FK_TOPIC`: column matched against the Topic filter, also used as the primary sort key
- `DE_PERSON`: person's display name, used as the secondary sort key and shown in the log line for each opened URL

## 🔍 How It Works

1. **Load & sort**: reads the Excel file and sorts rows by `FK_TOPIC` then `DE_PERSON`
2. **Filter**: keeps rows where `URL_LINKEDIN` is not null, then narrows by the selected Alert / Star / Circle / Topic values (a dropdown value of `Any` skips that filter; Circle's `None` matches null `FK_CIRCLE`)
3. **Rewrite URLs** (`update_linkedin_url`): if the URL contains `/company/`, appends `posts/?feedView=all`; otherwise appends `recent-activity/shares/` — so each tab lands on the profile's/company's recent activity feed rather than the bare profile page
4. **Open in chunks**: opens `webbrowser.open(url)` for `chunk_size` rows at a time, logging `Opening URL 00N (topic - person)` for each; when more rows remain, the dialog's button turns into "Continue (N more)" and processing pauses until you click it

## 📚 Dependencies

- `pandas`: reads the Excel workbook
- `easygui`: file-access error message box
- `tkinter` (standard library): the filter/progress dialog
- `webbrowser` (standard library): opens each LinkedIn URL in the default browser
