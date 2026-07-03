# Google

## 🚀 Overview

This folder holds two Google-API automations — Gmail→Drive email tracking and a weekly Google Photos album sharer for the kids — plus the setup and OAuth-scope-testing helpers that get a fresh Google Cloud project working with both of them.

## 📚 Documentation

- [`gmail_drive_automation.md`](gmail_drive_automation.md) — monitors Gmail for emails matching a search query and logs new ones (subject, date, from, content preview) into a Google Sheet in Drive, skipping ones already recorded.
- [`weekly_photo_automation.md`](weekly_photo_automation.md) — creates a per-child Google Photos album for the most recent complete Saturday–Friday week and emails the album link to configured recipients; documents the `albums.share()` 403 limitation and the semi-manual workaround.
- [`diag_check_google_config.md`](diag_check_google_config.md) — describes `diag_check_google_config.py`, a diagnostic that checks gcloud project/account, `client_secret.json`, `token.json` scopes, and live Photos/Gmail/Drive auth in one pass.

## 🔧 Setup & diagnostic helpers

Three scripts in this folder aren't automations themselves — they're the OAuth setup and verification tooling used to get credentials working before (or while debugging) the two automations above. The `diag_*` scripts are one-shot diagnostics, not part of a pytest test tier.

### `setup_helper_script.py`

Interactive first-time setup wizard for the **weekly photo automation**. It checks that `credentials.json` exists in the folder (printing Google Cloud Console instructions if not), then prompts interactively for each child's name and birth date, one or more email recipients (validated for basic `@`/domain shape), and a timezone (default `Europe/Madrid`). It assembles these into a config dict matching `weekly_photo_automation.py`'s expected shape (`auth`, `children`, `email`, `settings` with `max_photos_per_album: 500`, 3 retry attempts), shows a summary including each child's current age in weeks, and on confirmation writes it to `config.json` — backing up any existing `config.json` to `config.backup.<timestamp>.json` first if you say yes.

Run it once when setting up the weekly photo automation for the first time, or re-run it any time you want to regenerate `config.json` from scratch (e.g. adding a new child or recipient) rather than hand-editing the JSON:

```bash
python setup_helper_script.py
```

### `diag_all_scopes.py`

OAuth scope discovery tool. It deletes any existing `token.json`, then runs an `InstalledAppFlow` against `client_secret.json` requesting a large hardcoded list of scopes covering every Photos Library variant (`photoslibrary`, `.readonly`, `.appendonly`, `.sharing`, `.edit.appcreateddata`, `.readonly.appcreateddata`) and every Gmail variant (`send`, `compose`, `modify`, `readonly`, `metadata`, `insert`, `labels`, `settings.basic`, `settings.sharing`, the `addons.*` scopes, and full-access `mail.google.com`). It saves the resulting credentials to a **separate** `token_all_scopes.json` (not `token.json`, so it won't clobber a working token), then actually exercises the Photos API (list/create/share albums, list/search media items) and Gmail API (get profile, list labels, create+delete a draft, list messages) to see which operations the granted scopes actually unlock, and prints a minimal-scope recommendation at the end.

Run it when you don't know which scopes to enable on the OAuth consent screen, or when the real automations are failing with permission errors and you need to isolate exactly which scope combination works. If it reports success, copy the resulting token over the production one:

```bash
python diag_all_scopes.py
# on success:
cp token_all_scopes.json token.json
```

### `diag_gmail_drive.py`

Offline diagnostic suite for `gmail_drive_automation.py` — makes no Google API calls and needs no credentials. It runs four checks: that `config_gmail_drive.json.sample` exists and contains the `auth`/`spreadsheet`/`gmail`/`settings` sections the automation expects; that mock email dicts transform into the correct 4-column row shape with content truncated to 100 characters plus `...`; that a set of "existing subjects" correctly filters out duplicate emails from a mock incoming batch; and that a mock spreadsheet response has the expected `properties`/`sheets` structure with a 4-column grid. Exits 0 only if all four pass.

Run it after changing `gmail_drive_automation.py`'s data-processing logic (row building, dedup, or spreadsheet-shape assumptions), or as a fast sanity check before scheduling a run, since it needs no network access or live credentials:

```bash
python diag_gmail_drive.py
```
