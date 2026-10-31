# Google

## 🚀 Overview

This folder holds two Google automations — Gmail→Drive email tracking (Google APIs) and a weekly Google Photos album sharer for the kids (web UIs) — plus the diagnostic helpers used with them.

## 📚 Documentation

- [`gmail_drive_automation.md`](gmail_drive_automation.md) — monitors Gmail for emails matching a search query and logs new ones (subject, date, from, content preview) into a Google Sheet in Drive, skipping ones already recorded.
- [`photos_weekly_album.md`](photos_weekly_album.md) — builds the week's (Saturday–Friday) Google Photos album by driving the Photos web UI with Playwright in real Chrome, leaving screenshots out, then shares it by link and saves a Gmail draft to the configured list (sent only with `--send`; Gmail helper: `gmail_web.py`); run as `python -m google.photos_weekly_album` from the repo root, or on Saturday mornings by the app-launcher job `run_photos_weekly_album.bat` (`photos_weekly_album_job.py`), which sends a Telegram message with the outcome. Replaces the earlier Photos Library API script, which stopped working in March 2025 when the API could no longer see the whole library.
- [`diag_check_google_config.md`](diag_check_google_config.md) — describes `diag_check_google_config.py`, a diagnostic that checks gcloud project/account, `client_secret.json`, `token.json` scopes, and live Photos/Gmail/Drive auth in one pass.

## 🔧 Setup & diagnostic helpers

The `diag_*` scripts in this folder aren't automations themselves — they're one-shot diagnostics, not part of a pytest test tier.

### `diag_gmail_drive.py`

Offline diagnostic suite for `gmail_drive_automation.py` — makes no Google API calls and needs no credentials. It runs three checks: that `config_gmail_drive.json.sample` exists and contains the `auth`/`spreadsheet`/`gmail`/`settings` sections the automation expects; that the module's real `build_sheet_row()` turns mock email dicts into the correct 4-column row with content truncated to 100 characters plus `...`; and that its real `filter_new_emails()` drops duplicate emails (by stripped subject, header row ignored) from a mock incoming batch. Exits 0 only if all three pass.

Run it after changing `gmail_drive_automation.py`'s data-processing logic (row building, dedup, or spreadsheet-shape assumptions), or as a fast sanity check before scheduling a run, since it needs no network access or live credentials:

```bash
python diag_gmail_drive.py
```
