# parking

Parking reservation automation. Slots free up when someone cancels, so a small deterministic poller (no LLM) checks every 5-15 minutes (by time of day) and books any free slot on the days you need, then sends a Telegram message.

## Hard rules

- **Never call `deleteBooking`** or add any cancel path. The API client deliberately exposes only reads and `createBooking`.
- **Never book a day that already has an ACTIVE booking.** One booking per day, one `createBooking` call per day.
- **Never print or log the token** (`PARKING_TOKEN`); only its expiry may appear in logs.
- **`dry_run` defaults to true.** Turning it off is the owner's decision, never an agent's.
- Secrets and identifiers (site URL, API endpoint, plate, login user, chat id, token) live only in the gitignored `parking/.env`; behaviour lives in `parking/config.json`. No organisation name in any tracked file.

## How it works

Each run (`python -m parking.poller`, from the repo root):

1. Load `config.json` and `.env`; skip if not due (backoff / last run too recent) or if the time window says don't poll.
2. Read my bookings. A day with an ACTIVE booking is **never booked again**.
3. Work out the target days (configured weekdays inside the horizon, minus holidays) that have no booking.
4. For each centre x size (in config order, 1.5-3 s jittered pauses) ask which days are bookable.
5. For each uncovered target day with a free slot, book **that one day**, re-read my bookings to confirm, and notify.

No slot found is silent. It alerts on: booked, a free slot whose booking failed, login needed, token expiring (3 days), repeated errors. It stops and backs off (state in `parking/state/state.json`) on 401/403/429/5xx instead of pushing through.

## Setup

```powershell
Copy-Item parking\.env.example parking\.env      # then fill it in
& .\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

`.env` holds the booking URL, API endpoint, plate and notifier settings (see `.env.example`). Telegram goes through the fleet notifier (`fleet-config`'s `notify_send.py`); set `NOTIFY_CHAT` to send to one specific chat id, otherwise `NOTIFY_CATEGORY` decides.

### Monthly login

The site signs in with SSO and issues its own token, valid about 30 days.

```powershell
& .\.venv\Scripts\python.exe -m parking.login
```

A dedicated Chrome window opens (profile in `parking/.browser-profile/`): sign in by hand once. The token is read from the page and written to `parking/.env` as `PARKING_TOKEN`; it is never printed. When it is close to expiring the poller sends a Telegram "login needed" message. `--refresh` re-reads it headlessly from the saved profile.

## Config (`config.json`)

| Key | Meaning |
|---|---|
| `dry_run` | `true` (default): look and report only, never book. Go live locally with `PARKING_DRY_RUN=false` in `parking/.env` (only an explicit false/0/no counts), so the tracked file stays untouched |
| `weekdays` | days you need a spot, e.g. `["mon","thu","fri"]` |
| `horizon_days` | how far ahead to look |
| `holiday_region`, `skip_dates` | days never booked: the `holidays` calendar for that country/subdivision (ES/CT) plus your own `skip_dates` for local days the library lacks (e.g. `2026-09-24`, La Mercè). Review the list each year |
| `centers`, `sizes` | ids and names, tried in this order; `type` is `standard` |
| `treat_placeless_as_covered` | `true`: a day with an ACTIVE booking that has no slot number counts as booked |
| `poll_minutes`, `poll_windows`, `timezone` | default minutes between polls; ordered `{start, end, every_minutes}` windows override it (`0` = don't poll, end exclusive, may wrap midnight; first match wins); clock for both |
| `jitter_max_seconds`, `request_pause_seconds` | politeness toward the site |

## Schedule

Registered in the app-launcher Jobs tab as `parking-poll`, a 5-minute heartbeat running `parking/run-poll.bat` (hidden window). The launcher only supports a flat interval, so the poller decides per run whether it is due: `poll_windows` picks the interval for the time of day and each successful run sets `next_allowed` (the same field failure backoff uses); runs that arrive early exit before any network call. Nothing runs on the Mac Mini.

## Tests

```powershell
& .\.venv\Scripts\python.exe -m pytest -q parking
```

Part of the repo's `scripts/verify-before-ship.ps1` gate.
