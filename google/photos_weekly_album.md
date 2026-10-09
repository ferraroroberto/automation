# Weekly Google Photos album (`photos_weekly_album.py`)

Creates the week's Google Photos album by driving the Google Photos web UI with Playwright in real Chrome. It is deterministic, with no LLM in the loop. The week runs Saturday to Friday. Screenshots and screen recordings are left out. Sharing the album and emailing the link are later steps and are not done here.

Why a browser: since 31 March 2025 the Photos Library API only sees items the calling app uploaded, and it can't share albums. The Picker API needs a person to pick the items. Neither can build this album unattended. The older `weekly_photo_automation.py` in this folder is built on the Library API and no longer works for this job.

## Setup

1. Copy `photos_weekly_album.json.sample` to `photos_weekly_album.json` (gitignored) and fill it in:
   - `title_template`: the album title, with a `{name}` placeholder for each counter.
   - `counters`: for each placeholder, its `value` for the week starting on `week_start`, which must be a Saturday. Each counter goes up by one per Saturday-to-Friday week.
   - `profile_dir` (optional): the dedicated Chrome profile. Default `~/.photos-weekly-album-profile`, outside the repo.
   - `log_file` (optional): the run log. Default `google/photos_weekly_album.log` (gitignored by `*.log`).
2. Sign the dedicated profile in once, by hand, from the repo root:

   ```powershell
   & .\.venv\Scripts\python.exe -m google.photos_weekly_album --login
   ```

   A Chrome window opens on Google Photos. Sign in (including 2FA). The window closes once the library has loaded, with a 10-minute limit.

## Usage

Run from the repo root:

```powershell
& .\.venv\Scripts\python.exe -m google.photos_weekly_album --dry-run    # report only, creates nothing
& .\.venv\Scripts\python.exe -m google.photos_weekly_album              # create the album, print its URL
& .\.venv\Scripts\python.exe -m google.photos_weekly_album --from 2026-10-03 --to 2026-10-09 --dry-run
```

- With no `--from`/`--to`, the week is the most recent complete Saturday-to-Friday week. A Friday run takes the week before, because the current one isn't over. Run it on Saturday morning.
- `--dry-run` prints the week, the title, the items per day, the screenshots left out and the total. It selects items in the library to count them and then clears the selection.
- If an album with the computed title already exists, a real run reports it and exits 0 without creating anything. A dry run reports the existing album and still counts.
- Exit codes: `0` for done, no-op or dry run; `1` for a flow error or count mismatch; `2` for a config error; `3` when the profile isn't signed in.
- Every run appends one JSON line to the run log: week, title, per-day counts, screenshots excluded, selected count, album URL and count, and the outcome.

## Flow

1. Launch real Chrome (`channel="chrome"`) on the dedicated profile through the repo-root `browser_stealth.py` helper. It applies the fleet's human-like settings (1280×900, `navigator.webdriver` hidden, no automation infobar, sandbox on). If the profile is in use by another Chrome window, the helper waits 60/120/240/480 s and never kills it. The window is visible, because Google Photos doesn't render the grid in a hidden window.
2. On `/albums`, look for a card whose title equals the computed title. If one exists, stop (or, in a dry run, note it).
3. Open the **Screenshots & recordings** collection (its link is on `/albums`) and collect the labels of the week's items.
4. On the library timeline, scroll to each day of the week and click its day-header select-all checkbox. The checkbox only appears while the pointer is over the header row, so the script hovers the row first. Each click is a real pointer click, and the script checks `aria-checked="true"` afterwards, retrying once. The `N selected` counter is read before and after each click to get that day's count. A day header selects the whole day even when most of its tiles haven't rendered.
5. Sweep the week's tiles again and untick each one whose label matches a screenshot. The collection and the timeline give the same item different ids, so tiles are matched by label: type, orientation and the second the photo was taken.
6. Check completeness: the tiles seen in the week must equal the sum of the per-day counts (the sweep covered everything), and the counter must equal that sum minus the unticked screenshots.
7. Toolbar **Add to album** → **Album** → **New album**, type the title into `Edit album name`, then click **Done**. Finally, re-read the album card's item count on `/albums` and compare it with the selection.

## Selectors (English UI)

| What | Selector |
|---|---|
| Day select-all | `[role=checkbox][aria-label^="Select all photos from "]`, state in `aria-checked` |
| Item tile | `a[href*="/photo/"]`, `aria-label` like `Photo - Landscape - Oct 3, 2026, 8:23:55 AM` |
| Item checkbox | the `[role=checkbox]` under the tile with the same `aria-label` |
| Selection count | text `N selected` |
| Scroll container | the `c-wiz` with the largest `scrollHeight` (the window doesn't scroll) |
| Album list | `/albums`, card link `a[href*="/album/"]`, first line the title, then `N items` |

Day-header labels are relative ("Today", "Wednesday", "Sat, Oct 3"), so each header is dated by the first tile after it, never by its own text.

## Failure modes

| Symptom | Cause and handling |
|---|---|
| `not signed in` (exit 3) | The profile's Google session is missing or expired. Run `--login` again. |
| `window is hidden or minimised` | Chrome doesn't render the grid in a minimised window. Covered windows are fine because native occlusion detection is turned off for this profile. Keep the window on screen and re-run. |
| `Chrome profile ... still in use` | Another Chrome window has the dedicated profile open. Close it. The script waits about 15 minutes before giving up and never kills it. |
| `Locator.click: Timeout ... element is not visible` | A day's select-all checkbox has no size until the pointer is over its header row. The script hovers the row first when the checkbox isn't shown. If this comes back, Google changed what reveals the checkbox. Nothing has been created at that point. |
| `checkbox did not change after two clicks` | A click was dropped (seen once after a fresh navigation in testing) or the page changed. Re-run. Nothing has been created at that point. |
| `saw N tiles in the week but the day headers selected M` | The sweep didn't render every tile, so a screenshot might stay in. Nothing is created. Re-run with the window in front. |
| `selection counter reads ...` | An untick or a day click didn't land. Nothing is created. Re-run. |
| `share a timestamp label` | Two items from the same second with the same type and orientation can't be told apart, so the run aborts. Create that week's album by hand. |
| `screenshots not in the timeline` (info) | An item in the collection doesn't appear in the timeline (for example it's archived), so it was never selected. |
| `album created but it lists N items` (exit 1) | The album card's count differs from the selection after three reads. Check the album by hand. |
| Late items | Anything taken late on Friday or synced late from another device after the run is missed. Run on Saturday morning. |
| Google changes the UI | The selectors above stop matching, so the scans find nothing or the clicks time out. Update the selectors in the module. |

The UI must be in English: the selectors and the date labels are matched on English text.
