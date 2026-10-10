# Weekly Google Photos album (`photos_weekly_album.py`)

Creates the week's Google Photos album by driving the Google Photos web UI with Playwright in real Chrome, then shares it by link and saves a Gmail draft with the link to a configured list. It is deterministic, with no LLM in the loop. The week runs Saturday to Friday. Screenshots and screen recordings are left out. Nothing is sent unless asked: `--send`, or `"send": true` in the config.

Why a browser: since 31 March 2025 the Photos Library API only sees items the calling app uploaded, and it can't share albums. The Picker API needs a person to pick the items. Neither can build this album unattended. The older `weekly_photo_automation.py` in this folder is built on the Library API and no longer works for this job.

## Setup

1. Copy `photos_weekly_album.json.sample` to `photos_weekly_album.json` (gitignored) and fill it in:
   - `title_template`: the album title, with a `{name}` placeholder for each counter.
   - `counters`: for each placeholder, its `value` for the week starting on `week_start`, which must be a Saturday. Each counter goes up by one per Saturday-to-Friday week.
   - `profile_dir` (optional): the dedicated Chrome profile. Default `~/.photos-weekly-album-profile`, outside the repo.
   - `log_file` (optional): the run log. Default `google/photos_weekly_album.log` (gitignored by `*.log`).
   - `email` (optional): with no `email.to`, the run stops after the album, without sharing it or drafting anything.
     - `to`: a list of plain addresses, e.g. `["first.person@example.com", "second.person@example.com"]`. No display names, and no address twice.
     - `subject`: the subject template. It must use `{title}`, `{week_start}` or `{week_end}`, because each week's message is found again by its exact subject.
     - `body`: the body template, with `\n` for line breaks. It must contain `{link}`, and it may also use `{title}`, `{week_start}` and `{week_end}` (dates as `YYYY-MM-DD`).
     - `send`: `false` by default. `true` makes every run, a scheduled one included, send the week's draft. Leave it `false` unless you mean that.
2. Sign the dedicated profile in once, by hand, from the repo root:

   ```powershell
   & .\.venv\Scripts\python.exe -m google.photos_weekly_album --login
   ```

   A Chrome window opens on Google Photos. Sign in (including 2FA). The window closes once the library has loaded, with a 10-minute limit. The same Google sign-in covers Gmail. The draft goes out from this account.

## Usage

Run from the repo root:

```powershell
& .\.venv\Scripts\python.exe -m google.photos_weekly_album --dry-run    # report only, creates nothing
& .\.venv\Scripts\python.exe -m google.photos_weekly_album              # album + share link + Gmail draft
& .\.venv\Scripts\python.exe -m google.photos_weekly_album --send       # ... and send this week's draft
& .\.venv\Scripts\python.exe -m google.photos_weekly_album --from 2026-10-03 --to 2026-10-09 --dry-run
```

- With no `--from`/`--to`, the week is the most recent complete Saturday-to-Friday week. A Friday run takes the week before, because the current one isn't over. Run it on Saturday morning.
- `--dry-run` prints the week, the title, the items per day, the screenshots left out and the total. It selects items in the library to count them and then clears the selection.
- If an album with the computed title already exists, a real run reports it and doesn't create another. With `email.to` set, it carries on to the share link and the draft for that album. A dry run reports the existing album and still counts.
- `--dry-run` also prints the subject a real run would draft and the number of recipients. It doesn't share anything or open Gmail.
- Re-runs are safe. An album has at most one share link, which is reused. If the week's subject is already in Drafts, no second draft is made. If it is already in Sent, the run neither drafts nor sends.
- `--send` sends only the one draft whose subject is this week's, and only after checking that its body carries this run's share link and its recipients are the configured list. Without `--send` (and with `"send": false`), nothing is ever sent.
- Exit codes: `0` for done, no-op or dry run; `1` for a flow error or count mismatch; `2` for a config error; `3` when the profile isn't signed in; `4` when the Chrome window is hidden or minimised; `5` when a Google page didn't respond in time (a wait ran out, usually because Google changed its page).
- Every run appends one JSON line to the run log: week, title, per-day counts, screenshots excluded, selected count, album URL and count, the outcome, and (when email is configured) the share link and the mail outcome: `drafted`, `draft-exists`, `sent`, `already-sent` or `error`.

## Flow

1. Launch real Chrome (`channel="chrome"`) on the dedicated profile through the repo-root `browser_stealth.py` helper. It applies the fleet's human-like settings (1280×900, `navigator.webdriver` hidden, no automation infobar, sandbox on). If the profile is in use by another Chrome window, the helper waits 60/120/240/480 s and never kills it. The window is visible, because Google Photos doesn't render the grid in a hidden window.
2. On `/albums`, look for a card whose title equals the computed title. If one exists, stop (or, in a dry run, note it).
3. Open the **Screenshots & recordings** collection (its link is on `/albums`) and collect the labels of the week's items.
4. On the library timeline, scroll to each day of the week and click its day-header select-all checkbox. The checkbox only appears while the pointer is over the header row, so the script hovers the row first. Each click is a real pointer click, and the script checks `aria-checked="true"` afterwards, retrying once. The `N selected` counter is read before and after each click to get that day's count. A day header selects the whole day even when most of its tiles haven't rendered.
5. Sweep the week's tiles again and untick each one whose label matches a screenshot. The collection and the timeline give the same item different ids, so tiles are matched by label: type, orientation and the second the photo was taken.
6. Check completeness: the tiles seen in the week must equal the sum of the per-day counts (the sweep covered everything), and the counter must equal that sum minus the unticked screenshots.
7. Toolbar **Add to album** → **Album** → **New album**, type the title into `Edit album name`, then click **Done**. Finally, re-read the album card's item count on `/albums` and compare it with the selection.
8. (Only with `email.to`.) Open the album (a link-shared one redirects to `/share/<id>?key=...`; the script waits for that before clicking) and click **Share**. If the dialog offers **Create link**, click it: a second dialog, **Create link to share**, opens, and its own **Create link** button makes the link and then shows it in an input. Once a link exists, the Share dialog offers **Copy link** instead, so a re-run can't make a second link. The script takes the first link it finds in the dialogs (the confirmation's input, or the Share dialog's markup once a link exists), Google's responses, the page URL, or the clipboard after **Copy link**, which overwrites your clipboard. If it finds none, it logs the sources it tried and each dialog's buttons.
9. In Gmail (same profile and tab), search `in:sent subject:"<subject>"` and `in:draft subject:"<subject>"` and keep only rows whose subject is exactly the week's. Gmail's own subject search is fuzzy. If the subject is already in Sent, stop. If it is in Drafts once, reuse that draft. If it isn't there, **Compose**: type each address and press Tab, which turns it into a chip and closes the suggestion list that would otherwise cover the Subject field (a trailing comma leaves the text in the field and the list open), check that the recipient chips equal the configured list (otherwise discard and stop), type the subject, insert the body above any signature, then **Save & close** and confirm the draft is in Drafts.
10. (Only with `--send`.) Open the week's single draft, check that its body has the share link and its recipients match, click **Send**, and confirm the subject is now in Sent.

## Scheduled run (every Saturday morning)

`photos_weekly_album_job.py` is what the schedule runs, through `run_photos_weekly_album.bat` (same pattern as the other repo jobs: repo `.venv`, foreground, exit code passed through). It runs the album for the week that just ended, exactly as `python -m google.photos_weekly_album` does with no flags, then sends one short Telegram message through the fleet notifier. It never passes `--send`: mail goes out only when the config says `"send": true`. It shares and drafts only when the config lists `email.to`, as above.

**Needs an unlocked desktop.** Chrome opens a visible window, because Google Photos doesn't render the grid in a hidden one. Before starting Chrome the job checks that the desktop isn't locked and reports it if it is. If nobody is signed in at all, Windows never fires the task (app-launcher's tasks run only while the user is logged on), so nothing runs and no message comes; the Jobs tab shows the job as **not firing**.

| Outcome | Exit | Message (chat) |
|---|---|---|
| Album made | 0 | `Weekly album done: <title>`, the share link (else the album link), `Draft: saved, not sent` / `sent` / `not shared (no recipients configured)` (log) |
| Album already there | 0 | `Weekly album already there: <title>`, the link, the draft status (log) |
| Nothing to album | 0 | `No album this week: nothing to add for <dates>` (log) |
| Profile signed out | 3 | names the `--login` command (attention) |
| Desktop locked | 4 | `Unlock it, then run the job from the Jobs tab` (attention) |
| Chrome window hidden or minimised | 4 | `Keep it on screen, then run the job again` (attention) |
| A Google page didn't respond | 5 | `Google may have changed its pages: see the run log` (attention) |
| Config problem, count mismatch, any other error | 2 / 1 | the first 200 characters of the error, plus the album link if the album was made (attention) |

The "already there" case also covers a re-run after a successful run, which is safe. A failed run can be run again from the Jobs tab once the cause is fixed; nothing is duplicated. Every run, including the locked-desktop one, appends its line to the run log (`failure` names the kind when it failed).

**Notifier setup.** Success goes to the notifier's `log` chat, failures to `attention`. Put `NOTIFY_PYTHON` and `NOTIFY_SCRIPT` in the repo-root `.env` (see `.env.sample`; optional `NOTIFY_CHAT` sends everything to one chat). Without them the message is only logged and the job still exits with the right code.

**Register the job** in app-launcher's Jobs tab (**Add job**), or add this entry to the machine-local `config/jobs.json` (the live file is gitignored, so nothing here registers it). Replace `<repo>` with this checkout's absolute path:

```json
{
  "id": "photos-weekly-album",
  "name": "Weekly Photos album",
  "script_path": "<repo>\\google\\run_photos_weekly_album.bat",
  "args": "",
  "schedule": { "type": "weekly", "day": "SAT", "at": "09:00" },
  "cooldown_seconds": 3600,
  "max_runtime_seconds": 1800
}
```

Leave `alert_on_failure` off: the job already sends its own, more specific message for each failure. Don't set `session_less` or `visible`: the job needs the interactive desktop, and its Chrome window is its own. Before the first Saturday: sign the profile in once (Setup, step 2), fill `photos_weekly_album.json`, set the `NOTIFY_*` lines, and try the **Dry-run check** in the job's menu and a manual `--dry-run`.

## Selectors (English UI)

| What | Selector |
|---|---|
| Day select-all | `[role=checkbox][aria-label^="Select all photos from "]`, state in `aria-checked` |
| Item tile | `a[href*="/photo/"]`, `aria-label` like `Photo - Landscape - Oct 3, 2026, 8:23:55 AM` |
| Item checkbox | the `[role=checkbox]` under the tile with the same `aria-label` |
| Selection count | text `N selected` |
| Scroll container | the `c-wiz` with the largest `scrollHeight` (the window doesn't scroll) |
| Album list | `/albums`, card link `a[href*="/album/"]`, or `a[href*="/share/"]` once the album is shared by link, first line the title, then `N items` |
| Album share | button `Share` → dialog heading `Invite to album` → button `Create link` (no link yet) or `Copy link` (link exists); `Create link` → dialog heading `Create link to share` → button `Create link` → the link in an input next to button `Copy` |
| Gmail compose | button `Compose` → dialog with combobox `To recipients`, `input[name=subjectbox]`, textbox `Message Body`, buttons `Save & close`, `Send ...`, `Discard draft ...` |
| Gmail search rows | `#search/<query>`; rows `[role=main] tr.zA`, subject `span.bog`; empty state `No messages matched your search` |

Day-header labels are relative ("Today", "Wednesday", "Sat, Oct 3"), so each header is dated by the first tile after it, never by its own text.

## Failure modes

| Symptom | Cause and handling |
|---|---|
| `not signed in` (exit 3) | The profile's Google session is missing or expired. Run `--login` again. |
| `window is hidden or minimised` (exit 4) | Chrome doesn't render the grid in a minimised window. Covered windows are fine because native occlusion detection is turned off for this profile. Keep the window on screen and re-run. |
| `Chrome profile ... still in use` | Another Chrome window has the dedicated profile open. Close it. The script waits about 15 minutes before giving up and never kills it. |
| `Locator.click: Timeout ... element is not visible` | A day's select-all checkbox has no size until the pointer is over its header row. The script hovers the row first when the checkbox isn't shown. If this comes back, Google changed what reveals the checkbox. Nothing has been created at that point. |
| `checkbox did not change after two clicks` | A click was dropped (seen once after a fresh navigation in testing) or the page changed. Re-run. Nothing has been created at that point. |
| `saw N tiles in the week but the day headers selected M` | The sweep didn't render every tile, so a screenshot might stay in. Nothing is created. Re-run with the window in front. |
| `selection counter reads ...` | An untick or a day click didn't land. Nothing is created. Re-run. |
| `share a timestamp label` | Two items from the same second with the same type and orientation can't be told apart, so the run aborts. Create that week's album by hand. |
| `screenshots not in the timeline` (info) | An item in the collection doesn't appear in the timeline (for example it's archived), so it was never selected. |
| `album created but it lists N items` (exit 1) | The album card's count differs from the selection after three reads. Check the album by hand. |
| `the album is now shared by link, but the script could not read it` / `the album has a share link, but ...` | The link exists but wasn't found in any of the places checked; the warning just before it lists the sources tried and each dialog's buttons. Re-run: the dialog then offers **Copy link** and the link is read again. Or copy it from the album's Share dialog. |
| `the album's Share dialog gave no link to create or copy` | The dialog didn't offer **Create link** or **Copy link**, or the **Create link to share** confirmation never appeared. Photos changed the dialog: compare the buttons in the log with the selector table above. |
| `compose shows N recipients, not the M configured; draft discarded` | The typed addresses didn't all turn into chips, or autocomplete changed one. Nothing is saved. Check `email.to` and re-run. |
| `N drafts carry this week's subject` | More than one draft has the week's subject. Delete the extras by hand. The run never sends when that happens. |
| `the draft's body doesn't carry this run's share link` / `recipients differ` | The week's draft was edited, or comes from somewhere else. Nothing is sent. Fix or delete the draft and re-run. |
| `not signed in to Gmail` (exit 3) | The profile's Google session doesn't reach Gmail. Run `--login` again. |
| Late items | Anything taken late on Friday or synced late from another device after the run is missed. Run on Saturday morning. |
| Google changes the UI | The selectors above stop matching, so the scans find nothing or the clicks time out. Update the selectors in the module. |

The UI must be in English: the selectors and the date labels are matched on English text.
