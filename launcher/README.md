# 🚀 Remote Project Launcher

A tiny Flask app that lets you tap a button on your phone (over Tailscale) and have a project's launch bat fire up in a fresh CMD window on your PC. Intended use case: launch Claude Code in `--remote` mode on your home machine while you're out.

## How it works

1. The launcher scans the **parent directory of this repo** for `*.bat` files whose names contain the word `remote` (case-insensitive). Example layout:

    ```
    E:\automation\
    ├── automation\          (this repo)
    │   └── launcher\        (this app)
    ├── automation_remote.bat
    ├── notion_remote.bat
    └── client_x_remote.bat
    ```

   With that layout the UI shows three buttons: "Automation", "Notion", "Client X".

2. Tapping a button on the web UI hits `POST /launch`. The server runs:

    ```
    cmd /c start "" cmd /k <bat_path>
    ```

   which opens a brand new visible CMD window on the host, runs the bat, and leaves the window open so you can interact with it locally when you get back to the PC.

3. Optional password gate (set `LAUNCHER_PASSWORD` in `.env` to enable; leave it unset for open access on a trusted tailnet). Flask session cookie, per-session CSRF token. POST-only for everything that mutates state.

## Install

The launcher uses the repo's existing `.venv`. From the repo root:

```powershell
& .\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

`flask` is already in the root `requirements.txt`; `python-dotenv` was there before.

## Configure

Add to the root `.env`:

```env
# Optional — omit to disable the login screen entirely (fine on a private tailnet)
LAUNCHER_PASSWORD=pick-something-strong
LAUNCHER_SECRET_KEY=run-python-secrets-token_urlsafe-32
# Optional overrides:
# LAUNCHER_PROJECTS_DIR=E:\some\other\folder
# LAUNCHER_PORT=5050
# LAUNCHER_HOST=0.0.0.0
```

To generate a secret key:

```powershell
& .\.venv\Scripts\python.exe -c "import secrets; print(secrets.token_urlsafe(32))"
```

## Run

```powershell
.\launcher\launcher.bat
```

The server binds to `0.0.0.0:5050`. From your phone (joined to your tailnet), browse to:

```
http://<pc-tailscale-name>:5050
```

…or use the Tailscale IP (`100.x.x.x`).

If the bat's hardcoded paths don't match your install, edit `VENV_DIR` and `SCRIPT_DIR` at the top of `launcher.bat`.

## Generate BAT files

The **Generate BAT files** button at the bottom of the launcher index page opens `/generate`, which keeps your `*-remote.bat` files in sync with your `.code-workspace` files.

**Conventions:**
- Every `.code-workspace` file in `PROJECTS_DIR` maps to a `{name}-remote.bat` in the same directory.
- The bat opens a CMD window and runs `claude <flags>` in the workspace's project folder.
- Flags (e.g. `--remote-control --dangerously-skip-permissions --verbose --effort high`) are editable in the UI and persisted to `launcher/config.json`.

**Workflow for a new project:**
1. Create `E:\automation\my-project\` and a `my-project.code-workspace` file in `E:\automation\`.
2. Open the launcher → **Generate BAT files**.
3. `my-project-remote.bat` appears in the *New bat files* section — hit **Generate**.
4. The new bat shows up on the launcher home screen immediately.

**Sections on the generate page:**

| Section | Default | Action |
|---|---|---|
| New bat files | — | Always created |
| Existing bat files | unchecked | Tick to overwrite with current flags |
| Missing workspaces | checked | Creates a minimal `.code-workspace` for orphan bats |

## Apps tab (Streamlit launchers)

The home page has two tabs:

- **Cloud Code** — the original `*remote*.bat` list described above.
- **Apps** — Streamlit app launchers discovered by recursively scanning a configured root.

**How it works:**

1. Tap **Scan projects** on the Apps tab. The server walks `LAUNCHER_APPS_SCAN_ROOT` (default: the parent of this repo, same as `LAUNCHER_PROJECTS_DIR`), looking at every `*.bat` whose contents include `streamlit run`. Skipped: `.venv/`, `venv/`, `__pycache__/`, `node_modules/`, `certificates/`, `.git/`, `old/`.
2. New finds appear with checkboxes — submit to add them to `launcher/apps_config.json`. Each saved entry has a stable `id`, a display `name` (auto-derived from the parent folder, e.g. `system\grocery\launcher.bat` → "Grocery"), and the absolute `bat_path`.
3. The list is **static** after that — no rescan on every page load. Hit **Scan projects** again only when you add new Streamlit apps to the repo.
4. Each row has an **Edit** disclosure for **Rename** (override the auto display name) and **Remove**.
5. Tapping a saved app launches it in a new visible CMD window via the same `cmd /c start "" cmd /k <bat>` plumbing as Cloud Code.
6. The red **⛔ Kill :8501** button at the top of the Apps tab kills whatever process is currently listening on Streamlit's default port. Streamlit always picks 8501 when free, so the typical phone workflow is: **Kill :8501 → tap the next app**. Override the port via `LAUNCHER_STREAMLIT_PORT` in `.env` if your apps run elsewhere.

**Config files:**

- `launcher/apps_config.json` — your live list. **Gitignored.**
- `launcher/apps_config.sample.json` — example schema, committed.

**Optional override:**

```env
LAUNCHER_APPS_SCAN_ROOT=E:\automation
```

Defaults to the parent of this repo (matches `LAUNCHER_PROJECTS_DIR`). Narrow it (e.g. `E:\automation\automation`) if you only want this repo scanned, or point it elsewhere entirely.

## Auto-start at log on with Task Scheduler

1. Open **Task Scheduler** → **Create Task…** (not "Create Basic Task" — you need the advanced options).
2. **General** tab:
    - Name: `Remote Project Launcher`
    - **Run only when user is logged on** ✅ (required so the launcher can spawn visible CMD windows on your desktop)
    - **Run with highest privileges** is *not* needed
    - Configure for: Windows 10/11
3. **Triggers** tab → **New…**
    - Begin the task: **At log on**
    - Specific user: your Windows account
    - Delay task for: 30 seconds (gives the network and Tailscale time to come up)
4. **Actions** tab → **New…**
    - Action: **Start a program**
    - Program/script: `E:\automation\automation\launcher\launch_launcher.bat`
    - Start in (optional): `E:\automation\automation\launcher`
5. **Conditions** tab:
    - Uncheck **Start the task only if the computer is on AC power**
6. **Settings** tab:
    - **Allow task to be run on demand** ✅
    - **If the task fails, restart every:** 1 minute, attempt up to 3 times
    - **If the task is already running:** *Do not start a new instance*

To test without rebooting: select the task → **Run** in the right-hand pane.

To make the CMD window invisible at startup, change the action to launch `pythonw.exe` directly with `launcher.py` instead of the bat — but the bat is more debuggable when something goes wrong.

## Security notes

- Tailscale already gates network access; the password is a second factor in case a tailnet device is compromised.
- The launcher only ever runs bats from the configured projects directory — the `name` field from the form is validated against the discovered set, so it can't be coerced into running an arbitrary path.
- Session cookies are HTTP-only and `SameSite=Lax`. CSRF tokens are checked on every POST.
- Tailscale traffic is encrypted between nodes, so plaintext HTTP is fine here. **Don't expose port 5050 to the public internet.**

## Files

- `launcher.py` — Flask app
- `tray.py` — system-tray wrapper; starts the Flask server in a background thread and shows a green dot icon in the notification area
- `templates/login.html` — login form (shown only when `LAUNCHER_PASSWORD` is set)
- `templates/index.html` — Cloud Code tab (project list)
- `templates/apps.html` — Apps tab (Streamlit launcher list)
- `templates/generate.html` — Generate BAT files page
- `config.json` — persisted Claude Code launch flags (created on first generate run)
- `apps_config.json` — saved Streamlit apps (gitignored, created on first scan)
- `apps_config.sample.json` — committed schema example
- `launcher.bat` — start launcher + tray (use this normally)
- `tray.bat` — start tray only
- `README.md` — this file

## Duplicate-launch guard

`tray.py` checks whether the server port is already in use before starting. If a launcher is already running, a Windows balloon notification appears ("Already Running — Launcher server is already running.") and the second instance exits immediately. No duplicate servers, no duplicate tray icons.
