# 🚀 Remote Project Launcher

A tiny Flask app that lets you tap a button on your phone (joined to your tailnet) and have a project's launch bat fire up in a fresh CMD window on your PC. The original use case was launching Claude Code in `--remote` mode on the home machine while away from the desk; the **Apps** tab grew out of needing the same one-tap flow for Streamlit and FastAPI tools.

The launcher itself is a small Flask server. The tray wrapper (`tray.py`) supervises it so it survives reboots without a console window cluttering the desktop.

---

## What it does, in one screen

The web UI has two tabs:

- **Cloud Code** — every `*remote*.bat` in the projects directory becomes a button that opens Claude Code in remote mode for that project.
- **Apps** — Streamlit *and* FastAPI launchers discovered by scanning a configured root, persisted to `apps_config.json`, surfaced as one-tap buttons grouped by kind.

Tapping a button POSTs to the server, which spawns a brand-new visible CMD window via `cmd /c start "" cmd /k <bat_path>`. The window stays open so you can interact with it locally when you get back to the PC.

Optional password gate (`LAUNCHER_PASSWORD` in `.env`) plus per-session CSRF on every POST. Designed to run behind Tailscale, not on the open internet.

---

## Install

The launcher uses the repo's existing `.venv`. From the repo root:

```powershell
& .\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

`flask`, `pystray`, `pillow`, `psutil`, and `python-dotenv` are already in the root `requirements.txt`.

## Configure

Add to the root `.env`:

```env
# Optional — omit to disable the login screen entirely (fine on a private tailnet)
LAUNCHER_PASSWORD=pick-something-strong
LAUNCHER_SECRET_KEY=run-python-secrets-token_urlsafe-32

# Optional overrides
# LAUNCHER_PROJECTS_DIR=E:\some\other\folder
# LAUNCHER_APPS_SCAN_ROOT=E:\automation
# LAUNCHER_PORT=5050
# LAUNCHER_HOST=0.0.0.0
# LAUNCHER_STREAMLIT_PORT=8501
# LAUNCHER_WEBAPP_PORT=8443
# LAUNCHER_SSL_CERT=...
# LAUNCHER_SSL_KEY=...
```

To generate a secret key:

```powershell
& .\.venv\Scripts\python.exe -c "import secrets; print(secrets.token_urlsafe(32))"
```

## Run

```powershell
.\launcher\launcher.bat   # tray + server (normal use)
.\launcher\tray.bat       # same thing, kept for muscle memory
```

The server binds to `0.0.0.0:5050`. From your phone (joined to your tailnet):

```
http://<pc-tailscale-name>:5050
```

…or use the Tailscale IP (`100.x.x.x`).

If `tray.py` finds the port already busy it shows a Windows balloon ("Already Running") and exits — no duplicate servers, no duplicate icons.

---

## Cloud Code tab

Scans the **parent of this repo** (`LAUNCHER_PROJECTS_DIR`) for `*.bat` files whose names contain `remote` (case-insensitive). Example layout:

```
E:\automation\
├── automation\          (this repo)
│   └── launcher\        (this app)
├── automation_remote.bat
├── notion_remote.bat
└── client_x_remote.bat
```

Three buttons appear: "Automation", "Notion", "Client X". Tap → fresh CMD window runs the bat.

### Generate BAT files

The **Generate BAT files** button at the bottom of the Cloud Code tab opens `/generate`, which keeps `*-remote.bat` files in sync with `.code-workspace` files in the same directory.

- Every `.code-workspace` maps to a `{name}-remote.bat`.
- The bat opens a CMD window and runs `claude <flags>` in the workspace's project folder.
- Flags (e.g. `--remote-control --dangerously-skip-permissions --verbose --effort high`) are editable in the UI and persisted to `launcher/config.json`.

| Section | Default | Action |
|---|---|---|
| New bat files | — | Always created |
| Existing bat files | unchecked | Tick to overwrite with current flags |
| Missing workspaces | checked | Creates a minimal `.code-workspace` for orphan bats |

---

## Apps tab

The Apps tab lists every web app launcher you've saved. Tap → it spawns. The list is grouped by kind, and the tools we surface follow from what each kind actually does on disk.

### What gets detected

Hit **Scan projects**. The server walks `LAUNCHER_APPS_SCAN_ROOT` (default: parent of this repo, same as `LAUNCHER_PROJECTS_DIR`) and inspects every `*.bat`. The classifier is mutually exclusive — first match wins:

| Kind | Match rule | Example |
|---|---|---|
| **Streamlit** | body contains `streamlit run` | `accounting-quarterly\launch_app.bat` |
| **Tunnel** | filename contains `tunnel` AND body references `uvicorn` / `run_tunnel` / `cloudflared` | `voice-transcriber\webapp_tunnel.bat` |
| **Webapp** | body runs `uvicorn` (or imports `app.webapp.server`) | `voice-transcriber\webapp.bat` |
| *(skipped)* | none of the above | tray scripts, generic helper bats |

Skip directories during the walk: `.venv/`, `venv/`, `__pycache__/`, `node_modules/`, `certificates/`, `.git/`, `old/`.

The display name is auto-derived from the parent folder (`system\grocery\launcher.bat` → "Grocery"). A stable `id` is generated from the relative path so renames don't collide.

### Why the kind matters

The kind isn't decoration — it changes which tools the row gets:

- **Streamlit and Webapp** kinds are plain "tap to spawn" entries. The only difference is bookkeeping: they render in separate sections so you can scan the list at a glance.
- **Tunnel** kind adds an inline `📡 <URL>` row directly under the launch button. The URL is read from `<bat_dir>/webapp/last_tunnel_url.txt` at page render time — that file is the documented hook the tunnel script writes to. If the file is missing or empty the row shows "Tunnel not running"; refresh the page after the tunnel has had time to come up to reveal the link.
- The two **⛔ Kill :PORT** buttons at the top sit on the same endpoint with different port hidden fields. Streamlit always prefers `:8501`; FastAPI webapps in this constellation default to `:8443`. Override with `LAUNCHER_STREAMLIT_PORT` / `LAUNCHER_WEBAPP_PORT` if your stack picks something else.

### Edge cases worth knowing

- **Streamlit + cloudflared hybrids.** Some bats embed both `streamlit run` and `cloudflared tunnel` inline (e.g. `launch_server.bat` patterns in this repo). They classify as **Streamlit** — the cloudflared call is fire-and-forget and doesn't write a URL file we can surface, so the tunnel row would only confuse. If a hybrid bat ever does start writing `webapp/last_tunnel_url.txt`, we can revisit; until then, plain Streamlit is the honest classification.
- **Tray scripts are not picked up.** A bat that just spawns a tray icon (e.g. `voice-transcriber\tray.bat` runs that project's own `launcher.py tray`) has no uvicorn/streamlit/cloudflared fingerprint and no `tunnel` in its name. It correctly stays out of the list. The Apps tab is for HTTP launchers, not arbitrary tray spawners.
- **Existing entries from before the kind field.** `apps_config.json` entries that predate this feature have no `kind` — they default to `"streamlit"` at load time. No migration needed; rescan only when you want to pick up new finds.

### The flow

1. Tap **Scan projects** on the Apps tab.
2. New finds appear under **Streamlit** and **Web** sub-headers, all checked by default. Submit to add them to `launcher/apps_config.json`.
3. The list is **static** after that — no rescan on every page load. Re-scan only when a project lands.
4. Each card has an **Edit** disclosure (Rename / Remove). Rename overrides only the display name; the underlying bat path is the source of truth.
5. The Cloudflare-tunnel flow when remote: tap the tunnel app's launch button, wait ~5 s for cloudflared to print its URL, refresh the page — the `📡` row reveals the link, tap to open in mobile Safari.

### Config files

- `launcher/apps_config.json` — your live list. **Gitignored.** Each entry: `id`, `name`, `bat_path`, `kind`, `added_at`.
- `launcher/apps_config.sample.json` — committed schema example.

### Override the scan root

```env
LAUNCHER_APPS_SCAN_ROOT=E:\automation
```

Defaults to the parent of this repo. Narrow it (e.g. `E:\automation\automation`) if you only want this repo scanned, or point it elsewhere entirely.

---

## Auto-start at log on with Task Scheduler

1. Open **Task Scheduler** → **Create Task…** (not "Create Basic Task" — you need the advanced options).
2. **General** tab:
    - Name: `Remote Project Launcher`
    - **Run only when user is logged on** ✅ (required so the launcher can spawn visible CMD windows on your desktop)
    - Configure for: Windows 10/11
3. **Triggers** tab → **New…**
    - Begin the task: **At log on**
    - Specific user: your Windows account
    - Delay task for: 30 seconds (gives the network and Tailscale time to come up)
4. **Actions** tab → **New…**
    - Action: **Start a program**
    - Program/script: `E:\automation\automation\launcher\launcher.bat`
    - Start in (optional): `E:\automation\automation\launcher`
5. **Conditions** tab:
    - Uncheck **Start the task only if the computer is on AC power**
6. **Settings** tab:
    - **Allow task to be run on demand** ✅
    - **If the task fails, restart every:** 1 minute, attempt up to 3 times
    - **If the task is already running:** *Do not start a new instance*

To test without rebooting: select the task → **Run** in the right-hand pane.

---

## Security notes

- Tailscale already gates network access; the password is a second factor in case a tailnet device is compromised.
- The launcher only ever runs bats from the configured projects directory or saved Apps entries — the form `name`/`id` is validated against the discovered set, so it can't be coerced into running an arbitrary path.
- Session cookies are HTTP-only and `SameSite=Lax`. CSRF tokens are checked on every POST.
- The kill-port endpoint whitelists requests to the configured Streamlit and Webapp ports — arbitrary ports are rejected.
- Tailscale traffic is encrypted between nodes, so plaintext HTTP is fine here. **Don't expose port 5050 to the public internet.**

---

## Files

- `launcher.py` — Flask app
- `tray.py` — system-tray wrapper; starts the Flask server in a background thread, gray icon → green when listening
- `templates/login.html` — login form (shown only when `LAUNCHER_PASSWORD` is set)
- `templates/index.html` — Cloud Code tab (project list)
- `templates/apps.html` — Apps tab (Streamlit + Web app lists, kill-port row, scan/save flow)
- `templates/generate.html` — Generate BAT files page
- `config.json` — persisted Claude Code launch flags (created on first generate run)
- `apps_config.json` — saved Apps entries (gitignored, created on first scan)
- `apps_config.sample.json` — committed schema example
- `launcher.bat` / `tray.bat` — start launcher + tray
- `README.md` — this file
