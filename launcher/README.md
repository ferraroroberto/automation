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

3. Single password gate, Flask session cookie, per-session CSRF token. POST-only for everything that mutates state.

## Install

The launcher uses the repo's existing `.venv`. From the repo root:

```powershell
& .\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

`flask` is already in the root `requirements.txt`; `python-dotenv` was there before.

## Configure

Add to the root `.env`:

```env
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
.\launcher\launch_launcher.bat
```

The server binds to `0.0.0.0:5050`. From your phone (joined to your tailnet), browse to:

```
http://<pc-tailscale-name>:5050
```

…or use the Tailscale IP (`100.x.x.x`).

If the bat's hardcoded paths don't match your install, edit `VENV_DIR` and `SCRIPT_DIR` at the top of `launch_launcher.bat`.

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
- `templates/login.html`, `templates/index.html` — UI
- `launch_launcher.bat` — startup script for Task Scheduler
- `README.md` — this file
