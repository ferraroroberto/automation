# `system/wifi/` — Wi-Fi & local-network tools

Four independent Windows tools for working with Wi-Fi profiles and the
local network. All are invoked via the repo's `.venv`. None require
configuration beyond what's noted below.

## Tools

| Tool | Files | Purpose |
|------|-------|---------|
| Quick connect (single SSID) | `wifi_connect.py` `wifi_connect.bat` `wifi_connect.ps1` `wifi_connect.xml` | Connect to the network whose credentials are baked into `wifi_connect.xml`. Idempotent: skips reconnect if already on that SSID. |
| Saved-password export | `wifi_passwords.py` `wifi_passwords.json` | Dump every saved Wi-Fi profile + cleartext password to JSON. |
| BAT generator (GUI) | `wifi_bat_generator.py` `wifi_bat_generator.md` | Pick a saved network from a Tkinter list, save a portable self-contained `.bat` anywhere. Generated BATs work on any Windows PC. See [`wifi_bat_generator.md`](wifi_bat_generator.md). |
| Network device scanner | `network_scanner.py` `network_scanner.json` `network_scanner.md` | ARP-scan the local subnet to discover devices. Not strictly Wi-Fi-related — useful for finding routers/APs that don't reply to ping. See [`network_scanner.md`](network_scanner.md). |

## Quick start

From the repo root, in PowerShell:

```powershell
# Connect to the SSID configured in wifi_connect.xml (or just double-click wifi_connect.bat).
& .\.venv\Scripts\python.exe system\wifi\wifi_connect.py

# Export every saved Wi-Fi profile + password to wifi_passwords.json.
& .\.venv\Scripts\python.exe system\wifi\wifi_passwords.py

# Open the BAT generator GUI (recommended day-to-day tool).
& .\.venv\Scripts\python.exe system\wifi\wifi_bat_generator.py

# ARP-scan the LAN for active devices (requires Admin + Npcap).
& .\.venv\Scripts\python.exe system\wifi\network_scanner.py
```

## Typical workflow: "make me a BAT for this network"

1. Connect to the Wi-Fi network normally (Windows saves the profile).
2. Run `wifi_bat_generator.py`.
3. Pick the network from the list → **Save BAT…** → choose any location.
4. Double-click the generated `.bat` whenever you want to reconnect — on
   the same PC, on another PC, or after a Windows reinstall. The BAT
   carries the full credentials inside it.

## Setting up the local `wifi_connect` profile

`wifi_connect.xml` holds the SSID and passphrase used by `wifi_connect.py`. It
is a machine-local file and is **gitignored** — only `wifi_connect.xml.sample`
is tracked. Generate or update it by passing the values to `wifi_connect.ps1`
at call time:

```powershell
.\system\wifi\wifi_connect.ps1 -Ssid "MyNetwork" -Password "my-passphrase"
```

The script also reads `WIFI_SSID` / `WIFI_PASSWORD` from the environment when
the parameters are omitted. Alternatively, copy `wifi_connect.xml.sample` to
`wifi_connect.xml` and fill in the placeholders by hand. Never edit values into
`wifi_connect.ps1` itself — that file is tracked in git.

## Notes

- `wifi_connect.*`, `wifi_passwords.py`, and `wifi_bat_generator.py` use
  only `netsh` (built into Windows). No admin rights needed.
- All three reach `netsh` through `_netsh.py`, which captures raw bytes and
  decodes them explicitly (UTF-8 → OEM console page → ANSI → lossy UTF-8)
  instead of letting `subprocess`'s `text=True` decode with the *parent's*
  locale. Without that, running under a parent in Python's UTF-8 mode
  (`PYTHONUTF8=1` — a tray app, a wrapper `.bat`, a scheduled job) silently
  mangles or empties the output, and the tools then report "not connected" /
  "no profiles" / "export failed" for reasons unrelated to the network.
  `_netsh.py` also keeps "the query failed" distinct from "the answer is
  nothing", so an unreadable state is never shown as a confirmed negative.
  Guarded by `test_netsh_decoding.py`:

  ```powershell
  & .\.venv\Scripts\python.exe -m unittest discover -s system/wifi -p "test_*.py"
  ```
- `network_scanner.py` uses raw sockets via scapy. **Run as
  Administrator** and install Npcap first
  (<https://npcap.com/#download>). See `network_scanner.md` for details.
- Generated BATs from `wifi_bat_generator.py` contain the cleartext
  Wi-Fi password (base64-encoded inside the embedded XML profile).
  Treat them like credentials.
