# Smart Life / Tuya Light Controller

Control Smart Life / Tuya lights and plugs from the command line. Uses **Tuya Cloud first** (when configured) for fast, reliable control; falls back to **local network** (1s timeout) if cloud is not set up or the device is only on LAN.

## Setup

### 1. Create a Tuya IoT Cloud Project (one-time)

You need this to get device keys and (optionally) to use cloud control. Create a project to obtain **Access ID** and **Access Secret** and to link your Smart Life app.

1. Go to [iot.tuya.com](https://iot.tuya.com) and create an account.
2. Create a **Cloud Project** (select your data center, e.g. *Central Europe*).
3. Under **Devices > Link Tuya App Account**, link your **Smart Life** app account by scanning the QR code from the app (Me > top-right scan icon).
4. Note your **Access ID / Client ID** and **Access Secret / Client Secret** from the project overview page.

### 2. Discover Devices & Get Local Keys

```bash
python -m tinytuya wizard
```

The wizard will ask for your **Access ID** and **Access Secret**, then pull every device linked to your Smart Life account and write a `devices.json` file.

Copy (or move) the generated `devices.json` into the `smart_life/` folder:

```bash
cp devices.json smart_life/devices.json
```

### 3. (Optional) Refresh IPs and status

`devices.json` uses the same format as tinytuya’s snapshot (top-level `timestamp` and `devices` array). To refresh IPs and live status from the network, run:

```bash
python smart_life/light_control.py update
```

This runs `tinytuya snapshot` (feeding it a temporary list of devices so it works with our snapshot-format file) and overwrites `devices.json` with the result. You can also run `python -m tinytuya scan` to discover devices; add new ones to `devices.json` (id, key, name, ip, ver) as needed.

## Usage

```bash
# Turn a light ON
python smart_life/light_control.py on --light "luz despacho"

# Turn a light OFF
python smart_life/light_control.py off --light "luz despacho"

# Check current status
python smart_life/light_control.py status --light despacho

# List all configured devices
python smart_life/light_control.py list

# Scan network for Tuya devices
python smart_life/light_control.py scan

# Refresh devices.json from snapshot (update IPs and status; same format as snapshot.json)
python smart_life/light_control.py update

# Debug mode (verbose output)
python smart_life/light_control.py status --light despacho --debug
```

The `--light` parameter accepts a **device name** (partial, case-insensitive) or **full device ID**. Examples: `--light despacho`, `--light "luz despacho"`, `--light bfc158aece14a52035diwf`.

## devices.json Format

The file uses **snapshot format**: `{"timestamp": ..., "devices": [...]}`. Each device needs at least `id`, `key`, `name`, `ip`, and `ver` (or `version`):

```json
{
  "timestamp": 1234567890.0,
  "devices": [
    {
      "id": "DEVICE_ID_FROM_TUYA",
      "key": "LOCAL_KEY_FROM_WIZARD",
      "name": "luz despacho",
      "ip": "192.168.1.100",
      "ver": "3.3"
    }
  ]
}
```

Run `python smart_life/light_control.py update` to refresh IPs and fill in extra fields (e.g. `dps`, `origin`) from the network.

## Cloud (recommended)

With `cloud.json` configured, the script uses the **Tuya Cloud API first** — typically a few seconds per command. If cloud is not configured or the API call fails, it tries **local** control with a 1s timeout so it doesn’t hang.

1. Create `smart_life/cloud.json` (or copy `tinytuya.json` from the wizard into `smart_life/` as `cloud.json`) with your Tuya IoT credentials:

```json
{
  "apiRegion": "eu",
  "apiKey": "YOUR_ACCESS_ID",
  "apiSecret": "YOUR_ACCESS_SECRET"
}
```

2. Use **Access ID** and **Access Secret** from your [Tuya IoT project](https://iot.tuya.com) (project > Overview). Set `apiRegion` to your data center: `eu`, `us`, `cn`, `in`, etc.

3. Commands then run via cloud when the device is linked to your project. No extra message — it just works. Without `cloud.json`, only local control is tried (1s timeout).

`cloud.json` is in `.gitignore` so secrets are not committed.

## Troubleshooting

| Problem | Fix |
|---|---|
| `Cannot reach device` | Run `python smart_life/light_control.py update` to refresh IPs. If the device works in the Smart Life app but not locally, add `smart_life/cloud.json` (see **Cloud** above) for cloud control. |
| `devices.json not found` | Run `python -m tinytuya wizard` and copy the output file into `smart_life/`. Convert to snapshot format (see below) or run `update` after. |
| `Snapshot failed` / `TypeError` | Ensure `devices.json` is snapshot format `{"timestamp", "devices": [...]}`; the script writes a temp list for tinytuya automatically. |
| Wrong device toggled | Use `list` to check names/IDs, then use a more specific `--light` query or the device ID. |
| Protocol errors | Try changing `ver` in devices.json to `"3.1"`, `"3.4"`, or `"3.5"`. |
