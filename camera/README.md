# YI Camera Monitor

Custom desktop application for monitoring YI home cameras via RTSP streams, bypassing the official YI Home app.

## Overview

A PyQt6-based multi-camera viewer that connects directly to your YI cameras over your local network using RTSP. Features include:

- Multi-camera grid view with adjustable columns
- Per-camera snapshot and recording
- Fullscreen mode (double-click or button)
- Auto-reconnect with exponential backoff
- Dark theme UI
- Add/manage cameras at runtime
- Remote access via Tailscale VPN

## Tested Camera Models

| Camera | Hack Project | Status |
|---|---|---|
| YI Dome Guard 2K (3MP) 360 | [yi-hack-Allwinner-v2](https://github.com/roleoroleo/yi-hack-Allwinner-v2) | Supported (Allwinner chipset) |
| YI Outdoor 1080P Bullet | [yi-hack-v4](https://github.com/TheCrypt0/yi-hack-v4) | Fully supported (Hi3518e chipset) |
| YI 1080p Dome | [yi-hack-v4](https://github.com/TheCrypt0/yi-hack-v4) | Fully supported (Hi3518e chipset) |

## Setup

### 1. Install Dependencies

```bash
cd camera
pip install -r requirements.txt
```

### 2. Enable RTSP on Your YI Cameras (SD Card Method — No Firmware Flash)

YI cameras don't expose RTSP by default. The **yi-hack** projects add RTSP support by loading from an SD card. This is **fully reversible** — remove the SD card and the camera goes back to stock. No permanent firmware modification is needed.

#### Step 1 — Identify your hack variant

| Your Camera | Hack to Use | Download |
|---|---|---|
| YI Dome Guard 2K | yi-hack-Allwinner-v2 | [Releases](https://github.com/roleoroleo/yi-hack-Allwinner-v2/releases) |
| YI Outdoor 1080P | yi-hack-v4 | [Releases](https://github.com/TheCrypt0/yi-hack-v4/releases) |
| YI 1080p Dome | yi-hack-v4 | [Releases](https://github.com/TheCrypt0/yi-hack-v4/releases) |

#### Step 2 — Prepare the SD card

1. Use a micro SD card (8GB+ recommended, FAT32 formatted)
2. Download the correct release archive for your camera model
3. Extract **all files** to the root of the SD card

For **yi-hack-v4** (YI Outdoor + YI Dome 1080p):
```
SD Card Root/
├── home_<model>       (e.g. home_h20 for Dome, home_h30 for Outdoor)
├── rootfs_<model>
└── yi-hack-v4/
    └── ...
```

For **yi-hack-Allwinner-v2** (YI Dome Guard 2K):
```
SD Card Root/
├── home_<model>
├── rootfs_<model>
└── yi-hack/
    └── ...
```

#### Step 3 — Boot the camera with the SD card

1. Power off the camera
2. Insert the prepared SD card
3. Power on the camera
4. Wait for the LED to turn solid blue (1-2 minutes)
5. The hack loads automatically from the SD card on every boot

#### Step 4 — Find the camera IP and enable RTSP

1. Open your router's admin page and find the camera's IP address (or check the YI Home app → Camera Settings → Network Info)
2. Open a browser and go to `http://<CAMERA_IP>` — the yi-hack web interface will load
3. Go to the RTSP settings and **enable the RTSP server**
4. Note the RTSP URL shown (typically `rtsp://<CAMERA_IP>/ch0_0.h264`)

#### Step 5 — Verify the stream

Open **VLC** on your PC: Media → Open Network Stream → paste the RTSP URL. If you see the feed, you're ready.

#### Common RTSP URLs for YI cameras (after yi-hack)

```
rtsp://<IP>/ch0_0.h264         # High quality stream
rtsp://<IP>/ch0_1.h264         # Low quality stream (less bandwidth)
rtsp://<IP>:554/ch0_0.h264     # Explicit port
```

#### Troubleshooting

- **Dome Guard 2K RTSP drops**: Edit `/sd/yi-hack/model_suffix` on the SD card to contain `r35gb`, then reboot
- **Camera restarts on RTSP connect**: Try an older hack version (v0.2.6 is more stable for some Dome Guard models)
- **Stream works in VLC but not in monitor**: Try adding `?tcp` to the URL, or check the `--debug` output for codec errors
- **To revert**: Simply remove the SD card and reboot — the camera returns to stock

### 3. Configure Cameras

Edit `config.json` with your actual camera IPs:

```json
{
    "cameras": [
        {
            "name": "Luca",
            "rtsp_url": "rtsp://192.168.1.100/ch0_0.h264",
            "enabled": true
        },
        {
            "name": "Outdoor",
            "rtsp_url": "rtsp://192.168.1.101/ch0_0.h264",
            "enabled": true
        },
        {
            "name": "Living Room",
            "rtsp_url": "rtsp://192.168.1.102/ch0_0.h264",
            "enabled": true
        }
    ],
    "grid_columns": 2,
    "snapshot_dir": "snapshots",
    "recording_dir": "recordings",
    "reconnect_interval_sec": 5,
    "max_reconnect_attempts": 0,
    "stream_timeout_sec": 10,
    "window_width": 1280,
    "window_height": 720
}
```

**Config fields:**

| Field | Description |
|---|---|
| `cameras` | List of camera objects (`name`, `rtsp_url`, `enabled`) |
| `grid_columns` | Number of columns in the camera grid (1-6) |
| `snapshot_dir` | Directory for saved snapshots |
| `recording_dir` | Directory for saved recordings |
| `reconnect_interval_sec` | Base delay between reconnect attempts |
| `max_reconnect_attempts` | Max retries before giving up (0 = infinite) |
| `stream_timeout_sec` | Timeout for initial stream connection |

## Usage

```bash
# Basic launch
python monitor.py

# Custom config path
python monitor.py --config /path/to/my_config.json

# Debug mode (verbose logging)
python monitor.py --debug
```

### Controls

| Action | How |
|---|---|
| **Snapshot** | Click 📷 on a camera tile |
| **Record** | Click ⏺ to start, ⏹ to stop |
| **Fullscreen** | Double-click a camera tile or click ⛶ |
| **Exit fullscreen** | Press ESC or double-click |
| **Reconnect** | Click 🔄 on a camera tile |
| **Add camera** | Click ➕ Add Camera in the toolbar |
| **Adjust grid** | Change the Grid spinner in the toolbar |

## Remote Access (View Cameras From Anywhere)

By default, RTSP streams only work on your home Wi-Fi. To access cameras from work, your phone on mobile data, or anywhere else, set up **Tailscale** — a zero-config mesh VPN.

### Why Tailscale

- Free for personal use (up to 100 devices)
- No port forwarding, no dynamic DNS, no certificates
- Encrypted WireGuard tunnel between your devices
- Works behind NAT/firewalls automatically
- Takes 5 minutes to set up

### Setup

#### 1. Create a Tailscale account

Go to [tailscale.com](https://tailscale.com) and sign up (free tier is enough).

#### 2. Install on your home PC (the one running the monitor)

**Linux:**
```bash
curl -fsSL https://tailscale.com/install.sh | sh
sudo tailscale up
```

**Windows:**
Download from [tailscale.com/download/windows](https://tailscale.com/download/windows) and run the installer. Click "Log in" in the system tray.

**Mac:**
```bash
brew install --cask tailscale
```

After login, note the Tailscale IP assigned to your home PC (e.g. `100.x.y.z`). You can find it with:
```bash
tailscale ip -4
```

#### 3. Install on your phone / work laptop

- **Android/iOS**: Install "Tailscale" from the app store and log in with the same account
- **Work laptop**: Install Tailscale and log in

All devices on the same Tailscale account can reach each other via their Tailscale IPs.

#### 4. Update camera config for remote access

The cameras themselves stay on your local network. The monitor app runs on your home PC. To view from a remote device, you have two options:

**Option A — Run the monitor on your home PC, access via remote desktop**

1. Install Tailscale on home PC and remote device
2. Use any remote desktop tool (e.g. RustDesk, Parsec) to connect to your home PC via its Tailscale IP
3. The monitor app runs locally and you view it remotely

**Option B — Run the monitor on your remote device, tunnel RTSP through Tailscale**

1. Install Tailscale on your home PC and on your remote device
2. On your home PC, enable Tailscale subnet routing to expose your local network:
   ```bash
   sudo tailscale up --advertise-routes=192.168.1.0/24
   ```
3. In the Tailscale admin console, approve the subnet route
4. On your remote device, the cameras' local IPs (192.168.1.x) become reachable through the tunnel
5. Run the monitor app with the same `config.json` — the RTSP URLs work as if you were home

**Option C — Forward RTSP ports with Tailscale Funnel (advanced)**

For a mobile app in the future, you can use `tailscale serve` to expose specific ports. This is more relevant when building the mobile version.

### Security Notes

- Tailscale traffic is end-to-end encrypted (WireGuard)
- No camera streams are exposed to the public internet
- Only devices logged into your Tailscale account can connect
- You can add ACLs in the Tailscale admin console for extra control

## Architecture

```
camera/
├── monitor.py          # Entry point (CLI + app init)
├── main_window.py      # Main window, toolbar, grid management
├── camera_widget.py    # Individual camera tile (UI + controls)
├── camera_worker.py    # Background thread for RTSP capture
├── config.json         # Camera configuration
├── requirements.txt    # Python dependencies
└── README.md
```
