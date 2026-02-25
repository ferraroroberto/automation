# Xiaomi Camera Monitor

Custom desktop application for monitoring Xiaomi home cameras via RTSP streams, bypassing the official Xiaomi Home app.

## Overview

A PyQt6-based multi-camera viewer that connects directly to your Xiaomi cameras over your local network using RTSP. Features include:

- Multi-camera grid view with adjustable columns
- Per-camera snapshot and recording
- Fullscreen mode (double-click or button)
- Auto-reconnect with exponential backoff
- Dark theme UI
- Add/manage cameras at runtime

## Setup

### 1. Install Dependencies

```bash
cd camera
pip install -r requirements.txt
```

### 2. Enable RTSP on Your Xiaomi Cameras

Most Xiaomi cameras support RTSP natively. The method depends on your camera model:

**Option A — Native RTSP (Xiaomi Mijia / Mi Home Security Camera)**

1. Open the **Xiaomi Home** app on your phone
2. Go to the camera → Settings → **Network Info** → note the camera's IP address
3. Some models expose RTSP at: `rtsp://<CAMERA_IP>:554/stream1`
4. Test the URL in **VLC** (Media → Open Network Stream) before configuring

**Option B — Micam RTSP Bridge (recommended for newer models)**

If your camera doesn't natively expose RTSP, use [Micam](https://github.com/nicennnnnnnlee/Micam):

1. Run Micam via Docker on your local network
2. It bridges Xiaomi camera streams to standard RTSP URLs
3. Use the Micam-provided RTSP URLs in the config

**Option C — Custom Firmware (older models like Dafang/Chuangmi)**

For older Xiaomi cameras, flash custom firmware:
- [xiaomi-dafang-hacks](https://github.com/EliasKotlyar/Xiaomi-Dafang-Hacks)
- [chuangmi-720-hack](https://github.com/epalzeolithe/chuangmi-720-hack)

### 3. Configure Cameras

Edit `config.json`:

```json
{
    "cameras": [
        {
            "name": "Luca",
            "rtsp_url": "rtsp://192.168.1.100:554/stream1",
            "enabled": true
        },
        {
            "name": "Living Room",
            "rtsp_url": "rtsp://192.168.1.101:554/stream1",
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

### Finding Your Camera's RTSP URL

Common RTSP URL formats for Xiaomi cameras:

```
rtsp://<IP>:554/stream1
rtsp://<IP>:8554/live
rtsp://admin:password@<IP>:554/h264Preview_01_main
```

**Quick test with VLC:**
1. Open VLC → Media → Open Network Stream
2. Paste your RTSP URL
3. If it works in VLC, it will work in this monitor

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
