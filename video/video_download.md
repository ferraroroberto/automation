# Video Downloader

A unified GUI application for downloading videos from YouTube, HLS/M3U8 streams, and direct HTTP URLs. Supports batch downloads, resume capability for direct URLs, and custom headers for protected streams.

## Features

- **YouTube** – Downloads via yt-dlp; supports playlists and `?t=` start time in URLs
- **HLS/M3U8** – Downloads live/streaming video via ffmpeg; optional Referer/User-Agent for protected sources (e.g. LinkedIn)
- **Direct URL** – Downloads direct .mp4 (or other) links via HTTP with resume support
- **Auto-detect** – Automatically selects the appropriate method based on URL format

## Requirements

- Python 3.x
- `yt-dlp` – `pip install yt-dlp`
- `requests` – `pip install requests`
- **ffmpeg** – Required for HLS. Install separately:
  - Windows: `choco install ffmpeg` or [ffmpeg.org](https://ffmpeg.org/)
  - Linux: `sudo apt-get install ffmpeg`
  - macOS: `brew install ffmpeg`

## Running the Application

**From project root:**
```bash
python video/video_download.py
```

**Or double-click:** `video_download.bat`

## Usage

1. **Choose mode** – Auto-detect, YouTube, HLS/M3U8, or Direct URL
2. **Paste URLs** – One URL per line in the URLs tab
3. **HLS Options** (optional) – If using HLS, set Referer and User-Agent if the server requires them
4. **Select folder** – Where videos will be saved
5. **Click Download**

---

## Examples

### YouTube – single video
```
https://www.youtube.com/watch?v=dQw4w9WgXcQ
```

### YouTube – video with start time
```
https://www.youtube.com/watch?v=dQw4w9WgXcQ&t=42
```
*Downloads from 42 seconds into the video.*

### YouTube – batch (multiple videos)
```
https://www.youtube.com/watch?v=VIDEO_ID_1
https://youtu.be/VIDEO_ID_2
https://www.youtube.com/watch?v=VIDEO_ID_3&t=10
```

### HLS/M3U8 – LinkedIn live stream
1. Set mode to **HLS/M3U8** (or use Auto-detect)
2. Paste the `.m3u8` URL, e.g.:
   ```
   https://live.licdn.com/bitmovinneuprod/.../video_1024.m3u8
   ```
3. In **HLS Options** tab, set:
   - Referer: `https://www.linkedin.com/`
   - User-Agent: `Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36`
4. Select output folder and download

### Direct URL – MP4 file
```
https://example.com/path/to/video.mp4
```
*Uses HTTP with resume support; interrupted downloads can be continued.*

### Mixed batch (Auto-detect mode)
```
https://www.youtube.com/watch?v=abc123
https://example.com/direct-video.mp4
https://stream.example.com/live/video.m3u8
```
*Each URL is processed with the appropriate method based on its format.*

---

## URL detection rules (Auto-detect mode)

| URL pattern       | Method   |
|-------------------|----------|
| `youtube.com`, `youtu.be` | YouTube (yt-dlp) |
| `.m3u8` in URL    | HLS (ffmpeg)     |
| Any other HTTP(S) | Direct (requests) |

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "FFmpeg not found" | Install ffmpeg and ensure it's in your PATH |
| HLS download fails (403/401) | Set Referer and User-Agent in HLS Options tab |
| Direct download stalls | Server may not support Range (resume); try HLS or another source |
| YouTube "Video unavailable" | yt-dlp may need updating: `pip install -U yt-dlp` |
