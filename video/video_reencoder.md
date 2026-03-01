# Video Re-encoder

Re-encode videos to a target file size using FFmpeg. Useful for reducing file size while keeping reasonable quality. Uses GPU (NVIDIA NVENC) when available; falls back to CPU encoding. Supports GUI only (run via bat or Python).

## Features

- **Target file size** – Specify desired output size in MB; bitrate is calculated from duration
- **Multiple formats** – MP4, MKV, WebM, AVI, MOV
- **GPU acceleration** – Uses NVIDIA NVENC for MP4/MKV/MOV when available
- **Progress bar** – Live progress during encoding

## Requirements

- **FFmpeg** – Must be installed and in your PATH
- **ffprobe** – For duration detection (included with FFmpeg)
- Python 3.x (tkinter for GUI)
- Optional: NVIDIA GPU with NVENC for faster encoding

## Running the Application

**GUI:**
```bash
python video/video_reencoder.py
```

**Double-click:** `video_reencoder.bat`

## GUI Usage

1. **Browse** – Select the video file to re-encode.
2. **Format** – Choose output format (mp4, mkv, webm, avi, mov).
3. **Target size (MB)** – Enter the desired file size in megabytes.
4. **Re-encode** – Click the button. Progress bar shows encoding progress.
5. Output is saved as `{original_name}_reencoded.{format}` in the same folder.

## Examples

### Reduce 200 MB video to 50 MB
1. Select the video file.
2. Set format to **mp4**.
3. Set target size to **50**.
4. Click Re-encode.

### Convert to WebM for web
1. Select the video.
2. Set format to **webm**.
3. Enter target size (e.g. **30** for 30 MB).
4. Click Re-encode.

## Output

- **MP4, MKV, MOV** – H.264 video, AAC audio (or NVENC if GPU available)
- **WebM** – VP9 video, Opus audio
- **AVI** – MPEG-4 video, MP3 audio

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "FFmpeg not found" | Install FFmpeg and add it to your PATH |
| "Could not determine duration" | Check file is a valid video; ensure ffprobe is available |
| "Target size too small" | Increase target MB; minimum depends on video duration |
| Slow encoding | Use NVIDIA GPU for faster NVENC encoding |
