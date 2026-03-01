# Video Trimmer

A GUI and CLI tool for trimming videos with FFmpeg. Supports GPU acceleration (NVIDIA NVENC) when available, and precise time control down to 0.1 seconds. Output files include EBU R128 audio normalization and H.264/AAC for compatibility.

## Features

- **GUI** – Select video, choose Trim or Cut middle part, set times with text boxes or +/- 0.1s buttons, queue multiple jobs
- **CLI** – Full command-line support for automation and scripting
- **GPU acceleration** – Uses NVIDIA NVENC when available; falls back to CPU (libx264)
- **Precise timing** – Supports `mm:ss` and `mm:ss.s` (e.g. `1:30.5`) formats
- **Progress bar** – Live progress during encoding (in GUI)

## Requirements

- **FFmpeg** – Must be installed and in your PATH
- **ffprobe** – Usually included with FFmpeg (for duration detection)
- Python 3.x (tkinter for GUI)
- Optional: NVIDIA GPU with NVENC for faster encoding

## Running the Application

**GUI (default):**
```bash
python video/video_trim.py
```
Or: `python video/video_trim.py --gui`

**Double-click:** `video_trim.bat`

**CLI (trim):**
```bash
python video/video_trim.py -f "path/to/video.mp4" -s 1:30 -e 2:45
```

**CLI (cut middle section):**
```bash
python video/video_trim.py -f "path/to/video.mp4" --cut -s 1:00 -e 1:02
```
*Removes 1:00 to 1:02 (2 seconds) and joins the remaining parts.*

## Time Format

| Format   | Example | Description                 |
|----------|---------|-----------------------------|
| `mm:ss`  | `1:30`  | 1 minute 30 seconds         |
| `mm:ss.s`| `1:30.5`| 1 minute 30.5 seconds       |
| `hh:mm:ss`| `1:05:30`| 1 hour 5 min 30 seconds   |
| `hh:mm:ss.s`| `1:05:30.2`| With fraction of second |

## CLI Examples

### Basic trim
```bash
python video_trim.py -f "C:\Videos\recording.mp4" -s 0:10 -e 2:30
```
*Trims from 10 seconds to 2 min 30 sec. Output: `recording_trim_0-10_2-30.mp4`*

### Precise 0.1s trim
```bash
python video_trim.py -f video.mp4 -s 1:23.4 -e 3:45.8
```
*Trims from 1:23.4 to 3:45.8.*

### Interactive (file dialogs)
```bash
python video_trim.py
```
*Opens GUI. No args = GUI mode.*

### With only file specified
```bash
python video_trim.py -f video.mp4
```
*Prompts for start and end times via dialogs.*

### Force GUI from command line
```bash
python video_trim.py --gui
```

## GUI Usage

1. **Browse** – Select a video file. Duration is shown automatically.
2. **Operation** – Choose:
   - **1) Trim** – Keep video from start to end
   - **2) Cut middle part** – Remove a section and join the remaining parts
3. **Start / End** (or "Cut from" / "Cut to" for cut middle) – Enter times or use **+0.1** / **-0.1** buttons.
4. **Add to queue** – Add the job. You can mix trims and cut-middle jobs.
5. **Process queue** – Run all queued jobs. Progress bar shows encoding progress.
6. **Repeat** – After processing, the queue clears. Add more and process again as needed.

## Output Files

Output is saved in the same folder as the input:

- **Trim:** `{original_name}_trim_{start}_{end}.{ext}` (e.g. `input_video_trim_0-10_2-30.mp4`)
- **Cut middle:** `{original_name}_cut_{cut_start}_{cut_end}.{ext}` (e.g. `input_video_cut_1-00_1-02.mp4`)

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "FFmpeg not found" | Install FFmpeg and add it to your PATH |
| "Could not detect duration" | Ensure ffprobe is available; check file is a valid video |
| End before start | End time must be greater than start time |
| Slow encoding | Use NVIDIA GPU; CPU fallback is slower |
