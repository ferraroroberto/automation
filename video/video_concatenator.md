# Video Concatenator

Merge multiple videos into a single file using FFmpeg. Uses stream copy (no re-encoding) by default for fast processing. Supports both GUI and command-line interfaces.

## Features

- **GUI** – Add/remove/reorder videos, choose output path, progress bar
- **CLI** – Batch or scripted concatenation
- **Stream copy** – No re-encoding; preserves quality and is very fast
- **Fallback** – Uses filter_complex re-encoding if concat demuxer fails (e.g. incompatible formats)

## Requirements

- **FFmpeg** – Must be installed and in your PATH
- Python 3.x (tkinter for GUI)

## Running the Application

**GUI (default):**
```bash
python video/video_concatenator.py
```
Or: `python video/video_concatenator.py --gui`

**Double-click:** `video_concatenator.bat`

**CLI:**
```bash
python video/video_concatenator.py -o output.mp4 video1.mp4 video2.mp4 video3.mp4
```

## CLI Examples

### Basic concatenation
```bash
python video_concatenator.py -o merged.mp4 part1.mp4 part2.mp4 part3.mp4
```
*Merges the three files in order. Output: `merged.mp4`*

### Sort alphabetically
```bash
python video_concatenator.py -o merged.mp4 clip_c.mp4 clip_a.mp4 clip_b.mp4 --sort
```
*Concatenates in alphabetical order: clip_a, clip_b, clip_c*

### Interactive (file dialogs)
```bash
python video_concatenator.py
```
*Opens GUI. No args = GUI mode.*

## GUI Usage

1. **Add videos** – Click "Add videos..." and select files. Order in the list = concatenation order.
2. **Reorder** – Use "Move up" and "Move down" to change order.
3. **Remove** – Select an item and click "Remove selected".
4. **Output** – Choose where to save the merged file.
5. **Concatenate** – Click the button. Progress bar shows status.

## Output

The concatenated video preserves the format of the first input. All streams (video, audio, subtitles) are copied when possible. If the concat demuxer fails (e.g. different codecs), the tool automatically tries filter_complex (re-encoding) as a fallback.

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "FFmpeg not found" | Install FFmpeg and add it to your PATH |
| Concatenation fails | Videos may have incompatible codecs; filter_complex fallback will re-encode |
| Wrong order | Use Move up/down in GUI, or ensure correct order in CLI |
