# Audio Scripts

Tools for pulling, converting, trimming, and normalizing audio, plus a small text-collation helper. Most are Tkinter file-dialog utilities meant to be double-clicked or run with no arguments; `audio_extractor_core.py` also has a proper CLI for unattended use over a directory of video files.

## 🎬 audio_extractor_core.py

Extracts audio tracks from video files (`.mkv`, `.mp4`) using FFmpeg/FFprobe. For each video it probes all audio streams, writes a **unified** mix of every track (`{name}_audio_unified.mp3`, `libmp3lame` at 192k — a single-track file is just copied through), and, when a video has more than one audio track, additionally writes each track individually (`{name}_audio_track{N}.mp3`).

```bash
# Process every .mkv/.mp4 in a directory (default: current directory)
python audio/audio_extractor_core.py "C:\Videos"

# Open a file-picker dialog to choose specific videos instead
python audio/audio_extractor_core.py --gui
python audio/audio_extractor_core.py -g
```

Arguments:
- `directory` (positional, optional): directory to scan for video files. Defaults to `.`.
- `--gui` / `-g`: skip directory scanning and open a Tkinter multi-file picker (filtered to `*.mkv`/`*.mp4`) instead.

Requires FFmpeg and FFprobe on `PATH`.

## 🖥️ audio_extractor_gui.py

Tkinter GUI front-end for the same extraction logic as `audio_extractor_core.py`: a "Select Video Files" button (filtered to `.mkv`/`.mp4`), an "Extract Audio" button, a live scrolling log panel fed from the `logging` module, and an indeterminate progress bar while extraction runs on a background thread. The "Extract unified audio (all tracks)" / "Extract individual tracks" checkboxes in the Options panel select which outputs are written; both are ticked by default, which reproduces the core script's behaviour (unified mix plus per-track files when a video has multiple audio tracks). Unticking both is refused with a warning.

```bash
python audio/audio_extractor_gui.py
```

No CLI arguments — it's GUI-only. Double-click `audio_extractor_gui.bat` to launch it via the repo's `.venv` without opening a terminal manually. Requires FFmpeg and FFprobe on `PATH`.

## 🔁 convert_ogg_mp3.py

Converts a single `.ogg` file to `.mp3` (192k bitrate) via `pydub`, saving the output alongside the source with the same base name. Resolves its own FFmpeg binary at runtime through `imageio_ffmpeg` (used for both encoding and probing), so it works without a separate FFmpeg install on `PATH`.

```bash
python audio/convert_ogg_mp3.py
```

No CLI arguments — opens a file-picker dialog filtered to `*.ogg`. Requires the `pydub` and `imageio-ffmpeg` Python packages.

## 📝 transcript_collate.py

Collates a folder of `.md` transcript files into a single `.txt` file. Prompts for a directory, sorts the `.md` files by the first run of digits found in each filename (so `part2.md` sorts before `part10.md`), concatenates their contents in that order (dropping a file's last line if it contains both `[` and `]` — e.g. a trailing timestamp/reference marker), and prompts for a save location for the combined output.

```bash
python audio/transcript_collate.py
```

No CLI arguments — directory-picker in, save-as dialog out. No FFmpeg dependency; pure Python + Tkinter.

## 🎚️ audio_normalize.py

Applies two-pass EBU R128 loudness normalization (FFmpeg's `loudnorm` filter, `TP=-1.5:LRA=11`) to a chosen audio file, re-encoding to AAC at 192k. Before normalizing it analyzes the source with FFmpeg's `ebur128` filter (integrated loudness, loudness range, true peak, sample peak) and shows a settings dialog with a target-loudness dropdown (`-23.0`, `-20.0`, `-16.0`, `-14.0` LUFS; default `-16.0`). After processing it shows a before/after comparison of all four metrics. Output is written next to the source as `{name}_normalized_{target}LUFS{ext}`.

```bash
# Interactive GUI mode: pick a file, choose a target LUFS, review results
python audio/audio_normalize.py

# Non-interactive test mode: run the analyze -> normalize -> analyze pipeline
# on a specific file at a fixed -16.0 LUFS target, no dialogs
python audio/audio_normalize.py --test "C:\Audio\clip.mp3"
```

Arguments:
- `--test [file]`: skips all dialogs. Uses the given path, or falls back to the `AUDIO_NORMALIZE_TEST_FILE` environment variable if no path is given; exits with an error if neither is set.

Accepts `.mp3 .wav .flac .aac .ogg .m4a .wma .aiff` in the file picker. Requires FFmpeg on `PATH`.

## ✂️ audio_trim.py

Trims an audio file to a start/end time range, applies `loudnorm` normalization, and re-encodes to AAC. Prompts for a file, then for start and end times validated against `mm:ss` or `hh:mm:ss` (re-prompting on invalid input). Output is always `.m4a`, named `{name}_trim_{start}_{end}.m4a` with `:` replaced by `-` in the time labels.

```bash
python audio/audio_trim.py
```

No CLI arguments — file picker, then two time-entry dialogs. Accepts `.mp3 .wav .flac .aac .ogg .m4a .wma .aiff` in the file picker. Requires FFmpeg on `PATH`.
