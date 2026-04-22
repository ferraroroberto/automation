# 🎙️ Voice Transcription

Local, fast, always-on voice-to-text backed by a shared `whisper.cpp` server.
Designed as a drop-in WhisperFlow replacement: tray-resident app, global
hotkey, instant record → transcribe → clipboard.

## 🚀 Overview

**What changed from the legacy version** (see `legacy/`):
- No more in-process Whisper / PyTorch / CUDA. All transcription happens over
  HTTP against a **local `whisper.cpp` server** on a fixed port.
- The server is defined once in `whisper_server/whisper_server.yaml` and the
  same folder drops unchanged into [`claude-local-calls`](https://github.com/ferraroroberto/claude-local-calls)
  so either project can own the server — the port collision is the mutual
  exclusion.
- New **system-tray mode** with a **global hotkey** (default `Ctrl+Alt+Space`)
  for day-to-day use.
- Structure reorganised per [`AGENTS_STRUCTURE.md`](../../AGENTS_STRUCTURE.md)
  and [`AGENTS_CLI.md`](../../AGENTS_CLI.md): clean `cli/`, `core/`, `gui/`,
  `whisper_server/`, `config/` split.

## 📋 Usage

### Day-to-day: tray + global hotkey

```bat
tray.bat
```

A microphone icon appears in the system tray. Press **Ctrl+Alt+Space** to
start recording; press it again to stop. The transcription lands in the
clipboard and a toast notification shows the first line.

Tray menu: Record / Open window / Start server / Stop server / Quit. Only
**Quit** stops the server (and only if this app started it).

### Stream Deck / one-shot

```bat
quick_record.bat
quick_record_english.bat
quick_record_spanish.bat
```

Records until Enter, transcribes, copies to clipboard, exits. No persistent
process. Use this as the Elgato Stream Deck command.

### Classic main window

```bat
transcribe_voice.bat
```

Tkinter window with server status, start/stop, language dropdown, big
Record button, and file-transcribe dialog.

### CLI

```bash
python launcher.py record [--language LANG] [--translate|--no-translate] [--start-server]
python launcher.py transcribe <file> [--language LANG] [--translate]
python launcher.py gui
python launcher.py tray
python launcher.py server start|stop|status|logs
```

All subcommands accept `--debug` and `--config path/to/config.json`.

## 🔧 Configuration

Two separate config files — app config vs. server config:

### `config/config.json` (app)

```json
{
  "language": "Spanish",
  "translate": true,
  "max_record_seconds": 300,
  "sample_rate": 16000,
  "preferred_mics": null,
  "machine_specific_mics": {
    "tower": ["el gato wave XLR (Elgato Wave XLR)"],
    "laptop": ["Micrófono (Realtek(R) Audio)"]
  },
  "hotkey": "<ctrl>+<alt>+<space>",
  "auto_copy": true,
  "auto_start_server": false,
  "log_level": "INFO"
}
```

The hotkey uses [`pynput.keyboard.GlobalHotKeys`](https://pynput.readthedocs.io/en/latest/keyboard.html#global-hotkeys)
syntax (angle-bracketed modifiers + a key, `+`-separated).

### `whisper_server/whisper_server.yaml` (server — shared)

```yaml
server:
  host: "127.0.0.1"
  bind_host: "0.0.0.0"
  port: 8090
binary:
  path: "vendor/whisper.cpp/whisper-server"
model:
  path: "vendor/whisper.cpp/models/ggml-small.bin"
args:
  - "--threads"
  - "4"
pid_file: ".whisper_server.pid"
```

**Keep this file identical in `claude-local-calls`** (copy-paste the whole
`whisper_server/` folder). That guarantees the same port, binary path, and
PID location on both sides.

## 🗂️ Layout

```
audio/transcribe_voice/
├── launcher.py                # thin entry (sys.path + CLI main)
├── cli/
│   ├── main.py                # argparse dispatcher
│   └── commands/              # record, transcribe, gui, tray, server
├── core/
│   ├── app_config.py          # AppConfig loader
│   ├── recorder.py            # sounddevice capture
│   └── transcription_client.py  # HTTP → whisper server
├── gui/
│   ├── app.py                 # tkinter main window
│   ├── tray.py                # pystray + pynput hotkey
│   └── recording_popup.py     # compact VU popup
├── whisper_server/            # ← copy this into claude-local-calls as-is
│   ├── manager.py             # spawn / kill / health / PID file
│   └── whisper_server.yaml    # shared config (SAME on both repos)
├── config/
│   └── config.json            # app-level config
├── *.bat                      # Windows launchers
├── requirements.txt
└── legacy/                    # old GUI & whisper-in-process code (to be removed)
```

## 📦 Install

```powershell
& .\.venv\Scripts\python.exe -m pip install -r audio\transcribe_voice\requirements.txt
```

**Whisper server binary** — put a built `whisper.cpp`'s `whisper-server` at
`vendor/whisper.cpp/whisper-server[.exe]` relative to the repo root, and a
GGML model at `vendor/whisper.cpp/models/ggml-small.bin`. See
[whisper.cpp build docs](https://github.com/ggerganov/whisper.cpp#quick-start).

## 🧪 First run

```bat
REM 1. Make sure the binary + model exist
dir vendor\whisper.cpp\whisper-server.exe
dir vendor\whisper.cpp\models\ggml-small.bin

REM 2. Start the server (manual check)
server.bat start
server.bat status

REM 3. Quick record test
quick_record.bat
```

If the server doesn't come up, run `server.bat logs` to see what
`whisper-server` printed.

## 🔗 See also

- [AGENTS.md](../../AGENTS.md)
- [AGENTS_CLI.md](../../AGENTS_CLI.md)
- [AGENTS_STRUCTURE.md](../../AGENTS_STRUCTURE.md)
- [ferraroroberto/claude-local-calls](https://github.com/ferraroroberto/claude-local-calls) — hub for local LLMs; same
  `whisper_server/` folder can live there unchanged.
