# QuickDeck

## 🚀 Overview
QuickDeck is a minimal always-on-top command launcher that behaves like a mini Stream Deck with configurable buttons. It's a tiny desktop app for Windows, macOS, and Linux that provides quick access to frequently used commands, URLs, and scripts.

## 📋 Usage
### Installation
```bash
pip install PySimpleGUI pyautogui pyperclip pillow
```

> **Note:** PySimpleGUI is no longer freely available on PyPI — it requires a paid license or registration. Use a separately-obtained PySimpleGUI installation, or migrate to the drop-in [FreeSimpleGUI](https://github.com/spyoungtech/FreeSimpleGUI) fork (`pip install FreeSimpleGUI` and change the import).

### Running
```bash
python quickdeck.py [optional path to config]
```

If no config is provided, it will look for `quickdeck.json` in the current directory.

### Button Types
- **shell**: Execute shell commands
- **url**: Open URLs in default browser
- **copy**: Copy text to clipboard
- **python**: Execute Python code
- **reload**: Reload configuration

## 🔧 Configuration
Create a `quickdeck.json` file with your button configuration. Use `quickdeck.example.json` as a starting template.

### Configuration Structure
```json
{
  "title": "🚀 QuickDeck",
  "geometry": "1000x300+40+40",
  "topmost": true,
  "borderless": true,
  "theme": {
    "bg": "#111827",
    "fg": "#E5E7EB",
    "button_bg": "#1F2937"
  },
  "grid": {
    "rows": 3,
    "cols": 3,
    "pad": 4,
    "button_height": 40,
    "emoji_size": 30,
    "text_size": 12,
    "text_position": "below"
  },
  "buttons": [
    {
      "label": "Browser",
      "emoji": "🌐",
      "type": "url",
      "url": "https://www.example.com"
    }
  ]
}
```

## 📊 Output
Logs to console with emoji indicators:
- ✅ Success messages
- ❌ Error messages
- ⚠️ Warning messages
- 📝 Info messages
