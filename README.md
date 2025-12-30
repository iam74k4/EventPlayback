# EventPlayback

![License](https://img.shields.io/github/license/iam74k4/EventPlayback?style=flat-square)
![Python](https://img.shields.io/badge/python-3.7+-blue?style=flat-square)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey?style=flat-square)

A lightweight application for recording and playing back mouse and keyboard input.

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Download](#download)
- [Installation](#installation)
- [Usage](#usage)
- [Controls](#controls)
- [Hotkeys](#hotkeys)
- [File Format](#file-format)
- [Notes](#notes)
- [Build](#build)
- [License](#license)

## Features

- Single file structure (approximately 800 lines)
- Minimal dependencies (pynput, keyboard, customtkinter)
- Modern dark theme UI
- Save and load macros in JSON format

## Requirements

- Python 3.7 or higher
- Windows (DPI Awareness supported)
- Administrator privileges may be required when using global hotkeys

## Download

### Executable File (Recommended)

**Download from GitHub Releases** (latest stable version):
- Download the latest version of `EventPlayback.exe` from the [Releases page](https://github.com/iam74k4/EventPlayback/releases)
- After downloading, you can use it by simply running the exe file

**Download from GitHub Actions Artifacts** (development version or manual builds):
- Select the latest build from the [Actions page](https://github.com/iam74k4/EventPlayback/actions)
- Download `EventPlayback-exe` from the "Artifacts" section
- Note: Artifacts are kept for 30 days only

### Run from Source Code

If Python is installed, you can also run directly from the source code.

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
python main.py
```

## Controls

| Button | Function |
|--------|----------|
| ● Record | Start recording after 3-second countdown |
| ■ Stop | Stop recording/playback |
| ▶ Play | Start playback after 3-second countdown |
| ×[number] | Loop count (0=infinite) |

## Hotkeys

| Key | Function |
|-----|----------|
| F9 | Start/Stop recording |
| F10 | Start/Stop playback |
| Escape | Stop |

## File Format

Macros are saved in JSON format. The structure is as follows:

```json
{
  "name": "Macro Name",
  "created_at": "2024-01-01T00:00:00",
  "events": [
    {
      "type": "mouse_move",
      "timestamp": 0.0,
      "x": 100,
      "y": 200
    }
  ]
}
```

## Notes

**Security Warning**: This application records all keyboard input. Please stop recording when entering passwords, credit card numbers, or other sensitive information. Recorded macro files may contain input content, so please manage them appropriately.

## Troubleshooting

### Hotkeys Not Working

- Try running with administrator privileges
- Check if other applications are using the same hotkeys
- Restart the application

### Recording Not Working Properly

- Check if mouse and keyboard inputs are being detected correctly
- Restart the application

### Playback Not Working Properly

- Check if the recorded macro file is in the correct format
- Check if other applications are interfering during playback
- If an error message is displayed, check the message shown in the GUI

## Build

### Local Build

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name EventPlayback main.py
```

### Automatic Build (GitHub Actions)

When you push a tag to GitHub, it will automatically build and upload the exe file to GitHub Releases.

```bash
# Create and push a tag
git tag v1.0.0
git push origin v1.0.0
```

Alternatively, you can manually run the workflow from the Actions tab on GitHub.

## License

MIT License
