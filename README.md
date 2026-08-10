# EventPlayback

![License](https://img.shields.io/github/license/iam74k4/EventPlayback?style=flat-square)
![Python](https://img.shields.io/badge/python-3.10+-blue?style=flat-square)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey?style=flat-square)

A lightweight application for recording and playing back mouse and keyboard input.

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Download](#download)
- [Installation](#installation)
- [Usage](#usage)
- [Controls](#controls)
- [Hotkeys](#hotkeys)
- [Settings](#settings)
- [File Format](#file-format)
- [Architecture](#architecture)
- [Development](#development)
- [Notes](#notes)
- [Troubleshooting](#troubleshooting)
- [Build](#build)
- [License](#license)

## Features

- Layered architecture with a GUI- and OS-independent core
- Minimal dependencies (pynput, customtkinter)
- Modern dark theme UI
- Save and load macros in JSON format
- Settings persist between runs (loop count, last directory, hotkeys)

## Requirements

- Python 3.10 or higher
- Windows, macOS or Linux (X11)
- Platform specific permissions:
  - **Windows**: no special privileges required
  - **macOS**: Accessibility and Input Monitoring permission must be granted in
    System Settings before input can be captured or replayed
  - **Linux**: an X11 session; Wayland does not expose the required APIs

## Download

### Executable File (Recommended)

**Download from GitHub Releases** (latest stable version):
- Download the latest version of `EventPlayback.exe` from the [Releases page](https://github.com/iam74k4/EventPlayback/releases)
- After downloading, you can use it by simply running the exe file

**Download from GitHub Actions Artifacts** (development version or manual builds):
- Select the latest build from the [Actions page](https://github.com/iam74k4/EventPlayback/actions)
- Download `EventPlayback-exe` from the "Artifacts" section
- Note: Artifacts are kept for 30 days only

Pre-built executables are provided for Windows only. On macOS and Linux, run
from source.

### Run from Source Code

If Python is installed, you can also run directly from the source code.

## Installation

```bash
pip install -r requirements.txt
```

Or install the package itself, which also provides an `eventplayback` command:

```bash
pip install -e .
```

## Usage

```bash
python main.py
```

Equivalently:

```bash
python -m eventplayback   # requires the package to be installed
```

## Controls

| Button | Function |
|--------|----------|
| ● Record | Start recording after a 3-second countdown |
| ■ Stop | Stop recording/playback, or cancel the countdown |
| ▶ Play | Start playback after a 3-second countdown |
| ×[number] | Loop count (0=infinite) |
| 📂 / 💾 | Load / save a macro |

## Hotkeys

| Key | Function |
|-----|----------|
| F9 | Start/Stop recording |
| F10 | Start/Stop playback |
| Escape | Stop |

Hotkeys are global and work while other applications have focus. Whichever keys
are bound are automatically excluded from recordings, so pressing F9 to stop
does not end up in the macro.

## Settings

Settings are stored as JSON and loaded at startup:

| Platform | Location |
|----------|----------|
| Windows | `%APPDATA%\EventPlayback\settings.json` |
| macOS | `~/Library/Application Support/EventPlayback/settings.json` |
| Linux | `$XDG_CONFIG_HOME/EventPlayback/settings.json` (default `~/.config`) |

```json
{
  "loop_count": 1,
  "countdown_seconds": 3,
  "always_on_top": true,
  "appearance_mode": "dark",
  "last_directory": "",
  "hotkeys": { "record": "f9", "play": "f10", "stop": "esc" }
}
```

Unreadable or out-of-range values fall back to the defaults, so a damaged
settings file never prevents the application from starting. Hotkeys accept
combinations such as `ctrl+shift+f9`.

## File Format

Macros are saved in JSON format. The structure is as follows:

```json
{
  "version": 1,
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

`type` is one of `mouse_move`, `mouse_click`, `mouse_scroll`, `key_press` or
`key_release`, and `timestamp` is seconds since the start of the recording.
Files written by earlier versions have no `version` field and are still
accepted.

## Architecture

```
src/eventplayback/
├── core/                    Recording and playback. No GUI, no OS input library.
│   ├── events.py            Event / EventType
│   ├── macro.py             Macro, validation and the file format
│   ├── recorder.py          Input capture
│   ├── player.py            PlaybackEngine (synchronous) + Player (threaded)
│   └── backends/
│       ├── base.py          InputSource / InputSynthesizer / HotkeyListener / Clock
│       ├── pynput_backend.py  The only module that imports pynput
│       └── fake.py          In-memory backends and a virtual clock, for tests
├── services/                Settings, macro files, hotkey registration
├── ui/
│   ├── state.py             State -> appearance, as a pure function
│   └── app.py               The customtkinter window
└── platform_support.py      Everything that has to know which OS this is
```

The input library and the clock are injected rather than constructed in place.
That is what allows the whole of `core` to be tested head-lessly on any
platform, and what keeps pynput confined to a single module.

## Development

```bash
pip install -r requirements-dev.txt
pytest        # unit tests; no display server required
ruff check .  # lint
```

The test suite exercises `core`, `services` and `ui.state`, none of which import
pynput or customtkinter.

## Notes

**Security Warning**: This application records all keyboard input. Please stop recording when entering passwords, credit card numbers, or other sensitive information. Recorded macro files may contain input content, so please manage them appropriately.

## Troubleshooting

### Hotkeys Not Working

- **macOS**: grant Accessibility and Input Monitoring permission in System
  Settings > Privacy & Security, then restart the application
- **Linux**: confirm the session is X11 (`echo $XDG_SESSION_TYPE`); Wayland is
  not supported
- Check if other applications are using the same hotkeys
- Rebind the conflicting hotkey in `settings.json`

### Recording Not Working Properly

- Check if mouse and keyboard inputs are being detected correctly
- Confirm the input permissions listed under [Requirements](#requirements)
- Restart the application

### Playback Not Working Properly

- Check if the recorded macro file is in the correct format
- Check if other applications are interfering during playback
- If an error message is displayed, check the message shown in the GUI

### Coordinates Are Off on a Scaled Display

Recording and playback must run at the same display scaling. On Windows the
application declares per-monitor DPI awareness at startup to keep coordinates
consistent.

## Build

### Local Build

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --paths src --name EventPlayback main.py
```

### Automatic Build (GitHub Actions)

Every push and pull request runs lint and tests. When you push a tag, the
workflow additionally builds and uploads the exe file to GitHub Releases.

```bash
# Create and push a tag
git tag v1.0.0
git push origin v1.0.0
```

Alternatively, you can manually run the workflow from the Actions tab on GitHub.

## License

MIT License
