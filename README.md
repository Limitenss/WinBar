# WinBar

WinBar is a custom floating taskbar for Windows built with Python and `customtkinter`.

It sits at the top or bottom of your screen as a compact overlay and gives you quick access to running apps, pinned apps, search, system stats, the clock, and a small settings menu. The goal is simple: make the desktop feel a little cleaner and a little more personal without trying to replace all of Windows.

## What It Does

- Shows open apps in a compact floating bar
- Lets you pin apps for quick launching
- Includes a built-in app search
- Displays CPU, RAM, time, and date
- Supports drag-and-drop layout editing
- Can sit at the top or bottom of the screen
- Hides the default Windows taskbar while WinBar is active

## Requirements

- Windows 10
- Python 3.11+ recommended

## Install

```bash
pip install -r requirements.txt
```

## Run

Normal launch:

```bash
pythonw main.pyw
```

Development mode with auto-restart on save:

```bash
python dev_runner.py
```

## Build

If you want to package it as an `.exe`:

```bash
pyinstaller Limitens_FloatingBar.spec
```

## Controls

- `Ctrl + E` toggles edit mode so you can rearrange the bar layout
- `Ctrl + Space` opens search
- Middle click a running app button to close that window

To exit the app, open the tray/settings menu on the bar and click `Quit WinBar`.

## Configuration

WinBar reads from `config.json`. If the file is missing, it creates one with default values.

You can use `config.json.template` as a starting point. A typical config looks like this:

```json
{
  "position": "Top",
  "width_percent": 95,
  "height": 45,
  "bg_color": "#1e1e1e",
  "opacity": 0.8,
  "layout": {
    "left": ["apps", "active_window"],
    "center": ["search"],
    "right": ["taskmanager", "sys_monitor", "clock", "tray"]
  },
  "pinned_apps": [
    {
      "name": "Example App",
      "path": "C:\\Path\\To\\App.exe"
    }
  ]
}
```

### Main Config Options

- `position`: `"Top"` or `"Bottom"`
- `width_percent`: how wide the bar should be relative to the screen
- `height`: bar height in pixels
- `bg_color`: background color
- `opacity`: window transparency
- `layout`: controls which widgets appear on the left, center, and right
- `pinned_apps`: apps you want permanently available on the bar

## Notes

- WinBar adds itself to Windows startup when it launches.
- The app is currently centered around one main file: `main.pyw`.
- The default Windows taskbar is restored when WinBar exits cleanly.

## Project Status

This is a personal Windows customization project and it is still evolving. Some parts are intentionally lightweight, and the focus is on making the bar feel good to use rather than turning it into a huge framework.

## License

This project is licensed under the MIT License. See `LICENSE` for details.
