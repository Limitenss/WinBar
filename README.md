# WinBar-v2

A beautiful, customizable floating top/bottom bar for Windows 10/11, built with Python and CustomTkinter.

![GitHub last commit](https://img.shields.io/github/last-commit/USER/WinBar-v2)
![GitHub license](https://img.shields.io/github/license/USER/WinBar-v2)

## Features

- **Pill-shaped design**: Modern aesthetics with rounded corners and transparency.
- **Modular Layout**: Configure widgets in left, center, and right zones.
- **Built-in Widgets**:
  - App Launcher (with overflow menu for >14 apps)
  - Active Window Tracker
  - System Monitor (CPU/RAM)
  - Search (Shortcut: `Ctrl+Space`)
  - Settings / Control Center
- **Hot-Reloading**: Use `dev_runner.py` for real-time development.
- **High DPI Support**: Crystal clear on 4K monitors.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/USER/WinBar-v2.git
   cd WinBar-v2
   ```

2. Create a virtual environment and install dependencies:
   ```bash
   python -m venv .venv
   source .venv/scripts/activate  # On Windows
   pip install -r requirements.txt
   ```

3. Run the application:
   ```bash
   python main.pyw
   ```

## Configuration

The application uses a `config.json` file. Copy the template to get started:
```bash
cp config.json.template config.json
```

## Building Executable

To create a standalone `.exe`:
```bash
pyinstaller --noconfirm --onefile --windowed --name "Limitens_FloatingBar" --add-data "C:/Users/Admin/AppData/Local/Programs/Python/Python310/lib/site-packages/customtkinter;customtkinter/" --icon="NONE" "main.pyw"
```

## License

Distributed under the MIT License. See `LICENSE` for more information.
