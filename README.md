# Ajazz AJ139 V2 Control Software

[日本語版 README](./README_JA.md)

Unofficial on-device control software for the [Ajazz AJ139 V2](https://139v2mc.yjx2012.com/) wireless mouse.

This project replaces the vendor web UI with a local Python application built around `hidapi` and Tkinter. It supports direct device configuration, tray battery monitoring, key remapping, and macro management without depending on a browser wrapper.

## Features
- Status page with firmware, online state, and battery information
- Performance settings
- DPI levels
- Polling rate
- Debounce time
- Lift-off distance
- Lighting / sleep settings
- Key Mapping tab for all 8 on-device button slots
- Mouse presets, media presets, normal keyboard keys, and modifier combos
- Extended function key support up to `F24`
- Macro tab with 32 macro slots
- On-device macro read / write / reset
- Macro recording and manual event insertion
- Inline macro editing for `Name`, `Action`, and `Delay`
- English / Japanese UI
- Raw HID debug log output
- System tray battery indicator

## Requirements
- Python 3.8+
- Windows recommended
- Python packages:
  - `hidapi`
  - `Pillow`
  - `pystray`
  - `pyinstaller` for building

Install dependencies with:

```powershell
python -m pip install -r requirements.txt
```

## Running

```powershell
python main.py
```

If `pystray` or another package appears missing, make sure you installed dependencies with the same Python you use to launch the app:

```powershell
python -m pip install -r requirements.txt
python main.py
```

## Building

Use the included batch file:

```powershell
build.bat
```

`build.bat` installs requirements into the active Python environment and then runs:

```powershell
python -m PyInstaller --noconfirm --onefile --windowed main.py
```

## Usage Notes
- Closing the main window hides the app to the system tray.
- `Refresh Status` loads the basic device state first.
- Macro data is loaded lazily when the `Macro` tab is opened.
- Heavy device operations run in the background so the UI stays responsive.
- The `Macro` tab shows a status message while loading, writing, or resetting.
- Some macro read chunks may come back empty on this device; the app handles that without blocking for long timeouts.

## Known Limitations
- Macro storage behavior is vendor-specific and was reverse engineered from the web app.
- Devices with empty or sparse macro pages may still load more slowly than normal settings.
- Full validation still depends on real hardware testing.

## Project Files
- [main.py](/D:/学校/python/AjazzAJ139V2/main.py): entry point
- [ui_app.py](/D:/学校/python/AjazzAJ139V2/ui_app.py): Tkinter UI and background task handling
- [ajazz_mouse.py](/D:/学校/python/AjazzAJ139V2/ajazz_mouse.py): HID protocol layer
- [build.bat](/D:/学校/python/AjazzAJ139V2/build.bat): Windows build helper
- [requirements.txt](/D:/学校/python/AjazzAJ139V2/requirements.txt): Python dependencies
