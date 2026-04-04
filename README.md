# Ajazz AJ139 V2 Control Software

This is an unofficial, on-device control software for the [Ajazz AJ139 V2](https://139v2mc.yjx2012.com/) wireless mouse, created to eliminate the dependency on web-based wrappers and allow system tray background monitoring.

![System Tray Battery Indicator](https://img.shields.io/badge/System%20Tray-Battery%20Indicator-brightgreen)

## Features
- **Cross-platform UI Options**: Modify settings directly from Python cleanly and locally.
- **System Tray Agent**: Running in the background with a system tray icon dynamically showing your real-time mouse battery percentage natively.
- **Full Customizability**:
  - DPI configuration (up to 6 levels, individually adjustable per specific axis/color)
  - Polling rate (125 / 250 / 500 / 1000 Hz)
  - Key Debounce Time
  - Lift Off Distance (LOD: 1mm or 2mm)
  - Customizable LED Sleeper and Lighting modes.
- **i18n Language Settings**: The application supports both English (Default) and Japanese locally.
- **Integrated Debugger**: Intercepts Hex raw communication logs between your computer and the mouse natively.

## Prerequisites

- **Python 3.8+**
- Packages used: `hidapi`, `pystray`, `Pillow`

## Installation & Build Instructions

If you just want the executable, you can easily build it using the included batch file.

1. **For Executable Build (Windows)**:
   Just run `build.bat`. It will automatically install all needed libraries and invoke PyInstaller to create a portable `.exe` file inside `dist/`. No further installation or python dependency is required to run the `main.exe` result!

2. **For Manual Python Run**:
   ```bash
   pip install -r requirements.txt
   python main.py
   ```

## Usage
1. Connect your Ajazz AJ139 V2 matching USB Dongle or Cable.
2. Launch `main.py` or the built `main.exe`.
3. Press **[Refresh Status]** to fetch real firmware / device information.
4. **Closing the Window**: Simply clicking "X" gracefully minimizes the program to the System Tray. Right click on the tray icon to `Open Settings` or firmly `Exit` the tracker.
