@echo off
setlocal

echo Installing requirements into the active Python environment...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo Building executable with PyInstaller...
REM Build as a single executable without console
python -m PyInstaller --noconfirm --onefile --windowed ^
  --add-data "constants.py;." ^
  --add-data "ui_helpers.py;." ^
  --add-data "ui_keymapping.py;." ^
  --add-data "ui_macro.py;." ^
  --add-data "ajazz_mouse.py;." ^
  main.py

echo.
echo Build complete. You can find the output in the "dist" folder.
pause
