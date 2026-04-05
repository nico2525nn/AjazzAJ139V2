@echo off
setlocal

echo Installing requirements into the active Python environment...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo Building executable with PyInstaller...
REM Build as a single executable without console
python -m PyInstaller --noconfirm --onefile --windowed main.py

echo.
echo Build complete. You can find the output in the "dist" folder.
pause
