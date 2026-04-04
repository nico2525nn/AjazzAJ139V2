@echo off
echo Installing requirements...
pip install -r requirements.txt

echo Building executable with PyInstaller...
pip install pyinstaller

REM Build as a single executable without console
pyinstaller --noconfirm --onefile --windowed main.py

echo.
echo Build complete. You can find the output in the "dist" folder.
pause
