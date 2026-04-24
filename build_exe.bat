@echo off
cd /d "%~dp0"

echo Installing PyInstaller...
python -m pip install pyinstaller --quiet

echo.
echo Building BT Switcher.exe ...
python -m PyInstaller ^
  --onefile ^
  --windowed ^
  --name "BT Switcher" ^
  --hidden-import pystray ^
  --hidden-import PIL ^
  --hidden-import PIL.Image ^
  --hidden-import PIL.ImageDraw ^
  --hidden-import keyboard ^
  bluetooth_switcher.py

if %errorLevel% neq 0 (
    echo.
    echo *** Build failed — see errors above ***
    pause
    exit /b 1
)

echo.
echo === Build complete! ===
echo Your standalone exe is at:
echo   %~dp0dist\BT Switcher.exe
echo.
echo You can copy just that one file to any Windows PC.
echo No Python required on the target machine.
echo.
pause
