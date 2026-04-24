@echo off
:: Bluetooth Headset Switcher launcher
:: Runs WITHOUT elevation so audio switching affects all user apps (Firefox, Spotify, etc.)
:: Elevation was causing Set-AudioDevice to apply only to the admin context, not the user session.

cd /d "%~dp0"

:: Install dependencies if not already installed
python -c "import pystray, PIL, keyboard" >nul 2>&1
if %errorLevel% neq 0 (
    echo Installing pystray...
    python -m pip install pystray
    echo Installing pillow...
    python -m pip install pillow
    echo Installing keyboard...
    python -m pip install keyboard
    echo All packages installed.
)

python bluetooth_switcher.py
if %errorLevel% neq 0 (
    echo.
    echo *** App crashed with the above error ***
    pause
)
