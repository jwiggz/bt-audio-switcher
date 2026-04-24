# BT Audio Switcher

A lightweight Windows system tray app for switching your default audio output device instantly — using keyboard shortcuts or a right-click menu.

Built for people who swap between Bluetooth headphones and speakers throughout the day and don't want to dig through Windows Sound Settings every time.

![Windows 11](https://img.shields.io/badge/Windows-11-0078D4?logo=windows)
![Python 3.9+](https://img.shields.io/badge/Python-3.9%2B-3776AB?logo=python)

---

## Features

- **Global hotkeys** — switch audio devices from any app, even when the window isn't focused
- **System tray icon** — right-click to switch, with a checkmark showing your current device
- **Works with any audio device** — Bluetooth headphones, speakers, USB DACs, HDMI — anything Windows sees
- **Standalone exe** — build once, share with anyone (no Python required on their machine)
- **Persistent settings** — hotkeys and devices saved automatically

---

## Build a standalone exe (Windows **without Python**)

Run **`build_exe.bat`** — it installs PyInstaller and produces `dist\BT Switcher.exe`. That single file runs on any Windows 10/11 PC with no Python required.

> **Note:** Windows SmartScreen may warn about an unsigned exe. Click **More info → Run anyway**.

---


## Quick start (with Python)

1. Install dependencies:
   ```
   pip install pystray pillow keyboard
   ```

2. Install the [AudioDeviceCmdlets](https://github.com/frgnca/AudioDeviceCmdlets) PowerShell module (one time):
   ```powershell
   Install-Module AudioDeviceCmdlets -Scope CurrentUser -Force
   ```

3. Double-click **`BT Switcher.vbs`** to launch (no console window), or run:
   ```
   python bluetooth_switcher.py
   ```

---

## Setup

1. Launch the app — a Bluetooth icon appears in the system tray
2. Right-click → **Settings**
3. Click **Scan audio devices** and add the devices you want to switch between
4. Select a device and click **Set hotkey** to assign a key combo
5. Press your hotkey from anywhere to switch instantly

---

## Requirements

- Windows 10 or 11
- Python 3.9+ (only needed if running from source)
- [AudioDeviceCmdlets](https://github.com/frgnca/AudioDeviceCmdlets) PowerShell module (auto-installed on first run)

---

## Tips

- **Auto-start on login:** Press `Win+R`, type `shell:startup`, and drop a shortcut to `BT Switcher.vbs` (or the exe) in that folder
- **After switching**, if a specific app (Spotify, Firefox, etc.) doesn't move to the new device, pause and unpause it — Windows re-routes the stream on resume
- Run the app **without** administrator privileges — elevation causes audio changes to apply only to the admin session, not your normal apps

---

## License

MIT
