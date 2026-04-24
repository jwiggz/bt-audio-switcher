#!/usr/bin/env python3
"""
Bluetooth Headset Switcher
Switches the Windows default audio output device via hotkeys.
Also connects Bluetooth if the device needs it.

Requirements: pip install pystray pillow keyboard
Run via run.bat (as Administrator).
"""

import ctypes
import io
import json
import re
import subprocess
import threading
import time
import tkinter as tk
import urllib.request
import zipfile
from tkinter import messagebox, ttk
from pathlib import Path

import keyboard
import pystray
from PIL import Image, ImageDraw


CONFIG_PATH = Path.home() / ".bluetooth_switcher.json"
APP_NAME = "BT Audio Switcher"
PS_CMD = ["powershell", "-NoProfile", "-NonInteractive", "-Command"]

# SoundVolumeView (NirSoft) — used to forcibly move per-app audio streams
TOOLS_DIR = Path.home() / ".bluetooth_switcher"
SVVIEW_PATH = TOOLS_DIR / "SoundVolumeView.exe"
SVVIEW_URL = "https://www.nirsoft.net/utils/soundvolumeview-x64.zip"


def ensure_svview():
    """Download SoundVolumeView.exe if not already cached. Returns True if available."""
    if SVVIEW_PATH.exists():
        return True
    try:
        TOOLS_DIR.mkdir(parents=True, exist_ok=True)
        with urllib.request.urlopen(SVVIEW_URL, timeout=20) as resp:
            data = resp.read()
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            for name in z.namelist():
                if name.lower().endswith("soundvolumeview.exe"):
                    with z.open(name) as src, open(SVVIEW_PATH, "wb") as dst:
                        dst.write(src.read())
                    break
        return SVVIEW_PATH.exists()
    except Exception:
        return False


def svview_set_default(device_name):
    """Set system default for all roles via SoundVolumeView."""
    if not SVVIEW_PATH.exists():
        return False
    try:
        subprocess.run(
            [str(SVVIEW_PATH), "/SetDefault", device_name, "all"],
            capture_output=True, timeout=10,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        return True
    except Exception:
        return False


def svview_force_all_apps(device_name):
    """
    Force every running audio app (Spotify, Firefox, Discord, SteelSeriesSonar,
    Steam, etc.) to route its audio to device_name via per-app routing.

    Per-app routing overrides the system default, so even if SteelSeries Sonar
    keeps resetting the system default back to Sonos, each app's audio will
    still go to AirPods.
    """
    if not SVVIEW_PATH.exists():
        return
    try:
        import tempfile
        tmp = Path(tempfile.gettempdir()) / "svview_sessions.txt"
        subprocess.run(
            [str(SVVIEW_PATH), "/stext", str(tmp)],
            capture_output=True, timeout=15,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        if not tmp.exists():
            return

        content = tmp.read_text(encoding="utf-8", errors="ignore")
        try:
            tmp.unlink()
        except Exception:
            pass

        # Parse every unique process name from audio sessions
        processes = set()
        for line in content.splitlines():
            line = line.strip()
            if line.lower().startswith("process name") and ":" in line:
                proc = line.split(":", 1)[1].strip()
                if proc and proc not in ("", "?", "System"):
                    processes.add(proc)

        # Set per-app audio default for each process
        for proc in processes:
            subprocess.run(
                [str(SVVIEW_PATH), "/SetAppDefault", device_name, "all", proc],
                capture_output=True, timeout=5,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
    except Exception:
        pass

BT_SKIP = [
    "microsoft", "bluetooth", "generic", "root", "hid",
    "hands-free", "audio gateway", "enumerator",
]

# Injected at the top of every AudioDeviceCmdlets script.
# Auto-installs the module if missing.
PS_AUDIO_INIT = r"""
if (-not (Get-Module AudioDeviceCmdlets -EA SilentlyContinue)) {
    if (-not (Get-Module -ListAvailable AudioDeviceCmdlets -EA SilentlyContinue)) {
        Install-Module AudioDeviceCmdlets -Force -Scope CurrentUser -AllowClobber -EA SilentlyContinue
    }
    Import-Module AudioDeviceCmdlets -EA SilentlyContinue
}
"""

WINRT_INIT = r"""
Add-Type -AssemblyName System.Runtime.WindowsRuntime
[void][Windows.Devices.Bluetooth.BluetoothDevice,Windows.Devices.Bluetooth,ContentType=WindowsRuntime]
[void][Windows.Devices.Enumeration.DeviceInformation,Windows.Devices.Enumeration,ContentType=WindowsRuntime]
[void][Windows.Devices.Bluetooth.Rfcomm.RfcommDeviceService,Windows.Devices.Bluetooth.Rfcomm,ContentType=WindowsRuntime]
$_ag = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
function WinAwait($t,$tp){ $a=$_ag.MakeGenericMethod($tp); $n=$a.Invoke($null,@($t)); $n.Wait(8000)|Out-Null; $n.Result }
"""


def load_config():
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"devices": {}}


def save_config(config):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2)


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


def run_ps(script, timeout=25):
    try:
        r = subprocess.run(PS_CMD + [script], capture_output=True, text=True,
                           timeout=timeout, creationflags=subprocess.CREATE_NO_WINDOW)
        return r.stdout.strip()
    except Exception as e:
        return f"ERROR: {e}"


# ---------------------------------------------------------------------------
# Windows Audio helpers  (primary mechanism for switching / status)
# ---------------------------------------------------------------------------

def get_audio_output_devices():
    """Return [(name, index)] for all Windows audio playback devices."""
    script = PS_AUDIO_INIT + r"""
$devs = Get-AudioDevice -List -EA SilentlyContinue | Where-Object { $_.Type -eq 'Playback' }
foreach ($d in $devs) { "$($d.Index)|$($d.Name)" }
"""
    out = run_ps(script, timeout=35)
    results = []
    for line in out.splitlines():
        parts = line.split("|", 1)
        if len(parts) == 2:
            idx, name = parts[0].strip(), parts[1].strip()
            if idx.isdigit() and name:
                results.append((name, idx))
    return results


def set_audio_endpoint_state(audio_name, enable: bool):
    """Enable or disable a Windows audio endpoint by its friendly name.
    This prevents background apps (Sonos, SteelSeries Sonar, etc.) from
    reclaiming the default audio device after we switch away from it."""
    safe = audio_name.replace("'", "''")
    action = "Enable-PnpDevice" if enable else "Disable-PnpDevice"
    script = f"""
$devs = Get-PnpDevice -Class AudioEndpoint -EA SilentlyContinue |
        Where-Object {{ $_.FriendlyName -like '*{safe}*' }}
if (-not $devs) {{
    # Try shorter keyword match
    foreach ($w in ('{safe}' -split ' ' | Where-Object {{ $_.Length -gt 3 }})) {{
        $devs = Get-PnpDevice -Class AudioEndpoint -EA SilentlyContinue |
                Where-Object {{ $_.FriendlyName -like "*$w*" }}
        if ($devs) {{ break }}
    }}
}}
if ($devs) {{
    $devs | ForEach-Object {{ {action} -InstanceId $_.InstanceId -Confirm:$false -EA SilentlyContinue }}
    Write-Output "OK: $($devs.Count) devices"
}} else {{
    Write-Output "NOTFOUND"
}}
"""
    return run_ps(script, timeout=15)


def get_current_audio_name():
    """Return the name of the current default Windows playback device."""
    script = PS_AUDIO_INIT + r"""
$d = Get-AudioDevice -Playback -EA SilentlyContinue
if ($d) { Write-Output $d.Name }
"""
    return run_ps(script, timeout=20).strip()


def set_default_audio(audio_name):
    """Set default playback + communications device using the stable device ID."""
    safe = audio_name.replace("'", "''")
    script = PS_AUDIO_INIT + f"""
$all = Get-AudioDevice -List -EA SilentlyContinue | Where-Object {{ $_.Type -eq 'Playback' }}
$dev = $all | Where-Object {{ $_.Name -like '*{safe}*' }} | Select-Object -First 1
if (-not $dev) {{
    foreach ($w in ('{safe}' -split ' ' | Where-Object {{ $_.Length -gt 3 }})) {{
        $dev = $all | Where-Object {{ $_.Name -like "*$w*" }} | Select-Object -First 1
        if ($dev) {{ break }}
    }}
}}
if ($dev) {{
    # Use stable device ID (not Index which can shift)
    Set-AudioDevice -ID $dev.ID | Out-Null
    # Also set as default communications device so all apps pick it up
    try {{ Set-AudioDevice -ID $dev.ID -DefaultCommunication | Out-Null }} catch {{}}
    # Verify the switch actually took
    $cur = (Get-AudioDevice -Playback -EA SilentlyContinue).Name
    Write-Output "OK: $cur"
}} else {{
    Write-Output "FAIL: no audio device matching '{safe}'"
}}
"""
    out = run_ps(script, timeout=35)
    return out.startswith("OK"), out


# ---------------------------------------------------------------------------
# Bluetooth helpers  (secondary — used to connect devices before audio switch)
# ---------------------------------------------------------------------------

def get_paired_bt_devices():
    script = r"""
$d = Get-PnpDevice -Class Bluetooth -EA SilentlyContinue |
     Where-Object { $_.Status -in @('OK','Error','Unknown','Degraded') } |
     Select-Object FriendlyName, InstanceId
$d | ConvertTo-Json -Depth 2 -Compress
"""
    raw = run_ps(script, timeout=15)
    if not raw or raw.startswith("ERROR"):
        return []
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return []
    if isinstance(data, dict):
        data = [data]
    results = []
    for d in data:
        name = (d.get("FriendlyName") or "").strip()
        iid = (d.get("InstanceId") or "").strip()
        if not name or not iid:
            continue
        if any(s in name.lower() for s in BT_SKIP):
            continue
        results.append((name, iid))
    return results


def extract_mac(iid):
    m = re.search(r'[_\\]([0-9A-Fa-f]{12})(?:[&\\]|$)', iid)
    return m.group(1).upper() if m else None


def bt_connect(instance_id):
    """Enable PnP entries + WinRT RFCOMM open to force BT link."""
    mac = extract_mac(instance_id)
    filt = ("{ $_.InstanceId -match '" + mac + "' }" if mac
            else "{ $_.InstanceId -eq '" + instance_id + "' }")
    # Enable PnP
    run_ps(
        "try { $d = Get-PnpDevice | Where-Object " + filt + "\n"
        "  $d | ForEach-Object { Enable-PnpDevice -InstanceId $_.InstanceId -Confirm:$false -EA SilentlyContinue }"
        "} catch {} \nWrite-Output 'done'",
        timeout=20,
    )
    if not mac:
        return True, "PnP enabled"
    # WinRT connect
    script = WINRT_INIT + """
try {
    $sel  = [Windows.Devices.Bluetooth.BluetoothDevice]::GetDeviceSelector()
    $devs = WinAwait ([Windows.Devices.Enumeration.DeviceInformation]::FindAllAsync($sel)) ([Windows.Devices.Enumeration.DeviceInformationCollection])
    $tgt  = $devs | Where-Object { $_.Id -match '""" + mac + """' } | Select-Object -First 1
    if (-not $tgt) { throw 'not found' }
    $bt   = WinAwait ([Windows.Devices.Bluetooth.BluetoothDevice]::FromIdAsync($tgt.Id)) ([Windows.Devices.Bluetooth.BluetoothDevice])
    $svc  = WinAwait ($bt.GetRfcommServicesAsync([Windows.Devices.Bluetooth.BluetoothCacheMode]::Uncached)) ([Windows.Devices.Bluetooth.Rfcomm.RfcommDeviceServicesResult])
    Write-Output "OK: $($bt.Name)"
} catch { Write-Output "FAIL: $_" }
"""
    out = run_ps(script, timeout=35)
    return out.startswith("OK"), out


def bt_disconnect(instance_id):
    mac = extract_mac(instance_id)
    filt = ("{ $_.InstanceId -match '" + mac + "' }" if mac
            else "{ $_.InstanceId -eq '" + instance_id + "' }")
    script = (
        "try { $d = Get-PnpDevice | Where-Object " + filt + "\n"
        "  if (-not $d) { throw 'none' }\n"
        "  $d | ForEach-Object { Disable-PnpDevice -InstanceId $_.InstanceId -Confirm:$false -EA SilentlyContinue }\n"
        '  Write-Output "OK: $($d.Count) disabled"\n'
        "} catch { Write-Output \"FAIL: $_\" }"
    )
    out = run_ps(script, timeout=25)
    return out.startswith("OK"), out


# ---------------------------------------------------------------------------
# Full connect: BT (if instance_id known) + set Windows audio default
# ---------------------------------------------------------------------------

def get_exact_device_name(audio_name):
    """Return the exact Windows audio device name that matches audio_name."""
    safe = audio_name.replace("'", "''")
    script = PS_AUDIO_INIT + f"""
$all = Get-AudioDevice -List -EA SilentlyContinue | Where-Object {{ $_.Type -eq 'Playback' }}
$dev = $all | Where-Object {{ $_.Name -like '*{safe}*' }} | Select-Object -First 1
if ($dev) {{ Write-Output $dev.Name }}
"""
    name = run_ps(script, timeout=20).strip()
    return name if name else audio_name



def connect_device(info: dict):
    """
    Switch the Windows audio default to the target device.
    Must run WITHOUT elevation (non-admin) so the change is visible to all
    user-session apps (Firefox, Spotify, Volume Mixer, etc.).
    """
    # audio_name is the Windows audio device name (e.g. "Headphones (AirPods Pro)")
    # Fall back to the device key name (_name) if audio_name wasn't set in config.
    audio_name = info.get("audio_name") or info.get("_name") or ""
    if not audio_name:
        return False, "No audio name configured — open Settings and re-add the device"

    ok, msg = set_default_audio(audio_name)
    time.sleep(1)

    cur = get_current_audio_name()
    # Accept partial match: "AirPods Pro" matches "Headphones (AirPods Pro)"
    # Use a keyword from the name for matching
    keyword = audio_name.split("(")[-1].rstrip(")").strip() if "(" in audio_name else audio_name
    if keyword.lower() in cur.lower():
        return True, f"OK: {cur}"
    return False, f"Could not switch — still on: {cur}"


# ---------------------------------------------------------------------------
# Hotkey manager
# ---------------------------------------------------------------------------

class HotkeyManager:
    def __init__(self):
        self._b = {}

    def register(self, hk, cb):
        self.unregister(hk)
        try:
            self._b[hk] = keyboard.add_hotkey(hk, cb, suppress=False)
        except Exception as e:
            print(f"[hk] {hk}: {e}")

    def unregister(self, hk):
        if hk in self._b:
            try:
                keyboard.remove_hotkey(self._b[hk])
            except Exception:
                pass
            del self._b[hk]

    def unregister_all(self):
        for hk in list(self._b):
            self.unregister(hk)


# ---------------------------------------------------------------------------
# Tray icon
# ---------------------------------------------------------------------------

def make_icon(size=64):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    r, cx, cy = size // 2 - 2, size // 2, size // 2
    draw.ellipse([cx-r, cy-r, cx+r, cy+r], fill=(0, 120, 215))
    s, lw = size/64, max(2, int(3*size/64))
    for seg in [((28,16),(28,48)),((28,16),(40,24)),((40,24),(28,32)),((28,32),(40,40)),((40,40),(28,48))]:
        draw.line([(seg[0][0]*s, seg[0][1]*s), (seg[1][0]*s, seg[1][1]*s)], fill="white", width=lw)
    return img


# ---------------------------------------------------------------------------
# Settings window
# ---------------------------------------------------------------------------

class SettingsWindow:
    def __init__(self, app):
        self.app = app
        self._win = None

    def open(self):
        if self._win and self._win.winfo_exists():
            self._win.lift(); self._win.focus_force(); return
        self._build()

    def close(self):
        if self._win and self._win.winfo_exists():
            self._win.destroy()
        self._win = None

    def _build(self):
        win = tk.Toplevel(self.app.root)
        win.title(APP_NAME + " - Settings")
        win.geometry("700x440")
        win.protocol("WM_DELETE_WINDOW", self.close)
        self._win = win

        hdr = tk.Frame(win, bg="#0078d4", height=56)
        hdr.pack(fill=tk.X); hdr.pack_propagate(False)
        tk.Label(hdr, text="BT Audio Switcher", font=("Segoe UI",13,"bold"),
                 bg="#0078d4", fg="white").pack(side=tk.LEFT, padx=16, pady=10)
        adm = "Admin: YES" if is_admin() else "Admin: NO"
        col = "#90ee90" if is_admin() else "#ffcc00"
        tk.Label(hdr, text=adm, font=("Segoe UI",9), bg="#0078d4", fg=col).pack(side=tk.RIGHT, padx=16)

        f = tk.Frame(win)
        f.pack(fill=tk.BOTH, expand=True, padx=10, pady=(8,0))
        cols = ("Active", "Device", "Audio device name", "Hotkey")
        self._tree = ttk.Treeview(f, columns=cols, show="headings", height=9)
        self._tree.heading("Active", text="")
        self._tree.heading("Device", text="Device")
        self._tree.heading("Audio device name", text="Windows audio name")
        self._tree.heading("Hotkey", text="Hotkey")
        self._tree.column("Active", width=28, anchor=tk.CENTER, stretch=False)
        self._tree.column("Device", width=170, anchor=tk.W)
        self._tree.column("Audio device name", width=230, anchor=tk.W)
        self._tree.column("Hotkey", width=130, anchor=tk.CENTER)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb = ttk.Scrollbar(f, orient=tk.VERTICAL, command=self._tree.yview)
        self._tree.configure(yscrollcommand=sb.set); sb.pack(side=tk.RIGHT, fill=tk.Y)

        bf = tk.Frame(win); bf.pack(pady=6)
        for lbl, cmd in [("Scan audio devices", self._scan_audio),
                         ("Scan BT devices",    self._scan_bt),
                         ("Set hotkey",         self._set_hotkey),
                         ("Connect now",        self._connect_now),
                         ("Disconnect BT",      self._disconnect_now),
                         ("Remove",             self._remove),
                         ("Refresh",            self._refresh_status)]:
            ttk.Button(bf, text=lbl, command=cmd).pack(side=tk.LEFT, padx=3)

        self._sv = tk.StringVar(value="Tip: use 'Scan audio devices' first, then assign hotkeys.")
        tk.Label(win, textvariable=self._sv, anchor=tk.W, font=("Segoe UI",9),
                 fg="#555").pack(fill=tk.X, padx=10, pady=(2,6))
        self._refresh()

    def _refresh(self):
        if not self._win: return
        cur = self.app.current_audio_name.lower()
        for r in self._tree.get_children(): self._tree.delete(r)
        for name, info in self.app.config["devices"].items():
            hk = info.get("hotkey") or "--"
            aname = info.get("audio_name") or ""
            active = "●" if (aname and aname.lower() in cur) else "○"
            self._tree.insert("", tk.END, iid=name, values=(active, name, aname, hk))

    def _status(self, msg):
        if self._win and self._win.winfo_exists():
            self._sv.set(msg)

    def _sel(self):
        s = self._tree.selection()
        return s[0] if s else None

    # Scan Windows audio output devices
    def _scan_audio(self):
        self._status("Scanning Windows audio output devices...")
        def worker():
            devs = get_audio_output_devices()
            self._win.after(0, lambda: self._pick_audio(devs, is_bt=False))
        threading.Thread(target=worker, daemon=True).start()

    def _pick_audio(self, devs, is_bt=False):
        if not devs:
            self._status("No audio devices found.")
            return
        sel = tk.Toplevel(self._win)
        sel.title("Add Audio Device")
        sel.geometry("440x300")
        sel.grab_set()
        tk.Label(sel, text="Select a Windows audio output device:",
                 font=("Segoe UI",10,"bold")).pack(pady=(12,4), padx=12, anchor=tk.W)
        lbf = tk.Frame(sel); lbf.pack(fill=tk.BOTH, expand=True, padx=12)
        lb = tk.Listbox(lbf, font=("Segoe UI",10), height=9)
        lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        lsb = ttk.Scrollbar(lbf, orient=tk.VERTICAL, command=lb.yview)
        lb.configure(yscrollcommand=lsb.set); lsb.pack(side=tk.RIGHT, fill=tk.Y)
        for name, _ in devs:
            lb.insert(tk.END, name)
        def add():
            idx = lb.curselection()
            if not idx: return
            name, _ = devs[idx[0]]
            if name not in self.app.config["devices"]:
                self.app.config["devices"][name] = {"audio_name": name, "hotkey": None}
            else:
                self.app.config["devices"][name]["audio_name"] = name
            save_config(self.app.config)
            self.app.reload_hotkeys()
            self.app.rebuild_tray_menu()
            self._refresh()
            self._status(f"Added: {name}")
            sel.destroy()
        ttk.Button(sel, text="Add", command=add).pack(pady=8)

    # Scan paired BT devices
    def _scan_bt(self):
        self._status("Scanning paired Bluetooth devices...")
        def worker():
            devs = get_paired_bt_devices()
            self._win.after(0, lambda: self._pick_bt(devs))
        threading.Thread(target=worker, daemon=True).start()

    def _pick_bt(self, devs):
        if not devs:
            self._status("No paired BT devices found.")
            return
        sel = tk.Toplevel(self._win)
        sel.title("Add BT Device")
        sel.geometry("440x300")
        sel.grab_set()
        tk.Label(sel, text="Select a paired Bluetooth device:",
                 font=("Segoe UI",10,"bold")).pack(pady=(12,4), padx=12, anchor=tk.W)
        lbf = tk.Frame(sel); lbf.pack(fill=tk.BOTH, expand=True, padx=12)
        lb = tk.Listbox(lbf, font=("Segoe UI",10), height=9)
        lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        lsb = ttk.Scrollbar(lbf, orient=tk.VERTICAL, command=lb.yview)
        lb.configure(yscrollcommand=lsb.set); lsb.pack(side=tk.RIGHT, fill=tk.Y)
        for name, _ in devs: lb.insert(tk.END, name)
        def add():
            idx = lb.curselection()
            if not idx: return
            name, iid = devs[idx[0]]
            entry = self.app.config["devices"].get(name, {})
            entry["instance_id"] = iid
            if not entry.get("audio_name"):
                entry["audio_name"] = name
            self.app.config["devices"][name] = entry
            save_config(self.app.config)
            self.app.reload_hotkeys()
            self.app.rebuild_tray_menu()
            self._refresh()
            self._status(f"Added BT device: {name}. Set audio name if it differs.")
            sel.destroy()
        ttk.Button(sel, text="Add", command=add).pack(pady=8)

    def _set_hotkey(self):
        name = self._sel()
        if not name:
            messagebox.showwarning("No selection", "Select a device first.", parent=self._win)
            return
        hw = tk.Toplevel(self._win)
        hw.title("Set Hotkey"); hw.geometry("340x185"); hw.resizable(False,False); hw.grab_set()
        tk.Label(hw, text="Hotkey for:", font=("Segoe UI",9)).pack(pady=(14,2))
        tk.Label(hw, text=name, font=("Segoe UI",11,"bold")).pack()
        tk.Label(hw, text="Press a key combo:", font=("Segoe UI",9), fg="gray").pack(pady=(8,4))
        captured = [None]
        hkv = tk.StringVar(value="Waiting...")
        tk.Label(hw, textvariable=hkv, font=("Consolas",13,"bold"),
                 bg="#f4f4f4", relief="sunken", padx=8, pady=5).pack(padx=20, fill=tk.X)
        MODS = {"ctrl","alt","shift","left ctrl","right ctrl","left alt","right alt",
                "left shift","right shift","left windows","right windows"}
        def on_key(ev):
            if ev.event_type != keyboard.KEY_DOWN or ev.name in MODS: return
            mods = [m for m in ("ctrl","alt","shift") if keyboard.is_pressed(m)]
            captured[0] = "+".join(mods + [ev.name]); hkv.set(captured[0])
        hook = keyboard.hook(on_key)
        def confirm():
            keyboard.unhook(hook)
            if captured[0]:
                old = self.app.config["devices"][name].get("hotkey")
                if old: self.app.hotkey_mgr.unregister(old)
                self.app.config["devices"][name]["hotkey"] = captured[0]
                save_config(self.app.config)
                self.app.reload_hotkeys(); self.app.rebuild_tray_menu()
                self._refresh(); self._status(f"Hotkey '{captured[0]}' set for {name}")
            hw.destroy()
        def cancel(): keyboard.unhook(hook); hw.destroy()
        bf2 = tk.Frame(hw); bf2.pack(pady=10)
        ttk.Button(bf2, text="Confirm", command=confirm).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf2, text="Cancel",  command=cancel).pack(side=tk.LEFT, padx=5)

    def _connect_now(self):
        name = self._sel()
        if not name: return
        info = self.app.config["devices"][name]
        self._status(f"Connecting {name}...")
        def worker():
            ok, msg = connect_device(info)
            time.sleep(1)
            self.app.current_audio_name = get_current_audio_name()
            self.app.rebuild_tray_menu()
            if ok:
                # msg is "OK: <actual Windows device name>"
                actual = msg[4:].strip() if msg.startswith("OK:") else name
                s = (f"Switched to: {actual}  "
                     f"(If audio didn't move, pause & unpause your music app)")
            else:
                s = f"Failed: {msg}"
            if self._win: self._win.after(0, lambda: (self._status(s), self._refresh()))
        threading.Thread(target=worker, daemon=True).start()

    def _disconnect_now(self):
        name = self._sel()
        if not name: return
        info = self.app.config["devices"][name]
        iid = info.get("instance_id")
        audio_name = (info.get("audio_name") or "").lower()
        if not iid:
            self._status(f"'{name}' has no BT entry — use Connect on another device to switch away.")
            return
        self._status(f"Disconnecting {name}...")
        def worker():
            ok, msg = bt_disconnect(iid)
            time.sleep(1)
            # If audio is still on this device, switch Windows to any other available output
            cur = get_current_audio_name()
            if audio_name and audio_name in cur.lower():
                all_devs = get_audio_output_devices()
                for dev_name, _ in all_devs:
                    if audio_name not in dev_name.lower():
                        set_default_audio(dev_name)
                        time.sleep(1)
                        break
            self.app.current_audio_name = get_current_audio_name()
            self.app.rebuild_tray_menu()
            s = f"Disconnected: {name}" if ok else f"BT disconnect: {msg}"
            if self._win: self._win.after(0, lambda: (self._status(s), self._refresh()))
        threading.Thread(target=worker, daemon=True).start()

    def _remove(self):
        name = self._sel()
        if not name: return
        old = self.app.config["devices"][name].get("hotkey")
        if old: self.app.hotkey_mgr.unregister(old)
        del self.app.config["devices"][name]
        save_config(self.app.config)
        self.app.rebuild_tray_menu(); self._refresh()
        self._status(f"Removed: {name}")

    def _refresh_status(self):
        self._status("Refreshing...")
        def worker():
            self.app.current_audio_name = get_current_audio_name()
            self.app.rebuild_tray_menu()
            if self._win: self._win.after(0, lambda: (self._refresh(), self._status("Refreshed.")))
        threading.Thread(target=worker, daemon=True).start()


# ---------------------------------------------------------------------------
# Main app
# ---------------------------------------------------------------------------

class App:
    def __init__(self):
        self.config = load_config()
        self.hotkey_mgr = HotkeyManager()
        self.root = tk.Tk()
        self.root.withdraw()
        self.root.title(APP_NAME)
        self.settings = SettingsWindow(self)
        self.icon = None
        self.current_audio_name = ""
        self._polling = True
        self.reload_hotkeys()

    def reload_hotkeys(self):
        self.hotkey_mgr.unregister_all()
        for name, info in self.config["devices"].items():
            hk = info.get("hotkey")
            if hk:
                def make_cb(n, i):
                    def cb():
                        self._notify(f"Switching audio to {n}...")
                        def do():
                            ok, msg = connect_device(i)
                            time.sleep(1)
                            self.current_audio_name = get_current_audio_name()
                            self.rebuild_tray_menu()
                            if ok:
                                actual = msg[4:].strip() if msg.startswith("OK:") else n
                                self._notify(f"Audio: {actual}")
                            else:
                                self._notify(f"Failed: {msg[:60]}")
                        threading.Thread(target=do, daemon=True).start()
                    return cb
                self.hotkey_mgr.register(hk, make_cb(name, dict(info)))

    def _poll(self):
        while self._polling:
            try:
                name = get_current_audio_name()
                if name != self.current_audio_name:
                    self.current_audio_name = name
                    self.rebuild_tray_menu()
            except Exception:
                pass
            time.sleep(7)

    def _make_menu(self):
        items = []
        cur = self.current_audio_name.lower()
        for name, info in self.config["devices"].items():
            hk = info.get("hotkey")
            label = f"{name}  [{hk}]" if hk else name
            aname = (info.get("audio_name") or "").lower()
            is_active = bool(aname and aname in cur)

            def make_action(i):
                def action(icon, item):
                    self._notify(f"Switching audio to {i['_name']}...")
                    def do():
                        ok, msg = connect_device(i)
                        time.sleep(1)
                        self.current_audio_name = get_current_audio_name()
                        self.rebuild_tray_menu()
                        if ok:
                            actual = msg[4:].strip() if msg.startswith("OK:") else i['_name']
                            self._notify(f"Audio: {actual}")
                    threading.Thread(target=do, daemon=True).start()
                return action

            def make_check(flag):
                def check(item): return flag
                return check

            inf = dict(info); inf["_name"] = name
            items.append(pystray.MenuItem(label, make_action(inf), checked=make_check(is_active)))

        if items:
            items.append(pystray.Menu.SEPARATOR)

        def open_settings(icon, item):
            self.root.after(0, self.settings.open)

        def quit_app(icon, item):
            self._quit()

        items.append(pystray.MenuItem("Settings", open_settings))
        items.append(pystray.Menu.SEPARATOR)
        items.append(pystray.MenuItem("Quit", quit_app))
        return pystray.Menu(*items)

    def rebuild_tray_menu(self):
        if self.icon:
            try:
                self.icon.menu = self._make_menu()
                self.icon.update_menu()
            except Exception:
                pass

    def _quit(self, icon=None, item=None):
        self._polling = False
        self.hotkey_mgr.unregister_all()
        if self.icon: self.icon.stop()
        self.root.quit()

    def _notify(self, msg):
        if self.icon:
            try: self.icon.notify(msg, APP_NAME)
            except Exception: pass

    def run(self):

        def init():
            ensure_svview()   # download SoundVolumeView before first switch
            self.current_audio_name = get_current_audio_name()
            self.rebuild_tray_menu()
        threading.Thread(target=init, daemon=True).start()
        threading.Thread(target=self._poll, daemon=True).start()

        def tray():
            self.icon = pystray.Icon("bt_sw", make_icon(64), APP_NAME, menu=self._make_menu())
            self.icon.run()
        threading.Thread(target=tray, daemon=True).start()
        self.root.mainloop()


if __name__ == "__main__":
    App().run()
