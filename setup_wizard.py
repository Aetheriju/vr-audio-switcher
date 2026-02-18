"""VR Audio Switcher — First-Run Setup Wizard.

Guides new users through device detection, svcl download, config generation,
Windows audio configuration, and shortcut creation.

Run once:  python setup_wizard.py
"""

import ctypes
import csv
import json
import os
import subprocess
import sys
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.resolve()
CONFIG_PATH = SCRIPT_DIR / "config.json"
VM_DEVICES_PATH = SCRIPT_DIR / "vm_devices.json"
SVCL_PATH = SCRIPT_DIR / "svcl.exe"
SVCL_URL = "https://www.nirsoft.net/utils/svcl-x64.zip"
VM_DLL = r"C:\Program Files (x86)\VB\Voicemeeter\VoicemeeterRemote64.dll"
VM_EXE = r"C:\Program Files (x86)\VB\Voicemeeter\voicemeeterpro.exe"

REQUIRED_PACKAGES = ["psutil", "pystray", "PIL"]  # PIL = Pillow

# Registry property keys for "Listen to this device"
LISTEN_PROP_GUID = "{24dbb0fc-9311-4b3d-9cf0-18ff155639d4}"
# Byte patterns captured from a working "Listen to this device" configuration
LISTEN_BYTES_1 = "0x0B,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0xFF,0xFF,0x00,0x00"
LISTEN_BYTES_2 = "0x0B,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00"


# ---------------------------------------------------------------------------
# VoiceMeeter device enumeration
# ---------------------------------------------------------------------------
class VMDeviceEnumerator:
    """Connect to VoiceMeeter and enumerate hardware devices."""

    TYPE_MAP = {1: "mme", 3: "wdm", 5: "ks"}

    def __init__(self):
        self._dll = ctypes.WinDLL(VM_DLL)
        ret = self._dll.VBVMR_Login()
        if ret not in (0, 1):
            raise RuntimeError(f"VoiceMeeter login failed (code {ret})")
        time.sleep(0.2)

    def input_devices(self) -> list[dict]:
        n = self._dll.VBVMR_Input_GetDeviceNumber()
        devs = []
        for i in range(n):
            dt = ctypes.c_long()
            name = ctypes.create_string_buffer(256)
            hwid = ctypes.create_string_buffer(256)
            if self._dll.VBVMR_Input_GetDeviceDescA(
                    ctypes.c_long(i), ctypes.byref(dt), name, hwid) == 0:
                devs.append({
                    "index": i,
                    "type": self.TYPE_MAP.get(dt.value, "?"),
                    "name": name.value.decode("utf-8", errors="replace"),
                })
        return devs

    def output_devices(self) -> list[dict]:
        n = self._dll.VBVMR_Output_GetDeviceNumber()
        devs = []
        for i in range(n):
            dt = ctypes.c_long()
            name = ctypes.create_string_buffer(256)
            hwid = ctypes.create_string_buffer(256)
            if self._dll.VBVMR_Output_GetDeviceDescA(
                    ctypes.c_long(i), ctypes.byref(dt), name, hwid) == 0:
                devs.append({
                    "index": i,
                    "type": self.TYPE_MAP.get(dt.value, "?"),
                    "name": name.value.decode("utf-8", errors="replace"),
                })
        return devs

    def close(self):
        try:
            self._dll.VBVMR_Logout()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# svcl helpers
# ---------------------------------------------------------------------------
def query_svcl_devices() -> list[dict]:
    """Run svcl.exe and return parsed device list."""
    if not SVCL_PATH.exists():
        return []
    tmp = SCRIPT_DIR / "_setup_query.csv"
    try:
        subprocess.run(
            [str(SVCL_PATH), "/scomma", str(tmp),
             "/Columns", "Name,Command-Line Friendly ID,Item ID,Direction,Type,Device State,Device Name"],
            capture_output=True, timeout=10,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        devices = []
        with open(tmp, newline="", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                devices.append({
                    "name": row.get("Name", "").strip(),
                    "friendly_id": row.get("Command-Line Friendly ID", "").strip(),
                    "item_id": row.get("Item ID", "").strip(),
                    "direction": row.get("Direction", "").strip(),
                    "type": row.get("Type", "").strip(),
                    "state": row.get("Device State", "").strip(),
                    "device_name": row.get("Device Name", "").strip(),
                })
        return devices
    except Exception:
        return []
    finally:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass


def find_svcl_device(devices: list[dict], name_contains: str,
                     direction: str, dev_type: str = "Device") -> dict | None:
    """Find a device in svcl output by name substring."""
    for d in devices:
        if (name_contains.lower() in d["name"].lower()
                and d["direction"] == direction
                and d["type"] == dev_type):
            return d
    return None


def extract_guid(item_id: str) -> str | None:
    """Extract the device GUID from an svcl Item ID like {0.0.1.00000000}.{guid}."""
    parts = item_id.split("}.{")
    if len(parts) == 2:
        return "{" + parts[1]
    return None


# ---------------------------------------------------------------------------
# "Listen to this device" via registry (requires admin)
# ---------------------------------------------------------------------------
def build_listen_ps_script(b2_guid: str, target_endpoint_id: str) -> str:
    """Build a PowerShell script that enables 'Listen to this device' on B2."""
    return f'''
$ErrorActionPreference = "Stop"
$keyPath = "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\MMDevices\\Audio\\Capture\\{b2_guid}\\Properties"

if (-not (Test-Path $keyPath)) {{
    Write-Error "Registry path not found: $keyPath"
    exit 1
}}

# Property 0: target render device endpoint ID
Set-ItemProperty -Path $keyPath -Name "{LISTEN_PROP_GUID},0" -Value "{target_endpoint_id}" -Type String

# Property 1: listen enabled flags
$bytes1 = [byte[]]@({LISTEN_BYTES_1})
Set-ItemProperty -Path $keyPath -Name "{LISTEN_PROP_GUID},1" -Value $bytes1 -Type Binary

# Property 2: additional flags
$bytes2 = [byte[]]@({LISTEN_BYTES_2})
Set-ItemProperty -Path $keyPath -Name "{LISTEN_PROP_GUID},2" -Value $bytes2 -Type Binary

Write-Host "Listen to device configured successfully"
exit 0
'''


def run_elevated_ps(script: str) -> tuple[bool, str]:
    """Run a PowerShell script with admin elevation (UAC prompt)."""
    script_path = SCRIPT_DIR / "_listen_setup.ps1"
    result_path = SCRIPT_DIR / "_listen_result.txt"
    try:
        # Write script that captures output
        full_script = script + f'\n"OK" | Out-File -FilePath "{result_path}" -Encoding utf8\n'
        script_path.write_text(full_script, encoding="utf-8")
        result_path.unlink(missing_ok=True)

        # Run elevated
        ret = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", "powershell.exe",
            f'-ExecutionPolicy Bypass -WindowStyle Hidden -File "{script_path}"',
            None, 0,  # SW_HIDE
        )
        if ret <= 32:
            return False, "UAC prompt was declined or elevation failed"

        # Wait for result (up to 10 seconds)
        for _ in range(20):
            time.sleep(0.5)
            if result_path.exists():
                return True, "OK"

        return False, "Timed out waiting for elevated script"
    except Exception as e:
        return False, str(e)
    finally:
        try:
            script_path.unlink(missing_ok=True)
        except OSError:
            pass
        try:
            result_path.unlink(missing_ok=True)
        except OSError:
            pass


# ---------------------------------------------------------------------------
# Shortcut creation
# ---------------------------------------------------------------------------
def create_shortcut(target: str, shortcut_path: str, args: str = "",
                    description: str = ""):
    """Create a Windows .lnk shortcut via PowerShell."""
    ps_script = SCRIPT_DIR / "_mkshortcut.ps1"
    ps_script.write_text(f'''
$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut("{shortcut_path}")
$sc.TargetPath = "{target}"
$sc.Arguments = '{args}'
$sc.WorkingDirectory = "{SCRIPT_DIR}"
$sc.Description = "{description}"
$sc.Save()
''', encoding="utf-8")
    subprocess.run(
        ["powershell", "-ExecutionPolicy", "Bypass", "-File", str(ps_script)],
        capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW,
    )
    ps_script.unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# Setup Wizard UI
# ---------------------------------------------------------------------------
class SetupWizard:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("VR Audio Switcher \u2014 Setup")
        self.root.resizable(False, True)
        self.root.minsize(0, 600)

        self.bg = "#1e1e1e"
        self.fg = "#e0e0e0"
        self.accent = "#4caf50"
        self.err = "#f44336"
        self.ok_fg = "#66bb6a"
        self.btn_bg = "#333333"
        self.btn_act = "#444444"
        self.desc_fg = "#888888"
        self.hint_bg = "#252525"
        self.root.configure(bg=self.bg)

        self._vm_inputs: list[dict] = []
        self._svcl_devices: list[dict] = []
        self._vm_launched_by_us = False

        self._build_ui()
        self._center()
        self._check_prerequisites()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        main = tk.Frame(self.root, bg=self.bg, padx=24, pady=16)
        main.pack(fill="both", expand=True)

        # Title
        tk.Label(main, text="VR Audio Switcher", bg=self.bg, fg=self.fg,
                 font=("Segoe UI", 16, "bold")).pack(anchor="w")
        tk.Label(main, text="One-time setup \u2014 takes about a minute",
                 bg=self.bg, fg=self.desc_fg,
                 font=("Segoe UI", 10)).pack(anchor="w", pady=(0, 14))

        # --- Step 1: Prerequisites ---
        self._sec(main, "STEP 1 \u2014 Prerequisites")
        self.chk_frame = tk.Frame(main, bg=self.bg)
        self.chk_frame.pack(fill="x", pady=(0, 4))

        self.chk_labels = {}
        self.chk_btns = {}
        for key, label, btn_text, btn_cmd in [
            ("python", "Python 3.10+", None, None),
            ("voicemeeter", "VoiceMeeter Banana", "Download",
             lambda: os.startfile("https://vb-audio.com/Voicemeeter/banana.htm")),
            ("svcl", "svcl.exe (NirSoft audio tool)", "Download",
             self._download_svcl),
            ("packages", "Python packages (psutil, pystray, Pillow)",
             "Install", self._install_deps),
        ]:
            row = tk.Frame(self.chk_frame, bg=self.bg)
            row.pack(fill="x", pady=1)
            status = tk.Label(row, text="\u2022", bg=self.bg, fg=self.desc_fg,
                              font=("Segoe UI", 10), width=2)
            status.pack(side="left")
            tk.Label(row, text=label, bg=self.bg, fg=self.fg,
                     font=("Segoe UI", 10)).pack(side="left")
            self.chk_labels[key] = status
            if btn_text:
                btn = tk.Button(
                    row, text=btn_text, bg=self.btn_bg, fg=self.fg,
                    activebackground=self.btn_act, activeforeground=self.fg,
                    relief="flat", padx=8, pady=1, font=("Segoe UI", 8),
                    cursor="hand2", command=btn_cmd,
                )
                btn.pack(side="right", padx=(4, 0))
                self.chk_btns[key] = btn

        # --- Before you begin ---
        self._sec(main, "BEFORE YOU BEGIN")
        ready_frame = tk.Frame(main, bg=self.hint_bg, padx=12, pady=8)
        ready_frame.pack(fill="x", pady=(0, 4))
        for instruction in [
            "\u2022 Turn on your VR headset and make sure it's connected",
            "\u2022 Make sure your speakers or headphones are plugged in",
            "\u2022 Have VRChat installed (you'll need it for one setting)",
        ]:
            tk.Label(ready_frame, text=instruction, bg=self.hint_bg,
                     fg=self.fg, font=("Segoe UI", 9),
                     anchor="w").pack(anchor="w", pady=1)

        # --- Step 2: Devices ---
        self._sec(main, "STEP 2 \u2014 Select Your Devices")
        dev_frame = tk.Frame(main, bg=self.bg)
        dev_frame.pack(fill="x", pady=(0, 4))

        # Music browser
        tk.Label(dev_frame, text="Music browser (plays Spotify, YouTube, etc.):",
                 bg=self.bg, fg=self.fg,
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 0))
        self.browser_var = tk.StringVar()
        browser_frame = tk.Frame(dev_frame, bg=self.bg)
        browser_frame.pack(fill="x", pady=(1, 2))
        for bname, bexe in [
            ("Chrome", "chrome.exe"), ("Firefox", "firefox.exe"),
            ("Edge", "msedge.exe"), ("Brave", "brave.exe"),
        ]:
            tk.Radiobutton(
                browser_frame, text=bname, variable=self.browser_var,
                value=bexe, bg=self.bg, fg=self.fg,
                selectcolor="#333333", activebackground=self.bg,
                activeforeground=self.fg, font=("Segoe UI", 9),
            ).pack(side="left", padx=(0, 12))
        self.browser_var.set("chrome.exe")  # default

        for label_text, var_name, combo_name in [
            ("Microphone (your physical mic):", "mic_var", "mic_combo"),
            ("VR headset audio (where you hear music in VR):",
             "vr_var", "vr_combo"),
        ]:
            tk.Label(dev_frame, text=label_text, bg=self.bg, fg=self.fg,
                     font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 0))
            var = tk.StringVar()
            setattr(self, var_name, var)
            combo = ttk.Combobox(dev_frame, textvariable=var,
                                 state="readonly", width=55)
            combo.pack(fill="x", pady=(1, 2))
            setattr(self, combo_name, combo)

        # Desktop speakers note
        tk.Label(dev_frame,
                 text="Desktop mode automatically uses your Windows default "
                      "speakers \u2014 no selection needed.",
                 bg=self.bg, fg=self.desc_fg,
                 font=("Segoe UI", 8)).pack(anchor="w", pady=(4, 0))

        self.refresh_btn = tk.Button(
            dev_frame, text="\u21bb Refresh Devices", bg=self.btn_bg,
            fg=self.fg, activebackground=self.btn_act,
            activeforeground=self.fg, relief="flat", padx=8, pady=3,
            font=("Segoe UI", 8), cursor="hand2",
            command=self._detect_devices,
        )
        self.refresh_btn.pack(anchor="w", pady=(4, 0))

        # --- Step 3: VRChat reminder ---
        self._sec(main, "STEP 3 \u2014 VRChat Mic (one-time, in-game)")
        hint = tk.Frame(main, bg=self.hint_bg, padx=12, pady=8)
        hint.pack(fill="x", pady=(0, 4))
        tk.Label(hint, text="After setup, open VRChat and go to:",
                 bg=self.hint_bg, fg=self.desc_fg,
                 font=("Segoe UI", 9)).pack(anchor="w")
        tk.Label(hint,
                 text="Settings \u2192 Audio \u2192 Microphone \u2192 "
                      "\"Voicemeeter Out B1\"",
                 bg=self.hint_bg, fg=self.fg,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(2, 0))
        tk.Label(hint, text="This tells VRChat to use VoiceMeeter as your mic "
                            "(voice + optional music).",
                 bg=self.hint_bg, fg=self.desc_fg,
                 font=("Segoe UI", 8)).pack(anchor="w", pady=(2, 0))

        # --- Status Log ---
        self.log_text = tk.Text(
            main, height=5, bg="#111111", fg=self.desc_fg,
            font=("Consolas", 9), relief="flat", wrap="word",
            insertbackground=self.fg,
        )
        self.log_text.pack(fill="x", pady=(10, 8))
        self.log_text.config(state="disabled")

        # --- Bottom buttons ---
        bottom = tk.Frame(main, bg=self.bg)
        bottom.pack(fill="x")

        self.setup_btn = tk.Button(
            bottom, text="\u2714  Set Up Everything",
            bg=self.accent, fg="#000", activebackground="#66bb6a",
            activeforeground="#000", relief="flat", padx=20, pady=10,
            font=("Segoe UI", 11, "bold"), cursor="hand2",
            command=self._setup_everything,
        )
        self.setup_btn.pack(side="left")

        self.launch_btn = tk.Button(
            bottom, text="\u25b6  Launch", bg=self.btn_bg, fg=self.fg,
            activebackground=self.btn_act, activeforeground=self.fg,
            relief="flat", padx=20, pady=10, font=("Segoe UI", 11),
            cursor="hand2", command=self._launch, state="disabled",
        )
        self.launch_btn.pack(side="right")

    def _sec(self, parent, text):
        tk.Label(parent, text=text, bg=self.bg, fg=self.accent,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(12, 2))

    def _center(self):
        self.root.update_idletasks()
        w, h = self.root.winfo_width(), self.root.winfo_height()
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        self.root.geometry(f"+{x}+{y}")

    def _log(self, msg):
        self.log_text.config(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")
        self.root.update_idletasks()

    def _set_check(self, key, ok):
        lbl = self.chk_labels[key]
        if ok:
            lbl.config(text="\u2713", fg=self.ok_fg)
            if key in self.chk_btns:
                self.chk_btns[key].config(state="disabled")
        else:
            lbl.config(text="\u2717", fg=self.err)

    # ------------------------------------------------------------------
    # Prerequisites
    # ------------------------------------------------------------------
    def _check_prerequisites(self):
        py_ok = sys.version_info >= (3, 10)
        self._set_check("python", py_ok)
        if py_ok:
            self._log(f"Python {sys.version.split()[0]}")

        vm_ok = Path(VM_DLL).exists()
        self._set_check("voicemeeter", vm_ok)
        self._log("VoiceMeeter Banana " + ("found" if vm_ok else "NOT FOUND"))

        svcl_ok = SVCL_PATH.exists()
        self._set_check("svcl", svcl_ok)
        if not svcl_ok:
            self._log("svcl.exe missing \u2014 click Download")

        pkg_ok = all(self._check_pkg(p) for p in REQUIRED_PACKAGES)
        self._set_check("packages", pkg_ok)
        if not pkg_ok:
            self._log("Some Python packages missing \u2014 click Install")

        if vm_ok:
            self._detect_devices()

    @staticmethod
    def _check_pkg(name):
        try:
            __import__(name)
            return True
        except ImportError:
            return False

    # ------------------------------------------------------------------
    # Fix actions
    # ------------------------------------------------------------------
    def _install_deps(self):
        self._log("Installing packages...")
        btn = self.chk_btns.get("packages")
        if btn:
            btn.config(state="disabled")

        def run():
            try:
                result = subprocess.run(
                    [sys.executable, "-m", "pip", "install", "-r",
                     str(SCRIPT_DIR / "requirements.txt")],
                    capture_output=True, text=True, timeout=120,
                )
                ok = result.returncode == 0
                self.root.after(0, lambda: self._set_check("packages", ok))
                self.root.after(0, lambda: self._log(
                    "Packages installed" if ok
                    else f"pip error: {result.stderr[:200]}"))
                if not ok and btn:
                    self.root.after(0, lambda: btn.config(state="normal"))
            except Exception as e:
                self.root.after(0, lambda: self._log(f"Install failed: {e}"))
                if btn:
                    self.root.after(0, lambda: btn.config(state="normal"))

        threading.Thread(target=run, daemon=True).start()

    def _download_svcl(self):
        self._log("Downloading svcl.exe...")
        btn = self.chk_btns.get("svcl")
        if btn:
            btn.config(state="disabled")

        def run():
            try:
                import urllib.request
                import zipfile

                zip_path = SCRIPT_DIR / "_svcl.zip"
                urllib.request.urlretrieve(SVCL_URL, str(zip_path))

                with zipfile.ZipFile(str(zip_path), "r") as zf:
                    for name in zf.namelist():
                        if name.lower() == "svcl.exe":
                            with zf.open(name) as src, \
                                 open(str(SVCL_PATH), "wb") as dst:
                                dst.write(src.read())
                            break

                zip_path.unlink(missing_ok=True)
                ok = SVCL_PATH.exists()
                self.root.after(0, lambda: self._set_check("svcl", ok))
                self.root.after(0, lambda: self._log(
                    "svcl.exe downloaded" if ok else "svcl.exe not found in zip"))
                if not ok and btn:
                    self.root.after(0, lambda: btn.config(state="normal"))
            except Exception as e:
                self.root.after(0, lambda: self._log(f"Download failed: {e}"))
                if btn:
                    self.root.after(0, lambda: btn.config(state="normal"))

        threading.Thread(target=run, daemon=True).start()

    # ------------------------------------------------------------------
    # Device detection
    # ------------------------------------------------------------------
    def _detect_devices(self):
        self._log("Detecting audio devices...")
        try:
            self._ensure_voicemeeter()
            vm = VMDeviceEnumerator()
            self._vm_inputs = vm.input_devices()
            vm_outputs = vm.output_devices()
            vm.close()

            # Mic dropdown (VoiceMeeter WDM input devices)
            wdm_inputs = [d for d in self._vm_inputs if d["type"] == "wdm"]
            mic_names = [d["name"] for d in wdm_inputs]
            self.mic_combo["values"] = mic_names
            if mic_names:
                # Auto-select a physical mic (skip Steam virtual devices)
                sel = 0
                for i, name in enumerate(mic_names):
                    nl = name.lower()
                    if ("microphone" in nl or "mic" in nl) \
                            and "steam" not in nl:
                        sel = i
                        break
                self.mic_combo.current(sel)

            # VR headset dropdown — physical output devices (exclude VoiceMeeter)
            wdm_outputs = [d for d in vm_outputs if d["type"] == "wdm"]
            vr_names = [d["name"] for d in wdm_outputs
                        if "voicemeeter" not in d["name"].lower()]
            self.vr_combo["values"] = vr_names
            if vr_names:
                # Auto-select Steam Streaming Speakers (not Microphone)
                sel = 0
                for i, name in enumerate(vr_names):
                    if "steam streaming speakers" in name.lower():
                        sel = i
                        break
                self.vr_combo.current(sel)

            # Also query svcl for endpoint IDs
            if SVCL_PATH.exists():
                self._svcl_devices = query_svcl_devices()

            self._log(f"Found {len(mic_names)} mics, {len(vr_names)} outputs")

        except Exception as e:
            self._log(f"Device detection failed: {e}")

    def _ensure_voicemeeter(self):
        """Launch VoiceMeeter if not already running."""
        try:
            import psutil
            for p in psutil.process_iter(["name"]):
                if p.info["name"] and p.info["name"].lower() == "voicemeeterpro.exe":
                    return
        except ImportError:
            pass
        if Path(VM_EXE).exists():
            self._log("Launching VoiceMeeter Banana...")
            subprocess.Popen([VM_EXE],
                             creationflags=subprocess.CREATE_NO_WINDOW)
            self._vm_launched_by_us = True
            time.sleep(4)

    # ------------------------------------------------------------------
    # Main setup action
    # ------------------------------------------------------------------
    def _setup_everything(self):
        self.setup_btn.config(state="disabled")
        self._log("")
        threading.Thread(target=self._setup_thread, daemon=True).start()

    def _ui(self, fn):
        """Schedule fn on the main thread."""
        self.root.after(0, fn)

    def _setup_thread(self):
        """Background thread: auto-install deps, detect devices, configure."""
        def log(msg):
            self._ui(lambda: self._log(msg))
        def set_check(key, ok):
            self._ui(lambda: self._set_check(key, ok))

        # ---- Phase 1: Auto-install missing dependencies ----

        # VoiceMeeter Banana
        if not Path(VM_DLL).exists():
            log("VoiceMeeter Banana not found \u2014 downloading installer...")
            try:
                import urllib.request
                installer = SCRIPT_DIR / "_VoicemeeterProSetup.exe"
                urllib.request.urlretrieve(
                    "https://download.vb-audio.com/Download_CABLE/"
                    "VoicemeeterProSetup.exe", str(installer))
                log("Launching VoiceMeeter installer...")
                subprocess.Popen([str(installer)])
                log("")
                log("Install VoiceMeeter, REBOOT your PC, then run "
                    "this wizard again.")
                self._ui(lambda: messagebox.showinfo(
                    "VoiceMeeter Required",
                    "VoiceMeeter Banana installer has been downloaded "
                    "and launched.\n\n"
                    "1. Complete the VoiceMeeter installation\n"
                    "2. REBOOT your PC\n"
                    "3. Run this setup wizard again\n\n"
                    "The wizard will continue from where you left off."))
            except Exception as e:
                log(f"Download failed: {e}")
                log("Install VoiceMeeter manually: "
                    "https://vb-audio.com/Voicemeeter/banana.htm")
            self._ui(lambda: self.setup_btn.config(state="normal"))
            return

        # svcl.exe
        if not SVCL_PATH.exists():
            log("Downloading svcl.exe (NirSoft audio tool)...")
            try:
                import urllib.request
                import zipfile

                zip_path = SCRIPT_DIR / "_svcl.zip"
                urllib.request.urlretrieve(SVCL_URL, str(zip_path))
                with zipfile.ZipFile(str(zip_path), "r") as zf:
                    for name in zf.namelist():
                        if name.lower() == "svcl.exe":
                            with zf.open(name) as src, \
                                 open(str(SVCL_PATH), "wb") as dst:
                                dst.write(src.read())
                            break
                zip_path.unlink(missing_ok=True)
                ok = SVCL_PATH.exists()
                set_check("svcl", ok)
                log("svcl.exe \u2713" if ok else "svcl download failed")
                if not ok:
                    self._ui(lambda: self.setup_btn.config(state="normal"))
                    return
            except Exception as e:
                log(f"svcl download failed: {e}")
                self._ui(lambda: self.setup_btn.config(state="normal"))
                return

        # Python packages
        if not all(self._check_pkg(p) for p in REQUIRED_PACKAGES):
            log("Installing Python packages...")
            try:
                result = subprocess.run(
                    [sys.executable, "-m", "pip", "install", "-r",
                     str(SCRIPT_DIR / "requirements.txt")],
                    capture_output=True, text=True, timeout=120,
                )
                ok = result.returncode == 0
                set_check("packages", ok)
                log("Packages installed \u2713" if ok
                    else f"pip error: {result.stderr[:200]}")
                if not ok:
                    self._ui(lambda: self.setup_btn.config(state="normal"))
                    return
            except Exception as e:
                log(f"Package install failed: {e}")
                self._ui(lambda: self.setup_btn.config(state="normal"))
                return

        # ---- Phase 2: Detect devices if not already done ----
        mic_name = self.mic_var.get()
        vr_name = self.vr_var.get()
        browser = self.browser_var.get()

        if not mic_name or not vr_name:
            log("Detecting audio devices...")
            detect_done = threading.Event()
            self._ui(lambda: (self._detect_devices(), detect_done.set()))
            detect_done.wait(timeout=15)
            time.sleep(0.5)
            mic_name = self.mic_var.get()
            vr_name = self.vr_var.get()

        if not mic_name:
            log("No microphone detected. Please select one and try again.")
            self._ui(lambda: self.setup_btn.config(state="normal"))
            return
        if not vr_name:
            log("No VR headset output detected. Please select one and "
                "try again.")
            self._ui(lambda: self.setup_btn.config(state="normal"))
            return

        # ---- Phase 3: Configure everything ----
        log("Setting up...")
        errors = []

        # 1. Find VoiceMeeter VAIO svcl ID
        vaio_id = None
        if SVCL_PATH.exists():
            if not self._svcl_devices:
                self._svcl_devices = query_svcl_devices()
            d = find_svcl_device(self._svcl_devices, "Voicemeeter Input",
                                 "Render")
            if d:
                vaio_id = d["friendly_id"]
        if not vaio_id:
            vaio_id = r"VB-Audio Voicemeeter VAIO\Device\Voicemeeter Input\Render"
        log(f"VoiceMeeter VAIO: {vaio_id[:50]}...")

        # 2. config.json
        config = {
            "poll_interval_seconds": 3,
            "steamvr_process": "vrserver.exe",
            "target_process": browser,
            "svcl_path": "svcl.exe",
            "vr_device": vaio_id,
            "debounce_seconds": 5,
            "music_strip": 3,
        }
        try:
            with open(CONFIG_PATH, "w") as f:
                json.dump(config, f, indent=2)
            log("config.json \u2713")
        except Exception as e:
            errors.append(f"config.json: {e}")

        # 3. vm_devices.json
        try:
            with open(VM_DEVICES_PATH, "w") as f:
                json.dump({"Strip[0]": mic_name}, f, indent=2)
            log(f"vm_devices.json \u2713 (mic: {mic_name})")
        except Exception as e:
            errors.append(f"vm_devices.json: {e}")

        # 4. "Listen to this device" on Voicemeeter Out B2 → VR headset
        listen_ok = self._configure_listen(vr_name)
        if listen_ok:
            log("Listen to this device \u2713 (B2 \u2192 "
                f"{vr_name})")
        else:
            log("Listen to this device: manual step needed (see below)")

        # 5. Shortcuts
        pythonw = Path(sys.executable).parent / "pythonw.exe"
        if not pythonw.exists():
            pythonw = Path(sys.executable)
        script = str(SCRIPT_DIR / "vr_audio_switcher.py")

        try:
            desktop = Path(os.environ["USERPROFILE"]) / "OneDrive" / "Desktop"
            if not desktop.exists():
                desktop = Path(os.environ["USERPROFILE"]) / "Desktop"
            create_shortcut(str(pythonw),
                            str(desktop / "VR Audio Switcher.lnk"),
                            args=f'"{script}"',
                            description="VR Audio Switcher")
            log("Desktop shortcut \u2713")
        except Exception as e:
            errors.append(f"Desktop shortcut: {e}")

        try:
            startup = (Path(os.environ["APPDATA"])
                       / "Microsoft" / "Windows" / "Start Menu"
                       / "Programs" / "Startup")
            create_shortcut(str(pythonw),
                            str(startup / "VR Audio Switcher.lnk"),
                            args=f'"{script}"',
                            description="VR Audio Switcher (auto-start)")
            log("Startup shortcut \u2713 (auto-launches on boot)")
        except Exception as e:
            errors.append(f"Startup shortcut: {e}")

        # 6. Shut down VoiceMeeter so tray app gets a fresh start
        if self._vm_launched_by_us:
            self._ui(lambda: self._shutdown_voicemeeter())

        if errors:
            for err in errors:
                log(f"Warning: {err}")

        if not listen_ok:
            log("")
            self._ui(lambda: self._show_manual_listen(vr_name))
        else:
            log("")
            log("Setup complete! Click Launch to start.")
            self._ui(lambda: self.launch_btn.config(
                state="normal", bg=self.accent, fg="#000"))

        self._ui(lambda: self.setup_btn.config(state="normal"))

    def _configure_listen(self, vr_output_name: str) -> bool:
        """Try to enable 'Listen to this device' on B2 via registry."""
        if not self._svcl_devices:
            if SVCL_PATH.exists():
                self._svcl_devices = query_svcl_devices()
            else:
                return False

        # Find Voicemeeter Out B2 capture device
        b2_dev = find_svcl_device(self._svcl_devices, "Voicemeeter Out B2",
                                  "Capture")
        if not b2_dev:
            self._log("Could not find Voicemeeter Out B2 in device list")
            return False

        b2_guid = extract_guid(b2_dev["item_id"])
        if not b2_guid:
            self._log(f"Could not extract B2 GUID from: {b2_dev['item_id']}")
            return False

        # Find the target VR render device — match by endpoint name
        # svcl "Name" for device entries is the endpoint name (e.g., "Speakers")
        # We need to match against VoiceMeeter's output device name format:
        # "Speakers (Steam Streaming Speakers)"
        # In svcl, the Name is "Speakers" and Device Name is "Steam Streaming Speakers"
        target_dev = None
        for d in self._svcl_devices:
            if (d["direction"] == "Render" and d["type"] == "Device"
                    and d["state"] == "Active"):
                # Match: "Name (Device Name)" should equal vr_output_name
                svcl_full = f"{d['name']} ({d['device_name']})"
                if svcl_full == vr_output_name:
                    target_dev = d
                    break
                # Also try matching just the device_name in the VR name
                if d["device_name"] and d["device_name"] in vr_output_name:
                    target_dev = d
                    break
        if not target_dev:
            self._log(f"Could not find VR device '{vr_output_name}' in svcl")
            return False

        target_endpoint_id = target_dev["item_id"]
        if not target_endpoint_id:
            return False

        # Build and run the PowerShell script with elevation
        ps = build_listen_ps_script(b2_guid, target_endpoint_id)
        self._log("Requesting admin permission for audio config...")
        ok, msg = run_elevated_ps(ps)
        return ok

    def _show_manual_listen(self, vr_name: str):
        """Show manual instructions if automatic config failed."""
        self._log("MANUAL STEP NEEDED:")
        self._log("1. Open Windows Sound Settings (Recording tab)")
        self._log("2. Right-click 'Voicemeeter Out B2' \u2192 Properties")
        self._log("3. Listen tab \u2192 Check 'Listen to this device'")
        self._log(f"4. Set playback to: {vr_name}")
        self._log("5. Click OK")
        self._log("")
        self._log("Then click Launch below.")
        self.launch_btn.config(state="normal", bg=self.accent, fg="#000")

        # Also offer to open Sound Settings
        messagebox.showinfo(
            "One Manual Step",
            "Automatic audio routing requires admin permission.\n\n"
            "Please do this one-time step:\n\n"
            "1. Open Windows Sound Settings \u2192 Recording tab\n"
            "2. Right-click 'Voicemeeter Out B2'\n"
            "3. Properties \u2192 Listen tab\n"
            "4. Check 'Listen to this device'\n"
            f"5. Set playback to: {vr_name}\n"
            "6. Click OK\n\n"
            "Click OK here, then click Launch in the wizard.",
        )
        # Open Recording tab
        subprocess.Popen(
            ["rundll32.exe", "Shell32.dll,Control_RunDLL", "mmsys.cpl,,1"],
            creationflags=subprocess.CREATE_NO_WINDOW,
        )

    def _shutdown_voicemeeter(self):
        """Shut down VoiceMeeter so the tray app can start it fresh."""
        try:
            import psutil
            for proc in psutil.process_iter(["name"]):
                if (proc.info["name"]
                        and proc.info["name"].lower() == "voicemeeterpro.exe"):
                    proc.kill()
                    self._log("VoiceMeeter shut down (tray app will restart it)")
                    break
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Launch
    # ------------------------------------------------------------------
    def _launch(self):
        pythonw = Path(sys.executable).parent / "pythonw.exe"
        if not pythonw.exists():
            pythonw = Path(sys.executable)
        script = str(SCRIPT_DIR / "vr_audio_switcher.py")
        subprocess.Popen([str(pythonw), script])
        self._log("Launched! Look for the tray icon in your system tray.")
        self.root.after(2000, self.root.destroy)

    def run(self):
        self.root.mainloop()


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    SetupWizard().run()
