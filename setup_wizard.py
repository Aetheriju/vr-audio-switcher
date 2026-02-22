"""VR Audio Switcher — First-Run Setup Wizard.

Guides new users through device detection, svcl download, config generation,
Windows audio configuration, and shortcut creation.

Run once:  python setup_wizard.py
"""

import ctypes
import csv
import json
import logging
import logging.handlers
import os
import platform
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
WIZARD_LOG_PATH = SCRIPT_DIR / "wizard.log"
VM_DOWNLOAD_URL = ("https://download.vb-audio.com/Download_CABLE/"
                   "VoicemeeterSetup_v2122.zip")
RESUME_SHORTCUT = (Path(os.environ.get("APPDATA", ""))
                   / "Microsoft" / "Windows" / "Start Menu" / "Programs"
                   / "Startup" / "VR Audio Switcher Setup (resume).lnk")

from vm_path import find_dll, find_exe, is_vm_process

REQUIRED_PACKAGES = ["psutil"]

# Registry property keys for "Listen to this device"
LISTEN_PROP_GUID = "{24dbb0fc-9311-4b3d-9cf0-18ff155639d4}"
LISTEN_BYTES_1 = "0x0B,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0xFF,0xFF,0x00,0x00"
LISTEN_BYTES_2 = "0x0B,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00"


# ---------------------------------------------------------------------------
# VoiceMeeter device enumeration
# ---------------------------------------------------------------------------
class VMDeviceEnumerator:
    """Connect to VoiceMeeter and enumerate hardware devices."""

    TYPE_MAP = {1: "mme", 3: "wdm", 5: "ks"}

    def __init__(self):
        dll_path = find_dll()
        if not dll_path:
            raise RuntimeError("VoiceMeeter DLL not found")
        self._dll = ctypes.WinDLL(str(dll_path))
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
             "/Columns", "Name,Command-Line Friendly ID,Item ID,"
                         "Direction,Type,Device State,Device Name"],
            capture_output=True, timeout=10,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        devices = []
        with open(tmp, newline="", encoding="utf-8-sig",
                  errors="replace") as f:
            for row in csv.DictReader(f):
                devices.append({
                    "name": row.get("Name", "").strip(),
                    "friendly_id": row.get("Command-Line Friendly ID",
                                           "").strip(),
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
    for d in devices:
        if (name_contains.lower() in d["name"].lower()
                and d["direction"] == direction
                and d["type"] == dev_type):
            return d
    return None


def extract_guid(item_id: str) -> str | None:
    parts = item_id.split("}.{")
    if len(parts) == 2:
        return "{" + parts[1]
    return None


# ---------------------------------------------------------------------------
# "Listen to this device" via registry (requires admin)
# ---------------------------------------------------------------------------
def build_listen_ps_script(b2_guid: str, target_endpoint_id: str) -> str:
    return f'''
$ErrorActionPreference = "Stop"
$keyPath = "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\MMDevices\\Audio\\Capture\\{b2_guid}\\Properties"

if (-not (Test-Path $keyPath)) {{
    Write-Error "Registry path not found: $keyPath"
    exit 1
}}

Set-ItemProperty -Path $keyPath -Name "{LISTEN_PROP_GUID},0" -Value "{target_endpoint_id}" -Type String

$bytes1 = [byte[]]@({LISTEN_BYTES_1})
Set-ItemProperty -Path $keyPath -Name "{LISTEN_PROP_GUID},1" -Value $bytes1 -Type Binary

$bytes2 = [byte[]]@({LISTEN_BYTES_2})
Set-ItemProperty -Path $keyPath -Name "{LISTEN_PROP_GUID},2" -Value $bytes2 -Type Binary

Write-Host "Listen to device configured successfully"
exit 0
'''


def run_elevated_ps(script: str) -> tuple[bool, str]:
    script_path = SCRIPT_DIR / "_listen_setup.ps1"
    result_path = SCRIPT_DIR / "_listen_result.txt"
    try:
        full_script = (script
                       + f'\n"OK" | Out-File -FilePath "{result_path}"'
                         ' -Encoding utf8\n')
        script_path.write_text(full_script, encoding="utf-8")
        result_path.unlink(missing_ok=True)

        ctypes.windll.shell32.ShellExecuteW.restype = ctypes.c_void_p
        ret = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", "powershell.exe",
            f'-ExecutionPolicy Bypass -WindowStyle Hidden '
            f'-File "{script_path}"',
            None, 0,
        )
        if (ret or 0) <= 32:
            return False, "UAC prompt was declined or elevation failed"

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
        ["powershell", "-ExecutionPolicy", "Bypass", "-File",
         str(ps_script)],
        capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW,
    )
    ps_script.unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# Setup Wizard UI (phased)
# ---------------------------------------------------------------------------
class SetupWizard:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("VR Audio Switcher \u2014 Setup")
        self.root.resizable(False, True)
        self.root.minsize(480, 400)

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
        self._resuming = False

        self._setup_file_logging()
        self._cleanup_resume_shortcut()
        self._build_ui()
        self._center()

        # If all prerequisites are met, skip to device config phase
        if self._all_prerequisites_met():
            self._show_phase("installing")
            threading.Thread(target=self._install_thread,
                             daemon=True).start()
        else:
            self._show_phase("start")

    # ------------------------------------------------------------------
    # Logging
    # ------------------------------------------------------------------
    def _setup_file_logging(self):
        self._file_log = logging.getLogger("setup_wizard")
        self._file_log.setLevel(logging.DEBUG)
        handler = logging.handlers.RotatingFileHandler(
            WIZARD_LOG_PATH, maxBytes=500_000, backupCount=1,
            encoding="utf-8")
        handler.setFormatter(logging.Formatter(
            "%(asctime)s [%(levelname)s] %(message)s"))
        self._file_log.addHandler(handler)
        self._file_log.info("=== Setup Wizard Started ===")
        self._file_log.info("Python %s on %s %s",
                            sys.version.split()[0], platform.system(),
                            platform.version())
        self._file_log.info("Script dir: %s", SCRIPT_DIR)

    def _log(self, msg):
        self.log_text.config(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")
        self.root.update_idletasks()
        if msg.strip():
            self._file_log.info(msg)

    def _ui(self, fn):
        self.root.after(0, fn)

    # ------------------------------------------------------------------
    # Resume shortcut cleanup
    # ------------------------------------------------------------------
    def _cleanup_resume_shortcut(self):
        try:
            if RESUME_SHORTCUT.exists():
                RESUME_SHORTCUT.unlink()
                self._resuming = True
                self._file_log.info("Resumed after reboot (deleted resume "
                                    "shortcut)")
        except OSError:
            pass

    # ------------------------------------------------------------------
    # Prerequisites check
    # ------------------------------------------------------------------
    def _all_prerequisites_met(self) -> bool:
        return (sys.version_info >= (3, 10)
                and find_dll() is not None
                and SVCL_PATH.exists()
                and all(self._check_pkg(p) for p in REQUIRED_PACKAGES))

    @staticmethod
    def _check_pkg(name):
        try:
            __import__(name)
            return True
        except ImportError:
            return False

    # ------------------------------------------------------------------
    # UI: Build all phases
    # ------------------------------------------------------------------
    def _build_ui(self):
        self._main = tk.Frame(self.root, bg=self.bg, padx=24, pady=16)
        self._main.pack(fill="both", expand=True)

        # Title (always visible)
        tk.Label(self._main, text="VR Audio Switcher", bg=self.bg,
                 fg=self.fg,
                 font=("Segoe UI", 16, "bold")).pack(anchor="w")
        self._subtitle = tk.Label(
            self._main, text="One-time setup", bg=self.bg,
            fg=self.desc_fg, font=("Segoe UI", 10))
        self._subtitle.pack(anchor="w", pady=(0, 14))

        # Phase frames (only one visible at a time)
        self._phase_start = self._build_phase_start()
        self._phase_installing = self._build_phase_installing()
        self._phase_configure = self._build_phase_configure()
        self._phase_done = self._build_phase_done()
        self._phase_reboot = self._build_phase_reboot()

        # Log (always visible at bottom)
        self.log_text = tk.Text(
            self._main, height=5, bg="#111111", fg=self.desc_fg,
            font=("Consolas", 9), relief="flat", wrap="word",
            insertbackground=self.fg,
        )
        self.log_text.pack(fill="x", pady=(10, 0), side="bottom")
        self.log_text.config(state="disabled")

    def _build_phase_start(self) -> tk.Frame:
        """Phase 1: Get Started screen."""
        f = tk.Frame(self._main, bg=self.bg)

        tk.Label(f, text="BEFORE YOU BEGIN",
                 bg=self.bg, fg=self.accent,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w",
                                                     pady=(8, 4))

        hint = tk.Frame(f, bg=self.hint_bg, padx=12, pady=8)
        hint.pack(fill="x", pady=(0, 16))
        for text in [
            "\u2022 Start SteamVR (or your VR runtime) so your headset "
            "audio is available",
            "\u2022 Make sure your speakers or headphones are active",
        ]:
            tk.Label(hint, text=text, bg=self.hint_bg, fg=self.fg,
                     font=("Segoe UI", 9), anchor="w",
                     wraplength=420, justify="left").pack(anchor="w",
                                                          pady=1)

        tk.Label(f,
                 text="This will install VoiceMeeter Banana, download "
                      "required tools, detect your audio devices, and "
                      "configure everything automatically.",
                 bg=self.bg, fg=self.desc_fg, font=("Segoe UI", 9),
                 wraplength=440, justify="left").pack(anchor="w",
                                                      pady=(0, 16))

        self._start_btn = tk.Button(
            f, text="\u2714  Set Up Everything",
            bg=self.accent, fg="#000", activebackground="#66bb6a",
            activeforeground="#000", relief="flat", padx=24, pady=12,
            font=("Segoe UI", 12, "bold"), cursor="hand2",
            command=self._on_setup_click,
        )
        self._start_btn.pack(anchor="w")

        return f

    def _build_phase_installing(self) -> tk.Frame:
        """Phase 2: Installing/progress screen."""
        f = tk.Frame(self._main, bg=self.bg)

        self._install_status = tk.Label(
            f, text="Setting up...", bg=self.bg, fg=self.fg,
            font=("Segoe UI", 11), anchor="w")
        self._install_status.pack(anchor="w", pady=(8, 4))

        # Checklist
        self._chk_frame = tk.Frame(f, bg=self.bg)
        self._chk_frame.pack(fill="x", pady=(4, 8))

        self._chk_labels = {}
        for key, label in [
            ("voicemeeter", "VoiceMeeter Banana"),
            ("svcl", "svcl.exe (audio routing tool)"),
            ("packages", "Python packages"),
            ("devices", "Audio device detection"),
        ]:
            row = tk.Frame(self._chk_frame, bg=self.bg)
            row.pack(fill="x", pady=1)
            status = tk.Label(row, text="\u2022", bg=self.bg,
                              fg=self.desc_fg,
                              font=("Segoe UI", 10), width=2)
            status.pack(side="left")
            tk.Label(row, text=label, bg=self.bg, fg=self.fg,
                     font=("Segoe UI", 10)).pack(side="left")
            self._chk_labels[key] = status

        return f

    def _build_phase_configure(self) -> tk.Frame:
        """Phase 3: Device selection."""
        f = tk.Frame(self._main, bg=self.bg)

        # Device dropdowns
        tk.Label(f, text="SELECT YOUR DEVICES",
                 bg=self.bg, fg=self.accent,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w",
                                                     pady=(8, 4))

        tk.Label(f,
                 text="All audio apps (browsers, Spotify, media players) "
                      "are managed automatically. VRChat is excluded.",
                 bg=self.bg, fg=self.ok_fg, wraplength=440,
                 justify="left",
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 8))

        for label_text, var_name, combo_name in [
            ("Microphone (your physical mic):", "mic_var", "mic_combo"),
            ("VR headset audio (where you hear music in VR):",
             "vr_var", "vr_combo"),
        ]:
            tk.Label(f, text=label_text, bg=self.bg, fg=self.fg,
                     font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 0))
            var = tk.StringVar()
            setattr(self, var_name, var)
            combo = ttk.Combobox(f, textvariable=var,
                                 state="readonly", width=55)
            combo.pack(fill="x", pady=(1, 2))
            setattr(self, combo_name, combo)

        tk.Label(f,
                 text="Desktop mode automatically uses your Windows "
                      "default speakers.",
                 bg=self.bg, fg=self.desc_fg,
                 font=("Segoe UI", 8)).pack(anchor="w", pady=(4, 0))

        self._refresh_btn = tk.Button(
            f, text="\u21bb Refresh Devices", bg=self.btn_bg,
            fg=self.fg, activebackground=self.btn_act,
            activeforeground=self.fg, relief="flat", padx=8, pady=3,
            font=("Segoe UI", 8), cursor="hand2",
            command=self._detect_devices,
        )
        self._refresh_btn.pack(anchor="w", pady=(4, 8))

        # VRChat mic reminder
        tk.Label(f, text="VRCHAT MIC (one-time, in-game)",
                 bg=self.bg, fg=self.accent,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w",
                                                     pady=(8, 4))
        hint = tk.Frame(f, bg=self.hint_bg, padx=12, pady=8)
        hint.pack(fill="x", pady=(0, 12))
        tk.Label(hint, text="After setup, open VRChat and go to:",
                 bg=self.hint_bg, fg=self.desc_fg,
                 font=("Segoe UI", 9)).pack(anchor="w")
        tk.Label(hint,
                 text='Settings \u2192 Audio \u2192 Microphone \u2192 '
                      '"Voicemeeter Out B1"',
                 bg=self.hint_bg, fg=self.fg,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w",
                                                      pady=(2, 0))

        # Finish button
        self._finish_btn = tk.Button(
            f, text="\u2714  Finish Setup",
            bg=self.accent, fg="#000", activebackground="#66bb6a",
            activeforeground="#000", relief="flat", padx=20, pady=10,
            font=("Segoe UI", 11, "bold"), cursor="hand2",
            command=self._on_finish_click,
        )
        self._finish_btn.pack(anchor="w", pady=(4, 0))

        return f

    def _build_phase_done(self) -> tk.Frame:
        """Phase 4: Setup complete."""
        f = tk.Frame(self._main, bg=self.bg)

        tk.Label(f, text="\u2713  Setup complete!",
                 bg=self.bg, fg=self.ok_fg,
                 font=("Segoe UI", 14, "bold")).pack(anchor="w",
                                                      pady=(20, 8))
        tk.Label(f,
                 text="The app will run in the background and activate "
                      "when SteamVR starts.",
                 bg=self.bg, fg=self.desc_fg,
                 font=("Segoe UI", 10)).pack(anchor="w", pady=(0, 16))

        self._launch_btn = tk.Button(
            f, text="\u25b6  Launch",
            bg=self.accent, fg="#000", activebackground="#66bb6a",
            activeforeground="#000", relief="flat", padx=24, pady=12,
            font=("Segoe UI", 12, "bold"), cursor="hand2",
            command=self._launch,
        )
        self._launch_btn.pack(anchor="w")

        return f

    def _build_phase_reboot(self) -> tk.Frame:
        """Reboot countdown screen."""
        f = tk.Frame(self._main, bg=self.bg)

        tk.Label(f, text="VoiceMeeter installed!",
                 bg=self.bg, fg=self.ok_fg,
                 font=("Segoe UI", 14, "bold")).pack(anchor="w",
                                                      pady=(20, 8))
        tk.Label(f,
                 text="Your PC needs to restart for the audio drivers "
                      "to activate. Setup will continue automatically "
                      "after the restart.",
                 bg=self.bg, fg=self.fg, font=("Segoe UI", 10),
                 wraplength=440, justify="left").pack(anchor="w",
                                                      pady=(0, 16))

        self._reboot_label = tk.Label(
            f, text="Restarting in 15 seconds...",
            bg=self.bg, fg=self.desc_fg, font=("Segoe UI", 10))
        self._reboot_label.pack(anchor="w", pady=(0, 12))

        btn_frame = tk.Frame(f, bg=self.bg)
        btn_frame.pack(anchor="w")

        self._reboot_now_btn = tk.Button(
            btn_frame, text="Restart Now",
            bg=self.accent, fg="#000", activebackground="#66bb6a",
            activeforeground="#000", relief="flat", padx=16, pady=8,
            font=("Segoe UI", 10, "bold"), cursor="hand2",
            command=lambda: self._do_reboot(),
        )
        self._reboot_now_btn.pack(side="left", padx=(0, 8))

        self._cancel_reboot_btn = tk.Button(
            btn_frame, text="Cancel",
            bg=self.btn_bg, fg=self.fg,
            activebackground=self.btn_act, activeforeground=self.fg,
            relief="flat", padx=16, pady=8,
            font=("Segoe UI", 10), cursor="hand2",
            command=self._cancel_reboot,
        )
        self._cancel_reboot_btn.pack(side="left")

        return f

    # ------------------------------------------------------------------
    # Phase switching
    # ------------------------------------------------------------------
    def _show_phase(self, phase: str):
        for f in [self._phase_start, self._phase_installing,
                  self._phase_configure, self._phase_done,
                  self._phase_reboot]:
            f.pack_forget()

        subtitles = {
            "start": "One-time setup",
            "installing": "Please wait...",
            "configure": "Almost done!",
            "done": "",
            "reboot": "Restart required",
        }
        self._subtitle.config(text=subtitles.get(phase, ""))

        frame = {
            "start": self._phase_start,
            "installing": self._phase_installing,
            "configure": self._phase_configure,
            "done": self._phase_done,
            "reboot": self._phase_reboot,
        }[phase]
        frame.pack(fill="both", expand=True, before=self.log_text)

    def _center(self):
        self.root.update_idletasks()
        w, h = self.root.winfo_width(), self.root.winfo_height()
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        self.root.geometry(f"+{x}+{y}")

    def _set_check(self, key, ok):
        lbl = self._chk_labels.get(key)
        if not lbl:
            return
        if ok:
            lbl.config(text="\u2713", fg=self.ok_fg)
        else:
            lbl.config(text="\u2717", fg=self.err)

    # ------------------------------------------------------------------
    # Phase 1: Set Up Everything clicked
    # ------------------------------------------------------------------
    def _on_setup_click(self):
        self._start_btn.config(state="disabled")
        self._show_phase("installing")
        threading.Thread(target=self._install_thread,
                         daemon=True).start()

    # ------------------------------------------------------------------
    # Phase 2: Install thread
    # ------------------------------------------------------------------
    def _install_thread(self):
        def log(msg):
            self._ui(lambda: self._log(msg))
        def check(key, ok):
            self._ui(lambda: self._set_check(key, ok))

        # --- VoiceMeeter ---
        if find_dll():
            check("voicemeeter", True)
            log("VoiceMeeter found")
        else:
            log("Downloading VoiceMeeter installer...")
            try:
                import urllib.request, zipfile
                vm_zip = SCRIPT_DIR / "_VoicemeeterSetup.zip"
                urllib.request.urlretrieve(VM_DOWNLOAD_URL, str(vm_zip))
                log("Extracting installer...")
                with zipfile.ZipFile(str(vm_zip), 'r') as zf:
                    exe_names = [n for n in zf.namelist()
                                 if n.lower().endswith('.exe')]
                    if not exe_names:
                        raise RuntimeError("No .exe found in VoiceMeeter ZIP")
                    zf.extract(exe_names[0], str(SCRIPT_DIR))
                    installer = SCRIPT_DIR / exe_names[0]
                vm_zip.unlink(missing_ok=True)
                log("Launching VoiceMeeter installer...")
                log("Complete the installer, then setup will continue.")
                import ctypes
                ret = ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", str(installer), None, None, 1)
                if ret <= 32:
                    raise RuntimeError(f"Failed to launch installer (code {ret})")
                # Wait for installer process to finish
                time.sleep(5)
                while True:
                    r = subprocess.run(
                        ["tasklist", "/FI",
                         f"IMAGENAME eq {installer.name}"],
                        capture_output=True, text=True)
                    if installer.name.lower() not in r.stdout.lower():
                        break
                    time.sleep(2)
                installer.unlink(missing_ok=True)

                time.sleep(3)
                if find_dll():
                    check("voicemeeter", True)
                    log("VoiceMeeter installed!")
                else:
                    check("voicemeeter", True)
                    log("VoiceMeeter installed! Restart needed.")
                    self._ui(lambda: self._start_reboot_countdown())
                    return
            except Exception as e:
                check("voicemeeter", False)
                log(f"VoiceMeeter download failed: {e}")
                log("Install manually: vb-audio.com/Voicemeeter/banana.htm")
                return

        # --- svcl.exe ---
        if SVCL_PATH.exists():
            check("svcl", True)
            log("svcl.exe found")
        else:
            log("Downloading svcl.exe...")
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
                if SVCL_PATH.exists():
                    check("svcl", True)
                    log("svcl.exe downloaded")
                else:
                    check("svcl", False)
                    log("svcl.exe missing after download. Your antivirus "
                        "may have quarantined it.")
                    log("Add an exception for svcl.exe and re-run setup.")
                    return
            except Exception as e:
                check("svcl", False)
                log(f"svcl download failed: {e}")
                return

        # --- Python packages ---
        if all(self._check_pkg(p) for p in REQUIRED_PACKAGES):
            check("packages", True)
            log("Python packages OK")
        else:
            log("Installing Python packages...")
            try:
                result = subprocess.run(
                    [sys.executable, "-m", "pip", "install", "-r",
                     str(SCRIPT_DIR / "requirements.txt")],
                    capture_output=True, text=True, timeout=120,
                )
                if result.returncode == 0:
                    check("packages", True)
                    log("Packages installed")
                else:
                    check("packages", False)
                    log(f"pip error: {result.stderr[:200]}")
                    return
            except Exception as e:
                check("packages", False)
                log(f"Package install failed: {e}")
                return

        # --- Device detection ---
        log("Detecting audio devices...")
        try:
            self._ensure_voicemeeter()
            vm = VMDeviceEnumerator()
            self._vm_inputs = vm.input_devices()
            vm_outputs = vm.output_devices()
            vm.close()

            wdm_inputs = [d for d in self._vm_inputs if d["type"] == "wdm"]
            mic_names = [d["name"] for d in wdm_inputs]

            wdm_outputs = [d for d in vm_outputs if d["type"] == "wdm"]
            vr_names = [d["name"] for d in wdm_outputs
                        if "voicemeeter" not in d["name"].lower()]

            if SVCL_PATH.exists():
                self._svcl_devices = query_svcl_devices()

            check("devices", True)
            log(f"Found {len(mic_names)} mics, {len(vr_names)} outputs")

            # Populate dropdowns on UI thread
            def populate():
                self.mic_combo["values"] = mic_names
                if mic_names:
                    sel = 0
                    for i, name in enumerate(mic_names):
                        nl = name.lower()
                        if ("microphone" in nl or "mic" in nl) \
                                and "steam" not in nl:
                            sel = i
                            break
                    self.mic_combo.current(sel)

                self.vr_combo["values"] = vr_names
                if vr_names:
                    sel = 0
                    for i, name in enumerate(vr_names):
                        if "steam streaming speakers" in name.lower():
                            sel = i
                            break
                    self.vr_combo.current(sel)

                self._show_phase("configure")

            self._ui(populate)

        except Exception as e:
            check("devices", False)
            log(f"Device detection failed: {e}")

    def _ensure_voicemeeter(self):
        try:
            import psutil
            for p in psutil.process_iter(["name"]):
                if p.info["name"] and is_vm_process(p.info["name"]):
                    return
        except ImportError:
            pass
        vm_exe = find_exe()
        if vm_exe:
            self._log(f"Launching VoiceMeeter ({vm_exe.name})...")
            subprocess.Popen([str(vm_exe)],
                             creationflags=subprocess.CREATE_NO_WINDOW)
            self._vm_launched_by_us = True
            time.sleep(4)

    # ------------------------------------------------------------------
    # Reboot flow
    # ------------------------------------------------------------------
    def _start_reboot_countdown(self):
        self._show_phase("reboot")
        self._reboot_seconds = 15
        self._reboot_cancelled = False
        self._tick_reboot()

    def _tick_reboot(self):
        if self._reboot_cancelled:
            return
        if self._reboot_seconds <= 0:
            self._do_reboot()
            return
        self._reboot_label.config(
            text=f"Restarting in {self._reboot_seconds} seconds...")
        self._reboot_seconds -= 1
        self.root.after(1000, self._tick_reboot)

    def _do_reboot(self):
        # Create resume shortcut so wizard auto-launches after reboot
        try:
            create_shortcut(
                sys.executable,
                str(RESUME_SHORTCUT),
                args=f'"{SCRIPT_DIR / "setup_wizard.py"}"',
                description="Resumes VR Audio Switcher setup after reboot")
        except Exception:
            pass
        self._log("Restarting PC...")
        subprocess.Popen(["shutdown", "/r", "/t", "3"])
        self.root.after(2000, self.root.destroy)

    def _cancel_reboot(self):
        self._reboot_cancelled = True
        self._reboot_label.config(
            text="Restart cancelled. Please restart your PC manually, "
                 "then run install.bat again.")
        self._reboot_now_btn.config(state="normal")
        self._cancel_reboot_btn.config(state="disabled")

    # ------------------------------------------------------------------
    # Phase 3: Finish Setup clicked
    # ------------------------------------------------------------------
    def _on_finish_click(self):
        self._finish_btn.config(state="disabled")
        threading.Thread(target=self._finish_thread,
                         daemon=True).start()

    def _finish_thread(self):
        def log(msg):
            self._ui(lambda: self._log(msg))

        mic_name = self.mic_var.get()
        vr_name = self.vr_var.get()

        if not mic_name:
            log("Please select a microphone.")
            self._ui(lambda: self._finish_btn.config(state="normal"))
            return
        if not vr_name:
            log("Please select a VR headset audio output.")
            self._ui(lambda: self._finish_btn.config(state="normal"))
            return

        log("Configuring...")
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
            vaio_id = (r"VB-Audio Voicemeeter VAIO\Device"
                       r"\Voicemeeter Input\Render")

        # 2. config.json
        config = {
            "poll_interval_seconds": 3,
            "vr_process": "vrserver.exe",
            "exclude_processes": ["vrchat.exe"],
            "svcl_path": "svcl.exe",
            "vr_device": vaio_id,
            "debounce_seconds": 5,
            "music_strip": 3,
            "vrchat_mic_confirmed": False,
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
            log(f"Microphone: {mic_name}")
        except Exception as e:
            errors.append(f"vm_devices.json: {e}")

        # 4. "Listen to this device" on B2
        listen_ok = self._configure_listen(vr_name)
        if listen_ok:
            log(f"Audio routing configured (B2 \u2192 {vr_name})")
        else:
            log("Audio routing: manual step needed")

        # 5. Shortcuts
        pythonw = Path(sys.executable).parent / "pythonw.exe"
        if not pythonw.exists():
            pythonw = Path(sys.executable)
        script = str(SCRIPT_DIR / "vr_audio_switcher.py")

        try:
            import winreg
            try:
                with winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER,
                    r"SOFTWARE\Microsoft\Windows\CurrentVersion"
                    r"\Explorer\User Shell Folders") as key:
                    desktop_raw, _ = winreg.QueryValueEx(key, "Desktop")
                    desktop = Path(os.path.expandvars(desktop_raw))
            except OSError:
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
            log("Startup shortcut \u2713")
        except Exception as e:
            errors.append(f"Startup shortcut: {e}")

        # 6. Shut down VoiceMeeter
        if self._vm_launched_by_us:
            self._ui(lambda: self._shutdown_voicemeeter())

        if errors:
            for err in errors:
                log(f"Warning: {err}")

        if not listen_ok:
            self._ui(lambda: self._show_manual_listen(vr_name))
        else:
            log("")
            log("Setup complete!")
            self._ui(lambda: self._show_phase("done"))

    def _configure_listen(self, vr_output_name: str) -> bool:
        if not self._svcl_devices:
            if SVCL_PATH.exists():
                self._svcl_devices = query_svcl_devices()
            else:
                return False

        b2_dev = find_svcl_device(self._svcl_devices, "Voicemeeter Out B2",
                                  "Capture")
        if not b2_dev:
            self._log("Could not find Voicemeeter Out B2")
            return False

        b2_guid = extract_guid(b2_dev["item_id"])
        if not b2_guid:
            return False

        target_dev = None
        for d in self._svcl_devices:
            if (d["direction"] == "Render" and d["type"] == "Device"
                    and d["state"] == "Active"):
                svcl_full = f"{d['name']} ({d['device_name']})"
                if svcl_full == vr_output_name:
                    target_dev = d
                    break
                if d["device_name"] and d["device_name"] in vr_output_name:
                    target_dev = d
                    break
        if not target_dev:
            return False

        target_endpoint_id = target_dev["item_id"]
        if not target_endpoint_id:
            return False

        ps = build_listen_ps_script(b2_guid, target_endpoint_id)
        self._log("Requesting admin permission for audio config...")
        ok, _ = run_elevated_ps(ps)
        return ok

    def _show_manual_listen(self, vr_name: str):
        self._log("MANUAL STEP NEEDED:")
        self._log("1. Open Windows Sound Settings (Recording tab)")
        self._log("2. Right-click 'Voicemeeter Out B2' \u2192 Properties")
        self._log("3. Listen tab \u2192 Check 'Listen to this device'")
        self._log(f"4. Set playback to: {vr_name}")
        self._log("5. Click OK")
        self._log("")
        self._log("Then click Launch below.")
        self._show_phase("done")

        messagebox.showinfo(
            "One Manual Step",
            "Automatic audio routing requires admin permission.\n\n"
            "Please do this one-time step:\n\n"
            "1. Open Windows Sound Settings \u2192 Recording tab\n"
            "2. Right-click 'Voicemeeter Out B2'\n"
            "3. Properties \u2192 Listen tab\n"
            "4. Check 'Listen to this device'\n"
            f"5. Set playback to: {vr_name}\n"
            "6. Click OK",
        )
        subprocess.Popen(
            ["rundll32.exe", "Shell32.dll,Control_RunDLL", "mmsys.cpl,,1"],
            creationflags=subprocess.CREATE_NO_WINDOW,
        )

    def _shutdown_voicemeeter(self):
        try:
            import psutil
            for proc in psutil.process_iter(["name"]):
                if proc.info["name"] and is_vm_process(proc.info["name"]):
                    proc.kill()
                    break
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Phase 4: Launch
    # ------------------------------------------------------------------
    def _launch(self):
        pythonw = Path(sys.executable).parent / "pythonw.exe"
        if not pythonw.exists():
            pythonw = Path(sys.executable)
        script = str(SCRIPT_DIR / "vr_audio_switcher.py")
        subprocess.Popen([str(pythonw), script])
        self._log("Launched!")
        self.root.after(2000, self.root.destroy)

    # ------------------------------------------------------------------
    # Device detection (for Refresh button in configure phase)
    # ------------------------------------------------------------------
    def _detect_devices(self):
        self._log("Refreshing devices...")
        try:
            self._ensure_voicemeeter()
            vm = VMDeviceEnumerator()
            self._vm_inputs = vm.input_devices()
            vm_outputs = vm.output_devices()
            vm.close()

            wdm_inputs = [d for d in self._vm_inputs if d["type"] == "wdm"]
            mic_names = [d["name"] for d in wdm_inputs]
            self.mic_combo["values"] = mic_names
            if mic_names:
                sel = 0
                for i, name in enumerate(mic_names):
                    nl = name.lower()
                    if ("microphone" in nl or "mic" in nl) \
                            and "steam" not in nl:
                        sel = i
                        break
                self.mic_combo.current(sel)

            wdm_outputs = [d for d in vm_outputs if d["type"] == "wdm"]
            vr_names = [d["name"] for d in wdm_outputs
                        if "voicemeeter" not in d["name"].lower()]
            self.vr_combo["values"] = vr_names
            if vr_names:
                sel = 0
                for i, name in enumerate(vr_names):
                    if "steam streaming speakers" in name.lower():
                        sel = i
                        break
                self.vr_combo.current(sel)

            if SVCL_PATH.exists():
                self._svcl_devices = query_svcl_devices()

            self._log(f"Found {len(mic_names)} mics, "
                      f"{len(vr_names)} outputs")

        except Exception as e:
            self._log(f"Device detection failed: {e}")

    def run(self):
        self.root.mainloop()


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    SetupWizard().run()
