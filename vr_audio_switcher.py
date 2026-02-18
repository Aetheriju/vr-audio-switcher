"""VR Audio Auto-Switcher — detects SteamVR and toggles Chrome's audio output."""

import csv
import ctypes
import json
import logging
import logging.handlers
import os
import subprocess
import sys
import threading
import time
from enum import Enum, auto
from pathlib import Path

import psutil
import pystray
from PIL import Image, ImageDraw

SCRIPT_DIR = Path(__file__).parent.resolve()
CONFIG_PATH = SCRIPT_DIR / "config.json"
LOG_PATH = SCRIPT_DIR / "switcher.log"
MUTEX_NAME = "Global\\VRAudioSwitcherMutex"
STATE_PATH = SCRIPT_DIR / "state.json"
VM_DEVICES_PATH = SCRIPT_DIR / "vm_devices.json"

ENFORCE_INTERVAL = 15  # seconds between enforcement cycles


# ---------------------------------------------------------------------------
# Single-instance lock
# ---------------------------------------------------------------------------
def acquire_single_instance():
    kernel32 = ctypes.windll.kernel32
    mutex = kernel32.CreateMutexW(None, False, MUTEX_NAME)
    if kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
        return None
    return mutex


# ---------------------------------------------------------------------------
# User mode: what the user has selected
# ---------------------------------------------------------------------------
class UserMode(Enum):
    DESKTOP = auto()    # Blue   — forced desktop/soundbar
    AUTO = auto()       # Green  — auto-detect SteamVR
    VR = auto()         # Red    — forced VR (music thru mic)
    SILENT_VR = auto()  # Yellow — VR but music only in headset


# Audio output: what Chrome is actually doing
class AudioOutput(Enum):
    DESKTOP = auto()
    VR = auto()


# Cycle order for left-click
MODE_CYCLE = [UserMode.DESKTOP, UserMode.AUTO, UserMode.SILENT_VR, UserMode.VR]

# Colors
MODE_COLORS = {
    UserMode.DESKTOP: (66, 133, 244, 255),    # Blue
    UserMode.AUTO: (76, 175, 80, 255),         # Green
    UserMode.VR: (234, 67, 53, 255),           # Red
    UserMode.SILENT_VR: (255, 193, 7, 255),    # Yellow
}


# ---------------------------------------------------------------------------
# Icon generation
# ---------------------------------------------------------------------------
def create_icon(user_mode: UserMode, size: int = 64) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    fill = MODE_COLORS.get(user_mode, (158, 158, 158, 255))
    margin = 4
    draw.ellipse([margin, margin, size - margin, size - margin], fill=fill)
    return img


def is_process_running(name: str) -> bool:
    name_lower = name.lower()
    for proc in psutil.process_iter(["name"]):
        try:
            if proc.info["name"] and proc.info["name"].lower() == name_lower:
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return False


# ---------------------------------------------------------------------------
# VoiceMeeter Remote API (controls Strip[3].B1 for mic routing)
# ---------------------------------------------------------------------------
class VoiceMeeterRemote:
    """Thin wrapper around VoicemeeterRemote64.dll for toggling mic routing."""

    def __init__(self):
        self._dll = None
        self._logged_in = False

    def _ensure_connected(self) -> bool:
        if self._logged_in:
            return True
        try:
            if self._dll is None:
                from vm_path import find_dll
                dll_path = find_dll()
                if not dll_path:
                    logging.error("VoiceMeeter DLL not found")
                    return False
                self._dll = ctypes.WinDLL(str(dll_path))
            ret = self._dll.VBVMR_Login()
            # 0 = OK, 1 = OK but VoiceMeeter not running (launched it)
            if ret in (0, 1):
                self._logged_in = True
                time.sleep(0.1)  # brief settle after login
                return True
            logging.warning("VoiceMeeter Login returned %d", ret)
            return False
        except Exception:
            logging.exception("VoiceMeeter connect failed")
            return False

    def set_strip_b1(self, strip: int, enabled: bool) -> bool:
        """Set Strip[N].B1 on/off. Returns True on success."""
        if not self._ensure_connected():
            return False
        param = f"Strip[{strip}].B1".encode("ascii")
        value = ctypes.c_float(1.0 if enabled else 0.0)
        try:
            ret = self._dll.VBVMR_SetParameterFloat(param, value)
            if ret == 0:
                return True
            logging.warning("VoiceMeeter SetParameterFloat returned %d", ret)
            return False
        except Exception:
            logging.exception("VoiceMeeter set_strip_b1 failed")
            self._logged_in = False  # force reconnect next time
            return False

    def set_param(self, param: str, value: float) -> bool:
        """Set any VoiceMeeter parameter by name."""
        if not self._ensure_connected():
            return False
        try:
            ret = self._dll.VBVMR_SetParameterFloat(
                param.encode("ascii"), ctypes.c_float(value))
            return ret == 0
        except Exception:
            logging.exception("VoiceMeeter set_param(%s) failed", param)
            self._logged_in = False
            return False

    def get_string_param(self, param: str) -> str | None:
        """Get a string parameter (e.g. Strip[0].device.name)."""
        if not self._ensure_connected():
            return None
        try:
            buf = ctypes.create_string_buffer(512)
            ret = self._dll.VBVMR_GetParameterStringA(
                param.encode("ascii"), buf)
            if ret == 0:
                val = buf.value.decode("utf-8", errors="replace").strip()
                return val if val else None
            return None
        except Exception:
            logging.exception("get_string_param(%s) failed", param)
            self._logged_in = False
            return None

    def set_string_param(self, param: str, value: str) -> bool:
        """Set a string parameter (e.g. Strip[0].device.wdm)."""
        if not self._ensure_connected():
            return False
        try:
            ret = self._dll.VBVMR_SetParameterStringA(
                param.encode("ascii"), value.encode("utf-8"))
            if ret == 0:
                return True
            logging.warning("set_string_param(%s) returned %d", param, ret)
            return False
        except Exception:
            logging.exception("set_string_param(%s) failed", param)
            self._logged_in = False
            return False

    def shutdown(self):
        """Gracefully shut down VoiceMeeter via API (lets it save settings)."""
        if not self._ensure_connected():
            return
        try:
            self._dll.VBVMR_SetParameterFloat(
                b"Command.Shutdown", ctypes.c_float(1.0))
            logging.info("VoiceMeeter shutdown command sent")
        except Exception:
            logging.exception("VoiceMeeter shutdown failed")

    def close(self):
        if self._logged_in and self._dll:
            try:
                self._dll.VBVMR_Logout()
            except Exception:
                pass
            self._logged_in = False


# ---------------------------------------------------------------------------
# Audio switcher (wraps svcl.exe)
# ---------------------------------------------------------------------------
class AudioSwitcher:
    # System processes that should never have their audio switched
    SYSTEM_EXCLUDE = {
        "vrchat.exe", "vrserver.exe", "vrmonitor.exe", "vrwebhelper.exe",
        "steamwebhelper.exe", "voicemeeterpro.exe", "voicemeeter.exe",
        "voicemeeter8.exe", "voicemeeter8x64.exe", "svchost.exe",
        "rundll32.exe", "audiodg.exe", "dwm.exe",
    }

    def __init__(self, config: dict):
        self.svcl_path = str(SCRIPT_DIR / config["svcl_path"])
        self.vr_device = config["vr_device"]
        self._multi_target = "exclude_processes" in config

        if self._multi_target:
            self._user_exclude = {
                p.lower() for p in config["exclude_processes"]}
        else:
            # Legacy single-target mode (old config with target_process)
            self._target = config.get("target_process", "chrome.exe")

        if not Path(self.svcl_path).exists():
            raise FileNotFoundError(f"svcl.exe not found at {self.svcl_path}")

    def _enumerate_audio_apps(self) -> set[str]:
        """Query svcl for all processes with active render audio sessions."""
        tmp = SCRIPT_DIR / "_enum_apps.csv"
        try:
            subprocess.run(
                [self.svcl_path, "/scomma", str(tmp),
                 "/Columns", "Name,Type,Direction,Process Path"],
                capture_output=True, timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            processes = set()
            with open(tmp, newline="",
                      encoding="utf-8-sig", errors="replace") as f:
                for row in csv.DictReader(f):
                    if (row.get("Type", "").strip() == "Application"
                            and row.get("Direction", "").strip() == "Render"):
                        proc_path = row.get("Process Path", "").strip()
                        if proc_path:
                            processes.add(Path(proc_path).name.lower())
            return processes
        except Exception:
            logging.debug("Audio app enumeration failed", exc_info=True)
            return set()
        finally:
            try:
                tmp.unlink(missing_ok=True)
            except OSError:
                pass

    def _svcl_set(self, device: str, process: str) -> bool:
        """Run svcl /SetAppDefault for a single process."""
        cmd = [self.svcl_path, "/Stdout",
               "/SetAppDefault", device, "all", process]
        try:
            result = subprocess.run(
                cmd, capture_output=True, text=True, timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            stdout = result.stdout.strip()
            return "1 item" in stdout or "items found" in stdout
        except (FileNotFoundError, subprocess.TimeoutExpired):
            return False

    def switch_to(self, output: AudioOutput) -> bool:
        """Set audio output for target app(s). Returns True if any succeeded."""
        device = self.vr_device if output == AudioOutput.VR \
            else "DefaultRenderDevice"

        if not self._multi_target:
            # Legacy single-target mode
            if not is_process_running(self._target):
                return False
            return self._svcl_set(device, self._target)

        # Multi-target: enumerate and switch all non-excluded audio apps
        apps = self._enumerate_audio_apps()
        exclude = self.SYSTEM_EXCLUDE | self._user_exclude
        targets = [p for p in apps if p not in exclude]

        if not targets:
            return False

        ok_count = 0
        for proc in targets:
            if self._svcl_set(device, proc):
                ok_count += 1
        logging.debug("Switched %d/%d audio apps to %s",
                      ok_count, len(targets), device[:30])
        return ok_count > 0


# ---------------------------------------------------------------------------
# SteamVR detector (polling thread)
# ---------------------------------------------------------------------------
class SteamVRDetector:
    def __init__(self, config: dict, on_change):
        self.process_name = config["steamvr_process"].lower()
        self.poll_interval = config["poll_interval_seconds"]
        self.debounce = config["debounce_seconds"]
        self.on_change = on_change
        self._stop = threading.Event()
        self._thread = None
        self._vr_running = None
        self._last_change = 0.0

    def start(self):
        self._thread = threading.Thread(target=self._poll, daemon=True)
        self._thread.start()

    def stop(self):
        self._stop.set()
        if self._thread:
            self._thread.join(timeout=10)

    def is_steamvr_running(self) -> bool:
        return is_process_running(self.process_name)

    def _poll(self):
        while not self._stop.is_set():
            try:
                running = self.is_steamvr_running()
                if running != self._vr_running:
                    now = time.time()
                    if now - self._last_change >= self.debounce:
                        self._vr_running = running
                        self._last_change = now
                        logging.info("SteamVR %s", "started" if running else "stopped")
                        self.on_change(running)
            except Exception:
                logging.exception("Poll error")
            self._stop.wait(self.poll_interval)


# ---------------------------------------------------------------------------
# Tray application
# ---------------------------------------------------------------------------
class VRAudioSwitcher:
    def __init__(self, config: dict):
        self.config = config
        self.audio = AudioSwitcher(config)
        self.detector = SteamVRDetector(config, self._on_steamvr_change)
        self.vm = VoiceMeeterRemote()
        self.music_strip = config.get("music_strip", 3)
        self._user_mode = UserMode.AUTO
        self._current_output = None
        self._mic_enabled = True  # Strip[3].B1 state
        self._confirmed = False   # True when svcl matched at least one app
        self._stop = threading.Event()
        self._lock = threading.Lock()
        self.icon = None

    def _desired_output(self) -> AudioOutput:
        """Determine what Chrome's output should be based on user mode."""
        if self._user_mode == UserMode.DESKTOP:
            return AudioOutput.DESKTOP
        if self._user_mode in (UserMode.VR, UserMode.SILENT_VR):
            return AudioOutput.VR
        # AUTO — follow SteamVR
        return AudioOutput.VR if self.detector.is_steamvr_running() else AudioOutput.DESKTOP

    def _desired_mic(self) -> bool:
        """Should music go through the VRChat mic (Strip[3].B1)?"""
        # Only Public (VR) enables mic routing; Auto and Private keep it off
        return self._user_mode == UserMode.VR

    def _apply(self, force: bool = False):
        """Switch Chrome's output and mic routing if needed."""
        with self._lock:
            desired = self._desired_output()
            desired_mic = self._desired_mic()

            need_switch = force or desired != self._current_output or not self._confirmed
            if need_switch:
                ok = self.audio.switch_to(desired)
                if ok:
                    changed = desired != self._current_output
                    self._current_output = desired
                    self._confirmed = True
                    if changed:
                        logging.info("Audio -> %s", desired.name)
                else:
                    self._confirmed = False

            # Sync music-to-mic routing (Strip[3].B1)
            if desired_mic != self._mic_enabled or force:
                ok = self.vm.set_strip_b1(self.music_strip, desired_mic)
                if ok:
                    if desired_mic != self._mic_enabled:
                        logging.info("Mic routing -> %s", "ON" if desired_mic else "OFF")
                    self._mic_enabled = desired_mic

            # Ensure mic and music bus routing flags stay set
            if force:
                self.vm.set_param("Strip[0].B1", 1.0)  # mic → B1
                self.vm.set_param("Strip[3].B2", 1.0)  # music → B2

            self._update_tray()

    def _enforce_loop(self):
        """Periodically re-apply audio and monitor VoiceMeeter health."""
        while not self._stop.is_set():
            self._stop.wait(ENFORCE_INTERVAL)
            if self._stop.is_set():
                break
            try:
                self._check_mode_request()
                self._apply(force=True)
                self._check_voicemeeter_health()
            except Exception:
                logging.exception("Enforce error")

    def _check_voicemeeter_health(self):
        """If VoiceMeeter crashed, restart it and restore settings."""
        from vm_path import is_vm_process, find_exe
        vm_running = any(
            is_vm_process(p.info.get("name", ""))
            for p in psutil.process_iter(["name"])
        )
        if vm_running:
            return
        logging.warning("VoiceMeeter not running — restarting...")
        vm_exe = find_exe()
        if not vm_exe:
            return
        subprocess.Popen([str(vm_exe)], creationflags=subprocess.CREATE_NO_WINDOW)
        time.sleep(4)
        # Reconnect and restore devices
        self.vm._logged_in = False
        self.vm._ensure_connected()
        if VM_DEVICES_PATH.exists():
            try:
                with open(VM_DEVICES_PATH) as f:
                    devs = json.load(f)
                for key, name in devs.items():
                    self.vm.set_string_param(f"{key}.device.wdm", name)
                logging.info("VoiceMeeter restarted — devices restored")
            except Exception:
                logging.exception("Device restore after VM restart failed")
        # Restore gains from vm_state.json
        vm_state_path = SCRIPT_DIR / "vm_state.json"
        if vm_state_path.exists():
            try:
                with open(vm_state_path) as f:
                    params = json.load(f)
                for k, v in params.items():
                    self.vm.set_param(k, float(v))
                logging.info("VoiceMeeter gains restored")
            except Exception:
                logging.exception("Gain restore after VM restart failed")

    def _on_steamvr_change(self, running: bool):
        """Called by detector when SteamVR starts/stops."""
        if not running and self._user_mode in (UserMode.VR, UserMode.SILENT_VR):
            self._user_mode = UserMode.AUTO
            logging.info("SteamVR gone while in VR mode — falling back to AUTO")
            self._write_state()
        if self._user_mode == UserMode.AUTO:
            self._apply()

    def _update_tray(self):
        if self.icon:
            self.icon.icon = create_icon(self._user_mode)
            labels = {
                UserMode.DESKTOP: "Desktop",
                UserMode.AUTO: f"Auto ({self._current_output.name})" if self._current_output else "Auto",
                UserMode.VR: "Public",
                UserMode.SILENT_VR: "Private",
            }
            self.icon.title = f"Audio: {labels[self._user_mode]}"

    def _write_state(self):
        """Write current mode to state.json for mixer communication."""
        try:
            state = {}
            if STATE_PATH.exists():
                try:
                    with open(STATE_PATH) as f:
                        state = json.load(f)
                except Exception:
                    pass
            state["current_mode"] = self._user_mode.name
            state.pop("requested_mode", None)
            with open(STATE_PATH, "w") as f:
                json.dump(state, f)
        except Exception:
            logging.exception("Failed to write state file")

    def _check_mode_request(self):
        """Check if the mixer requested a mode change via state.json."""
        if not STATE_PATH.exists():
            return
        try:
            with open(STATE_PATH) as f:
                state = json.load(f)
            requested = state.get("requested_mode")
            if not requested:
                return
            mode_map = {m.name: m for m in UserMode}
            if requested in mode_map:
                self._user_mode = mode_map[requested]
                logging.info("Mode request from mixer -> %s", requested)
                self._apply()
            # Clear the request
            state["requested_mode"] = None
            state["current_mode"] = self._user_mode.name
            with open(STATE_PATH, "w") as f:
                json.dump(state, f)
        except Exception:
            logging.exception("Failed to read mode request")

    def _cycle_mode(self, icon, item):
        """Left-click: cycle Desktop -> Auto -> VR -> Private -> Desktop..."""
        try:
            idx = MODE_CYCLE.index(self._user_mode)
        except ValueError:
            idx = -1
        self._user_mode = MODE_CYCLE[(idx + 1) % len(MODE_CYCLE)]
        logging.info("User mode -> %s", self._user_mode.name)
        self._apply()
        self._write_state()

    def _set_mode(self, mode: UserMode):
        def callback(icon, item):
            self._user_mode = mode
            logging.info("User mode -> %s", mode.name)
            self._apply()
            self._write_state()
        return callback

    def _is_mode(self, mode: UserMode):
        def check(item):
            return self._user_mode == mode
        return check

    def _open_mixer(self, icon, item):
        """Launch the mixer window as a separate process."""
        mixer_path = SCRIPT_DIR / "mixer.py"
        if mixer_path.exists():
            subprocess.Popen(
                [sys.executable, str(mixer_path)],
                creationflags=subprocess.CREATE_NO_WINDOW,
            )

    def _check_updates(self, icon, item):
        """Check for updates in background, prompt if available."""
        from updater import update_available, do_update, restart_app
        def _run():
            avail, local, remote = update_available()
            if not avail:
                import tkinter as tk
                from tkinter import messagebox
                root = tk.Tk(); root.withdraw()
                messagebox.showinfo("Up to Date",
                                    f"You're on the latest version (v{local}).")
                root.destroy()
                return
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk(); root.withdraw()
            if messagebox.askyesno(
                "Update Available",
                f"v{remote} is available (you have v{local}).\n\n"
                "Update now? The app will restart.",
            ):
                root.destroy()
                ok, msg = do_update()
                if ok:
                    icon.stop()
                    restart_app()
                else:
                    root2 = tk.Tk(); root2.withdraw()
                    messagebox.showwarning("Update Failed", msg)
                    root2.destroy()
            else:
                root.destroy()
        threading.Thread(target=_run, daemon=True).start()

    def _open_settings(self, icon, item):
        """Open a settings window for configuring excluded apps."""
        def _show():
            import tkinter as tk

            win = tk.Tk()
            win.title("VR Audio Switcher \u2014 Settings")
            win.resizable(False, False)
            win.attributes("-topmost", True)
            bg, fg = "#1e1e1e", "#e0e0e0"
            win.configure(bg=bg)

            tk.Label(win, text="Excluded Apps", bg=bg, fg="#4caf50",
                     font=("Segoe UI", 10, "bold")).pack(anchor="w",
                     padx=15, pady=(12, 4))
            tk.Label(win, text="Audio for these apps will NOT be managed:",
                     bg=bg, fg="#888", font=("Segoe UI", 8)).pack(
                     anchor="w", padx=15)

            list_frame = tk.Frame(win, bg=bg)
            list_frame.pack(fill="x", padx=15, pady=(4, 4))

            lb = tk.Listbox(list_frame, bg="#111", fg=fg, selectbackground="#333",
                            font=("Consolas", 9), height=6, relief="flat")
            lb.pack(side="left", fill="both", expand=True)

            user_exclude = list(self.config.get("exclude_processes", ["vrchat.exe"]))
            for p in user_exclude:
                lb.insert("end", p)

            btn_frame = tk.Frame(list_frame, bg=bg)
            btn_frame.pack(side="right", padx=(8, 0))

            def _add():
                from tkinter import simpledialog
                name = simpledialog.askstring(
                    "Add Exclusion", "Process name (e.g. discord.exe):",
                    parent=win)
                if name and name.strip():
                    name = name.strip().lower()
                    if not name.endswith(".exe"):
                        name += ".exe"
                    lb.insert("end", name)

            def _remove():
                sel = lb.curselection()
                if sel:
                    lb.delete(sel[0])

            tk.Button(btn_frame, text="+ Add", bg="#333", fg=fg,
                      relief="flat", padx=8, pady=2, font=("Segoe UI", 8),
                      command=_add).pack(pady=2)
            tk.Button(btn_frame, text="\u2212 Remove", bg="#333", fg=fg,
                      relief="flat", padx=8, pady=2, font=("Segoe UI", 8),
                      command=_remove).pack(pady=2)

            # System exclusions note
            tk.Label(win, text="System processes (svchost, audiodg, VoiceMeeter, "
                     "SteamVR) are always excluded automatically.",
                     bg=bg, fg="#666", font=("Segoe UI", 7),
                     wraplength=350).pack(anchor="w", padx=15, pady=(4, 8))

            def _save():
                items = [lb.get(i) for i in range(lb.size())]
                self.config["exclude_processes"] = items
                try:
                    with open(CONFIG_PATH, "w") as f:
                        json.dump(self.config, f, indent=2)
                    # Reload in AudioSwitcher
                    if self.audio._multi_target:
                        self.audio._user_exclude = {
                            p.lower() for p in items}
                    logging.info("Settings saved: exclude=%s", items)
                except Exception:
                    logging.exception("Failed to save settings")
                win.destroy()

            tk.Button(win, text="Save", bg="#4caf50", fg="#000",
                      relief="flat", padx=20, pady=6,
                      font=("Segoe UI", 10, "bold"),
                      command=_save).pack(pady=(0, 12))

            win.update_idletasks()
            w, h = win.winfo_width(), win.winfo_height()
            x = (win.winfo_screenwidth() - w) // 2
            y = (win.winfo_screenheight() - h) // 2
            win.geometry(f"+{x}+{y}")
            win.mainloop()

        threading.Thread(target=_show, daemon=True).start()

    def _vrchat_mic_help(self, icon, item):
        """Show VRChat mic setup instructions."""
        def _show():
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk(); root.withdraw()
            messagebox.showinfo(
                "VRChat Mic Setup",
                "To hear music and voice in VRChat, set your mic to "
                "VoiceMeeter:\n\n"
                "1. Open VRChat\n"
                "2. Settings \u2192 Audio \u2192 Microphone\n"
                "3. Select \"Voicemeeter Out B1\"\n\n"
                "This only needs to be done once. VRChat saves the setting.")
            root.destroy()
        threading.Thread(target=_show, daemon=True).start()

    def _check_vrchat_mic(self):
        """Show one-time VRChat mic reminder on first launch."""
        if self.config.get("vrchat_mic_confirmed", False):
            return
        def _show():
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk(); root.withdraw()
            messagebox.showinfo(
                "VRChat Mic Setup Required",
                "One more step! Open VRChat and set your microphone:\n\n"
                "Settings \u2192 Audio \u2192 Microphone \u2192 "
                "\"Voicemeeter Out B1\"\n\n"
                "This tells VRChat to use VoiceMeeter as your mic input "
                "(voice + optional music).\n\n"
                "Click OK to dismiss this reminder permanently.")
            root.destroy()
            self.config["vrchat_mic_confirmed"] = True
            try:
                with open(CONFIG_PATH, "w") as f:
                    json.dump(self.config, f, indent=2)
            except Exception:
                logging.exception("Failed to save VRChat mic flag")
        threading.Thread(target=_show, daemon=True).start()

    def _uninstall(self, icon, item):
        """Remove shortcuts, configs, and shut down cleanly."""
        def _do():
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk(); root.withdraw()
            if not messagebox.askyesno(
                "Uninstall VR Audio Switcher",
                "This will remove:\n"
                "\u2022 Desktop shortcut\n"
                "\u2022 Startup shortcut\n"
                "\u2022 Config and state files\n"
                "\u2022 Log files\n\n"
                "The program files will remain in case you want to "
                "reinstall.\n\nContinue?"):
                root.destroy()
                return
            root.destroy()

            # Restore audio before cleanup
            self.audio.switch_to(AudioOutput.DESKTOP)
            self.vm.set_strip_b1(self.music_strip, True)

            # Remove shortcuts
            import winreg
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                        r"SOFTWARE\Microsoft\Windows\CurrentVersion"
                        r"\Explorer\User Shell Folders") as key:
                    raw, _ = winreg.QueryValueEx(key, "Desktop")
                    desktop = Path(os.path.expandvars(raw))
            except OSError:
                desktop = Path(os.environ.get("USERPROFILE", "")) / "Desktop"

            startup = (Path(os.environ.get("APPDATA", ""))
                       / "Microsoft" / "Windows" / "Start Menu"
                       / "Programs" / "Startup")

            for shortcut in [
                desktop / "VR Audio Switcher.lnk",
                startup / "VR Audio Switcher.lnk",
            ]:
                try:
                    shortcut.unlink(missing_ok=True)
                except OSError:
                    pass

            # Remove config/state files
            for fn in ["config.json", "vm_devices.json", "vm_state.json",
                       "state.json", "presets.json",
                       "switcher.log", "switcher.log.1",
                       "wizard.log", "wizard.log.1",
                       "_enum_apps.csv"]:
                try:
                    (SCRIPT_DIR / fn).unlink(missing_ok=True)
                except OSError:
                    pass

            logging.info("Uninstall complete")
            self.vm.shutdown()
            time.sleep(2)
            self.vm.close()
            self._stop.set()
            self.detector.stop()
            icon.stop()

        threading.Thread(target=_do, daemon=True).start()

    def _quit(self, icon, item):
        from vm_path import is_vm_process
        self._stop.set()
        self.detector.stop()
        # Restore desktop audio and mic routing before exit
        self.audio.switch_to(AudioOutput.DESKTOP)
        self.vm.set_strip_b1(self.music_strip, True)
        # Save hardware device assignments BEFORE shutting down VoiceMeeter
        try:
            devices = {}
            for i in range(3):  # Hardware strips 0-2
                name = self.vm.get_string_param(f"Strip[{i}].device.name")
                if name:
                    devices[f"Strip[{i}]"] = name
            for i in range(3):  # Hardware buses A1-A3
                name = self.vm.get_string_param(f"Bus[{i}].device.name")
                if name:
                    devices[f"Bus[{i}]"] = name
            if devices:
                with open(VM_DEVICES_PATH, "w") as f:
                    json.dump(devices, f, indent=2)
                logging.info("Saved %d device assignments", len(devices))
        except Exception:
            logging.exception("Failed to save device config")
        # Kill the mixer if it's running (match exact path)
        mixer_path = str(SCRIPT_DIR / "mixer.py")
        for proc in psutil.process_iter(["name", "cmdline"]):
            try:
                if proc.info["name"] and proc.info["name"].lower() in (
                        "python.exe", "pythonw.exe"):
                    cmdline = proc.info.get("cmdline") or []
                    if any(mixer_path in arg for arg in cmdline):
                        proc.kill()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        # Gracefully shut down VoiceMeeter, then force-kill if needed
        self.vm.shutdown()
        time.sleep(3)
        for proc in psutil.process_iter(["name"]):
            try:
                if is_vm_process(proc.info.get("name", "")):
                    proc.kill()
                    break
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        self.vm.close()
        icon.stop()

    def run(self):
        menu = pystray.Menu(
            # Hidden default item — left-click cycles modes
            pystray.MenuItem("Cycle", self._cycle_mode, default=True, visible=False),
            pystray.MenuItem("Desktop", self._set_mode(UserMode.DESKTOP),
                             checked=self._is_mode(UserMode.DESKTOP), radio=True),
            pystray.MenuItem("Auto-detect", self._set_mode(UserMode.AUTO),
                             checked=self._is_mode(UserMode.AUTO), radio=True),
            pystray.MenuItem("Private", self._set_mode(UserMode.SILENT_VR),
                             checked=self._is_mode(UserMode.SILENT_VR), radio=True),
            pystray.MenuItem("Public", self._set_mode(UserMode.VR),
                             checked=self._is_mode(UserMode.VR), radio=True),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Mixer", self._open_mixer),
            pystray.MenuItem("Settings", self._open_settings),
            pystray.MenuItem("VRChat Mic Help", self._vrchat_mic_help),
            pystray.MenuItem("Check for Updates", self._check_updates),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Uninstall", self._uninstall),
            pystray.MenuItem("Quit", self._quit),
        )

        self.icon = pystray.Icon(
            name="vr_audio_switcher",
            icon=create_icon(UserMode.AUTO),
            title="Audio: Starting...",
            menu=menu,
        )

        def on_setup(icon):
            try:
                icon.visible = True
                logging.info("Tray icon ready")
                self.detector.start()
                self._apply()
                self._write_state()
                self._check_vrchat_mic()
                threading.Thread(target=self._enforce_loop, daemon=True).start()
                threading.Thread(target=self._update_check_loop,
                                 daemon=True).start()
            except Exception:
                logging.exception("Setup error")

        self.icon.run(setup=on_setup)

    def _update_check_loop(self):
        """Periodically check for updates in the background (every 6 hours)."""
        UPDATE_INTERVAL = 6 * 3600  # 6 hours
        while not self._stop.wait(UPDATE_INTERVAL):
            try:
                from updater import update_available
                avail, local, remote = update_available()
                if avail:
                    logging.info("Background update check: v%s available "
                                 "(have v%s)", remote, local)
                    if self.icon:
                        self.icon.notify(
                            f"Update v{remote} available — right-click "
                            f"tray icon \u2192 Check for Updates",
                            "VR Audio Switcher")
            except Exception:
                logging.debug("Background update check failed", exc_info=True)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    mutex = acquire_single_instance()
    if mutex is None:
        # Already running — notify user instead of silent exit
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk(); root.withdraw()
            messagebox.showinfo(
                "VR Audio Switcher",
                "Already running! Look for the icon in your system tray.\n\n"
                "Right-click the tray icon for options.")
            root.destroy()
        except Exception:
            pass
        sys.exit(0)

    # Check for updates before anything else
    from updater import check_and_prompt
    if not check_and_prompt():
        sys.exit(0)

    def _show_error_and_exit(msg):
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk(); root.withdraw()
            messagebox.showerror("VR Audio Switcher", msg)
            root.destroy()
        except Exception:
            pass
        sys.exit(1)

    if not CONFIG_PATH.exists():
        _show_error_and_exit(
            "No config found.\n\n"
            "Please run the setup wizard first:\n"
            "Double-click install.bat or run setup_wizard.py")

    with open(CONFIG_PATH) as f:
        config = json.load(f)

    if not config.get("vr_device"):
        _show_error_and_exit(
            "Audio devices not configured.\n\n"
            "Please run the setup wizard first:\n"
            "Double-click install.bat or run setup_wizard.py")

    # Rotate log at 1 MB, keep 1 backup
    file_handler = logging.handlers.RotatingFileHandler(
        LOG_PATH, maxBytes=1_000_000, backupCount=1,
    )
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[file_handler, logging.StreamHandler()],
    )

    # Ensure VoiceMeeter is running before we start
    from vm_path import is_vm_process, find_exe
    vm_running = any(
        is_vm_process(p.info.get("name", ""))
        for p in psutil.process_iter(["name"])
    )
    if not vm_running:
        vm_exe = find_exe()
        if vm_exe:
            logging.info("Launching VoiceMeeter (%s)...", vm_exe.name)
            subprocess.Popen([str(vm_exe)],
                             creationflags=subprocess.CREATE_NO_WINDOW)
            time.sleep(4)
        else:
            logging.warning("VoiceMeeter not found — install it first")

    logging.info("VR Audio Switcher starting...")
    app = VRAudioSwitcher(config)

    # Restore hardware device assignments (mic input, bus outputs)
    vm_devices = {}
    if VM_DEVICES_PATH.exists():
        try:
            with open(VM_DEVICES_PATH) as f:
                vm_devices = json.load(f)
        except Exception:
            logging.exception("Failed to read vm_devices.json")

    for key, name in vm_devices.items():
        ok = app.vm.set_string_param(f"{key}.device.wdm", name)
        logging.info("Set %s.device.wdm = %s (%s)", key, name,
                     "OK" if ok else "FAIL")
    if vm_devices:
        time.sleep(1)

    # Restore mixer settings (last saved, or neutral defaults for new users)
    VM_DEFAULTS = {
        "Strip[3].Gain": 0.0, "Bus[3].Gain": 0.0, "Bus[4].Gain": 0.0,
        "Strip[0].Gain": 0.0,
        "Strip[3].eqgain1": 0.0, "Strip[3].eqgain2": 0.0, "Strip[3].eqgain3": 0.0,
    }
    vm_state_path = SCRIPT_DIR / "vm_state.json"
    vm_params = dict(VM_DEFAULTS)
    if vm_state_path.exists():
        try:
            with open(vm_state_path) as f:
                vm_params = json.load(f)
        except Exception:
            logging.exception("Failed to read vm_state.json, using defaults")

    # Always ensure routing flags are set (these don't get saved by the mixer)
    vm_params["Strip[0].B1"] = 1.0   # mic → B1 (VRChat)
    vm_params["Strip[3].B2"] = 1.0   # music → B2 (headset monitoring)
    vm_params["Strip[0].A1"] = 0.0   # mic doesn't go to hardware out
    vm_params["Strip[3].A1"] = 0.0   # music doesn't go to hardware out

    for param, value in vm_params.items():
        try:
            app.vm.set_param(param, float(value))
        except (ValueError, TypeError):
            logging.warning("Skipping invalid VM param %s=%s", param, value)
    time.sleep(0.5)  # let VoiceMeeter process params before _apply()
    # Re-apply routing flags a second time (VoiceMeeter can overwrite on startup)
    app.vm.set_param("Strip[0].B1", 1.0)
    app.vm.set_param("Strip[3].B2", 1.0)
    logging.info("Applied %d VM params (routing + %s)",
                 len(vm_params), "saved" if vm_state_path.exists() else "defaults")

    app.run()


if __name__ == "__main__":
    main()
