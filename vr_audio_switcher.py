"""VR Audio Auto-Switcher: detects VR and routes per-app audio output.

Lifecycle: runs silently watching for SteamVR. When VR starts, boots
VoiceMeeter, opens the mixer UI, and routes audio. When VR stops (or
the user closes the UI), shuts everything down and goes back to watching.
"""

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

SCRIPT_DIR = Path(__file__).parent.resolve()
CONFIG_PATH = SCRIPT_DIR / "config.json"
LOG_PATH = SCRIPT_DIR / "switcher.log"
MUTEX_NAME = "Global\\VRAudioSwitcherMutex"
STATE_PATH = SCRIPT_DIR / "state.json"
VM_DEVICES_PATH = SCRIPT_DIR / "vm_devices.json"

ENFORCE_INTERVAL = 5  # seconds between enforcement cycles
SPLASH_DONE = SCRIPT_DIR / "_splash_done"
SPLASH_STATUS = SCRIPT_DIR / "_splash_status"


def _splash_update(text):
    """Update the loading splash status text."""
    try:
        SPLASH_STATUS.write_text(text, encoding="utf-8")
    except OSError:
        pass


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
    AUTO = auto()       # Green  - auto-detect VR
    VR = auto()         # Red    — forced VR (music thru mic)
    SILENT_VR = auto()  # Yellow — VR but music only in headset


# Audio output: what apps are actually doing
class AudioOutput(Enum):
    DESKTOP = auto()
    VR = auto()


# Cycle order for mode buttons
MODE_CYCLE = [UserMode.DESKTOP, UserMode.AUTO, UserMode.SILENT_VR, UserMode.VR]


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
# VoiceMeeter Remote API
# ---------------------------------------------------------------------------
class VoiceMeeterRemote:
    """Wrapper around VoicemeeterRemote64.dll with auto-reconnect."""

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

    def get(self, param: str) -> float:
        """Get a float parameter (used by mixer sliders)."""
        if not self._ensure_connected():
            return 0.0
        try:
            self._dll.VBVMR_IsParametersDirty()
            buf = ctypes.c_float()
            ret = self._dll.VBVMR_GetParameterFloat(
                param.encode("ascii"), ctypes.byref(buf))
            if ret != 0:
                return 0.0
            return round(buf.value, 1)
        except Exception:
            logging.exception("VoiceMeeter get(%s) failed", param)
            self._logged_in = False
            return 0.0

    def set(self, param: str, value: float):
        """Set a float parameter (used by mixer sliders)."""
        self.set_param(param, float(value))

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

    # VR headset audio device name patterns (matched case-insensitive)
    VR_HEADSET_PATTERNS = {
        "steam streaming", "index", "vive", "rift", "oculus",
        "quest", "virtual desktop", "pico", "bigscreen",
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

    # Patterns that indicate a real speaker/soundbar/headphone (high priority)
    SPEAKER_PATTERNS = {
        "speaker", "headphone", "soundbar", "realtek",
        "samsung", "sony", "bose", "jbl", "harman",
        "logitech", "creative", "sonos",
    }
    # Patterns that indicate monitor/display audio (low priority)
    DISPLAY_AUDIO_PATTERNS = {
        "display audio", "nvidia high definition", "amd high definition",
        "intel display", "hdmi", "displayport",
    }

    def _find_desktop_device(self) -> str:
        """Find the best non-headset render device for Desktop mode."""
        tmp = SCRIPT_DIR / "_enum_devices.csv"
        try:
            subprocess.run(
                [self.svcl_path, "/scomma", str(tmp),
                 "/Columns", "Name,Type,Direction,Device State,"
                 "Command-Line Friendly ID"],
                capture_output=True, timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            speakers = []       # real speakers/soundbars (tier 1)
            other = []          # unknown devices (tier 2)
            display_audio = []  # monitor audio outputs (tier 3)
            with open(tmp, newline="",
                      encoding="utf-8-sig", errors="replace") as f:
                for row in csv.DictReader(f):
                    if (row.get("Type", "").strip() != "Device"
                            or row.get("Direction", "").strip() != "Render"
                            or row.get("Device State", "").strip() != "Active"):
                        continue
                    name = row.get("Name", "").strip()
                    fid = row.get("Command-Line Friendly ID", "").strip()
                    if not fid:
                        continue
                    name_lower = name.lower()
                    fid_lower = fid.lower()
                    # Skip VoiceMeeter virtual devices
                    if "voicemeeter" in name_lower or "vb-audio" in fid_lower:
                        continue
                    # Skip VR headset devices entirely
                    combined = name_lower + " " + fid_lower
                    if any(p in combined for p in self.VR_HEADSET_PATTERNS):
                        continue
                    # Classify remaining devices by priority
                    if any(p in combined
                           for p in self.SPEAKER_PATTERNS):
                        speakers.append(fid)
                    elif any(p in combined
                             for p in self.DISPLAY_AUDIO_PATTERNS):
                        display_audio.append(fid)
                    else:
                        other.append(fid)
            # Pick best available in priority order
            for tier in (speakers, other, display_audio):
                if tier:
                    logging.debug("Desktop device selected: %s (from %d candidates)",
                                  tier[0][:40], len(tier))
                    return tier[0]
            return "DefaultRenderDevice"
        except Exception:
            logging.debug("Desktop device enumeration failed", exc_info=True)
            return "DefaultRenderDevice"
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
            else self._find_desktop_device()

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
# VR detector (polling thread)
# ---------------------------------------------------------------------------
class VRDetector:
    def __init__(self, config: dict, on_change):
        self.process_name = config.get("vr_process",
                                       config.get("steamvr_process",
                                                   "vrserver.exe")).lower()
        self.poll_interval = config["poll_interval_seconds"]
        self.debounce = config["debounce_seconds"]
        self.on_change = on_change
        self._stop = threading.Event()
        self._thread = None
        self._vr_running = None
        self._last_change = 0.0

    def start(self):
        self._stop.clear()
        self._thread = threading.Thread(target=self._poll, daemon=True)
        self._thread.start()

    def stop(self):
        self._stop.set()
        if self._thread:
            self._thread.join(timeout=10)

    def is_vr_running(self) -> bool:
        return is_process_running(self.process_name)

    def _poll(self):
        while not self._stop.is_set():
            try:
                running = self.is_vr_running()
                if running != self._vr_running:
                    now = time.time()
                    if now - self._last_change >= self.debounce:
                        self._vr_running = running
                        self._last_change = now
                        logging.info("VR %s", "started" if running else "stopped")
                        self.on_change(running)
            except Exception:
                logging.exception("Poll error")
            self._stop.wait(self.poll_interval)


# ---------------------------------------------------------------------------
# Main application
# ---------------------------------------------------------------------------
class VRAudioSwitcher:
    def __init__(self, config: dict):
        self.config = config
        self.audio = AudioSwitcher(config)
        self.detector = VRDetector(config, self._on_vr_change)
        self.vm = VoiceMeeterRemote()
        self.music_strip = config.get("music_strip", 3)
        self._user_mode = UserMode.AUTO
        self._current_output = None
        self._mic_enabled = True  # Strip[3].B1 state
        self._confirmed = False   # True when svcl matched at least one app
        self._stop = threading.Event()
        self._lock = threading.Lock()
        self._vm_ready = threading.Event()
        # Session lifecycle
        self._session_active = False
        self._in_cleanup = False
        self._vr_start_event = threading.Event()
        self._user_quit = False
        self._vm_dialog_shown = False
        self.ui = None

    @staticmethod
    def _minimize_voicemeeter():
        """Hide VoiceMeeter's window so only our UI is visible."""
        from vm_path import is_vm_process
        user32 = ctypes.windll.user32
        SW_MINIMIZE = 6
        for proc in psutil.process_iter(["name", "pid"]):
            try:
                if not is_vm_process(proc.info.get("name", "")):
                    continue
                pid = proc.info["pid"]
                WNDENUMPROC = ctypes.WINFUNCTYPE(
                    ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
                def _enum_cb(hwnd, _lp):
                    tid_pid = ctypes.c_ulong()
                    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(tid_pid))
                    if tid_pid.value == pid and user32.IsWindowVisible(hwnd):
                        user32.ShowWindow(hwnd, SW_MINIMIZE)
                    return True
                user32.EnumWindows(WNDENUMPROC(_enum_cb), 0)
                break
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue

    def _init_voicemeeter(self):
        """Connect to VoiceMeeter and restore device/param settings."""
        # Restore hardware device assignments
        if VM_DEVICES_PATH.exists():
            try:
                with open(VM_DEVICES_PATH) as f:
                    devs = json.load(f)
                for key, name in devs.items():
                    ok = self.vm.set_string_param(f"{key}.device.wdm", name)
                    logging.info("Set %s.device.wdm = %s (%s)", key, name,
                                 "OK" if ok else "FAIL")
                if devs:
                    time.sleep(1)
            except Exception:
                logging.exception("Failed to restore device assignments")

        # Restore mixer settings
        VM_DEFAULTS = {
            "Strip[3].Gain": 0.0, "Bus[3].Gain": 0.0, "Bus[4].Gain": 0.0,
            "Strip[0].Gain": 0.0,
            "Strip[3].eqgain1": 0.0, "Strip[3].eqgain2": 0.0,
            "Strip[3].eqgain3": 0.0,
        }
        vm_state_path = SCRIPT_DIR / "vm_state.json"
        vm_params = dict(VM_DEFAULTS)
        if vm_state_path.exists():
            try:
                with open(vm_state_path) as f:
                    vm_params = json.load(f)
            except Exception:
                logging.exception("Failed to read vm_state.json, using defaults")

        vm_params["Strip[0].B1"] = 1.0
        vm_params["Strip[3].B2"] = 1.0
        vm_params["Strip[0].A1"] = 0.0
        vm_params["Strip[3].A1"] = 0.0

        for param, value in vm_params.items():
            try:
                self.vm.set_param(param, float(value))
            except (ValueError, TypeError):
                logging.warning("Skipping invalid VM param %s=%s", param, value)
        time.sleep(0.5)
        self.vm.set_param("Strip[0].B1", 1.0)
        self.vm.set_param("Strip[3].B2", 1.0)
        logging.info("Applied %d VM params (routing + %s)",
                     len(vm_params),
                     "saved" if vm_state_path.exists() else "defaults")
        self._vm_ready.set()

    def _desired_output(self) -> AudioOutput:
        """Determine what audio output should be based on user mode."""
        if self._user_mode == UserMode.DESKTOP:
            return AudioOutput.DESKTOP
        if self._user_mode in (UserMode.VR, UserMode.SILENT_VR):
            return AudioOutput.VR
        # AUTO - follow VR state
        return AudioOutput.VR if self.detector.is_vr_running() else AudioOutput.DESKTOP

    def _desired_mic(self) -> bool:
        """Should music go through the VRChat mic (Strip[3].B1)?"""
        return self._user_mode == UserMode.VR

    def _apply(self, force: bool = False):
        """Switch audio output and mic routing if needed."""
        if not self._vm_ready.is_set():
            return
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
                self.vm.set_param("Strip[0].B1", 1.0)  # mic -> B1
                self.vm.set_param("Strip[3].B2", 1.0)  # music -> B2

            # Notify UI to refresh mode display
            self._notify_ui()

    def _notify_ui(self):
        """Thread-safe UI notification to refresh mode/VR display."""
        if self.ui and self.ui.root:
            try:
                self.ui.root.after(0, self.ui._update_mode_display)
            except Exception:
                pass

    def _reload_config_if_changed(self):
        """Re-read config.json to pick up Settings tab changes."""
        try:
            if CONFIG_PATH.exists():
                with open(CONFIG_PATH) as f:
                    fresh = json.load(f)
                if self.audio._multi_target:
                    new_exclude = fresh.get("exclude_processes", [])
                    self.audio._user_exclude = {p.lower() for p in new_exclude}
        except Exception:
            pass

    def _enforce_loop(self):
        """Periodically re-apply audio and monitor VoiceMeeter health."""
        while not self._stop.is_set():
            self._stop.wait(ENFORCE_INTERVAL)
            if self._stop.is_set():
                break
            try:
                self._reload_config_if_changed()
                self._apply(force=True)
                self._check_voicemeeter_health()
            except Exception:
                logging.exception("Enforce error")

    def _check_voicemeeter_health(self):
        """Detect if VoiceMeeter closed during session and prompt user."""
        if not self._session_active or self._vm_dialog_shown:
            return
        from vm_path import is_vm_process
        vm_running = any(
            is_vm_process(p.info.get("name", ""))
            for p in psutil.process_iter(["name"])
        )
        if vm_running:
            return
        logging.warning("VoiceMeeter stopped during active session")
        self._vm_dialog_shown = True
        if self.ui and self.ui.root:
            try:
                self.ui.root.after(0, self.ui._show_vm_closed_dialog)
            except Exception:
                pass

    def restart_voicemeeter(self):
        """Restart VoiceMeeter and restore all settings."""
        from vm_path import find_exe
        vm_exe = find_exe()
        if not vm_exe:
            logging.error("Cannot restart VoiceMeeter — exe not found")
            return
        logging.info("Restarting VoiceMeeter...")
        # Clean up stale DLL state before reconnecting
        self.vm.close()
        time.sleep(1)
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        si.wShowWindow = 7  # SW_SHOWMINNOACTIVE
        subprocess.Popen([str(vm_exe)], startupinfo=si)
        time.sleep(6)  # VM needs time to fully init its API
        self._minimize_voicemeeter()
        # Retry login a few times — VM API may not be ready immediately
        for attempt in range(5):
            if self.vm._ensure_connected():
                break
            time.sleep(1)
        # Restore devices
        if VM_DEVICES_PATH.exists():
            try:
                with open(VM_DEVICES_PATH) as f:
                    devs = json.load(f)
                for key, name in devs.items():
                    self.vm.set_string_param(f"{key}.device.wdm", name)
            except Exception:
                logging.exception("Device restore after VM restart failed")
        # Restore gains
        vm_state_path = SCRIPT_DIR / "vm_state.json"
        if vm_state_path.exists():
            try:
                with open(vm_state_path) as f:
                    params = json.load(f)
                for k, v in params.items():
                    self.vm.set_param(k, float(v))
            except Exception:
                logging.exception("Gain restore after VM restart failed")
        self._vm_dialog_shown = False
        logging.info("VoiceMeeter restarted and restored")

    def _on_vr_change(self, running: bool):
        """Called by detector when VR starts/stops."""
        if self._in_cleanup:
            return  # ignore transitions during session teardown
        if running and not self._session_active:
            logging.info("VR detected — signaling session start")
            self._vr_start_event.set()
        elif not running and self._session_active:
            logging.info("VR stopped — closing session")
            if self.ui and self.ui.root:
                try:
                    self.ui.root.after(0, self.ui._force_close)
                except Exception:
                    pass

    def _write_state(self):
        """Write current mode to state.json for persistence."""
        try:
            state = {}
            if STATE_PATH.exists():
                try:
                    with open(STATE_PATH) as f:
                        state = json.load(f)
                except Exception:
                    pass
            state["current_mode"] = self._user_mode.name
            state["vr_active"] = self.detector.is_vr_running()
            with open(STATE_PATH, "w") as f:
                json.dump(state, f)
        except Exception:
            logging.exception("Failed to write state file")

    # ------------------------------------------------------------------
    # Public methods for UI
    # ------------------------------------------------------------------
    def set_user_mode(self, mode_name: str):
        """Set mode by name string. Called by UI buttons and presets."""
        mode_map = {m.name: m for m in UserMode}
        if mode_name in mode_map:
            self._user_mode = mode_map[mode_name]
            logging.info("User mode -> %s", mode_name)
            self._write_state()
            threading.Thread(target=self._apply, daemon=True).start()

    def get_mode_name(self) -> str:
        return self._user_mode.name

    def get_output_name(self) -> str:
        return self._current_output.name if self._current_output else "?"

    def close_steamvr(self):
        """Force-close SteamVR and wait for it to actually die."""
        logging.info("Closing SteamVR...")
        for proc_name in ("vrmonitor.exe", "vrserver.exe", "vrcompositor.exe"):
            try:
                subprocess.run(
                    ["taskkill", "/F", "/IM", proc_name],
                    capture_output=True, timeout=10,
                    creationflags=subprocess.CREATE_NO_WINDOW,
                )
            except Exception:
                pass
        # Wait up to 15 seconds for SteamVR to actually stop
        for i in range(15):
            if not self.detector.is_vr_running():
                logging.info("SteamVR stopped after %ds", i)
                return
            time.sleep(1)
        logging.warning("SteamVR still running after 15s wait")

    # ------------------------------------------------------------------
    # Session lifecycle
    # ------------------------------------------------------------------
    def run(self):
        """Main loop: watch for SteamVR, run sessions, repeat."""
        self.detector.start()

        while not self._user_quit:
            logging.info("Waiting for SteamVR...")
            self._vr_start_event.clear()

            # Check if VR is already running right now
            if self.detector.is_vr_running():
                self._vr_start_event.set()

            # Block until VR starts (or app quits)
            while not self._user_quit and not self._vr_start_event.is_set():
                self._vr_start_event.wait(timeout=5)

            if self._user_quit:
                break

            # Brief debounce — make sure VR is actually stable
            time.sleep(1)
            if not self.detector.is_vr_running():
                continue

            logging.info("SteamVR detected — starting session")
            self._start_vr_session()
            self.ui.run()  # tkinter mainloop blocks until window closes
            self._end_vr_session()

            # Post-session cooldown — avoid re-detecting a dying SteamVR
            time.sleep(5)

        self.detector.stop()

    def _start_vr_session(self):
        """Boot VoiceMeeter, init audio, open UI."""
        self._session_active = True
        self._stop.clear()
        self._vm_ready.clear()
        self._confirmed = False
        self._vm_dialog_shown = False
        self._user_mode = UserMode.AUTO
        self._current_output = None

        # Show boot splash
        splash_script = SCRIPT_DIR / "splash.py"
        if splash_script.exists():
            try:
                subprocess.Popen(
                    [sys.executable, str(splash_script)],
                    creationflags=subprocess.CREATE_NO_WINDOW,
                )
            except Exception:
                pass

        # Launch VoiceMeeter if not already running
        _splash_update("Starting VoiceMeeter...")
        from vm_path import is_vm_process, find_exe
        vm_running = any(
            is_vm_process(p.info.get("name", ""))
            for p in psutil.process_iter(["name"])
        )
        vm_just_launched = False
        if not vm_running:
            vm_exe = find_exe()
            if vm_exe:
                logging.info("Launching VoiceMeeter (%s)...", vm_exe.name)
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                si.wShowWindow = 7  # SW_SHOWMINNOACTIVE
                subprocess.Popen([str(vm_exe)], startupinfo=si)
                vm_just_launched = True
            else:
                logging.warning("VoiceMeeter not found — install it first")

        # Initialize VoiceMeeter API
        _splash_update("Initializing audio...")
        if vm_just_launched:
            time.sleep(4)
            self._minimize_voicemeeter()
        self._init_voicemeeter()
        self._apply()

        # Start enforce loop for this session
        threading.Thread(target=self._enforce_loop, daemon=True).start()

        # First launch check
        first_launch = not self.config.get("first_launch_done", False)
        if first_launch:
            self.config["first_launch_done"] = True
            try:
                with open(CONFIG_PATH, "w") as f:
                    json.dump(self.config, f, indent=2)
            except Exception:
                pass

        # Create the UI
        _splash_update("Opening mixer...")
        from mixer import MixerApp
        initial_tab = "guide" if first_launch else "mixer"
        self.ui = MixerApp(self, initial_tab=initial_tab)

        # Dismiss boot splash — UI is ready
        try:
            SPLASH_DONE.touch()
        except OSError:
            pass

    def _end_vr_session(self):
        """Shut down VoiceMeeter, restore audio, clean up."""
        self._in_cleanup = True
        self._session_active = False
        self._stop.set()  # stop enforce loop

        # Show shutdown splash
        splash_script = SCRIPT_DIR / "splash.py"
        if splash_script.exists():
            try:
                subprocess.Popen(
                    [sys.executable, str(splash_script), "Shutting down..."],
                    creationflags=subprocess.CREATE_NO_WINDOW,
                )
            except Exception:
                pass

        # Wait for any background _apply thread to finish
        with self._lock:
            pass

        # Restore desktop audio before VM shuts down
        _splash_update("Restoring audio...")
        self.audio.switch_to(AudioOutput.DESKTOP)
        time.sleep(0.5)
        self.audio.switch_to(AudioOutput.DESKTOP)  # double-tap

        # Save hardware device assignments
        _splash_update("Saving settings...")
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

        # Shut down VoiceMeeter
        _splash_update("Closing VoiceMeeter...")
        self.vm.shutdown()
        time.sleep(2)
        self.vm.close()
        self.vm._logged_in = False

        # Wait for SteamVR to fully stop (in case close_steamvr was called)
        _splash_update("Waiting for SteamVR...")
        for i in range(20):
            if not self.detector.is_vr_running():
                break
            if i == 10:
                # Second attempt to force-kill if still alive
                logging.warning("SteamVR still running after 10s, force-killing...")
                for proc_name in ("vrserver.exe", "vrmonitor.exe"):
                    try:
                        subprocess.run(
                            ["taskkill", "/F", "/IM", proc_name],
                            capture_output=True, timeout=5,
                            creationflags=subprocess.CREATE_NO_WINDOW,
                        )
                    except Exception:
                        pass
            time.sleep(1)

        # Dismiss shutdown splash — all cleanup complete
        try:
            SPLASH_DONE.touch()
        except OSError:
            pass

        self.ui = None
        self._in_cleanup = False
        self.detector._vr_running = None  # reset for next detection cycle
        logging.info("VR session ended — back to watching")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    mutex = acquire_single_instance()
    if mutex is None:
        ctypes.windll.user32.MessageBoxW(
            0,
            "Already running! Check your taskbar.",
            "VR Audio Switcher", 0x40)  # MB_ICONINFORMATION
        sys.exit(0)

    def _show_error_and_exit(msg):
        ctypes.windll.user32.MessageBoxW(
            0, msg, "VR Audio Switcher", 0x10)  # MB_ICONERROR
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

    # Boot splash — instant feedback on launch
    splash_script = SCRIPT_DIR / "splash.py"
    if splash_script.exists():
        try:
            subprocess.Popen(
                [sys.executable, str(splash_script)],
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
        except Exception:
            pass

    # Rotate log at 1 MB, keep 1 backup
    file_handler = logging.handlers.RotatingFileHandler(
        LOG_PATH, maxBytes=1_000_000, backupCount=1,
    )
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[file_handler, logging.StreamHandler()],
    )

    # Forced auto-update on boot
    _splash_update("Checking for updates...")
    try:
        from updater import update_available, do_update, restart_app
        avail, local_ver, remote_ver = update_available()
        if avail:
            _splash_update(f"Downloading update v{remote_ver}...")
            logging.info("Update available: v%s -> v%s, auto-updating...",
                         local_ver, remote_ver)
            ok, msg = do_update()
            if ok:
                logging.info("Update applied, restarting...")
                try:
                    SPLASH_DONE.touch()
                except OSError:
                    pass
                if mutex:
                    ctypes.windll.kernel32.CloseHandle(mutex)
                restart_app()  # does not return
            else:
                logging.warning("Auto-update failed: %s — continuing with "
                                "current version", msg)
    except Exception:
        logging.debug("Auto-update check failed", exc_info=True)

    # Dismiss boot splash — now watching silently
    _splash_update("Ready! Waiting for SteamVR...")
    time.sleep(1.5)
    try:
        SPLASH_DONE.touch()
    except OSError:
        pass

    logging.info("VR Audio Switcher starting (watching for SteamVR)...")
    app = VRAudioSwitcher(config)
    app.run()


if __name__ == "__main__":
    main()
