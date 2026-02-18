"""VR Audio Auto-Switcher — detects SteamVR and toggles Chrome's audio output."""

import csv
import ctypes
import json
import logging
import logging.handlers
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

    DLL_PATH = r"C:\Program Files (x86)\VB\Voicemeeter\VoicemeeterRemote64.dll"

    def __init__(self):
        self._dll = None
        self._logged_in = False

    def _ensure_connected(self) -> bool:
        if self._logged_in:
            return True
        try:
            if self._dll is None:
                self._dll = ctypes.WinDLL(self.DLL_PATH)
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
    # Device ID prefixes to skip when auto-detecting the desktop device.
    # These are virtual devices or built-in monitor speakers — not what
    # the user considers "desktop speakers."
    DEFAULT_EXCLUDE = [
        "Steam Streaming",
        "VB-Audio Voicemeeter",
        "NVIDIA High Definition Audio",
        "HD Audio Driver for Display Audio",
        "Realtek",
    ]

    def __init__(self, config: dict):
        self.svcl_path = str(SCRIPT_DIR / config["svcl_path"])
        self.target = config["target_process"]
        self.desktop_device = config["desktop_device"]  # "auto" or a specific ID
        self.vr_device = config["vr_device"]
        self.desktop_exclude = config.get("desktop_exclude", self.DEFAULT_EXCLUDE)

        if not Path(self.svcl_path).exists():
            raise FileNotFoundError(f"svcl.exe not found at {self.svcl_path}")

    def _find_desktop_device(self) -> str | None:
        """Find a real physical audio device, skipping virtual/internal ones."""
        tmp = SCRIPT_DIR / "_default_query.csv"
        try:
            subprocess.run(
                [self.svcl_path, "/scomma", str(tmp),
                 "/Columns", "Command-Line Friendly ID,Direction,Type,Device State"],
                capture_output=True, timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            with open(tmp, newline="", encoding="utf-8-sig") as f:
                for row in csv.DictReader(f):
                    if (row.get("Direction", "").strip() != "Render"
                            or row.get("Type", "").strip() != "Device"
                            or row.get("Device State", "").strip() != "Active"):
                        continue
                    fid = row["Command-Line Friendly ID"].strip()
                    if any(ex in fid for ex in self.desktop_exclude):
                        continue
                    return fid  # first real hardware device
        except Exception:
            logging.exception("Failed to query desktop device")
        finally:
            try:
                tmp.unlink(missing_ok=True)
            except OSError:
                pass
        return None

    def switch_to(self, output: AudioOutput) -> bool:
        """Set Chrome's audio output. Returns True if svcl matched 1+ items."""
        if not is_process_running(self.target):
            return False  # Chrome not running — nothing to switch

        if output == AudioOutput.VR:
            device = self.vr_device
        elif self.desktop_device.lower() == "auto":
            device = self._find_desktop_device()
            if not device:
                return False  # no real desktop device found
        else:
            device = self.desktop_device

        cmd = [self.svcl_path, "/Stdout", "/SetAppDefault", device, "all", self.target]
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            stdout = result.stdout.strip()
            if "1 item" in stdout or "items found" in stdout:
                return True
            # 0 items — device not available
            return False
        except FileNotFoundError:
            logging.error("svcl.exe not found at %s", self.svcl_path)
            return False
        except subprocess.TimeoutExpired:
            logging.error("svcl.exe timed out")
            return False


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
        self._confirmed = False   # True when svcl actually matched Chrome
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
        """Periodically re-apply audio to catch Chrome restarts, new tabs, etc."""
        while not self._stop.is_set():
            self._stop.wait(ENFORCE_INTERVAL)
            if self._stop.is_set():
                break
            try:
                self._check_mode_request()
                self._apply(force=True)
            except Exception:
                logging.exception("Enforce error")

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
                UserMode.DESKTOP: "Desktop (Soundbar)",
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
        idx = MODE_CYCLE.index(self._user_mode)
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

    def _quit(self, icon, item):
        self._stop.set()
        self.detector.stop()
        # Restore desktop audio and mic routing before exit
        self.audio.switch_to(AudioOutput.DESKTOP)
        self.vm.set_strip_b1(self.music_strip, True)
        # Save hardware device assignments before shutdown
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
        # Kill the mixer if it's running
        for proc in psutil.process_iter(["name", "cmdline"]):
            try:
                if proc.info["name"] and proc.info["name"].lower() in ("python.exe", "pythonw.exe"):
                    cmdline = proc.info.get("cmdline") or []
                    if any("mixer.py" in arg for arg in cmdline):
                        proc.kill()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        # Gracefully shut down VoiceMeeter (saves settings), then force-kill if needed
        self.vm.shutdown()
        self.vm.close()
        time.sleep(2)
        for proc in psutil.process_iter(["name"]):
            if proc.info["name"] and proc.info["name"].lower() == "voicemeeterpro.exe":
                proc.kill()
                break
        icon.stop()

    def run(self):
        menu = pystray.Menu(
            # Hidden default item — left-click cycles modes
            pystray.MenuItem("Cycle", self._cycle_mode, default=True, visible=False),
            pystray.MenuItem("Desktop (Soundbar)", self._set_mode(UserMode.DESKTOP),
                             checked=self._is_mode(UserMode.DESKTOP), radio=True),
            pystray.MenuItem("Auto-detect", self._set_mode(UserMode.AUTO),
                             checked=self._is_mode(UserMode.AUTO), radio=True),
            pystray.MenuItem("Private", self._set_mode(UserMode.SILENT_VR),
                             checked=self._is_mode(UserMode.SILENT_VR), radio=True),
            pystray.MenuItem("Public", self._set_mode(UserMode.VR),
                             checked=self._is_mode(UserMode.VR), radio=True),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Mixer", self._open_mixer),
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
                threading.Thread(target=self._enforce_loop, daemon=True).start()
            except Exception:
                logging.exception("Setup error")

        self.icon.run(setup=on_setup)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    mutex = acquire_single_instance()
    if mutex is None:
        sys.exit(0)

    if not CONFIG_PATH.exists():
        print(f"Config not found at {CONFIG_PATH}. Creating default...")
        default = {
            "poll_interval_seconds": 3,
            "steamvr_process": "vrserver.exe",
            "target_process": "chrome.exe",
            "svcl_path": "svcl.exe",
            "desktop_device": "",
            "vr_device": "",
            "debounce_seconds": 5,
            "music_strip": 3,
        }
        CONFIG_PATH.write_text(json.dumps(default, indent=2))
        print("Edit desktop_device and vr_device in config.json, then rerun.")
        sys.exit(1)

    with open(CONFIG_PATH) as f:
        config = json.load(f)

    if not config.get("vr_device"):
        print("vr_device not set in config.json.")
        sys.exit(1)
    if not config.get("desktop_device"):
        print("desktop_device not set in config.json.")
        sys.exit(1)

    # Rotate log at 1 MB, keep 1 backup
    file_handler = logging.handlers.RotatingFileHandler(
        LOG_PATH, maxBytes=1_000_000, backupCount=1,
    )
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[file_handler, logging.StreamHandler()],
    )

    # Ensure VoiceMeeter Banana is running before we start
    VM_EXE = r"C:\Program Files (x86)\VB\Voicemeeter\voicemeeterpro.exe"
    vm_running = any(p.name().lower() == "voicemeeterpro.exe"
                     for p in psutil.process_iter(["name"]))
    if not vm_running:
        logging.info("Launching VoiceMeeter Banana...")
        subprocess.Popen([VM_EXE], creationflags=subprocess.CREATE_NO_WINDOW)
        time.sleep(4)  # VoiceMeeter needs time to fully initialize

    logging.info("VR Audio Switcher starting...")
    app = VRAudioSwitcher(config)

    # Restore hardware device assignments (mic input, bus outputs)
    VM_DEVICE_DEFAULTS = {
        "Strip[0]": "Microphone (Steam Streaming Microphone)",
    }
    vm_devices = dict(VM_DEVICE_DEFAULTS)
    if VM_DEVICES_PATH.exists():
        try:
            with open(VM_DEVICES_PATH) as f:
                saved = json.load(f)
            vm_devices.update(saved)
        except Exception:
            logging.exception("Failed to read vm_devices.json, using defaults")

    for key, name in vm_devices.items():
        ok = app.vm.set_string_param(f"{key}.device.wdm", name)
        logging.info("Set %s.device.wdm = %s (%s)", key, name, "OK" if ok else "FAIL")
    time.sleep(1)  # let VoiceMeeter process device changes

    # Restore mixer settings into VoiceMeeter (last saved, or sweet-spot defaults)
    VM_DEFAULTS = {
        "Strip[3].Gain": -31.0, "Bus[3].Gain": -1.0, "Bus[4].Gain": 1.0,
        "Strip[0].Gain": 4.0,
        "Strip[3].eqgain1": 12.0, "Strip[3].eqgain2": 12.0, "Strip[3].eqgain3": -8.0,
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
        app.vm.set_param(param, float(value))
    time.sleep(0.5)  # let VoiceMeeter process params before _apply()
    # Re-apply routing flags a second time (VoiceMeeter can overwrite on startup)
    app.vm.set_param("Strip[0].B1", 1.0)
    app.vm.set_param("Strip[3].B2", 1.0)
    logging.info("Applied %d VM params (routing + %s)",
                 len(vm_params), "saved" if vm_state_path.exists() else "defaults")

    app.run()


if __name__ == "__main__":
    main()
