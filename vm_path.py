"""VoiceMeeter path detection — finds DLL and EXE regardless of install location.

Searches the Windows registry, then falls back to common paths.
Supports VoiceMeeter Basic, Banana, and Potato.
"""

import winreg
from pathlib import Path

# Process names for each VoiceMeeter variant (in priority order)
VM_PROCESS_NAMES = [
    "voicemeeterpro.exe",   # Banana
    "voicemeeter8x64.exe",  # Potato (64-bit)
    "voicemeeter8.exe",     # Potato (32-bit)
    "voicemeeter.exe",      # Basic
]

# Default install paths (fallback if registry fails)
_DEFAULT_DIRS = [
    Path(r"C:\Program Files (x86)\VB\Voicemeeter"),
    Path(r"C:\Program Files\VB\Voicemeeter"),
]

DLL_NAME = "VoicemeeterRemote64.dll"


def _find_from_registry() -> Path | None:
    """Try to find VoiceMeeter install dir from the Windows registry."""
    # VoiceMeeter registers its install path under Uninstall keys
    uninstall_key = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(hive, uninstall_key) as key:
                i = 0
                while True:
                    try:
                        subkey_name = winreg.EnumKey(key, i)
                        i += 1
                        if "voicemeeter" not in subkey_name.lower() \
                                and "VB:" not in subkey_name:
                            continue
                        with winreg.OpenKey(key, subkey_name) as subkey:
                            try:
                                loc, _ = winreg.QueryValueEx(
                                    subkey, "UninstallString")
                                # UninstallString is like
                                # "C:\Program Files (x86)\VB\Voicemeeter\uninst..."
                                p = Path(loc).parent
                                if (p / DLL_NAME).exists():
                                    return p
                            except OSError:
                                pass
                    except OSError:
                        break
        except OSError:
            continue
    return None


def find_dll() -> Path | None:
    """Find VoicemeeterRemote64.dll."""
    # Try registry first
    reg_dir = _find_from_registry()
    if reg_dir:
        dll = reg_dir / DLL_NAME
        if dll.exists():
            return dll

    # Fallback to default paths
    for d in _DEFAULT_DIRS:
        dll = d / DLL_NAME
        if dll.exists():
            return dll

    return None


def find_exe() -> Path | None:
    """Find the VoiceMeeter executable (Banana > Potato > Basic)."""
    # Try registry first
    reg_dir = _find_from_registry()
    dirs = ([reg_dir] if reg_dir else []) + _DEFAULT_DIRS

    for d in dirs:
        for name in VM_PROCESS_NAMES:
            exe = d / name
            if exe.exists():
                return exe

    return None


def is_vm_process(process_name: str) -> bool:
    """Check if a process name is any VoiceMeeter variant."""
    if not process_name:
        return False
    return process_name.lower() in VM_PROCESS_NAMES
