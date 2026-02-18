"""VR Audio Switcher — Auto-updater.

Checks GitHub for a newer VERSION on startup. If an update is available,
prompts the user. Uses the GitHub API (no CDN cache delay). Also provides
a background check that can be called from the tray menu.
"""

import json
import logging
import subprocess
import sys
import threading
import time
import urllib.request
import urllib.error
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.resolve()
VERSION_PATH = SCRIPT_DIR / "VERSION"
REPO = "Aetheriju/vr-audio-switcher"
API_URL = f"https://api.github.com/repos/{REPO}/contents/VERSION"
REQUIREMENTS = SCRIPT_DIR / "requirements.txt"
CHECK_TIMEOUT = 5  # seconds — don't hang startup on slow network

log = logging.getLogger(__name__)


def local_version() -> str:
    """Read the local VERSION file."""
    try:
        return VERSION_PATH.read_text(encoding="utf-8").strip()
    except Exception:
        return "0.0.0"


def remote_version() -> str | None:
    """Fetch the latest VERSION from GitHub API. Returns None on failure."""
    try:
        req = urllib.request.Request(
            API_URL,
            headers={
                "User-Agent": "vr-audio-switcher",
                "Accept": "application/vnd.github.v3+json",
            },
        )
        with urllib.request.urlopen(req, timeout=CHECK_TIMEOUT) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        # GitHub API returns base64-encoded content
        if "content" not in data:
            return None
        import base64
        content = base64.b64decode(data["content"]).decode("utf-8").strip()
        return content
    except Exception as e:
        log.debug("Update check failed: %s", e)
        return None


def _parse_version(v: str) -> tuple[int, ...]:
    """Parse '1.2.3' into (1, 2, 3) for comparison."""
    try:
        return tuple(int(x) for x in v.split("."))
    except (ValueError, AttributeError):
        return (0, 0, 0)


def update_available() -> tuple[bool, str, str]:
    """Check if an update is available.

    Returns (is_available, local_ver, remote_ver).
    """
    local = local_version()
    remote = remote_version()
    if remote is None:
        return False, local, ""
    return _parse_version(remote) > _parse_version(local), local, remote


def do_update() -> tuple[bool, str]:
    """Pull latest code and update pip packages.

    Returns (success, message).
    """
    errors = []

    # Try git pull first (if this is a git repo)
    git_dir = SCRIPT_DIR / ".git"
    if git_dir.exists():
        try:
            result = subprocess.run(
                ["git", "pull", "--ff-only"],
                cwd=str(SCRIPT_DIR),
                capture_output=True, text=True, timeout=30,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            if result.returncode == 0:
                log.info("git pull: %s", result.stdout.strip())
            else:
                errors.append(f"git pull failed: {result.stderr.strip()}")
        except Exception as e:
            errors.append(f"git pull error: {e}")
    else:
        # Not a git repo — download and extract ZIP
        try:
            import zipfile
            import io

            zip_url = f"https://github.com/{REPO}/archive/refs/heads/main.zip"
            req = urllib.request.Request(
                zip_url, headers={"User-Agent": "vr-audio-switcher"})
            with urllib.request.urlopen(req, timeout=30) as resp:
                data = resp.read()

            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                prefix = "vr-audio-switcher-main/"
                for info in zf.infolist():
                    if info.is_dir():
                        continue
                    rel = info.filename
                    if rel.startswith(prefix):
                        rel = rel[len(prefix):]
                    if not rel:
                        continue
                    # Only update code files, not user configs
                    if rel.endswith((".py", ".bat", ".txt", ".md")) \
                            or rel in ("VERSION", ".gitignore", "LICENSE"):
                        dest = SCRIPT_DIR / rel
                        dest.parent.mkdir(parents=True, exist_ok=True)
                        dest.write_bytes(zf.read(info.filename))

            log.info("Downloaded and extracted latest code from GitHub")
        except Exception as e:
            errors.append(f"ZIP download failed: {e}")

    # Update pip packages
    if REQUIREMENTS.exists():
        try:
            result = subprocess.run(
                [sys.executable, "-m", "pip", "install", "-r",
                 str(REQUIREMENTS), "--upgrade", "--quiet"],
                capture_output=True, text=True, timeout=120,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            if result.returncode == 0:
                log.info("pip packages updated")
            else:
                errors.append(f"pip update: {result.stderr[:200]}")
        except Exception as e:
            errors.append(f"pip error: {e}")

    if errors:
        return False, "; ".join(errors)
    return True, "Update complete"


def restart_app():
    """Restart the current script by launching a new process and exiting."""
    python = sys.executable
    subprocess.Popen([python] + sys.argv,
                     creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    # Give the new process time to start before we release the mutex
    time.sleep(1)
    sys.exit(0)


def check_and_prompt() -> bool:
    """Check for updates and prompt user via tkinter dialog.

    Returns True if the app should continue booting, False if it should exit.
    Called early in main() before the tray icon is created.
    """
    is_available, local, remote = update_available()
    if not is_available:
        return True

    log.info("Update available: %s -> %s", local, remote)

    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()

        choice = messagebox.askyesnocancel(
            "VR Audio Switcher \u2014 Update Available",
            f"A new version is available: v{remote}\n"
            f"You have: v{local}\n\n"
            "Yes = Update now (recommended)\n"
            "No = Skip this time\n"
            "Cancel = Quit",
        )
        root.destroy()

        if choice is None:
            return False
        elif choice:
            ok, msg = do_update()
            if ok:
                restart_app()
                return False
            else:
                root2 = tk.Tk()
                root2.withdraw()
                messagebox.showwarning(
                    "Update Failed",
                    f"Couldn't update automatically:\n{msg}\n\n"
                    "The app will start with the current version.",
                )
                root2.destroy()
                return True
        else:
            return True

    except Exception as e:
        log.warning("Update prompt failed: %s", e)
        return True


def check_in_background(notify_callback=None):
    """Run update check on a background thread.

    If an update is found, calls notify_callback(local, remote) on the
    calling thread (if provided). This is used by the tray app to show
    a notification without blocking.
    """
    def _check():
        is_available, local, remote = update_available()
        if is_available and notify_callback:
            notify_callback(local, remote)

    threading.Thread(target=_check, daemon=True).start()
