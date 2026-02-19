"""Lightweight loading splash — appears instantly on boot for user feedback.

Usage:
    python splash.py                  # shows "Starting up..."
    python splash.py "Shutting down..."  # shows custom initial text
"""

import sys
import tkinter as tk
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.resolve()
DONE_SIGNAL = SCRIPT_DIR / "_splash_done"
STATUS_FILE = SCRIPT_DIR / "_splash_status"


def main():
    # Clean stale signals from a previous crash
    for f in (DONE_SIGNAL, STATUS_FILE):
        try:
            f.unlink(missing_ok=True)
        except OSError:
            pass

    root = tk.Tk()
    root.overrideredirect(True)
    root.configure(bg="#1e1e1e")
    root.attributes("-topmost", True)

    # Thin green accent border
    border = tk.Frame(root, bg="#4caf50", padx=1, pady=1)
    border.pack(fill="both", expand=True)
    inner = tk.Frame(border, bg="#1e1e1e")
    inner.pack(fill="both", expand=True)

    tk.Label(inner, text="VR Audio Switcher", bg="#1e1e1e", fg="#4caf50",
             font=("Segoe UI", 14, "bold")).pack(pady=(20, 5))
    initial_text = sys.argv[1] if len(sys.argv) > 1 else "Starting up..."
    status_lbl = tk.Label(inner, text=initial_text, bg="#1e1e1e",
                          fg="#888888", font=("Segoe UI", 10))
    status_lbl.pack(pady=(0, 20))

    # Size and center
    w, h = 320, 110
    x = (root.winfo_screenwidth() - w) // 2
    y = (root.winfo_screenheight() - h) // 2
    root.geometry(f"{w}x{h}+{x}+{y}")

    def poll():
        # Check if main process signaled us to close
        if DONE_SIGNAL.exists():
            for f in (DONE_SIGNAL, STATUS_FILE):
                try:
                    f.unlink(missing_ok=True)
                except OSError:
                    pass
            root.destroy()
            return
        # Update status text if main process wrote a new message
        try:
            if STATUS_FILE.exists():
                text = STATUS_FILE.read_text(encoding="utf-8").strip()
                if text:
                    status_lbl.config(text=text)
        except Exception:
            pass
        root.after(200, poll)

    root.after(30000, root.destroy)  # safety net — auto-close after 30s
    root.after(200, poll)
    root.mainloop()


if __name__ == "__main__":
    main()
