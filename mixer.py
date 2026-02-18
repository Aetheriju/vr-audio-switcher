"""VR Audio Mixer — percentage-based controls centered on your sweet spot."""

import ctypes
import json
import math
import sys
import time
import tkinter as tk
from tkinter import ttk, simpledialog
from pathlib import Path

from vm_path import find_dll
MUTEX_NAME = "Global\\VRAudioMixerMutex"
PRESETS_PATH = Path(__file__).parent.resolve() / "presets.json"
STATE_PATH = Path(__file__).parent.resolve() / "state.json"
VM_STATE_PATH = Path(__file__).parent.resolve() / "vm_state.json"

# Sweet-spot defaults — the "0%" center for each slider (in dB)
DEFAULTS = {
    "others": -32.0,
    "me":     -30.0,
    "voice":    3.0,
    "bass":    12.0,
    "mid":     12.0,
    "treble":  -8.0,
}

# dB floor for each parameter (-100% maps here)
FLOORS = {
    "others": -60.0, "me": -60.0, "voice": -60.0,
    "bass": -12.0, "mid": -12.0, "treble": -12.0,
}

# dB ceiling (hardware max)
CEILS = {
    "others": 12.0, "me": 12.0, "voice": 12.0,
    "bass": 12.0, "mid": 12.0, "treble": 12.0,
}

MODE_COLORS = {
    "DESKTOP": "#42a5f5", "AUTO": "#4caf50",
    "VR": "#f44336", "SILENT_VR": "#fdd835", None: "#555555",
}
MODE_LABELS = {
    "DESKTOP": "Desktop", "AUTO": "Auto",
    "VR": "Public", "SILENT_VR": "Private", None: "Any",
}

DEFAULT_PRESETS = {
    "Default": {
        "others": 0, "me": 0, "voice": 0,
        "bass": 100, "mid": 100, "treble": 100, "mode": None,
    },
    "Quiet for Others": {
        "others": -50, "me": 0, "voice": 0,
        "bass": 100, "mid": 100, "treble": 100, "mode": None,
    },
    "Solo Listen": {
        "others": -100, "me": 0, "voice": 25,
        "bass": 100, "mid": 100, "treble": 100, "mode": "SILENT_VR",
    },
}


# ---------------------------------------------------------------------------
# dB  <->  percentage conversion
# ---------------------------------------------------------------------------
EQ_KEYS = {"bass", "mid", "treble"}


def pct_to_db(pct, key):
    """Convert percentage to dB.

    Volume keys (-100..+100): 0% = sweet-spot, +100% = 2x amplitude, -100% = silent.
    EQ keys    (0..100):      100% = sweet-spot, 0% = floor (-12 dB).
    """
    default, floor, ceil = DEFAULTS[key], FLOORS[key], CEILS[key]
    if key in EQ_KEYS:
        # Linear 0..100 → floor..default
        return floor + (max(0, min(100, pct)) / 100.0) * (default - floor)
    # Volume: amplitude scaling
    if pct <= -100:
        return floor
    amp = 1.0 + pct / 100.0
    if amp <= 0:
        return floor
    db = default + 20.0 * math.log10(amp)
    return max(floor, min(ceil, db))


def db_to_pct(db, key):
    """Convert dB back to percentage."""
    default, floor = DEFAULTS[key], FLOORS[key]
    if key in EQ_KEYS:
        # Linear floor..default → 0..100
        span = default - floor
        if span == 0:
            return 100.0
        return max(0.0, min(100.0, (db - floor) / span * 100.0))
    # Volume: amplitude scaling
    if db <= floor:
        return -100.0
    amp = 10.0 ** ((db - default) / 20.0)
    return max(-100.0, min(100.0, (amp - 1.0) * 100.0))


# ---------------------------------------------------------------------------
# State file — communication with tray app
# ---------------------------------------------------------------------------
def read_current_mode():
    if STATE_PATH.exists():
        try:
            with open(STATE_PATH) as f:
                return json.load(f).get("current_mode")
        except Exception:
            pass
    return None


def request_mode(mode):
    state = {}
    if STATE_PATH.exists():
        try:
            with open(STATE_PATH) as f:
                state = json.load(f)
        except Exception:
            pass
    state["requested_mode"] = mode
    with open(STATE_PATH, "w") as f:
        json.dump(state, f)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def acquire_mutex():
    kernel32 = ctypes.windll.kernel32
    mutex = kernel32.CreateMutexW(None, False, MUTEX_NAME)
    if kernel32.GetLastError() == 183:
        return None
    return mutex


def load_presets():
    if PRESETS_PATH.exists():
        try:
            with open(PRESETS_PATH) as f:
                return json.load(f)
        except Exception:
            pass
    return dict(DEFAULT_PRESETS)


def save_presets(presets):
    with open(PRESETS_PATH, "w") as f:
        json.dump(presets, f, indent=2)


# ---------------------------------------------------------------------------
# VoiceMeeter
# ---------------------------------------------------------------------------
class VM:
    def __init__(self):
        dll_path = find_dll()
        if not dll_path:
            raise RuntimeError("VoiceMeeter DLL not found — is VoiceMeeter installed?")
        self.dll = ctypes.WinDLL(str(dll_path))
        ret = self.dll.VBVMR_Login()
        if ret not in (0, 1):
            raise RuntimeError(f"VoiceMeeter Login failed ({ret})")
        time.sleep(0.2)

    def get(self, param):
        self.dll.VBVMR_IsParametersDirty()
        buf = ctypes.c_float()
        ret = self.dll.VBVMR_GetParameterFloat(param.encode(), ctypes.byref(buf))
        if ret != 0:
            return 0.0
        return round(buf.value, 1)

    def set(self, param, value):
        self.dll.VBVMR_SetParameterFloat(param.encode(), ctypes.c_float(value))

    def close(self):
        self.dll.VBVMR_Logout()


# ---------------------------------------------------------------------------
# Mixer UI
# ---------------------------------------------------------------------------
class MixerApp:
    KEYS = ["others", "me", "voice", "bass", "mid", "treble"]

    def __init__(self, vm):
        self.vm = vm
        self.presets = load_presets()

        self.root = tk.Tk()
        self.root.title("VR Audio Mixer")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)

        # --- Read VoiceMeeter state, reconstruct conceptual dB, convert to % ---
        self._pct = self._read_vm()

        # --- Theme ---
        self.bg        = "#1e1e1e"
        self.fg        = "#e0e0e0"
        self.accent    = "#4caf50"
        self.accent2   = "#ff9800"
        self.desc_fg   = "#888888"
        self.btn_bg    = "#333333"
        self.btn_act   = "#444444"
        self.btn_danger = "#b71c1c"

        self.root.configure(bg=self.bg)
        self._setup_styles()

        main = ttk.Frame(self.root, style="D.TFrame", padding=15)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text="VR Audio Mixer", style="D.TLabel",
                  font=("Segoe UI", 14, "bold")).pack(pady=(0, 5))

        # --- Volume sliders ---
        self._sec(main, "Volume")
        self._vars, self._lbls = {}, {}
        for key, label, desc in [
            ("others", "Music for Others", "What friends hear in VRChat"),
            ("me",     "Music for Me",     "What you hear in your headset"),
            ("voice",  "My Voice",         "Your mic volume in VRChat"),
        ]:
            self._make_slider(main, key, label, desc, self.accent)

        # --- EQ sliders ---
        self._sec(main, "Music EQ")
        for key, label, desc in [
            ("bass",   "Bass",   "Low frequencies"),
            ("mid",    "Mid",    "Mid frequencies"),
            ("treble", "Treble", "High frequencies"),
        ]:
            self._make_slider(main, key, label, desc, self.accent2)

        # --- Presets ---
        self._sec(main, "Presets")
        self.preset_frame = ttk.Frame(main, style="D.TFrame")
        self.preset_frame.pack(fill="x", pady=(0, 5))
        self._rebuild_presets()

        btn_row = ttk.Frame(main, style="D.TFrame")
        btn_row.pack(fill="x")
        tk.Button(
            btn_row, text="+ Save Current", bg=self.accent, fg="#000",
            activebackground="#66bb6a", activeforeground="#000",
            relief="flat", padx=10, pady=5, font=("Segoe UI", 9, "bold"),
            cursor="hand2", command=self._save_preset,
        ).pack(side="left")
        tk.Button(
            btn_row, text="\u21bb Sync", bg=self.btn_bg, fg=self.fg,
            activebackground=self.btn_act, activeforeground=self.fg,
            relief="flat", padx=10, pady=5, font=("Segoe UI", 9),
            cursor="hand2", command=self._refresh_from_vm,
        ).pack(side="right")

        self.root.bind("<Escape>", lambda e: self.root.destroy())
        self.root.update_idletasks()
        w, h = self.root.winfo_width(), self.root.winfo_height()
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        self.root.geometry(f"+{x}+{y}")


    # ------------------------------------------------------------------
    # Styles
    # ------------------------------------------------------------------
    def _setup_styles(self):
        s = ttk.Style()
        s.theme_use("clam")
        s.configure("D.TFrame",  background=self.bg)
        s.configure("D.TLabel",  background=self.bg, foreground=self.fg,
                    font=("Segoe UI", 10))
        s.configure("Desc.TLabel", background=self.bg, foreground=self.desc_fg,
                    font=("Segoe UI", 8))
        s.configure("Sec.TLabel", background=self.bg, foreground=self.desc_fg,
                    font=("Segoe UI", 9, "bold"))
        s.configure("V.TLabel", background=self.bg, foreground=self.accent,
                    font=("Segoe UI", 11, "bold"), width=6, anchor="center")
        s.configure("VEQ.TLabel", background=self.bg, foreground=self.accent2,
                    font=("Segoe UI", 11, "bold"), width=6, anchor="center")
        s.configure("P.TLabel", background=self.bg, foreground=self.desc_fg,
                    font=("Segoe UI", 8))
        s.configure("D.Horizontal.TScale", background=self.bg,
                    troughcolor="#2d2d2d")
        s.configure("DragHL.TFrame", background="#333333")

    def _sec(self, parent, text):
        f = ttk.Frame(parent, style="D.TFrame")
        f.pack(fill="x", pady=(12, 0))
        ttk.Label(f, text=f"\u2014 {text} \u2014",
                  style="Sec.TLabel").pack(anchor="center")

    @staticmethod
    def _fmt(val, is_eq=False):
        if is_eq:
            return f"{val}%"
        return "0%" if val == 0 else f"{val:+d}%"

    # ------------------------------------------------------------------
    # Slider builder
    # ------------------------------------------------------------------
    def _make_slider(self, parent, key, label, desc, color):
        is_eq = key in EQ_KEYS
        fr = ttk.Frame(parent, style="D.TFrame")
        fr.pack(fill="x", pady=(6, 0))

        top = ttk.Frame(fr, style="D.TFrame")
        top.pack(fill="x")
        ttk.Label(top, text=label, style="D.TLabel",
                  font=("Segoe UI", 10, "bold")).pack(side="left")
        sty = "VEQ.TLabel" if color == self.accent2 else "V.TLabel"
        lbl = ttk.Label(top, text=self._fmt(self._pct[key], is_eq), style=sty)
        lbl.pack(side="right")
        self._lbls[key] = lbl

        ttk.Label(fr, text=desc, style="Desc.TLabel").pack(anchor="w", pady=(0,1))

        var = tk.IntVar(value=self._pct[key])
        self._vars[key] = var

        lo, hi = (0, 100) if is_eq else (-100, 100)
        ttk.Scale(fr, from_=lo, to=hi, orient="horizontal",
                  variable=var, style="D.Horizontal.TScale",
                  command=self._on(key)).pack(fill="x")

    def _on(self, key):
        """Return a slider callback for *key*."""
        is_eq = key in EQ_KEYS
        def cb(val):
            p = 5 * round(float(val) / 5)
            self._pct[key] = p
            self._vars[key].set(p)
            self._lbls[key].config(text=self._fmt(p, is_eq))
            self._sync_vm()
        return cb

    # ------------------------------------------------------------------
    # Read VoiceMeeter -> percentages
    # ------------------------------------------------------------------
    def _read_vm(self):
        """Read current VoiceMeeter state and convert to percentages."""
        s3   = self.vm.get("Strip[3].Gain")
        bus3 = self.vm.get("Bus[3].Gain")
        bus4 = self.vm.get("Bus[4].Gain")
        s0   = self.vm.get("Strip[0].Gain")
        raw_db = {
            "others": s3 + bus3,
            "me":     s3 + bus4,
            "voice":  s0 + bus3,
            "bass":   self.vm.get("Strip[3].eqgain1"),
            "mid":    self.vm.get("Strip[3].eqgain2"),
            "treble": self.vm.get("Strip[3].eqgain3"),
        }
        return {k: 5 * round(db_to_pct(raw_db[k], k) / 5) for k in self.KEYS}

    def _refresh_from_vm(self):
        """Re-read VoiceMeeter state and update all sliders."""
        live = self._read_vm()
        for k in self.KEYS:
            self._pct[k] = live[k]
            self._vars[k].set(live[k])
            self._lbls[k].config(text=self._fmt(live[k], k in EQ_KEYS))

    # ------------------------------------------------------------------
    # Apply percentages -> VoiceMeeter
    # ------------------------------------------------------------------
    def _sync_vm(self):
        o = pct_to_db(self._pct["others"], "others")
        m = pct_to_db(self._pct["me"],     "me")
        v = pct_to_db(self._pct["voice"],  "voice")

        VM_MAX, VM_MIN = 12.0, -60.0

        # Distribute music across Strip[3] + Bus[3] / Bus[4]
        s3 = max(min((o + m) / 2.0, VM_MAX), max(o, m) - VM_MAX)
        bus3 = o - s3
        bus4 = m - s3
        s0 = max(VM_MIN, min(VM_MAX, v - bus3))

        params = {
            "Strip[3].Gain":    round(s3, 1),
            "Bus[3].Gain":      round(bus3, 1),
            "Bus[4].Gain":      round(bus4, 1),
            "Strip[0].Gain":    round(s0, 1),
            "Strip[3].eqgain1": round(pct_to_db(self._pct["bass"],   "bass"),   1),
            "Strip[3].eqgain2": round(pct_to_db(self._pct["mid"],    "mid"),    1),
            "Strip[3].eqgain3": round(pct_to_db(self._pct["treble"], "treble"), 1),
        }
        for param, value in params.items():
            self.vm.set(param, value)

        # Persist so the tray app can restore after VoiceMeeter restart
        try:
            with open(VM_STATE_PATH, "w") as f:
                json.dump(params, f)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Presets
    # ------------------------------------------------------------------
    def _rebuild_presets(self):
        for w in self.preset_frame.winfo_children():
            w.destroy()

        self._preset_rows = []  # (name, row_frame) for drag-and-drop

        for name, vals in self.presets.items():
            mode = vals.get("mode")
            color = MODE_COLORS.get(mode, "#555555")

            row = ttk.Frame(self.preset_frame, style="D.TFrame")
            row.pack(fill="x", pady=1)
            row._preset_name = name
            self._preset_rows.append((name, row))

            # Mode dot
            dot = tk.Canvas(row, width=12, height=12, bg=self.bg,
                            highlightthickness=0, cursor="hand2")
            dot.create_oval(1, 1, 11, 11, fill=color, outline=color)
            dot.pack(side="left", padx=(0, 4), pady=2)

            # Name label (click = apply, drag = reorder)
            name_lbl = tk.Label(
                row, text=name, bg=self.btn_bg, fg=self.fg,
                relief="flat", padx=6, pady=2, font=("Segoe UI", 9),
                cursor="hand2", anchor="w",
            )
            name_lbl.pack(side="left")

            # Level summary
            o = int(vals.get("others", 0))
            m = int(vals.get("me", 0))
            v = int(vals.get("voice", 0))
            ml = MODE_LABELS.get(mode, "Any")
            info_lbl = ttk.Label(
                row, style="P.TLabel",
                text=f"  {ml}   O:{o:+d}%  M:{m:+d}%  V:{v:+d}%",
            )
            info_lbl.pack(side="left", padx=(4, 0))

            # Bind drag to row + all non-button children
            for w in (row, dot, name_lbl, info_lbl):
                w.bind("<ButtonPress-1>",   lambda e, n=name: self._drag_start(e, n))
                w.bind("<B1-Motion>",       self._drag_motion)
                w.bind("<ButtonRelease-1>", self._drag_end)

            # Double-click name to rename
            name_lbl.bind("<Double-1>", lambda e, n=name: self._rename_preset(n))

            # Delete button
            tk.Button(
                row, text="\u00d7", bg=self.btn_bg, fg="#666",
                activebackground=self.btn_danger, activeforeground=self.fg,
                relief="flat", padx=4, pady=1, font=("Segoe UI", 9),
                cursor="hand2",
                command=lambda n=name: self._delete_preset(n),
            ).pack(side="right")

            # Overwrite (save) button — floppy disk icon
            tk.Button(
                row, text="\U0001f4be", bg=self.btn_bg, fg="#888",
                activebackground=self.accent, activeforeground="#000",
                relief="flat", padx=4, pady=1, font=("Segoe UI", 9),
                cursor="hand2",
                command=lambda n=name: self._overwrite_preset(n),
            ).pack(side="right")

    # ------------------------------------------------------------------
    # Drag-and-drop reordering
    # ------------------------------------------------------------------
    _DRAG_THRESHOLD = 8   # pixels before drag activates
    _DRAG_OFFSET_X = 8    # floating row pops right
    _DRAG_OFFSET_Y = -4   # floating row pops up

    def _drag_start(self, event, name):
        self._drag_name = name
        self._drag_start_y = event.y_root
        self._drag_active = False
        self._drag_float = None
        self._drag_placeholder = None

    def _drag_activate(self, event):
        """Create floating ghost row and dark placeholder."""
        name = self._drag_name
        self._drag_active = True

        # Find original row
        orig_row = None
        orig_idx = 0
        for i, (nm, row) in enumerate(self._preset_rows):
            if nm == name:
                orig_row = row
                orig_idx = i
                break
        if orig_row is None:
            return

        row_w = orig_row.winfo_width()
        row_h = orig_row.winfo_height()
        row_x = orig_row.winfo_rootx()
        row_y = orig_row.winfo_rooty()
        self._drag_grab_offset = event.y_root - row_y

        # --- Floating ghost row (Toplevel) ---
        floater = tk.Toplevel(self.root)
        floater.overrideredirect(True)
        floater.attributes("-topmost", True)
        floater.attributes("-alpha", 0.85)
        floater.configure(bg="#2a2a2a")

        vals = self.presets[name]
        mode = vals.get("mode")
        color = MODE_COLORS.get(mode, "#555555")

        ff = tk.Frame(floater, bg="#2a2a2a")
        ff.pack(fill="x", padx=2, pady=2)

        c = tk.Canvas(ff, width=12, height=12, bg="#2a2a2a",
                      highlightthickness=0)
        c.create_oval(1, 1, 11, 11, fill=color, outline=color)
        c.pack(side="left", padx=(4, 4), pady=2)

        tk.Label(ff, text=name, bg="#2a2a2a", fg=self.fg,
                 font=("Segoe UI", 9)).pack(side="left")

        o = int(vals.get("others", 0))
        m = int(vals.get("me", 0))
        v = int(vals.get("voice", 0))
        ml = MODE_LABELS.get(mode, "Any")
        tk.Label(ff, text=f"  {ml}   O:{o:+d}%  M:{m:+d}%  V:{v:+d}%",
                 bg="#2a2a2a", fg=self.desc_fg,
                 font=("Segoe UI", 8)).pack(side="left", padx=(4, 0))

        fx = row_x + self._DRAG_OFFSET_X
        fy = event.y_root - self._drag_grab_offset + self._DRAG_OFFSET_Y
        floater.geometry(f"{row_w}x{row_h + 4}+{fx}+{fy}")
        self._drag_float = floater
        self._drag_float_x = fx

        # --- Hide original row, insert dark placeholder ---
        orig_row.pack_forget()

        ph = tk.Frame(self.preset_frame, bg="#2d2d2d", height=row_h)
        ph.pack_propagate(False)
        self._drag_placeholder = ph

        # Order of non-dragged items (stays constant during drag)
        self._drag_order = [nm for nm, _ in self._preset_rows if nm != name]
        self._drag_ph_idx = -1  # force first repack
        self._repack_rows(orig_idx)

    def _repack_rows(self, ph_idx):
        """Repack all non-dragged rows with placeholder at ph_idx."""
        if ph_idx == self._drag_ph_idx:
            return
        # Forget everything
        row_map = {nm: row for nm, row in self._preset_rows}
        for nm in self._drag_order:
            row_map[nm].pack_forget()
        self._drag_placeholder.pack_forget()

        # Repack with placeholder at the right position
        for i, nm in enumerate(self._drag_order):
            if i == ph_idx:
                self._drag_placeholder.pack(fill="x", pady=1)
            row_map[nm].pack(fill="x", pady=1)
        if ph_idx >= len(self._drag_order):
            self._drag_placeholder.pack(fill="x", pady=1)

        self._drag_ph_idx = ph_idx
        self.preset_frame.update_idletasks()

    def _insertion_index(self, y_root):
        """Which position should the placeholder be at?"""
        visible = [(nm, row) for nm, row in self._preset_rows
                   if nm != self._drag_name]
        if not visible:
            return 0
        for i, (nm, row) in enumerate(visible):
            mid = row.winfo_rooty() + row.winfo_height() / 2
            if y_root < mid:
                return i
        return len(visible)

    def _drag_motion(self, event):
        if not hasattr(self, "_drag_name") or not self._drag_name:
            return
        if not self._drag_active:
            if abs(event.y_root - self._drag_start_y) < self._DRAG_THRESHOLD:
                return
            self._drag_activate(event)
            return

        # Move floating ghost
        if self._drag_float:
            fy = event.y_root - self._drag_grab_offset + self._DRAG_OFFSET_Y
            self._drag_float.geometry(f"+{self._drag_float_x}+{fy}")

        # Dynamically reorder placeholder
        self._repack_rows(self._insertion_index(event.y_root))

    def _drag_end(self, event):
        if not hasattr(self, "_drag_name") or not self._drag_name:
            return
        source = self._drag_name
        self._drag_name = None
        if not getattr(self, "_drag_active", False):
            self._apply_preset(source)  # click, not drag
            return

        # Clean up floating row
        if self._drag_float:
            self._drag_float.destroy()
            self._drag_float = None

        ph_idx = getattr(self, "_drag_ph_idx", 0)

        if self._drag_placeholder:
            self._drag_placeholder.destroy()
            self._drag_placeholder = None

        # Build final order and apply
        order = list(self._drag_order)
        order.insert(ph_idx, source)
        self.presets = {k: self.presets[k] for k in order}
        save_presets(self.presets)
        self._rebuild_presets()

    def _apply_preset(self, name):
        vals = self.presets.get(name)
        if not vals:
            return
        for k in self.KEYS:
            self._pct[k] = vals.get(k, 100 if k in EQ_KEYS else 0)
            self._vars[k].set(self._pct[k])
            self._lbls[k].config(text=self._fmt(self._pct[k], k in EQ_KEYS))
        self._sync_vm()

        mode = vals.get("mode")
        if mode:
            request_mode(mode)

    def _save_preset(self):
        name = simpledialog.askstring("Save Preset", "Preset name:",
                                      parent=self.root)
        if not name or not name.strip():
            return
        name = name.strip()
        self.presets[name] = {k: self._pct[k] for k in self.KEYS}
        self.presets[name]["mode"] = read_current_mode()
        save_presets(self.presets)
        self._rebuild_presets()

    def _overwrite_preset(self, name):
        """Overwrite an existing preset with the current slider values + mode."""
        if name not in self.presets:
            return
        self.presets[name] = {k: self._pct[k] for k in self.KEYS}
        self.presets[name]["mode"] = read_current_mode()
        save_presets(self.presets)
        # Clear any pending mode request so the tray doesn't switch
        try:
            if STATE_PATH.exists():
                with open(STATE_PATH) as f:
                    state = json.load(f)
                state.pop("requested_mode", None)
                with open(STATE_PATH, "w") as f:
                    json.dump(state, f)
        except Exception:
            pass
        self.root.after(10, self._rebuild_presets)

    def _rename_preset(self, name):
        """Rename a preset, preserving its position in the list."""
        if name not in self.presets:
            return
        new_name = simpledialog.askstring("Rename Preset", "New name:",
                                          initialvalue=name, parent=self.root)
        if not new_name or not new_name.strip() or new_name.strip() == name:
            return
        new_name = new_name.strip()
        # Rebuild dict with new key in same position
        self.presets = {
            (new_name if k == name else k): v
            for k, v in self.presets.items()
        }
        save_presets(self.presets)
        self._rebuild_presets()

    def _delete_preset(self, name):
        if name in self.presets:
            del self.presets[name]
            save_presets(self.presets)
            self._rebuild_presets()

    def _on_close(self):
        try:
            self.vm.close()
        except Exception:
            pass
        self.root.destroy()

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    mutex = acquire_mutex()
    if mutex is None:
        sys.exit(0)
    try:
        vm = VM()
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        from tkinter import messagebox
        messagebox.showerror("VR Audio Mixer",
                             f"Could not connect to VoiceMeeter.\n\n"
                             f"Make sure VoiceMeeter Banana is running.\n\n{e}")
        sys.exit(1)
    MixerApp(vm).run()


if __name__ == "__main__":
    main()
