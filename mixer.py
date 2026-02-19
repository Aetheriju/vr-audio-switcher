"""VR Audio Switcher — tabbed UI with mixer, guide, settings, and help."""

import json
import logging
import math
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, simpledialog
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.resolve()
PRESETS_PATH = SCRIPT_DIR / "presets.json"
VM_STATE_PATH = SCRIPT_DIR / "vm_state.json"
CONFIG_PATH = SCRIPT_DIR / "config.json"

# Sweet-spot defaults — the "0%" center for each slider (in dB)
DEFAULTS = {
    "others":   0.0,
    "me":       0.0,
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
# Helpers
# ---------------------------------------------------------------------------
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
# Tabbed UI
# ---------------------------------------------------------------------------
class MixerApp:
    KEYS = ["others", "me", "voice", "bass", "mid", "treble"]

    def __init__(self, app, initial_tab="mixer"):
        self.app = app
        self.vm = app.vm
        self.presets = load_presets()
        self._closing = False

        self.root = tk.Tk()
        self.root.title("VR Audio Switcher")
        self.root.resizable(False, False)

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

        # --- Notebook with tabs ---
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=2, pady=2)

        # Build all tabs
        self._build_mixer_tab()
        self._build_guide_tab()
        self._build_settings_tab()
        self._build_help_tab()

        # Select requested tab
        tab_index = {"mixer": 0, "guide": 1, "settings": 2, "help": 3}
        self.notebook.select(tab_index.get(initial_tab, 0))

        # Initial mode/VR display
        self._update_mode_display()

        self.root.bind("<Escape>", lambda e: self._on_close())
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
        # Notebook tab styles
        s.configure("TNotebook", background=self.bg, borderwidth=0)
        s.configure("TNotebook.Tab", background="#333333", foreground=self.fg,
                    padding=[14, 6], font=("Segoe UI", 10))
        s.map("TNotebook.Tab",
              background=[("selected", self.bg), ("active", "#444444")],
              foreground=[("selected", self.accent), ("active", self.fg)])

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
    # Mode display (called by app._notify_ui via root.after)
    # ------------------------------------------------------------------
    def _update_mode_display(self):
        """Refresh mode buttons, VR status, and title bar."""
        current = self.app.get_mode_name()
        vr_active = self.app.detector.is_vr_running()

        # Highlight active mode button
        for mode_name, btn in self._mode_btns.items():
            color = MODE_COLORS.get(mode_name, "#555555")
            if mode_name == current:
                # Active: bright bg + sunken look
                btn.config(bg=color, fg="#000", relief="groove", bd=2)
            else:
                # Inactive: dimmed
                btn.config(bg="#2a2a2a", fg=color, relief="flat", bd=0)

        # VR status label
        if vr_active:
            self._vr_lbl.config(text="\u25cf VR Active", foreground="#4caf50")
        else:
            self._vr_lbl.config(text="\u25cb VR Off", foreground="#555555")

        # Title bar
        mode_label = MODE_LABELS.get(current, "")
        if mode_label:
            self.root.title(f"VR Audio Switcher \u2014 {mode_label}")

    def _set_mode(self, mode_name):
        """Set mode via the app and let it handle apply + UI refresh."""
        self.app.set_user_mode(mode_name)

    # ------------------------------------------------------------------
    # Mixer tab
    # ------------------------------------------------------------------
    def _build_mixer_tab(self):
        tab = ttk.Frame(self.notebook, style="D.TFrame")
        self.notebook.add(tab, text="  Mixer  ")

        main = ttk.Frame(tab, style="D.TFrame", padding=15)
        main.pack(fill="both", expand=True)

        # --- Mode selector ---
        mode_frame = ttk.Frame(main, style="D.TFrame")
        mode_frame.pack(fill="x", pady=(0, 8))

        self._mode_btns = {}
        modes = [
            ("DESKTOP",   "Desktop",  "#42a5f5"),
            ("AUTO",      "Auto",     "#4caf50"),
            ("SILENT_VR", "Private",  "#fdd835"),
            ("VR",        "Public",   "#f44336"),
        ]
        for mode_name, label, color in modes:
            btn = tk.Button(
                mode_frame, text=label, bg="#2a2a2a", fg=color,
                activebackground=color, activeforeground="#000",
                relief="flat", padx=12, pady=4,
                font=("Segoe UI", 9, "bold"), cursor="hand2",
                command=lambda m=mode_name: self._set_mode(m),
            )
            btn.pack(side="left", padx=(0, 4))
            self._mode_btns[mode_name] = btn

        # VR status indicator (right side)
        self._vr_lbl = ttk.Label(
            mode_frame, text="\u25cb VR Off", style="D.TLabel",
            foreground="#555555", font=("Segoe UI", 9))
        self._vr_lbl.pack(side="right")

        # --- Volume sliders ---
        self._sec(main, "Volume")
        self._vars, self._lbls = {}, {}
        for key, label, desc in [
            ("others", "Music for Others", "What friends hear (Public mode only)"),
            ("me",     "Music for Me",     "What you hear in your VR headset"),
            ("voice",  "My Voice",         "Your voice volume to others in VRChat"),
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
        self._sec(main, "Presets (click to apply, drag to reorder)")
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

    # ------------------------------------------------------------------
    # Guide tab
    # ------------------------------------------------------------------
    def _build_guide_tab(self):
        tab = ttk.Frame(self.notebook, style="D.TFrame")
        self.notebook.add(tab, text="  Guide  ")

        frame = ttk.Frame(tab, style="D.TFrame", padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Welcome to VR Audio Switcher",
                  style="D.TLabel",
                  font=("Segoe UI", 13, "bold"),
                  foreground=self.accent).pack(pady=(0, 10))

        sections = [
            ("How It Works",
             "This app runs silently in the background and springs\n"
             "to life when SteamVR starts. It launches VoiceMeeter,\n"
             "routes your music apps through it so friends in VRChat\n"
             "can hear what you're playing, and shuts everything\n"
             "down cleanly when VR stops."),
            ("First Time Setup",
             "1. Open VRChat > Settings > Audio > Microphone\n"
             "2. Select \"Voicemeeter Out B1\" as your mic\n"
             "3. Play music in any app. Done!\n"
             "(See the Help tab for more details)"),
            ("Modes",
             "\u2022 Blue (Desktop): all audio plays through speakers\n"
             "\u2022 Green (Auto): switches with VR automatically\n"
             "\u2022 Yellow (Private): music in your headset only\n"
             "\u2022 Red (Public): music in headset + VRChat mic\n\n"
             "Use the mode buttons at the top of the Mixer tab\n"
             "to switch between modes."),
            ("Mixer Tab",
             "\u2022 Music for Others: what friends hear in VRChat\n"
             "\u2022 Music for Me: your personal headset volume\n"
             "\u2022 My Voice: your mic volume in VRChat\n"
             "\u2022 Presets: save and load your favorite settings"),
            ("Good to Know",
             "\u2022 Everything boots and shuts down with SteamVR\n"
             "\u2022 Closing this window will also close SteamVR\n"
             "   and VoiceMeeter (you'll be asked to confirm)\n"
             "\u2022 The app starts with Windows automatically\n"
             "\u2022 Settings are saved and persist between sessions"),
        ]

        for title, body in sections:
            ttk.Label(frame, text=title, style="D.TLabel",
                      font=("Segoe UI", 9, "bold"),
                      foreground=self.accent).pack(
                      fill="x", pady=(10, 2), anchor="w")
            ttk.Label(frame, text=body, style="D.TLabel",
                      font=("Segoe UI", 9),
                      foreground="#bbb").pack(fill="x", anchor="w")

    # ------------------------------------------------------------------
    # Settings tab
    # ------------------------------------------------------------------
    def _build_settings_tab(self):
        tab = ttk.Frame(self.notebook, style="D.TFrame")
        self.notebook.add(tab, text=" Settings ")

        frame = ttk.Frame(tab, style="D.TFrame", padding=15)
        frame.pack(fill="both", expand=True)

        # Load config
        config = {}
        if CONFIG_PATH.exists():
            try:
                with open(CONFIG_PATH) as f:
                    config = json.load(f)
            except Exception:
                pass

        # --- Excluded Apps ---
        ttk.Label(frame, text="Excluded Apps", style="D.TLabel",
                  font=("Segoe UI", 10, "bold"),
                  foreground=self.accent).pack(anchor="w", pady=(0, 4))
        ttk.Label(frame, text="These apps keep their normal audio output.\n"
                  "Everything else gets routed through VoiceMeeter.",
                  style="Desc.TLabel").pack(anchor="w")

        list_frame = ttk.Frame(frame, style="D.TFrame")
        list_frame.pack(fill="x", pady=(4, 4))

        lb = tk.Listbox(list_frame, bg="#111", fg=self.fg,
                        selectbackground="#333",
                        font=("Consolas", 9), height=6, relief="flat")
        lb.pack(side="left", fill="both", expand=True)

        for p in config.get("exclude_processes", ["vrchat.exe"]):
            lb.insert("end", p)

        btn_frame = ttk.Frame(list_frame, style="D.TFrame")
        btn_frame.pack(side="right", padx=(8, 0))

        def _add():
            name = simpledialog.askstring(
                "Add Exclusion", "Process name (e.g. discord.exe):",
                parent=self.root)
            if name and name.strip():
                name = name.strip().lower()
                if not name.endswith(".exe"):
                    name += ".exe"
                lb.insert("end", name)

        def _remove():
            sel = lb.curselection()
            if sel:
                lb.delete(sel[0])

        tk.Button(btn_frame, text="+ Add", bg=self.btn_bg, fg=self.fg,
                  relief="flat", padx=8, pady=2, font=("Segoe UI", 8),
                  command=_add).pack(pady=2)
        tk.Button(btn_frame, text="\u2212 Remove", bg=self.btn_bg, fg=self.fg,
                  relief="flat", padx=8, pady=2, font=("Segoe UI", 8),
                  command=_remove).pack(pady=2)

        ttk.Label(frame, text="System processes (svchost, audiodg, VoiceMeeter, "
                  "SteamVR) are always excluded automatically.",
                  style="Desc.TLabel", wraplength=350).pack(
                  anchor="w", pady=(4, 8))

        def _save():
            # Re-read config to avoid clobbering other keys
            fresh = {}
            try:
                with open(CONFIG_PATH) as f:
                    fresh = json.load(f)
            except Exception:
                pass
            fresh["exclude_processes"] = [lb.get(i) for i in range(lb.size())]
            try:
                with open(CONFIG_PATH, "w") as f:
                    json.dump(fresh, f, indent=2)
            except Exception:
                pass
            save_btn.config(text="Saved!", bg="#66bb6a")
            self.root.after(1500, lambda: save_btn.config(
                text="Save", bg=self.accent))

        save_btn = tk.Button(frame, text="Save", bg=self.accent, fg="#000",
                             relief="flat", padx=20, pady=6,
                             font=("Segoe UI", 10, "bold"),
                             command=_save)
        save_btn.pack(pady=(8, 0))

        # --- Updates ---
        ttk.Label(frame, text="Updates", style="D.TLabel",
                  font=("Segoe UI", 10, "bold"),
                  foreground=self.accent).pack(anchor="w", pady=(12, 4))

        update_frame = ttk.Frame(frame, style="D.TFrame")
        update_frame.pack(fill="x")

        self._update_status = ttk.Label(
            update_frame, text="Updates are checked automatically on boot.",
            style="Desc.TLabel")
        self._update_status.pack(side="left")

        update_btn = tk.Button(
            update_frame, text="Check Now", bg=self.btn_bg, fg=self.fg,
            activebackground=self.btn_act, activeforeground=self.fg,
            relief="flat", padx=10, pady=3, font=("Segoe UI", 8),
            cursor="hand2", command=self._check_for_updates)
        update_btn.pack(side="right")

    def _check_for_updates(self):
        """Manual update check from settings tab."""
        self._update_status.config(text="Checking...", foreground=self.fg)
        self.root.update_idletasks()
        try:
            from updater import update_available, do_update, restart_app
            avail, local_ver, remote_ver = update_available()
            if avail:
                self._update_status.config(
                    text=f"Update v{remote_ver} available! Downloading...",
                    foreground=self.accent)
                self.root.update_idletasks()
                ok, msg = do_update()
                if ok:
                    self._update_status.config(
                        text="Updated! Restarting...",
                        foreground=self.accent)
                    self.root.update_idletasks()
                    self.root.after(500, restart_app)
                else:
                    self._update_status.config(
                        text=f"Update failed: {msg}",
                        foreground="#f44336")
            else:
                self._update_status.config(
                    text=f"You're on the latest version (v{local_ver}).",
                    foreground=self.accent)
        except Exception as e:
            self._update_status.config(
                text=f"Check failed: {e}", foreground="#f44336")

    # ------------------------------------------------------------------
    # Help tab
    # ------------------------------------------------------------------
    def _build_help_tab(self):
        tab = ttk.Frame(self.notebook, style="D.TFrame")
        self.notebook.add(tab, text="  Help  ")

        frame = ttk.Frame(tab, style="D.TFrame", padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="VRChat Mic Setup",
                  style="D.TLabel",
                  font=("Segoe UI", 13, "bold"),
                  foreground=self.accent).pack(pady=(0, 10))

        help_text = (
            "For friends to hear your music in VRChat, you need\n"
            "to set your in-game mic to VoiceMeeter:\n\n"
            "1. Launch VRChat and get into a world\n"
            "2. Open Settings \u2192 Audio \u2192 Microphone\n"
            "3. Select \"Voicemeeter Out B1\"\n"
            "4. Switch to Public mode to test\n\n"
            "This only needs to be done once. VRChat remembers it."
        )

        ttk.Label(frame, text=help_text, style="D.TLabel",
                  font=("Segoe UI", 10), wraplength=380,
                  justify="left", foreground="#bbb").pack(anchor="w")

        ttk.Label(frame, text="Troubleshooting",
                  style="D.TLabel",
                  font=("Segoe UI", 10, "bold"),
                  foreground=self.accent).pack(anchor="w", pady=(20, 5))

        trouble_text = (
            "\u2022 Friends can't hear music?\n"
            "   Make sure you're in Public mode (red button)\n"
            "\u2022 No sound at all?\n"
            "   Check that VRChat mic is set to \"Voicemeeter Out B1\"\n"
            "\u2022 Music too loud/quiet for others?\n"
            "   Adjust \"Music for Others\" in the Mixer tab\n"
            "\u2022 An app's audio keeps switching when it shouldn't?\n"
            "   Add it to excluded apps in the Settings tab"
        )

        ttk.Label(frame, text=trouble_text, style="D.TLabel",
                  font=("Segoe UI", 9),
                  foreground="#bbb").pack(anchor="w")

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

        # Persist so settings survive VoiceMeeter restarts
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

            # Overwrite (save) button
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
            self.app.set_user_mode(mode)

    def _save_preset(self):
        name = simpledialog.askstring("Save Preset", "Preset name:",
                                      parent=self.root)
        if not name or not name.strip():
            return
        name = name.strip()
        self.presets[name] = {k: self._pct[k] for k in self.KEYS}
        self.presets[name]["mode"] = self.app.get_mode_name()
        save_presets(self.presets)
        self._rebuild_presets()

    def _overwrite_preset(self, name):
        """Overwrite an existing preset with the current slider values + mode."""
        if name not in self.presets:
            return
        self.presets[name] = {k: self._pct[k] for k in self.KEYS}
        self.presets[name]["mode"] = self.app.get_mode_name()
        save_presets(self.presets)
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

    # ------------------------------------------------------------------
    # Close / confirmation dialogs
    # ------------------------------------------------------------------
    def _on_close(self):
        """X button clicked — show themed confirmation dialog."""
        if self._closing:
            return
        self._show_close_confirmation()

    def _force_close(self):
        """Close without confirmation (called when VR stops)."""
        if self._closing:
            return
        self._closing = True
        self.root.destroy()

    def _show_close_confirmation(self):
        """Themed dark confirmation popup for shutting down."""
        dlg = tk.Toplevel(self.root)
        dlg.overrideredirect(True)
        dlg.configure(bg="#1e1e1e")
        dlg.attributes("-topmost", True)
        dlg.grab_set()

        # Red accent border
        border = tk.Frame(dlg, bg="#f44336", padx=2, pady=2)
        border.pack(fill="both", expand=True)
        inner = tk.Frame(border, bg="#1e1e1e")
        inner.pack(fill="both", expand=True)

        tk.Label(
            inner, text="Shut Down VR Audio Switcher?",
            bg="#1e1e1e", fg="#f44336",
            font=("Segoe UI", 13, "bold"),
        ).pack(pady=(20, 10))

        tk.Label(
            inner,
            text=(
                "Your VRChat microphone is routed through VoiceMeeter.\n"
                "Closing this will also shut down VoiceMeeter and SteamVR."
            ),
            bg="#1e1e1e", fg="#aaaaaa",
            font=("Segoe UI", 10), justify="center",
        ).pack(padx=25)

        btn_frame = tk.Frame(inner, bg="#1e1e1e")
        btn_frame.pack(pady=(18, 20))

        def _yes():
            dlg.destroy()
            self.app.close_steamvr()
            self._closing = True
            self.root.destroy()

        def _no():
            dlg.destroy()

        tk.Button(
            btn_frame, text="Shut Down Everything",
            bg="#f44336", fg="#ffffff",
            activebackground="#d32f2f", activeforeground="#ffffff",
            relief="flat", padx=16, pady=6,
            font=("Segoe UI", 10, "bold"), cursor="hand2",
            command=_yes,
        ).pack(side="left", padx=8)

        tk.Button(
            btn_frame, text="Cancel",
            bg="#333333", fg="#e0e0e0",
            activebackground="#444444", activeforeground="#e0e0e0",
            relief="flat", padx=16, pady=6,
            font=("Segoe UI", 10), cursor="hand2",
            command=_no,
        ).pack(side="left", padx=8)

        # Center on parent window
        dlg.update_idletasks()
        dw, dh = dlg.winfo_width(), dlg.winfo_height()
        px, py = self.root.winfo_x(), self.root.winfo_y()
        pw, ph = self.root.winfo_width(), self.root.winfo_height()
        x = px + (pw - dw) // 2
        y = py + (ph - dh) // 2
        dlg.geometry(f"+{x}+{y}")

        dlg.bind("<Escape>", lambda e: _no())

    def _show_vm_closed_dialog(self):
        """Themed dialog when VoiceMeeter closes unexpectedly."""
        if self._closing:
            return

        dlg = tk.Toplevel(self.root)
        dlg.overrideredirect(True)
        dlg.configure(bg="#1e1e1e")
        dlg.attributes("-topmost", True)
        dlg.grab_set()

        # Orange accent border
        border = tk.Frame(dlg, bg="#ff9800", padx=2, pady=2)
        border.pack(fill="both", expand=True)
        inner = tk.Frame(border, bg="#1e1e1e")
        inner.pack(fill="both", expand=True)

        tk.Label(
            inner, text="VoiceMeeter Closed",
            bg="#1e1e1e", fg="#ff9800",
            font=("Segoe UI", 13, "bold"),
        ).pack(pady=(20, 10))

        tk.Label(
            inner,
            text=(
                "VoiceMeeter was closed but your VRChat mic\n"
                "needs it to work. What would you like to do?"
            ),
            bg="#1e1e1e", fg="#aaaaaa",
            font=("Segoe UI", 10), justify="center",
        ).pack(padx=25)

        btn_frame = tk.Frame(inner, bg="#1e1e1e")
        btn_frame.pack(pady=(18, 20))

        def _restart():
            dlg.destroy()
            threading.Thread(
                target=self.app.restart_voicemeeter, daemon=True
            ).start()

        def _shutdown():
            dlg.destroy()
            self.app.close_steamvr()
            self._closing = True
            self.root.destroy()

        tk.Button(
            btn_frame, text="Restart VoiceMeeter",
            bg="#4caf50", fg="#000000",
            activebackground="#66bb6a", activeforeground="#000000",
            relief="flat", padx=16, pady=6,
            font=("Segoe UI", 10, "bold"), cursor="hand2",
            command=_restart,
        ).pack(side="left", padx=8)

        tk.Button(
            btn_frame, text="Shut Down Everything",
            bg="#f44336", fg="#ffffff",
            activebackground="#d32f2f", activeforeground="#ffffff",
            relief="flat", padx=16, pady=6,
            font=("Segoe UI", 10), cursor="hand2",
            command=_shutdown,
        ).pack(side="left", padx=8)

        # Center on parent window
        dlg.update_idletasks()
        dw, dh = dlg.winfo_width(), dlg.winfo_height()
        px, py = self.root.winfo_x(), self.root.winfo_y()
        pw, ph = self.root.winfo_width(), self.root.winfo_height()
        x = px + (pw - dw) // 2
        y = py + (ph - dh) // 2
        dlg.geometry(f"+{x}+{y}")

        dlg.bind("<Escape>", lambda e: _restart())

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()
