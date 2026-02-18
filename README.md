# VR Audio Switcher

**One-click audio management for VRChat + Steam Link.**

You play VRChat on Quest via Steam Link. You want your Spotify/YouTube music to play through your VRChat mic so friends can hear it — but you also want independent volume control over what *they* hear vs. what *you* hear, and you want everything to switch back to your desktop speakers automatically when you close SteamVR. That's exactly what this does.

## What It Does

- **Auto-detects SteamVR** — switches Chrome's audio output to VoiceMeeter when SteamVR launches, back to your speakers when it closes
- **4 audio modes** with one-click tray icon switching:
  - **Desktop** (blue) — music plays through your speakers, VoiceMeeter bypassed
  - **Auto** (green) — follows SteamVR state, defaults to Private behavior
  - **Private** (yellow) — music in your headset only, mic stays clean
  - **Public** (red) — music plays through your VRChat mic AND your headset
- **Mixer UI** with independent sliders for:
  - Music volume for others (what friends hear)
  - Music volume for you (what you hear in your headset)
  - Your voice volume in VRChat
  - Bass / Mid / Treble EQ
- **Preset system** — save, load, rename, reorder audio profiles with drag-and-drop
- **Full VoiceMeeter lifecycle** — auto-launches VoiceMeeter on startup, persists all settings (gains, EQ, device assignments) across restarts, shuts down cleanly on quit

## How It Works

```
                    ┌─────────────────────────────┐
  Physical Mic ───► │  VoiceMeeter Banana          │
  (Steam Streaming  │                              │
   Microphone)      │  Strip[0] (Mic) ──► B1 ──────┼──► VRChat (mic input)
                    │                              │
  Chrome ─────────► │  Strip[3] (VAIO) ──► B2 ─────┼──► Quest Headset
  (YouTube/Spotify) │         └──► B1 (if Public) ─┼──► VRChat (music in mic)
                    │                              │
                    │  Mixer controls Strip/Bus    │
                    │  gains independently         │
                    └─────────────────────────────┘
```

**Desktop mode:** Chrome outputs directly to your speakers (Soundbar, headphones, etc). VoiceMeeter runs silently in the background with no effect.

**VR modes:** Chrome outputs to VoiceMeeter's virtual input (VAIO). VoiceMeeter routes the audio to B2 (your headset via "Listen to this device") and optionally B1 (VRChat mic). The Mixer UI controls the gain split so you can make music quieter for friends while keeping it loud for yourself.

## Requirements

- **Windows 10/11**
- **Python 3.10+** — [python.org/downloads](https://www.python.org/downloads/)
- **VoiceMeeter Banana** — [vb-audio.com/Voicemeeter/banana](https://vb-audio.com/Voicemeeter/banana.htm) (free)
- **Steam Link** + **Quest headset** (or any VR setup with Steam Streaming audio devices)

## Quick Start

### 1. Download
```bash
git clone https://github.com/Aetheriju/vr-audio-switcher.git
cd vr-audio-switcher
```

Or just [download the ZIP](https://github.com/Aetheriju/vr-audio-switcher/archive/refs/heads/main.zip) and extract it.

### 2. Double-click `install.bat`

That's it. The installer handles everything automatically:
- Downloads and installs **Python** if you don't have it
- Downloads **svcl.exe** (NirSoft per-app audio routing tool)
- Installs Python packages
- Downloads and runs the **VoiceMeeter Banana** installer if needed
- Detects your audio devices
- Configures "Listen to this device" for VR audio routing
- Creates desktop and startup shortcuts

### 3. You're done

The tray icon appears in your system tray. Left-click to cycle modes, right-click for the full menu. Launch the Mixer from the right-click menu to fine-tune volumes and save presets.

## Manual Setup

If you prefer to configure manually instead of using the wizard:

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Download [svcl-x64.zip](https://www.nirsoft.net/utils/svcl-x64.zip) from NirSoft, extract `svcl.exe` into the project directory.

3. Find your device IDs by running:
   ```bash
   svcl.exe /scomma "" /Columns "Name,Command-Line Friendly ID,Type,Direction,Device State"
   ```

4. Create `config.json`:
   ```json
   {
     "poll_interval_seconds": 3,
     "steamvr_process": "vrserver.exe",
     "target_process": "chrome.exe",
     "svcl_path": "svcl.exe",
     "desktop_device": "auto",
     "vr_device": "<VoiceMeeter VAIO Friendly ID from svcl>",
     "debounce_seconds": 5,
     "music_strip": 3
   }
   ```
   - `desktop_device`: Set to `"auto"` to auto-detect your physical speakers, or paste a specific svcl Friendly ID.
   - `vr_device`: The svcl Friendly ID for "Voicemeeter Input" (look for `VB-Audio Voicemeeter VAIO\Device\Voicemeeter Input\Render`).

5. Create `vm_devices.json` with your mic device name (as shown in VoiceMeeter):
   ```json
   {
     "Strip[0]": "Microphone (Your Mic Device Name)"
   }
   ```

6. Set up VoiceMeeter Banana:
   - **Strip[0]** hardware input: your microphone
   - **Strip[3]** (VAIO): this is where Chrome audio arrives — no setup needed
   - Enable "Listen to this device" on **Voicemeeter Out B2** in Windows Sound settings, targeting your Quest/Steam Streaming Speakers

7. Run:
   ```bash
   pythonw vr_audio_switcher.py
   ```

## VoiceMeeter "Listen to this device" Setup

This is the one Windows setting that bridges VoiceMeeter's B2 output to your Quest headset:

1. Open **Windows Sound Settings** → **More sound settings**
2. Go to the **Recording** tab
3. Find **Voicemeeter Out B2** (VB-Audio VoiceMeeter VAIO)
4. Right-click → **Properties** → **Listen** tab
5. Check **"Listen to this device"**
6. Set playback device to **Steam Streaming Speakers** (your Quest audio)
7. Click OK

This setting persists across reboots — you only need to do it once.

## Modes Explained

| Mode | Icon | Chrome Output | Music in Headset | Music in VRChat Mic | Use Case |
|------|------|--------------|-----------------|--------------------|----|
| Desktop | Blue | Your speakers | No | No | Normal desktop use |
| Auto | Green | Follows SteamVR | When VR active | No | Default — hands-free switching |
| Private | Yellow | VoiceMeeter | Yes | No | Personal listening in VR |
| Public | Red | VoiceMeeter | Yes | Yes | DJ mode — friends hear your music |

## Mixer

Right-click the tray icon → **Mixer** to open the volume control panel.

**Volume sliders** (-100% to +100%):
- **Music for Others** — what friends hear through your VRChat mic (Public mode only)
- **Music for Me** — what you hear in your Quest headset
- **My Voice** — your mic volume in VRChat

**EQ sliders** (0% to 100%):
- Bass, Mid, Treble — shape the music before it hits VRChat

**Presets:**
- Click a preset to load it instantly
- Drag presets to reorder them
- Double-click a preset name to rename it
- Use the save icon to overwrite a preset with current settings
- "Save Current" creates a new preset from your current slider positions

## Troubleshooting

**No audio after VoiceMeeter restart:**
The app automatically saves and restores VoiceMeeter device assignments and gain settings. If audio is missing, try: right-click tray → Quit, then relaunch from the desktop shortcut.

**Mic not working in VRChat:**
- Make sure VRChat's mic input is set to **Voicemeeter Out B1** (not your physical mic)
- Check that Strip[0] in VoiceMeeter shows your physical mic as the input device
- Verify Strip[0].B1 routing is enabled (the app enforces this every 15 seconds)

**Chrome audio not switching:**
- Chrome must be running when the switch happens
- The app uses `svcl.exe` for per-app audio routing — make sure it's in the project directory
- Check `switcher.log` for errors

**"svcl.exe not found":**
Run `python setup_wizard.py` to download it, or manually download [svcl-x64.zip](https://www.nirsoft.net/utils/svcl-x64.zip) and extract `svcl.exe` to the project directory.

**Antivirus flags svcl.exe:**
NirSoft tools are sometimes flagged as false positives by antivirus software. svcl.exe is a legitimate Windows audio management utility. You may need to add an exception in your antivirus.

**Tray icon not visible:**
Windows may hide new tray icons. Click the **^** arrow in the system tray to find it, then drag it to the visible area.

## Architecture

| File | Purpose |
|------|---------|
| `vr_audio_switcher.py` | Main tray app — SteamVR detection, mode switching, VoiceMeeter control |
| `mixer.py` | Mixer UI — volume/EQ sliders, presets, drag-and-drop |
| `setup_wizard.py` | First-run setup — device detection, config generation, svcl download |
| `config.json` | User config — device IDs, polling settings |
| `vm_devices.json` | Persisted VoiceMeeter device assignments |
| `vm_state.json` | Persisted VoiceMeeter gain/EQ values |
| `presets.json` | User-saved mixer presets |
| `state.json` | Runtime communication between tray app and mixer |

## Credits

- **[VoiceMeeter Banana](https://vb-audio.com/Voicemeeter/banana.htm)** by VB-Audio — the audio routing engine
- **[SoundVolumeCommandLine (svcl)](https://www.nirsoft.net/utils/sound_volume_command_line.html)** by NirSoft — per-app audio device switching

## License

MIT — see [LICENSE](LICENSE).
