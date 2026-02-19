# VR Audio Switcher

**One-click audio management for VRChat.**

You play VRChat on PC and want music from *any* app — Spotify, YouTube, SoundCloud, VLC, whatever you use — to play through your VRChat mic so friends can hear it. But you also want independent volume control over what *they* hear vs. what *you* hear, and you want everything to switch back to your desktop speakers automatically when you close SteamVR. That's exactly what this does.

Works with any VR headset — Quest (via Steam Link or Virtual Desktop), Index, Vive, Rift, or anything else with its own audio output.

## What It Does

- **Auto-detects SteamVR** — switches all your audio apps (Spotify, Chrome, VLC, etc.) to VoiceMeeter when SteamVR launches, back to your speakers when it closes. VRChat audio is never touched.
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
- **Settings menu** — configure excluded apps, adjust polling/debounce from the tray
- **Auto-updater** — checks for new versions in the background, one-click update from the tray
- **Clean uninstall** — remove shortcuts, configs, and state files from the tray menu

## How It Works

```
                      ┌──────────────────────────────┐
  Physical Mic ─────► │  VoiceMeeter Banana           │
  (your headset mic)  │                               │
                      │  Strip[0] (Mic) ──► B1 ───────┼──► VRChat (mic input)
                      │                               │
  Any Audio App ────► │  Strip[3] (VAIO) ──► B2 ──────┼──► VR Headset (you hear)
  (Spotify, YouTube,  │         └──► B1 (if Public) ──┼──► VRChat (friends hear)
   VLC, browser, etc) │                               │
                      │  Mixer controls Strip/Bus     │
                      │  gains independently          │
                      └──────────────────────────────┘

  VRChat is automatically excluded — its audio always stays on your headset.
  All other apps with active audio are switched together.
```

**Desktop mode:** All your apps output directly to your speakers (soundbar, headphones, etc). VoiceMeeter runs silently in the background with no effect.

**VR modes:** Every audio app on your system (except VRChat) is automatically routed to VoiceMeeter's virtual input (VAIO). VoiceMeeter sends the audio to B2 (your headset via "Listen to this device") and optionally B1 (VRChat mic). It doesn't matter if you're using Spotify, a browser, VLC, or anything else — they all get switched together. The Mixer UI controls the gain split so you can make music quieter for friends while keeping it loud for yourself.

## Requirements

- **Windows 10/11**
- **Python 3.10+** — [python.org/downloads](https://www.python.org/downloads/)
- **VoiceMeeter Banana** — [vb-audio.com/Voicemeeter/banana](https://vb-audio.com/Voicemeeter/banana.htm) (free)
- **Any VR headset** — Quest, Index, Vive, Rift, etc. Just needs its own audio output device that shows up in Windows

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

The tray icon appears in your system tray. Left-click to cycle modes, right-click for the full menu including Mixer, Settings, VRChat Mic Help, Check for Updates, and Uninstall.

**First launch:** You'll get a one-time reminder to set your VRChat microphone to "Voicemeeter Out B1". This is the only in-game step needed.

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
     "exclude_processes": ["vrchat.exe"],
     "svcl_path": "svcl.exe",
     "vr_device": "<VoiceMeeter VAIO Friendly ID from svcl>",
     "debounce_seconds": 5,
     "music_strip": 3,
     "vrchat_mic_confirmed": false
   }
   ```
   - `exclude_processes`: Apps whose audio should NOT be switched (VRChat is always excluded). All other audio apps are switched automatically.
   - `vr_device`: The svcl Friendly ID for "Voicemeeter Input" (look for `VB-Audio Voicemeeter VAIO\Device\Voicemeeter Input\Render`).
   - You can also add apps to the exclusion list later from the tray Settings menu.

5. Create `vm_devices.json` with your mic device name (as shown in VoiceMeeter):
   ```json
   {
     "Strip[0]": "Microphone (Your Mic Device Name)"
   }
   ```

6. Set up VoiceMeeter Banana:
   - **Strip[0]** hardware input: your microphone
   - **Strip[3]** (VAIO): this is where all app audio arrives — no setup needed
   - Enable "Listen to this device" on **Voicemeeter Out B2** in Windows Sound settings, targeting your VR headset's audio output

7. Run:
   ```bash
   pythonw vr_audio_switcher.py
   ```

## VoiceMeeter "Listen to this device" Setup

The setup wizard configures this automatically (with a UAC prompt). If it fails or you need to do it manually:

1. Open **Windows Sound Settings** → **More sound settings**
2. Go to the **Recording** tab
3. Find **Voicemeeter Out B2** (VB-Audio VoiceMeeter VAIO)
4. Right-click → **Properties** → **Listen** tab
5. Check **"Listen to this device"**
6. Set playback device to your **VR headset's audio output** (e.g. Steam Streaming Speakers, Index speakers, etc.)
7. Click OK

This setting persists across reboots — you only need to do it once.

## Modes Explained

| Mode | Icon | Audio Apps Output | Music in Headset | Music in VRChat Mic | Use Case |
|------|------|------------------|-----------------|--------------------|----|
| Desktop | Blue | Your speakers | No | No | Normal desktop use |
| Auto | Green | Follows SteamVR | When VR active | No | Default — hands-free switching |
| Private | Yellow | VoiceMeeter | Yes | No | Personal listening in VR |
| Public | Red | VoiceMeeter | Yes | Yes | DJ mode — friends hear your music |

## Mixer

Right-click the tray icon → **Mixer** to open the volume control panel.

**Volume sliders** (-100% to +100%):
- **Music for Others** — what friends hear through your VRChat mic (Public mode only)
- **Music for Me** — what you hear in your VR headset
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

**Audio apps not switching:**
- At least one audio app (Spotify, Chrome, etc.) must be playing or have an active audio session
- The app uses `svcl.exe` for per-app audio routing — make sure it's in the project directory
- Check `switcher.log` for errors
- Open Settings from the tray menu to check if the app is accidentally excluded

**"svcl.exe not found":**
Run `python setup_wizard.py` to download it, or manually download [svcl-x64.zip](https://www.nirsoft.net/utils/svcl-x64.zip) and extract `svcl.exe` to the project directory.

**Antivirus flags svcl.exe:**
NirSoft tools are sometimes flagged as false positives by antivirus software. svcl.exe is a legitimate Windows audio management utility. You may need to add an exception in your antivirus.

**Tray icon not visible:**
Windows may hide new tray icons. Click the **^** arrow in the system tray to find it, then drag it to the visible area.

## Architecture

| File | Purpose |
|------|---------|
| `vr_audio_switcher.py` | Main tray app — SteamVR detection, mode switching, VoiceMeeter control, settings, uninstall |
| `mixer.py` | Mixer UI — volume/EQ sliders, presets, drag-and-drop |
| `setup_wizard.py` | First-run setup — device detection, config generation, svcl download |
| `updater.py` | Auto-updater — checks GitHub releases, downloads and applies updates |
| `vm_path.py` | VoiceMeeter detection — finds DLL/EXE via Windows registry |
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
