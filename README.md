# VR Audio Switcher

**One-click audio management for VRChat.**

You play VRChat on PC and want music from *any* app (Spotify, YouTube, SoundCloud, VLC, whatever) to play through your VRChat mic so friends can hear it. You also want independent volume control over what *they* hear vs. what *you* hear, and you want everything to switch back to your desktop speakers automatically when you take off the headset. That's exactly what this does.

Works with any VR headset: Quest (via Steam Link or Virtual Desktop), Index, Vive, Rift, or anything else with its own audio output.

## What It Does

- **Auto-detects VR** by watching for a configurable process (defaults to `vrserver.exe`, works with SteamVR, Virtual Desktop, and others). Boots VoiceMeeter when VR starts, shuts it down when VR stops. Switches all your audio apps automatically.
- **3 audio modes** with buttons in the mixer UI:
  - **Desktop** (green): music plays through your speakers, VoiceMeeter bypassed
  - **Private** (yellow): music in your headset only, mic stays clean
  - **Public** (red): music plays through your VRChat mic AND your headset
- **Mixer UI** with independent sliders for:
  - Music volume for others (what friends hear)
  - Music volume for you (what you hear in your headset)
  - Your voice volume in VRChat
  - Bass / Mid / Treble EQ
- **Preset system** for saving, loading, renaming, and reordering audio profiles with drag-and-drop
- **VR-synced lifecycle**: everything boots and shuts down with SteamVR automatically
- **Tabbed interface**: Mixer, Guide, Settings, and Help all in one window
- **Auto-updater** that checks for new versions on boot and from the Settings tab
- **Clean uninstall** via `uninstall.bat`

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

  VRChat is automatically excluded. Its audio always stays on your headset.
  All other apps with active audio are switched together.
```

**Desktop mode:** All your apps output directly to your speakers (soundbar, headphones, etc). VoiceMeeter runs silently in the background with no effect.

**VR modes:** Every audio app on your system (except VRChat) is automatically routed to VoiceMeeter's virtual input (VAIO). VoiceMeeter sends the audio to B2 (your headset via "Listen to this device") and optionally B1 (VRChat mic). Doesn't matter if you're using Spotify, a browser, VLC, or anything else. They all get switched together. The Mixer UI controls the gain split so you can make music quieter for friends while keeping it loud for yourself.

## Requirements

- **Windows 10/11**
- **Python 3.10+**: [python.org/downloads](https://www.python.org/downloads/)
- **VoiceMeeter Banana** (free): [vb-audio.com/Voicemeeter/banana](https://vb-audio.com/Voicemeeter/banana.htm)
- **Any VR headset**: Quest, Index, Vive, Rift, etc. Just needs its own audio output device that shows up in Windows

## Quick Start

### 1. Download

[**Download VR Audio Switcher (.zip)**](https://github.com/Aetheriju/vr-audio-switcher/archive/refs/heads/main.zip)

### 2. Unblock and extract

Windows blocks files downloaded from the internet. You need to unblock the ZIP **before extracting**:

1. Open your **Downloads** folder
2. Find the file called **vr-audio-switcher-main.zip** (it may show as just **vr-audio-switcher-main** if extensions are hidden)
3. **Right-click the ZIP file** (not a folder, the file you just downloaded). A menu will appear. Click **Properties**
4. At the bottom of the General tab, check the **Unblock** checkbox → click **OK**
5. Now right-click the ZIP file again → **Extract All** → **Extract**
6. Open the extracted **vr-audio-switcher-main** folder (double-click into it)

You should see files like `install.bat`, `README.md`, `setup_wizard.py`, etc. If file extensions are hidden, they'll show as `install`, `README`, `setup_wizard`, etc. That's fine.

> **If you already extracted without unblocking:** Delete the extracted folder, go back to the ZIP in Downloads, unblock it (steps 3-4 above), then extract again.

### 4. Run the installer

Double-click **install** (it may show as `install.bat`, same file, Windows just hides the extension sometimes). A green terminal window will open and the installer handles everything automatically:

- Downloads and installs **Python** if you don't have it
- Downloads **svcl.exe** (NirSoft per-app audio routing tool)
- Installs Python packages
- Downloads and runs the **VoiceMeeter Banana** installer if needed
- Detects your audio devices
- Configures "Listen to this device" for VR audio routing
- Creates desktop and startup shortcuts

### 5. You're done

The app runs silently in the background and springs to life when SteamVR starts. It launches VoiceMeeter, opens the mixer UI, routes all your audio, and shuts everything down cleanly when VR stops.

**First launch:** The Guide tab opens automatically with setup instructions. Set your VRChat microphone to "Voicemeeter Out B1" (this only needs to be done once).

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
     "vr_process": "vrserver.exe",
     "exclude_processes": ["vrchat.exe"],
     "svcl_path": "svcl.exe",
     "vr_device": "<VoiceMeeter VAIO Friendly ID from svcl>",
     "debounce_seconds": 5,
     "music_strip": 3,
     "vrchat_mic_confirmed": false
   }
   ```
   - `vr_process`: The process name that indicates VR is running. Defaults to `vrserver.exe` (SteamVR). Virtual Desktop and most PCVR launchers start this process automatically. If your setup uses something different, change it here.
   - `exclude_processes`: Apps whose audio should NOT be switched (VRChat is always excluded). All other audio apps are switched automatically.
   - `vr_device`: The svcl Friendly ID for "Voicemeeter Input" (look for `VB-Audio Voicemeeter VAIO\Device\Voicemeeter Input\Render`).
   - You can also add apps to the exclusion list later from the Settings tab.

5. Create `vm_devices.json` with your mic device name (as shown in VoiceMeeter):
   ```json
   {
     "Strip[0]": "Microphone (Your Mic Device Name)"
   }
   ```

6. Set up VoiceMeeter Banana:
   - **Strip[0]** hardware input: your microphone
   - **Strip[3]** (VAIO): this is where all app audio arrives, no setup needed
   - Enable "Listen to this device" on **Voicemeeter Out B2** in Windows Sound settings, targeting your VR headset's audio output

7. Run:
   ```bash
   pythonw vr_audio_switcher.py
   ```

## VoiceMeeter "Listen to this device" Setup

The setup wizard configures this automatically (with a UAC prompt). If it fails or you need to do it manually:

1. Open **Windows Sound Settings** > **More sound settings**
2. Go to the **Recording** tab
3. Find **Voicemeeter Out B2** (VB-Audio VoiceMeeter VAIO)
4. Right-click > **Properties** > **Listen** tab
5. Check **"Listen to this device"**
6. Set playback device to your **VR headset's audio output** (e.g. Steam Streaming Speakers, Index speakers, etc.)
7. Click OK

This setting persists across reboots. You only need to do it once.

## Modes Explained

| Mode | Button | Audio Apps Output | Music in Headset | Music in VRChat Mic | Use Case |
|------|--------|------------------|-----------------|--------------------|----|
| Desktop | Green | Your speakers | No | No | Normal desktop use |
| Private | Yellow | VoiceMeeter | Yes | No | Personal listening in VR |
| Public | Red | VoiceMeeter | Yes | Yes | DJ mode, friends hear your music |

## Mixer

The mixer opens automatically when SteamVR starts. Use the tabs to switch between Mixer, Guide, Settings, and Help.

**Volume sliders** (-100% to +100%):
- **Music for Others**: what friends hear through your VRChat mic (Public mode only)
- **Music for Me**: what you hear in your VR headset
- **My Voice**: your mic volume in VRChat

**EQ sliders** (0% to 100%):
- Bass, Mid, Treble. Shape the music before it hits VRChat.

**Presets:**
- Click a preset to load it instantly
- Drag presets to reorder them
- Double-click a preset name to rename it
- Use the save icon to overwrite a preset with current settings
- "Save Current" creates a new preset from your current slider positions

## Troubleshooting

**No audio after VoiceMeeter restart:**
The app automatically saves and restores VoiceMeeter device assignments and gain settings. If audio is missing, close the app (X button > Shut Down Everything) and relaunch.

**Mic not working in VRChat:**
- Make sure VRChat's mic input is set to **Voicemeeter Out B1** (not your physical mic)
- Check that Strip[0] in VoiceMeeter shows your physical mic as the input device
- Verify Strip[0].B1 routing is enabled (the app enforces this every 5 seconds)

**Audio apps not switching:**
- At least one audio app (Spotify, Chrome, etc.) must be playing or have an active audio session
- The app uses `svcl.exe` for per-app audio routing. Make sure it's in the project directory.
- Check `switcher.log` for errors
- Open the Settings tab to check if the app is accidentally excluded

**"svcl.exe not found":**
Run `python setup_wizard.py` to download it, or manually download [svcl-x64.zip](https://www.nirsoft.net/utils/svcl-x64.zip) and extract `svcl.exe` to the project directory.

**Antivirus flags svcl.exe:**
NirSoft tools are sometimes flagged as false positives by antivirus software. svcl.exe is a legitimate Windows audio management utility. You may need to add an exception in your antivirus.

## Architecture

| File | Purpose |
|------|---------|
| `vr_audio_switcher.py` | Core app: VR-synced lifecycle, VoiceMeeter control, audio switching |
| `mixer.py` | Tabbed UI: mixer sliders, guide, settings, help |
| `splash.py` | Boot/shutdown splash screen with status updates |
| `setup_wizard.py` | First-run setup: device detection, config generation, svcl download |
| `updater.py` | Auto-updater: checks GitHub releases, downloads and applies updates |
| `vm_path.py` | VoiceMeeter detection: finds DLL/EXE via Windows registry |
| `config.json` | User config: device IDs, polling settings |
| `vm_devices.json` | Persisted VoiceMeeter device assignments |
| `vm_state.json` | Persisted VoiceMeeter gain/EQ values |
| `presets.json` | User-saved mixer presets |
| `state.json` | Runtime state (current mode, VR active) |

## Credits

- **[VoiceMeeter Banana](https://vb-audio.com/Voicemeeter/banana.htm)** by VB-Audio, the audio routing engine
- **[SoundVolumeCommandLine (svcl)](https://www.nirsoft.net/utils/sound_volume_command_line.html)** by NirSoft, per-app audio device switching

## License

GPL v3. Free for everyone, forever. See [LICENSE](LICENSE).
