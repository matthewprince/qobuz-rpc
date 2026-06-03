# Qobuz RPC

Discord Rich Presence for the Qobuz music streaming service. Shows what you're listening to with real per-track quality, album art, and a progress timer.

## What it does

- Detects what Qobuz is playing via the window title
- Looks up the track on the Qobuz API for actual quality metadata (bit depth, sample rate)
- Falls back to iTunes Search API if Qobuz API is unavailable
- Pushes album art, song title, artist, album name, and quality to Discord
- Tracks session stats (time listened, songs played)

## Quality detection

Quality is pulled from the Qobuz catalog per track, not a static setting. If you switch from a CD quality album to Hi-Res, the Discord status updates automatically.

## Requirements

- Windows 10/11
- Python 3.10+ (3.12 through 3.14 all work; every dependency ships prebuilt wheels)
- Discord desktop app
- Qobuz desktop app

## Setup

```
pip install -r requirements.txt
```

Or run `setup.bat`.

### Discord Application

1. Go to [Discord Developer Portal](https://discord.com/developers/applications)
2. Create a new application (name it whatever you want, e.g. "Qobuz")
3. Copy the **Application ID**
4. Optionally upload a Qobuz logo as `qobuz_icon` under Rich Presence > Art Assets

### Configure

Copy `config.example.json` to `config.json`, or just run the app and it'll create one.

Run `start.bat` (GUI) or `python qobuz_rpc_cli.py --setup` (CLI).

Enter your Discord Application ID, Qobuz email, and password. The password is MD5 hashed locally and stored in the **Windows Credential Manager** (not `config.json`). The hash is what Qobuz authenticates against, so keeping it out of plain files is deliberate; an existing hash from an older `config.json` is migrated into Credential Manager automatically on first run. (`config.json` is gitignored regardless.)

Qobuz credentials are optional. Without them the app falls back to iTunes for metadata (no per-track quality detection).

### Run

**GUI:** `start.bat` or `python qobuz_rpc.py`

**CLI:** `start_cli.bat` or `python qobuz_rpc_cli.py`

## Building an EXE

```
build.bat
```

Outputs `dist/QobuzRPC.exe` (GUI) and `dist/QobuzRPC-CLI.exe` (console). Requires PyInstaller.

## Options

- **Use Windows media session (SMTC)** - read Qobuz's media session for exact metadata and real position; falls back to the window title when off or unavailable
- **Auto-connect on launch** - connects automatically when the app starts
- **Minimize to tray on close** - hides to system tray instead of quitting
- **Start with Windows** - creates a startup script in your Startup folder

## How it works

By default the app reads the Windows media session (SMTC) that Qobuz publishes, which gives the exact track title, artist, album, real playback position, and play/pause state. If that's unavailable (Qobuz not reporting to SMTC, or `winrt` not installed) it falls back to reading the Qobuz desktop window title ("Track - Artist"). Either way it then searches the Qobuz API (`/track/search`) for that track to get the real `maximum_bit_depth` and `maximum_sampling_rate` from the catalog, plus album art from the Qobuz CDN. If the Qobuz API fails, it falls back to the iTunes Search API. You can turn the media-session source off in Settings to force the window-title method.

API credentials (`app_id` and `app_secret`) are extracted dynamically from the Qobuz web player's `bundle.js`, same method used by [QobuzApiSharp](https://github.com/DJDoubleD/QobuzApiSharp).

## Credits

Inspired by [Qobuz-RPC](https://github.com/Seeyaflying/Qobuz-RPC) by Seeyaflying.

Qobuz API integration ported from [QobuzApiSharp](https://github.com/DJDoubleD/QobuzApiSharp) by DJDoubleD.

## Disclaimer

This project is not affiliated with or endorsed by Qobuz. It uses the Qobuz API but does not include any app IDs or secrets. Credentials are fetched client-side from publicly available JavaScript. See [Qobuz API Terms of Use](http://static.qobuz.com/apps/api/QobuzAPI-TermsofUse.pdf).
