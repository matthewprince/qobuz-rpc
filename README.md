# Qobuz RPC

Discord Rich Presence for the Qobuz music streaming service. Shows what you're listening to with real per-track quality, album art, and a progress timer. Runs on **Windows and Linux**.

## What it does

- Detects what Qobuz is playing from the system media session
- Looks up the track on the Qobuz API for actual quality metadata (bit depth, sample rate)
- Falls back to iTunes Search API if Qobuz API is unavailable
- Pushes album art, song title, artist, album name, and quality to Discord
- Tracks session stats (time listened, songs played)

## Quality detection

Quality is pulled from the Qobuz catalog per track, not a static setting. If you switch from a CD quality album to Hi-Res, the Discord status updates automatically.

## Requirements

- Python 3.10+ (3.12 through 3.14 all work)
- Discord desktop app
- **Windows 10/11**: the Qobuz desktop app
- **Linux**: a player that publishes MPRIS (the Qobuz web player in a browser, or a third-party client) and a D-Bus session bus. The GUI also needs Tk (`sudo apt-get install python3-tk` or your distro's equivalent).

## Setup

```
pip install -r requirements.txt
```

`requirements.txt` uses platform markers, so pip only installs the bits for your OS (pywin32/winrt/pystray on Windows, jeepney on Linux).

Windows users can run `setup.bat` instead; Linux/macOS users can run `bash setup.sh`.

### Credentials (mostly automatic)

In the common case you don't enter anything.

- **Discord**: a Discord application id is a public client id, not a secret, so the app ships with a built-in one. You don't register your own. The presence uses the album cover as its image, so no Discord art assets need to be uploaded. (Maintainers/self-builders: set `DEFAULT_DISCORD_APP_ID` in `qobuz_core.py` to your own "Qobuz" app's id.)
- **Qobuz (Windows)**: if the Qobuz desktop app is installed and signed in (which it is, since you use it to play music), the app reuses that existing session token automatically. No email or password.

Optional fallbacks, only if the automatic path isn't available (the Linux web player, or no Qobuz desktop app):

- Enter your Qobuz email and password once in settings or via `python qobuz_rpc_cli.py --setup`. The password is MD5 hashed locally; on Windows the hash goes in the **Credential Manager**, on Linux into `config.json` (created `0600`, owner-only). `config.json` is gitignored.
- With no Qobuz auth at all, the app still runs and falls back to the iTunes Search API for metadata and art (no per-track quality badge).
- You can also override the Discord id with your own in settings.

`config.json` is created automatically on first run (or copy `config.example.json`).

### Run

| | Windows | Linux |
|---|---|---|
| GUI | `start.bat` or `python qobuz_rpc.py` | `bash start.sh` or `python3 qobuz_rpc.py` |
| CLI | `start_cli.bat` or `python qobuz_rpc_cli.py` | `bash start_cli.sh` or `python3 qobuz_rpc_cli.py` |

On Linux the CLI is the lightest option (no Tk needed). `pipx install` from source works too.

## Linux: which player gets tracked

There is no official Qobuz desktop app for Linux, so the "now playing" data comes from whatever MPRIS player is active, usually your browser playing [play.qobuz.com](https://play.qobuz.com). The app picks a player in this order:

1. A player whose bus name or track art/URL mentions `qobuz`
2. A player whose bus name contains your configured **MPRIS player** substring (e.g. `firefox`, `chromium`)
3. The first player that is currently playing

If you also use other MPRIS players (Spotify, VLC, etc.) and want to pin one, set `mpris_player` in `config.json` or the GUI field to a substring of its bus name.

## Building a standalone binary

- **Windows**: `build.bat` outputs `dist/QobuzRPC.exe` (GUI) and `dist/QobuzRPC-CLI.exe` (console).
- **Linux**: `bash build.sh` outputs `dist/QobuzRPC` and `dist/QobuzRPC-CLI`.

Both require PyInstaller. Linux binaries are tied to the build machine's glibc; for distribution, running from source or `pipx` is usually friendlier.

## Options

- **Use media session** - read the system media session for exact metadata and real position. This is SMTC on Windows and MPRIS on Linux; falls back to the Qobuz window title on Windows when off or unavailable.
- **Auto-connect on launch** - connects automatically when the app starts
- **Minimize to tray on close** - hides to system tray instead of quitting (Windows; Linux only if a `pystray` backend is installed)
- **Start with Windows / Start on login** - registers an autostart entry (Startup `.vbs` on Windows, `~/.config/autostart/*.desktop` on Linux)

## How it works

The app reads the system media session that the player publishes, which gives the exact track title, artist, album, real playback position, and play/pause state. On Windows that's SMTC (the Qobuz desktop app's session); on Linux it's MPRIS over the D-Bus session bus. On Windows, if SMTC is unavailable it falls back to reading the Qobuz desktop window title ("Track - Artist").

Either way it then searches the Qobuz API (`/track/search`) for that track to get the real `maximum_bit_depth` and `maximum_sampling_rate` from the catalog, plus album art from the Qobuz CDN. If the Qobuz API fails, it falls back to the iTunes Search API.

API credentials (`app_id` and `app_secret`) are extracted dynamically from the Qobuz web player's `bundle.js`, same method used by [QobuzApiSharp](https://github.com/DJDoubleD/QobuzApiSharp).

### Code layout

- `qobuz_core.py` - shared logic: config, credentials, the Qobuz API, the iTunes fallback, the pure state machine, and the platform "now playing" sources (SMTC + window title on Windows, MPRIS on Linux)
- `qobuz_rpc.py` - the Tk GUI
- `qobuz_rpc_cli.py` - the console version

## Credits

Inspired by [Qobuz-RPC](https://github.com/Seeyaflying/Qobuz-RPC) by Seeyaflying.

Qobuz API integration ported from [QobuzApiSharp](https://github.com/DJDoubleD/QobuzApiSharp) by DJDoubleD.

## Disclaimer

This project is not affiliated with or endorsed by Qobuz. It uses the Qobuz API but does not include any app IDs or secrets. Credentials are fetched client-side from publicly available JavaScript. See [Qobuz API Terms of Use](http://static.qobuz.com/apps/api/QobuzAPI-TermsofUse.pdf).
