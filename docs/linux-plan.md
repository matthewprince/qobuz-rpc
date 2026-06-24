# Linux port - design and plan

Status: design only, nothing implemented yet. The runtime path cannot be tested on the dev machine (Windows), so the plan ends with a smoke test on a real Linux box before merge.

## Context / reality

- There is no official Qobuz desktop app for Linux. People play Qobuz on Linux through the **web player** (play.qobuz.com) in a browser, or a third-party client.
- The cross-platform "what's playing" standard is **MPRIS** (`org.mpris.MediaPlayer2`) over the D-Bus session bus. Browsers (Chromium, Firefox) publish an MPRIS player when they are playing audio, and most native media players do too.
- `pypresence` (Discord RPC) already works on Linux: it connects to the Discord IPC socket under `$XDG_RUNTIME_DIR` (or `/tmp`).
- The architecture already fits: a Linux source just needs to return the same `sample` dict that SMTC/scraper return, then the existing tested `decide()` and the existing Discord push handle the rest.

## Already cross-platform (no change needed)

`decide()`, `parse`, `fmt`, `quality_str`, `QobuzAPI`, the iTunes fallback, the Discord push (pypresence), the Tkinter GUI, Pillow.

## Windows-only, must be gated behind `sys.platform == "win32"`

- imports: `win32gui`, `win32process` (window scraping), `win32cred` (creds), `winrt` (SMTC)
- `qobuz_title()` window-title scraper
- `smtc_now()`
- `set_autostart()` (writes a Startup `.vbs`)
- `SetCurrentProcessExplicitAppUserModelID` and taskbar icon (already in try/except)

Current blocker: the GUI hard-imports `win32gui` and calls `sys.exit()` if it is missing, so the module will not even import on Linux. First fix is to import the win32/winrt modules only on Windows.

## New Linux pieces

### 1. `mpris_now()` - the Linux now-playing source

- D-Bus library: **jeepney** (pure-Python, has a blocking API, pip-installable, no system libraries to compile). Alternatives considered: `dbus-next` (async, would need an event loop the way SMTC does) and `dbus-python` (needs libdbus headers + a compile step). jeepney's blocking API fits the synchronous monitor loop best.
- Logic per poll: list bus names beginning with `org.mpris.MediaPlayer2.`; for the chosen one read the `org.mpris.MediaPlayer2.Player` properties:
  - `PlaybackStatus` -> Playing / Paused / Stopped
  - `Metadata` -> `xesam:title`, `xesam:artist` (array, join), `xesam:album`, `mpris:length` (microseconds -> duration), `xesam:url`, `mpris:artUrl`
  - `Position` (microseconds -> position). MPRIS Position is current-at-read, so `tstart = now - pos`, same as SMTC. No drift correction needed.
- Returns the same shape: `{status, title, artist, album, pos, dur}` or `None`.

### 2. Player matching - the main open problem

On Windows the SMTC filter matches an app id containing "qobuz" (the desktop app). On Linux there is no Qobuz app, so the MPRIS player is usually the **browser** ("chromium", "firefox"), whose bus name does not contain "qobuz" and whose `xesam:url` may be a blob or site URL. Proposed resolution order:

1. A player whose bus name or `Metadata` url/artUrl contains "qobuz" (works if the web player exposes a qobuz CDN/site URL; uncertain, needs checking on a real box).
2. A user-configured `mpris_player` setting (substring of the bus name, e.g. "firefox", "chromium", or a specific client). Default empty.
3. Fall back to the first player whose `PlaybackStatus` is Playing.

Expose `mpris_player` in Settings so the user can pin their browser/client.

### 3. Autostart

Linux branch of `set_autostart()` writes `~/.config/autostart/QobuzRPC.desktop` (with `X-GNOME-Autostart-enabled=true`) instead of the Startup `.vbs`.

### 4. Credentials

`win32cred` is `None` off-Windows, so `cred_set()` already returns False and `set_pw_hash()` already falls back to storing the hash in `config.json`. Add `chmod 0600` on `config.json` on non-Windows for a little protection. Future option: `keyring` / Secret Service.

### 5. Source selection in `_sample()`

- Windows: `smtc_now()` (if enabled) -> window-title scraper -> None
- Linux: `mpris_now()` (if enabled) -> None (no scraper equivalent)

### 6. requirements.txt environment markers

So pip does not try to install Windows-only wheels on Linux (and vice versa):

```
pywin32>=306;                         sys_platform == "win32"
winrt-Windows.Media.Control>=2.0;     sys_platform == "win32"
winrt-Windows.Foundation>=2.0;        sys_platform == "win32"
winrt-Windows.Foundation.Collections>=2.0; sys_platform == "win32"
jeepney>=0.8;                         sys_platform == "linux"
```

(`pystray` is cross-platform but the Linux tray needs a backend such as `python3-Xlib` or PyGObject+AppIndicator; treat the tray as optional on Linux.)

## CI

Add `ubuntu-latest` to the matrix:
- CLI imports fine on Linux (no tkinter).
- GUI imports `tkinter`; ubuntu runners need a `sudo apt-get install -y python3-tk` step, or guard the GUI import test to Windows.
- Run pytest on Linux too. The pure tests (parse/fmt/quality_str/decide) are platform-independent and will catch any cross-platform regression in `decide()`.

This proves the refactor keeps Linux importable and the logic correct, even though the MPRIS to Discord loop is only smoke-tested by hand.

## What can and cannot be verified before merge

- Can verify (CI + dev machine): module imports on Linux, pure-logic tests, no Windows regression, the Windows exe still builds.
- Cannot verify here (needs a Linux box + Discord + Qobuz playing in a browser): that `mpris_now()` picks the right player and the presence renders. The big unknown is player matching.

## Open questions

1. How do you play Qobuz on Linux - browser web player (which browser?) or a specific client? This sets the default for player matching.
2. OK to add an "MPRIS player" setting (substring match) so it is user-selectable?
3. Tray on Linux: worth the extra backend dependency, or drop the tray on Linux for the first version?
4. Linux packaging: a PyInstaller one-file binary (build.sh mirroring build.bat), or just run-from-source / pipx?

## Rollout

Branch `linux-mpris` -> implement -> CI green on ubuntu + windows -> you smoke-test on Linux -> merge to main.
