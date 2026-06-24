# Shared core for Qobuz RPC - config, credentials, the Qobuz API, the iTunes
# fallback, the pure state machine, and the platform "now playing" sources.
# Both the GUI (qobuz_rpc.py) and the CLI (qobuz_rpc_cli.py) import from here so
# the logic lives in exactly one place. Windows-only and Linux-only bits are
# gated behind sys.platform so the module imports cleanly on either OS.

import asyncio, json, os, re, sys, time

try: import requests
except ImportError: print("[!] pip install requests"); sys.exit(1)

IS_WIN = sys.platform.startswith("win")
IS_LINUX = sys.platform.startswith("linux")
IS_MAC = sys.platform == "darwin"

# --- Windows-only imports (window scraping, Credential Manager, SMTC) ---------
if IS_WIN:
    try: import win32gui, win32process
    except Exception: win32gui = win32process = None
    try: import win32cred
    except Exception: win32cred = None
    try:
        from winrt.windows.media.control import GlobalSystemMediaTransportControlsSessionManager as _SMTCMgr
        HAS_SMTC = True
    except Exception:
        HAS_SMTC = False
else:
    win32gui = win32process = win32cred = None
    HAS_SMTC = False

# --- Linux-only imports (D-Bus / MPRIS via jeepney) ---------------------------
if IS_LINUX:
    try:
        from jeepney import DBusAddress, new_method_call
        from jeepney.io.blocking import open_dbus_connection
        HAS_MPRIS = True
    except Exception:
        HAS_MPRIS = False
else:
    HAS_MPRIS = False


# figure out where we actually live (handles PyInstaller temp dir)
if getattr(sys, "frozen", False):
    SCRIPT_DIR = os.path.dirname(sys.executable)
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_PATH = os.path.join(SCRIPT_DIR, "config.json")

# A Discord application id is a public client id (not a secret), so the app can
# ship its own and users never have to register one. A per-user
# config["discord_app_id"] overrides it. Leave "" to fall back to asking.
DEFAULT_DISCORD_APP_ID = "1519198716417020004"   # shared "Qobuz" Discord application

# The Qobuz desktop app stores its logged-in session token scoped to a desktop
# app id. Reusing that token lets search authenticate with zero typing. These
# ids are known to pair with the desktop token; extraction adds any others.
QOBUZ_DESKTOP_APP_IDS = ["304027809", "425621600"]

DEFAULT_CFG = {
    "discord_app_id": "", "qobuz_email": "", "qobuz_pw_hash": "",
    "quality_label": "Hi-Res 24-Bit / 96 kHz", "update_interval": 3,
    "show_quality_badge": True, "fallback_cover": "",
    "auto_connect": False, "minimize_to_tray": False, "start_with_windows": False,
    "use_smtc": True,            # Windows: read the SMTC media session
    "use_mpris": True,           # Linux: read the MPRIS media session
    "mpris_player": "",          # Linux: substring of the MPRIS bus name to pin (e.g. "firefox")
    "use_desktop_session": True, # reuse the Qobuz desktop app's login instead of asking for one
}


def load_cfg():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH) as f: return {**DEFAULT_CFG, **json.load(f)}
    # first run - create config.json from example or defaults
    example = os.path.join(SCRIPT_DIR, "config.example.json")
    if os.path.exists(example):
        with open(example) as f: cfg = {**DEFAULT_CFG, **json.load(f)}
    else:
        cfg = dict(DEFAULT_CFG)
    save_cfg(cfg)
    return cfg

def save_cfg(c):
    with open(CONFIG_PATH, "w") as f: json.dump(c, f, indent=2)
    # config.json may hold the password hash on platforms without a keystore;
    # lock it down to the owner there.
    if not IS_WIN:
        try: os.chmod(CONFIG_PATH, 0o600)
        except OSError: pass


# --- credentials --------------------------------------------------------------
# On Windows the qobuz password hash lives in the Credential Manager, never in
# config.json. Off-Windows there is no keystore wired up yet, so cred_set returns
# False and the hash falls back into config.json (which save_cfg chmods to 0600).
CRED_TARGET = "QobuzRPC"

def cred_get():
    if not (IS_WIN and win32cred): return ""
    try:
        c = win32cred.CredRead(CRED_TARGET, win32cred.CRED_TYPE_GENERIC)
        b = c.get("CredentialBlob") or b""
        return b.decode("utf-16-le") if isinstance(b, (bytes, bytearray)) else str(b)
    except Exception:
        return ""

def cred_set(h):
    if not (IS_WIN and win32cred) or not h: return False
    try:
        win32cred.CredWrite({"Type": win32cred.CRED_TYPE_GENERIC, "TargetName": CRED_TARGET,
            "UserName": "qobuz", "CredentialBlob": h, "Persist": win32cred.CRED_PERSIST_LOCAL_MACHINE}, 0)
        return True
    except Exception:
        return False

def get_pw_hash(cfg):
    h = cred_get()
    if h: return h
    legacy = (cfg.get("qobuz_pw_hash") or "").strip()   # migrate an old hash from config.json
    if legacy:
        if cred_set(legacy): cfg["qobuz_pw_hash"] = ""; save_cfg(cfg)
        return legacy
    return ""

def set_pw_hash(cfg, h):
    if cred_set(h): cfg["qobuz_pw_hash"] = ""
    else: cfg["qobuz_pw_hash"] = h


# --- existing Qobuz desktop session (zero-config auth) ------------------------
# The Qobuz desktop app (Windows) is an Electron app that keeps the logged-in
# user and session token in %APPDATA%/Qobuz/localuser.json. Reusing that token
# lets the app authenticate catalog search without the user typing anything.
# There is no official Qobuz desktop app on Linux/macOS, so this is Windows-only.
def _desktop_userfile():
    if IS_WIN:
        p = os.path.join(os.environ.get("APPDATA", ""), "Qobuz", "localuser.json")
        return p if os.path.exists(p) else None
    return None

def read_desktop_session():
    # {"token","email","login","sub"} from the Qobuz desktop app, or None
    p = _desktop_userfile()
    if not p: return None
    try:
        with open(p, encoding="utf-8") as f: d = json.load(f)
        tok = (d.get("token") or "").strip()
        if not tok: return None
        return {"token": tok,
            "email": (d.get("email") or d.get("login") or "").strip(),
            "login": (d.get("login") or "").strip(),
            "sub": ((d.get("credential") or {}).get("description") or "").strip()}
    except Exception:
        return None

def _scan_desktop_app_ids():
    # best-effort: pull app ids out of the desktop install's JS so a version bump
    # that rotates the id still works. The known ids are always tried first.
    out = []
    if not IS_WIN: return out
    import glob
    base = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Qobuz")
    try:
        for p in glob.glob(os.path.join(base, "**", "*.js"), recursive=True):
            try:
                with open(p, encoding="utf-8", errors="ignore") as f: txt = f.read()
            except Exception:
                continue
            for m in re.findall(r'appId["\':=\s]{1,4}(\d{6,})', txt):
                if m not in out: out.append(m)
    except Exception:
        pass
    return out


# --- autostart ----------------------------------------------------------------
def _autostart_target():
    # the command to relaunch this app at login, as (exe, script-or-None)
    if getattr(sys, "frozen", False):
        return os.path.abspath(sys.executable), None
    script = os.path.abspath(sys.argv[0]) if sys.argv and sys.argv[0] else ""
    return sys.executable, script

def set_autostart(on):
    try:
        if IS_WIN:
            startup = os.path.join(os.environ.get("APPDATA", ""),
                "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
            vbs_path = os.path.join(startup, "QobuzRPC.vbs")
            if on:
                os.makedirs(startup, exist_ok=True)
                if getattr(sys, "frozen", False):
                    exe = os.path.abspath(sys.executable)
                    vbs = f'Set s = CreateObject("WScript.Shell")\ns.Run """{exe}""", 0, False'
                else:
                    pyw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
                    if not os.path.exists(pyw): pyw = "pythonw"
                    script = os.path.abspath(sys.argv[0])
                    vbs = f'Set s = CreateObject("WScript.Shell")\ns.Run """{pyw}"" ""{script}""", 0, False'
                with open(vbs_path, "w") as f: f.write(vbs)
            elif os.path.exists(vbs_path):
                os.remove(vbs_path)
        elif IS_LINUX:
            d = os.path.expanduser("~/.config/autostart")
            desktop = os.path.join(d, "QobuzRPC.desktop")
            if on:
                os.makedirs(d, exist_ok=True)
                exe, script = _autostart_target()
                cmd = exe if script is None else f"{exe} {script}"
                with open(desktop, "w") as f:
                    f.write("[Desktop Entry]\nType=Application\nName=Qobuz RPC\n"
                        f"Exec={cmd}\nX-GNOME-Autostart-enabled=true\nTerminal=false\n")
            elif os.path.exists(desktop):
                os.remove(desktop)
        # macOS: not implemented in this version
    except OSError:
        pass


# --- qobuz api - ported from QobuzApiSharp by DJDoubleD -----------------------
class QobuzAPI:
    BASE = "https://www.qobuz.com/api.json/0.2"
    WEB = "https://play.qobuz.com"

    def __init__(self):
        self.app_id = None
        self.user_auth_token = None
        self.s = requests.Session()
        self.s.headers["User-Agent"] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/110.0"
        self._bun = None

    def init(self, log=print):
        try:
            log("Fetching web player...")
            html = self.s.get(f"{self.WEB}/login", timeout=15).text
            m = re.search(r'<script src="(/resources/\d+\.\d+\.\d+-[a-z]\d+/bundle\.js)"', html)
            if not m: m = re.search(r'<script[^>]+src="(/resources/[^"]*bundle[^"]*\.js)"', html)
            if not m:
                log("Couldn't find bundle.js"); return False
            log("Downloading bundle.js...")
            self._bun = self.s.get(f"{self.WEB}{m.group(1)}", timeout=20).text
            m2 = re.search(r'production:\{api:\{appId:"([^"]+)",appSecret:', self._bun)
            if not m2:
                log("App ID not found in bundle"); return False
            self.app_id = m2.group(1)
            self.s.headers["X-App-Id"] = self.app_id
            log(f"App ID: {self.app_id}")
            return True
        except Exception as e:
            log(f"Init failed: {e}"); return False

    def login(self, email, pw_md5, log=print):
        if not self.app_id: return False
        try:
            r = self.s.get(f"{self.BASE}/user/login", params={"email": email, "password": pw_md5}, timeout=15)
            if not r.ok:
                try: log(f"Login failed: {r.json().get('message', r.status_code)}")
                except: log(f"Login failed: HTTP {r.status_code}")
                return False
            d = r.json()
            tok = d.get("user_auth_token")
            if not tok:
                log("No auth token in response"); return False
            self.user_auth_token = tok
            self.s.headers["X-User-Auth-Token"] = tok
            u = d.get("user", {})
            log(f"Logged in: {u.get('display_name') or u.get('login') or email}")
            sub = u.get("credential", {}).get("label", "")
            if sub: log(f"Subscription: {sub}")
            return True
        except Exception as e:
            log(f"Login error: {e}"); return False

    def _probe_app_id(self, app_id):
        # does (the current token + this app_id) authenticate catalog search?
        try:
            r = self.s.get(f"{self.BASE}/track/search",
                params={"query": "a", "limit": "1"}, timeout=10, headers={"X-App-Id": app_id})
            if r.status_code == 200:
                self.app_id = app_id
                self.s.headers["X-App-Id"] = app_id
                return True
        except Exception:
            pass
        return False

    def use_desktop_session(self, session):
        # authenticate by reusing the Qobuz desktop app's existing token, paired
        # with whichever desktop app id it was issued under
        tok = (session or {}).get("token")
        if not tok: return False
        self.user_auth_token = tok
        self.s.headers["X-User-Auth-Token"] = tok
        for aid in QOBUZ_DESKTOP_APP_IDS:
            if self._probe_app_id(aid): return True
        for aid in _scan_desktop_app_ids():
            if aid not in QOBUZ_DESKTOP_APP_IDS and self._probe_app_id(aid): return True
        self.user_auth_token = None
        self.s.headers.pop("X-User-Auth-Token", None)
        return False

    def search(self, title, artist):
        if not self.app_id: return None
        try:
            r = self.s.get(f"{self.BASE}/track/search",
                params={"query": f"{artist} {title}", "limit": "5", "offset": "0"}, timeout=10)
            r.raise_for_status()
            items = r.json().get("tracks", {}).get("items", [])
            if not items: return None

            # try to match artist exactly
            best = items[0]
            for t in items:
                if (t.get("performer") or {}).get("name", "").lower() == artist.lower():
                    best = t; break

            alb = best.get("album") or {}
            img = alb.get("image") or {}
            cover = img.get("mega") or img.get("extralarge") or img.get("large") or img.get("small") or ""
            if cover and not cover.startswith("http"):
                cover = f"https:{cover}" if cover.startswith("//") else ""

            ql = quality_str(best.get("maximum_bit_depth"), best.get("maximum_sampling_rate"))

            track_id = best.get("id")
            album_id = alb.get("id")
            artist_id = (best.get("performer") or {}).get("id") or (alb.get("artist") or {}).get("id")
            track_url = f"https://play.qobuz.com/track/{track_id}" if track_id else ""
            album_url = f"https://play.qobuz.com/album/{album_id}" if album_id else ""
            artist_url = f"https://play.qobuz.com/artist/{artist_id}" if artist_id else ""

            return {
                "title": best.get("title") or title,
                "artist": (best.get("performer") or {}).get("name") or artist,
                "album": alb.get("title") or "",
                "cover": cover or None,
                "duration_ms": int((best.get("duration") or 0) * 1000),
                "quality": ql, "src": "Qobuz",
                "track_url": track_url, "album_url": album_url, "artist_url": artist_url,
            }
        except:
            return None


def authenticate_qobuz(qz, cfg, log=print):
    # Set up a QobuzAPI for catalog search and return the auth mode used:
    #   "desktop" - reused the Qobuz desktop app's existing session (no typing)
    #   "login"   - manual email + password fallback
    #   "none"    - no Qobuz auth; caller should use the iTunes fallback
    if not qz.init(log=log):
        return "none"
    web_app_id = qz.app_id   # init() set this from the web player bundle
    if cfg.get("use_desktop_session", True):
        sess = read_desktop_session()
        if sess and qz.use_desktop_session(sess):
            log(f"Using Qobuz desktop session: {sess.get('email') or sess.get('login') or 'signed in'}")
            return "desktop"
    # restore the web app id for the manual-login fallback
    qz.app_id = web_app_id
    qz.s.headers["X-App-Id"] = web_app_id
    qz.user_auth_token = None
    qz.s.headers.pop("X-User-Auth-Token", None)
    email = cfg.get("qobuz_email", "").strip()
    pw = get_pw_hash(cfg)
    if email and pw and qz.login(email, pw, log=log):
        return "login"
    return "none"


# --- itunes fallback ----------------------------------------------------------
_it_cache = {}
def itunes_lookup(artist, track):
    k = f"{artist}||{track}".lower()
    if k in _it_cache: return _it_cache[k]
    if len(_it_cache) > 200: _it_cache.pop(next(iter(_it_cache)))
    try:
        r = requests.get("https://itunes.apple.com/search",
            params={"term": f"{artist} {track}", "entity": "song", "limit": "5"}, timeout=6)
        items = r.json().get("results", [])
        if not items:
            _it_cache[k] = None; return None
        best = items[0]
        for i in items:
            if i.get("artistName", "").lower() == artist.lower(): best = i; break
        art = best.get("artworkUrl100", "").replace("100x100bb", "600x600bb")
        out = {
            "title": best.get("trackName", track), "artist": best.get("artistName", artist),
            "album": best.get("collectionName", ""), "cover": art or None,
            "duration_ms": best.get("trackTimeMillis", 0), "quality": "", "src": "iTunes",
        }
        _it_cache[k] = out; return out
    except:
        _it_cache[k] = None; return None


# --- image cache --------------------------------------------------------------
_img_cache = {}
def get_img(url):
    if not url: return None
    if url in _img_cache: return _img_cache[url]
    if len(_img_cache) > 128: _img_cache.pop(next(iter(_img_cache)))
    try:
        r = requests.get(url, timeout=8); r.raise_for_status()
        _img_cache[url] = r.content; return r.content
    except:
        _img_cache[url] = None; return None


# --- misc helpers -------------------------------------------------------------
def fmt(s):
    m, s = divmod(int(max(0, s)), 60); h, m = divmod(m, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m}:{s:02d}"

def quality_str(bd, sr):
    bd = bd or 0; sr = sr or 0
    if sr > 1000: sr /= 1000
    if not (bd and sr): return ""
    return f"{'Hi-Res' if bd >= 24 else 'CD'} {int(bd)}-Bit / {sr:g} kHz"

def parse(t):
    if not t or t.strip().lower() == "qobuz": return None
    p = t.split(" - ", 1)
    if len(p) == 2 and p[0].strip() and p[1].strip():
        return {"title": p[0].strip(), "artist": p[1].strip()}
    return None

def decide(st, sample, now):
    # pure state machine, shared by every source, so it can be tested.
    # st keys: tkey, tstart, tdur, playing, pause_pos, prev_status
    # sample: {status,title,artist,album,pos,dur} or None (nothing playing / source gone)
    # returns (event, new_st); event in: gone, pause, new, resume, seek, loop, tick, None
    st = dict(st)
    if sample is None:
        ev = "gone" if (st["playing"] or st["tkey"]) else None
        st["playing"] = False; st["prev_status"] = "gone"
        return ev, st
    if sample["status"] == "paused":
        ev = "pause" if st["playing"] else None
        pos = sample.get("pos")
        pp = pos if pos is not None else (now - st["tstart"] if st["tstart"] > 0 else 0.0)
        st["playing"] = False; st["tstart"] = 0.0
        st["pause_pos"] = max(0.0, pp); st["prev_status"] = "paused"
        return ev, st
    k = f"{sample['title']}|{sample['artist']}"
    pos = sample.get("pos")
    was = st["prev_status"]
    st["playing"] = True; st["prev_status"] = "playing"
    if k != st["tkey"]:
        st["tkey"] = k; st["tstart"] = now - (pos or 0.0)
        return "new", st
    if was == "paused":
        st["tstart"] = now - (pos if pos is not None else st["pause_pos"])
        return "resume", st
    if pos is not None:
        nt = now - pos
        if abs(nt - st["tstart"]) > 2: st["tstart"] = nt; return "seek", st
        return "tick", st
    if st["tdur"] > 0 and st["tstart"] > 0 and now - st["tstart"] > st["tdur"]/1000 + 5:
        st["tstart"] = now
        return "loop", st
    return "tick", st


# --- Windows source: window title --------------------------------------------
def qobuz_title():
    try: import psutil
    except Exception: return None
    if win32gui is None: return None
    pids = []
    for p in psutil.process_iter(["pid", "name"]):
        try:
            if p.info["name"] and p.info["name"].lower() == "qobuz.exe":
                pids.append(p.info["pid"])
        except: pass
    if not pids: return None
    titles = []
    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd): return
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid in pids:
                t = win32gui.GetWindowText(hwnd)
                if t and len(t) > 1: titles.append(t)
        except: pass
    try: win32gui.EnumWindows(cb, None)
    except: pass
    return max(titles, key=len) if titles else None


# --- Windows source: SMTC media session --------------------------------------
# exact title/artist/album + real position & play/pause for the qobuz session.
_smtc_loop = None
def _smtc_run(coro):
    global _smtc_loop
    if _smtc_loop is None: _smtc_loop = asyncio.new_event_loop()
    return _smtc_loop.run_until_complete(coro)

async def _smtc_read():
    mgr = await _SMTCMgr.request_async()
    sess = None
    for s in mgr.get_sessions():
        if "qobuz" in (s.source_app_user_model_id or "").lower(): sess = s; break
    if sess is None: return None
    status = int(sess.get_playback_info().playback_status)   # 4 = playing, 5 = paused
    if status not in (4, 5): return None
    props = await sess.try_get_media_properties_async()
    title = (props.title or "").strip()
    if not title: return None
    tl = sess.get_timeline_properties()
    pos = tl.position.total_seconds() if tl.position else 0.0
    dur = tl.end_time.total_seconds() if tl.end_time else 0.0
    if status == 4 and tl.last_updated_time:   # nudge a stale position up to wall-clock
        pos += max(0.0, time.time() - tl.last_updated_time.timestamp())
    if dur > 0: pos = min(pos, dur)
    return {"status": "playing" if status == 4 else "paused", "title": title,
        "artist": props.artist or "", "album": props.album_title or "", "pos": pos, "dur": dur}

def smtc_now():
    if not HAS_SMTC: return None
    try: return _smtc_run(_smtc_read())
    except: return None


# --- Linux source: MPRIS media session ---------------------------------------
# Reads the same shape SMTC does from an org.mpris.MediaPlayer2 player over the
# D-Bus session bus. There is no Qobuz desktop app on Linux, so the player is
# usually a browser (web player) or a third-party client; see _mpris_choose.
_mpris_state = {"conn": None}

def _mpris_conn():
    if _mpris_state["conn"] is None:
        _mpris_state["conn"] = open_dbus_connection(bus="SESSION")
    return _mpris_state["conn"]

def _mpris_reset():
    c = _mpris_state.get("conn"); _mpris_state["conn"] = None
    try:
        if c: c.close()
    except Exception: pass

def _mpris_list(conn):
    addr = DBusAddress("/org/freedesktop/DBus", bus_name="org.freedesktop.DBus",
        interface="org.freedesktop.DBus")
    reply = conn.send_and_get_reply(new_method_call(addr, "ListNames"))
    return [n for n in reply.body[0] if n.startswith("org.mpris.MediaPlayer2.")]

def _mpris_getall(conn, name):
    # org.freedesktop.DBus.Properties.GetAll on the Player interface.
    # Returns the raw a{sv} dict {prop: (signature, value)}.
    addr = DBusAddress("/org/mpris/MediaPlayer2", bus_name=name,
        interface="org.freedesktop.DBus.Properties")
    reply = conn.send_and_get_reply(
        new_method_call(addr, "GetAll", "s", ("org.mpris.MediaPlayer2.Player",)))
    return reply.body[0]

def _unwrap(v):
    # jeepney returns variants as (signature, value) tuples
    return v[1] if isinstance(v, tuple) and len(v) == 2 else v

def mpris_sample(status, metadata, position_us):
    # pure transform from MPRIS properties to the unified sample dict, so it can
    # be unit-tested without a live D-Bus. metadata values are already unwrapped.
    if status not in ("Playing", "Paused"): return None
    title = (metadata.get("xesam:title") or "").strip()
    if not title: return None
    artist = metadata.get("xesam:artist") or ""
    if isinstance(artist, (list, tuple)):
        artist = ", ".join(a for a in artist if a)
    artist = (artist or "").strip()
    album = (metadata.get("xesam:album") or "").strip()
    dur = (metadata.get("mpris:length") or 0) / 1_000_000.0
    pos = (position_us or 0) / 1_000_000.0
    if dur > 0: pos = min(pos, dur)
    return {"status": "playing" if status == "Playing" else "paused",
        "title": title, "artist": artist, "album": album, "pos": pos, "dur": dur}

def _mpris_choose(players, pref):
    # players: list of {name, status, metadata, pos}. Resolution order:
    #   1. a player whose bus name or track url/artUrl mentions qobuz
    #   2. a player whose bus name contains the configured `mpris_player` substring
    #   3. the first player that is actually Playing
    players = [p for p in players if p]
    if not players: return None
    for p in players:
        hay = (p["name"] + " " + str(p["metadata"].get("xesam:url", "")) + " "
            + str(p["metadata"].get("mpris:artUrl", ""))).lower()
        if "qobuz" in hay: return p
    pref = (pref or "").strip().lower()
    if pref:
        for p in players:
            if pref in p["name"].lower(): return p
    for p in players:
        if p["status"] == "Playing": return p
    return None

def mpris_now(cfg):
    if not (IS_LINUX and HAS_MPRIS): return None
    try:
        conn = _mpris_conn()
        names = _mpris_list(conn)
        if not names: return None
        players = []
        for n in names:
            try:
                ap = {k: _unwrap(v) for k, v in _mpris_getall(conn, n).items()}
                meta = {k: _unwrap(v) for k, v in (ap.get("Metadata") or {}).items()}
                players.append({"name": n, "status": ap.get("PlaybackStatus"),
                    "metadata": meta, "pos": ap.get("Position")})
            except Exception:
                continue
        chosen = _mpris_choose(players, cfg.get("mpris_player", ""))
        if not chosen: return None
        return mpris_sample(chosen["status"], chosen["metadata"], chosen["pos"])
    except Exception:
        _mpris_reset()   # drop the cached connection so the next poll reconnects
        return None


# --- unified now-playing ------------------------------------------------------
def now_playing(cfg):
    # returns the unified sample dict {status,title,artist,album,pos,dur} or None
    if IS_WIN:
        if cfg.get("use_smtc", True):
            s = smtc_now()
            if s is not None: return s
        raw = qobuz_title()
        if raw is None: return None
        p = parse(raw)
        if p: return {"status": "playing", "title": p["title"], "artist": p["artist"],
            "album": None, "pos": None, "dur": None}
        return {"status": "paused", "title": "", "artist": "", "album": None, "pos": None, "dur": None}
    if IS_LINUX:
        if cfg.get("use_mpris", True):
            return mpris_now(cfg)
        return None
    return None
