import base64, hashlib, json, re, threading, time, sys, os, io
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

try: import requests
except ImportError: print("[!] pip install requests"); sys.exit(1)
try: from pypresence import Presence, ActivityType, StatusDisplayType
except ImportError: print("[!] pip install pypresence"); sys.exit(1)
try: import psutil
except ImportError: print("[!] pip install psutil"); sys.exit(1)
try: import win32gui, win32process
except ImportError: print("[!] pip install pywin32"); sys.exit(1)
try: from PIL import Image, ImageTk, ImageDraw
except ImportError: print("[!] pip install Pillow"); sys.exit(1)

try:
    import pystray
    HAS_TRAY = True
except ImportError:
    HAS_TRAY = False

# figure out where we actually live (handles PyInstaller temp dir)
if getattr(sys, 'frozen', False):
    SCRIPT_DIR = os.path.dirname(sys.executable)
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_PATH = os.path.join(SCRIPT_DIR, "config.json")
ICON_ICO = os.path.join(SCRIPT_DIR, "icon.ico")
ICON_PNG = os.path.join(SCRIPT_DIR, "icon.png")

STARTUP_DIR = os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
STARTUP_VBS = os.path.join(STARTUP_DIR, "QobuzRPC.vbs")

DEFAULT_CFG = {
    "discord_app_id": "", "qobuz_email": "", "qobuz_pw_hash": "",
    "quality_label": "Hi-Res 24-Bit / 96 kHz", "update_interval": 3,
    "show_quality_badge": True, "fallback_cover": "",
    "auto_connect": False, "minimize_to_tray": False, "start_with_windows": False,
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

def set_autostart(on):
    try:
        if on:
            os.makedirs(STARTUP_DIR, exist_ok=True)
            if getattr(sys, 'frozen', False):
                # EXE mode - just run the exe
                exe = os.path.abspath(sys.executable)
                vbs = f'Set s = CreateObject("WScript.Shell")\ns.Run """{exe}""", 0, False'
            else:
                script = os.path.abspath(__file__)
                vbs = f'Set s = CreateObject("WScript.Shell")\ns.Run "pythonw ""{script}""", 0, False'
            with open(STARTUP_VBS, "w") as f: f.write(vbs)
        elif os.path.exists(STARTUP_VBS):
            os.remove(STARTUP_VBS)
    except OSError:
        pass


# qobuz api - ported from QobuzApiSharp by DJDoubleD

class QobuzAPI:
    BASE = "https://www.qobuz.com/api.json/0.2"
    WEB = "https://play.qobuz.com"

    def __init__(self):
        self.app_id = None
        self.app_secret = None
        self.user_auth_token = None
        self.s = requests.Session()
        self.s.headers["User-Agent"] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/110.0"
        self._bun = None

    def init(self, log=print):
        try:
            log("Fetching web player...")
            html = self.s.get(f"{self.WEB}/login", timeout=15).text

            # grab bundle.js url
            m = re.search(r'<script src="(/resources/\d+\.\d+\.\d+-[a-z]\d+/bundle\.js)"', html)
            if not m: m = re.search(r'<script[^>]+src="(/resources/[^"]*bundle[^"]*\.js)"', html)
            if not m:
                log("Couldn't find bundle.js")
                return False

            log("Downloading bundle.js...")
            self._bun = self.s.get(f"{self.WEB}{m.group(1)}", timeout=20).text

            # app id - production env
            m2 = re.search(r'production:\{api:\{appId:"([^"]+)",appSecret:', self._bun)
            if not m2:
                log("App ID not found in bundle")
                return False
            self.app_id = m2.group(1)
            self.s.headers["X-App-Id"] = self.app_id
            log(f"App ID: {self.app_id}")

            # app secret - seed + timezone + info/extras -> base64
            sm = re.search(r'\):[a-z]\.initialSeed\("([^"]+)",window\.utimezone\.([a-z]+)\)', self._bun)
            if sm:
                seed, tz = sm.group(1), sm.group(2)
                tz_cap = tz[0].upper() + tz[1:]
                im = re.search(r'timezones:\[.*?name:".*?/' + tz_cap + r'",info:"([^"]*)",extras:"([^"]*)"', self._bun)
                if im:
                    enc = seed + im.group(1) + im.group(2)
                    if len(enc) > 44:
                        try:
                            self.app_secret = base64.b64decode(enc[:-44]).decode("utf-8")
                            log("App secret extracted")
                        except: pass

            return True
        except Exception as e:
            log(f"Init failed: {e}")
            return False

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
                log("No auth token in response")
                return False
            self.user_auth_token = tok
            self.s.headers["X-User-Auth-Token"] = tok
            u = d.get("user", {})
            log(f"Logged in: {u.get('display_name') or u.get('login') or email}")
            sub = u.get("credential", {}).get("label", "")
            if sub: log(f"Subscription: {sub}")
            return True
        except Exception as e:
            log(f"Login error: {e}")
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

            bd = best.get("maximum_bit_depth") or 0
            sr = best.get("maximum_sampling_rate") or 0
            if sr > 1000: sr /= 1000

            ql = ""
            if bd and sr:
                ql = f"Hi-Res {int(bd)}-Bit / {sr:g} kHz" if bd >= 24 else f"CD {int(bd)}-Bit / {sr:g} kHz"

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


# itunes fallback

_it_cache = {}
def itunes_lookup(artist, track):
    k = f"{artist}||{track}".lower()
    if k in _it_cache: return _it_cache[k]
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


# image cache
_img_cache = {}
def get_img(url):
    if not url: return None
    if url in _img_cache: return _img_cache[url]
    try:
        r = requests.get(url, timeout=8); r.raise_for_status()
        _img_cache[url] = r.content; return r.content
    except:
        _img_cache[url] = None; return None


# window title stuff

def qobuz_title():
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

def parse(t):
    if not t or t.strip().lower() == "qobuz": return None
    p = t.split(" - ", 1)
    if len(p) == 2: return {"title": p[0].strip(), "artist": p[1].strip()}
    if p[0].strip().lower() != "qobuz": return {"title": p[0].strip(), "artist": "Unknown Artist"}
    return None


# misc helpers

def fmt(s):
    m, s = divmod(int(max(0, s)), 60); h, m = divmod(m, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m}:{s:02d}"

def mk_rounded(data, sz, r):
    img = Image.open(io.BytesIO(data)).resize((sz, sz), Image.LANCZOS)
    mask = Image.new("L", (sz, sz), 0)
    ImageDraw.Draw(mask).rounded_rectangle([(0, 0), (sz, sz)], radius=r, fill=255)
    out = Image.new("RGBA", (sz, sz), (0, 0, 0, 0))
    out.paste(img.convert("RGBA"), mask=mask)
    return ImageTk.PhotoImage(out)

def mk_placeholder(sz, r):
    img = Image.new("RGBA", (sz, sz), (26, 26, 30, 255))
    d = ImageDraw.Draw(img)
    cx, cy, br = sz//2, sz//2, sz//6
    d.ellipse([cx-br, cy+br//2, cx+br, cy+br//2+br*2], fill=(60, 60, 68))
    d.rectangle([cx+br-4, cy-br*2, cx+br, cy+br], fill=(60, 60, 68))
    mask = Image.new("L", (sz, sz), 0)
    ImageDraw.Draw(mask).rounded_rectangle([(0, 0), (sz, sz)], radius=r, fill=255)
    out = Image.new("RGBA", (sz, sz), (0, 0, 0, 0))
    out.paste(img, mask=mask)
    return ImageTk.PhotoImage(out)


# colors
BG      = "#0c0c0e"
CARD    = "#151518"
CARD2   = "#1a1a1e"
BORDER  = "#252528"
TXT     = "#e8e8ec"
DIM     = "#a0a0aa"
MUTED   = "#5a5a62"
BLUE    = "#3DA8FE"
BLUE_HI = "#62b8ff"
GREEN   = "#34d399"
RED     = "#f87171"
AMBER   = "#fbbf24"


class App:
    def __init__(self):
        self.cfg = load_cfg()
        self.qobuz = QobuzAPI()
        self.rpc = None
        self.rpc_ok = False
        self.monitoring = False

        # track state
        self.tkey = None
        self.tcover = None
        self.talbum = None
        self.tqual = ""
        self.tdur = 0
        self.tstart = 0.0
        self.turls = {}
        self.prev_raw = None
        self.playing = False
        self.qobuz_ok = False

        # stats
        self.sess_start = 0.0
        self.songs = 0
        self.listen_s = 0.0
        self.ltick = 0.0

        self.tray = None

        # fix taskbar icon grouping
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("qobuz.rpc")
        except: pass

        self.root = tk.Tk()
        self.root.title("Qobuz RPC")
        self.root.configure(bg=BG)
        self.root.resizable(False, True)
        self.root.geometry("480x920")
        self.root.minsize(480, 700)
        self.root.protocol("WM_DELETE_WINDOW", self._close)

        if os.path.exists(ICON_ICO):
            try: self.root.iconbitmap(ICON_ICO)
            except: pass
        if os.path.exists(ICON_PNG):
            try: self.root.iconphoto(True, ImageTk.PhotoImage(Image.open(ICON_PNG)))
            except: pass

        self.ph = mk_placeholder(150, 12)
        self.cov_img = None
        self._ui()
        self._load_fields()
        self._tick()

        if self.cfg.get("auto_connect"):
            self.root.after(500, self._start)

    def _ui(self):
        r = self.root

        # header
        hdr = tk.Frame(r, bg=BG, padx=24, pady=16); hdr.pack(fill="x")
        nf = tk.Frame(hdr, bg=BG); nf.pack(side="left")
        tk.Label(nf, text="Qobuz", font=("Segoe UI Semibold", 20), fg=TXT, bg=BG).pack(side="left")
        tk.Label(nf, text="RPC", font=("Segoe UI Light", 20), fg=BLUE, bg=BG).pack(side="left", padx=(5,0))
        sf = tk.Frame(hdr, bg=BG); sf.pack(side="right")
        self.dot = tk.Canvas(sf, width=8, height=8, bg=BG, highlightthickness=0)
        self.dot.pack(side="right", padx=(6,0))
        self._mkdot(MUTED)
        self.st_lbl = tk.Label(sf, text="Idle", font=("Segoe UI", 9), fg=MUTED, bg=BG)
        self.st_lbl.pack(side="right")

        tk.Frame(r, bg=BORDER, height=1).pack(fill="x", padx=24)

        # now playing card
        cpd = tk.Frame(r, bg=BG, padx=24, pady=18); cpd.pack(fill="x")
        card = tk.Frame(cpd, bg=CARD, highlightbackground=BORDER, highlightthickness=1); card.pack(fill="x")
        ci = tk.Frame(card, bg=CARD, padx=18, pady=18); ci.pack(fill="x")

        row = tk.Frame(ci, bg=CARD); row.pack(fill="x")
        self.cov = tk.Label(row, image=self.ph, bg=CARD, borderwidth=0); self.cov.pack(side="left")

        inf = tk.Frame(row, bg=CARD, padx=16); inf.pack(side="left", fill="both", expand=True)
        self.np = tk.Label(inf, text="Nothing Playing", font=("Segoe UI", 8, "bold"), fg=MUTED, bg=CARD, anchor="w")
        self.np.pack(anchor="w", pady=(0,5))
        self.l_title = tk.Label(inf, text=" ", font=("Segoe UI Semibold", 14), fg=TXT, bg=CARD, anchor="w", wraplength=230, justify="left")
        self.l_title.pack(anchor="w")
        self.l_artist = tk.Label(inf, text=" ", font=("Segoe UI", 10), fg=DIM, bg=CARD, anchor="w", wraplength=230, justify="left")
        self.l_artist.pack(anchor="w", pady=(2,0))
        self.l_album = tk.Label(inf, text=" ", font=("Segoe UI", 9), fg=MUTED, bg=CARD, anchor="w", wraplength=230, justify="left")
        self.l_album.pack(anchor="w", pady=(1,0))
        self.l_qual = tk.Label(inf, text="", font=("Segoe UI Semibold", 8), fg=BLUE, bg=CARD, anchor="w")
        self.l_qual.pack(anchor="w", pady=(6,0))

        # progress
        pf = tk.Frame(ci, bg=CARD); pf.pack(fill="x", pady=(14,0))
        self.bar = tk.Canvas(pf, height=3, bg=BORDER, highlightthickness=0, bd=0)
        self.bar.pack(fill="x", pady=(0,5))
        tf = tk.Frame(pf, bg=CARD); tf.pack(fill="x")
        self.l_pos = tk.Label(tf, text="0:00", font=("Segoe UI", 8), fg=MUTED, bg=CARD); self.l_pos.pack(side="left")
        self.l_dur = tk.Label(tf, text="0:00", font=("Segoe UI", 8), fg=MUTED, bg=CARD); self.l_dur.pack(side="right")

        # stats bar
        stf = tk.Frame(r, bg=CARD2, padx=24, pady=9); stf.pack(fill="x", padx=24)
        self.s_sess = tk.Label(stf, text="Session  0:00:00", font=("Consolas", 9), fg=DIM, bg=CARD2); self.s_sess.pack(side="left")
        self.s_songs = tk.Label(stf, text="0 songs", font=("Consolas", 9), fg=DIM, bg=CARD2); self.s_songs.pack(side="right")
        self.s_listen = tk.Label(stf, text="Listened  0:00:00", font=("Consolas", 9), fg=BLUE, bg=CARD2); self.s_listen.pack()

        # connect button
        bf = tk.Frame(r, bg=BG, padx=24); bf.pack(fill="x", pady=(12,12))
        self.btn = tk.Canvas(bf, height=42, bg=BLUE, highlightthickness=0, cursor="hand2"); self.btn.pack(fill="x")
        self.btn.bind("<Configure>", lambda e: self._rbtn())
        self.btn.bind("<Button-1>", self._toggle)
        self.btn.bind("<Enter>", lambda e: self.btn.configure(bg=BLUE_HI if not self.monitoring else "#ef4444"))
        self.btn.bind("<Leave>", lambda e: self.btn.configure(bg=RED if self.monitoring else BLUE))
        self._btxt = "Connect"

        # settings
        tk.Frame(r, bg=BORDER, height=1).pack(fill="x", padx=24)
        sh = tk.Frame(r, bg=BG, padx=24); sh.pack(fill="x", pady=(14,8))
        tk.Label(sh, text="Settings", font=("Segoe UI Semibold", 11), fg=TXT, bg=BG).pack(anchor="w")

        sf = tk.Frame(r, bg=BG, padx=24); sf.pack(fill="x")

        self.v_app = tk.StringVar()
        self._mkfield(sf, "Discord Application ID", self.v_app)
        self.v_email = tk.StringVar()
        self._mkfield(sf, "Qobuz Email", self.v_email)
        self.v_pw = tk.StringVar()
        self._mkfield(sf, "Qobuz Password", self.v_pw, show="\u2022")

        tk.Label(sf, text="Fallback Quality", font=("Segoe UI", 9), fg=DIM, bg=BG).pack(anchor="w", pady=(4,2))
        self.v_ql = tk.StringVar()
        ttk.Combobox(sf, textvariable=self.v_ql, values=[
            "Hi-Res 24-Bit / 192 kHz", "Hi-Res 24-Bit / 96 kHz",
            "CD 16-Bit / 44.1 kHz", "MP3 320 kbps",
        ], font=("Segoe UI", 9)).pack(fill="x", pady=(0,10), ipady=3)

        # checkboxes
        ckf = tk.Frame(sf, bg=BG); ckf.pack(fill="x", pady=(2,10))
        self.v_auto = tk.BooleanVar()
        self.v_tray = tk.BooleanVar()
        self.v_startup = tk.BooleanVar()
        for v, txt in [(self.v_auto, "Auto-connect on launch"), (self.v_tray, "Minimize to tray on close"), (self.v_startup, "Start with Windows")]:
            tk.Checkbutton(ckf, text=txt, variable=v, font=("Segoe UI", 9), fg=DIM, bg=BG,
                selectcolor=CARD2, activebackground=BG, activeforeground=DIM).pack(anchor="w")

        # save btn
        svr = tk.Frame(sf, bg=BG); svr.pack(fill="x", pady=(0,10))
        sv = tk.Canvas(svr, height=32, bg=CARD2, width=120, highlightthickness=1, highlightbackground=BORDER, cursor="hand2")
        sv.pack(side="right")
        sv.create_text(60, 16, text="Save Settings", font=("Segoe UI", 9), fill=TXT)
        sv.bind("<Button-1>", self._save)
        sv.bind("<Enter>", lambda e: sv.configure(bg=CARD))
        sv.bind("<Leave>", lambda e: sv.configure(bg=CARD2))

        # log
        tk.Frame(r, bg=BORDER, height=1).pack(fill="x", padx=24)
        lf = tk.Frame(r, bg=BG, padx=24); lf.pack(fill="both", expand=True, pady=(10,16))
        self.logw = tk.Text(lf, height=4, font=("Consolas", 8), bg=CARD, fg=MUTED,
            relief="flat", borderwidth=0, wrap="word", highlightbackground=BORDER,
            highlightthickness=1, state="disabled", padx=10, pady=8)
        self.logw.pack(fill="both", expand=True)

    def _mkfield(self, parent, label, var, show=None):
        tk.Label(parent, text=label, font=("Segoe UI", 9), fg=DIM, bg=BG).pack(anchor="w", pady=(4,2))
        kw = dict(textvariable=var, font=("Segoe UI", 10), bg=CARD2, fg=TXT,
            insertbackground=TXT, relief="flat", borderwidth=0, highlightbackground=BORDER, highlightthickness=1)
        if show: kw["show"] = show
        tk.Entry(parent, **kw).pack(fill="x", pady=(0,8), ipady=5, ipadx=8)

    def _mkdot(self, c):
        self.dot.delete("all"); self.dot.create_oval(0, 0, 8, 8, fill=c, outline=c)

    def _rbtn(self):
        self.btn.delete("all")
        self.btn.create_text(self.btn.winfo_width()//2, 21, text=self._btxt, font=("Segoe UI Semibold", 10), fill="#fff")

    def _setbtn(self, t, bg):
        self._btxt = t; self.btn.configure(bg=bg); self._rbtn()

    # config <-> ui
    def _load_fields(self):
        self.v_app.set(self.cfg.get("discord_app_id", ""))
        self.v_email.set(self.cfg.get("qobuz_email", ""))
        self.v_ql.set(self.cfg.get("quality_label", "Hi-Res 24-Bit / 96 kHz"))
        self.v_auto.set(self.cfg.get("auto_connect", False))
        self.v_tray.set(self.cfg.get("minimize_to_tray", False))
        self.v_startup.set(self.cfg.get("start_with_windows", False))
        # show dots if password hash exists so user knows it's saved
        if self.cfg.get("qobuz_pw_hash", ""):
            self.v_pw.set("saved")
            self._pw_is_placeholder = True
        else:
            self._pw_is_placeholder = False

    def _read_fields(self):
        self.cfg["discord_app_id"] = self.v_app.get().strip()
        self.cfg["qobuz_email"] = self.v_email.get().strip()
        self.cfg["quality_label"] = self.v_ql.get().strip()
        self.cfg["auto_connect"] = self.v_auto.get()
        self.cfg["minimize_to_tray"] = self.v_tray.get()
        self.cfg["start_with_windows"] = self.v_startup.get()
        pw = self.v_pw.get().strip()
        # only hash if user typed a new password (not the placeholder)
        if pw and not (pw == "saved" and self._pw_is_placeholder):
            self.cfg["qobuz_pw_hash"] = hashlib.md5(pw.encode()).hexdigest()
            self.v_pw.set("saved")
            self._pw_is_placeholder = True

    def _save(self, _e=None):
        self._read_fields()
        save_cfg(self.cfg)
        set_autostart(self.cfg.get("start_with_windows", False))
        self.log("Settings saved")

    def log(self, msg):
        self.logw.configure(state="normal")
        self.logw.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.logw.see("end")
        self.logw.configure(state="disabled")

    # discord rpc
    def _connect_rpc(self):
        aid = self.cfg.get("discord_app_id", "").strip()
        if not aid: self.log("No Discord App ID"); return False
        try:
            self.rpc = Presence(aid); self.rpc.connect()
            self.rpc_ok = True; self.log("Discord connected"); return True
        except Exception as e:
            self.log(f"Discord failed: {e}"); return False

    def _disconnect_rpc(self):
        if self.rpc and self.rpc_ok:
            try: self.rpc.clear(); self.rpc.close()
            except: pass
        self.rpc_ok = False

    def _push_rpc(self, title, artist, album, cover, quality):
        if not self.rpc_ok: return
        kw = {
            "activity_type": ActivityType.LISTENING,
            "status_display_type": StatusDisplayType.DETAILS,
            "details": title[:128],
            "large_image": cover or self.cfg.get("fallback_cover") or "qobuz_icon",
        }
        if artist: kw["state"] = f"by {artist}"[:128]
        if album: kw["large_text"] = album[:128]
        if self.tdur and self.tstart > 0:
            kw["start"] = int(self.tstart)
            kw["end"] = int(self.tstart + self.tdur / 1000)
        elif self.tstart > 0:
            kw["start"] = int(self.tstart)
        if self.cfg.get("show_quality_badge", True) and quality:
            kw["small_image"] = "qobuz_icon"; kw["small_text"] = quality

        btns = []
        urls = self.turls or {}
        if urls.get("track"): btns.append({"label": "Open Track", "url": urls["track"]})
        if urls.get("album"): btns.append({"label": "Open Album", "url": urls["album"]})
        elif urls.get("artist"): btns.append({"label": "Open Artist", "url": urls["artist"]})
        if btns: kw["buttons"] = btns[:2]

        try: self.rpc.update(**kw)
        except Exception as e: self.log(f"RPC error: {e}"); self.rpc_ok = False

    # 1s tick for progress + stats
    def _tick(self):
        try:
            now = time.time()

            if self.playing and self.ltick > 0:
                dt = now - self.ltick
                if 0 < dt < 10: self.listen_s += dt
            self.ltick = now if self.playing else 0

            # only touch UI widgets if window is actually visible
            visible = False
            try: visible = self.root.winfo_viewable()
            except: pass

            if visible and self.tkey and self.tstart > 0 and self.playing:
                el = now - self.tstart
                ds = self.tdur / 1000.0
                self.l_pos.config(text=fmt(el))
                self.bar.delete("all")
                w = self.bar.winfo_width()
                if w > 1:
                    if ds > 0:
                        self.l_dur.config(text=fmt(ds))
                        self.bar.create_rectangle(0, 0, int(min(el/ds, 1)*w), 3, fill=BLUE, outline="")
                    else:
                        self.l_dur.config(text="")
                        self.bar.create_rectangle(0, 0, w, 3, fill=BLUE, outline="")

            if visible and self.sess_start > 0:
                self.s_sess.config(text=f"Session  {fmt(now - self.sess_start)}")
                self.s_listen.config(text=f"Listened  {fmt(self.listen_s)}")
                self.s_songs.config(text=f"{self.songs} song{'s' if self.songs != 1 else ''}")
        except:
            pass

        self.root.after(1000, self._tick)

    # main loop
    def _monitor(self):
        while self.monitoring:
            try:
                raw = qobuz_title()

                if raw is None:
                    if self.playing or self.tkey:
                        self.root.after(0, lambda: self.log("Qobuz not detected"))
                        self._reset(); self.root.after(0, self._clear_np)
                    try:
                        if self.rpc and self.rpc_ok: self.rpc.clear()
                    except: pass

                else:
                    p = parse(raw)
                    if p:
                        k = f"{p['title']}|{p['artist']}"
                        self.playing = True

                        flick = self.prev_raw and not parse(self.prev_raw) and k == self.tkey
                        looped = k == self.tkey and self.tdur > 0 and self.tstart > 0 and time.time() - self.tstart > self.tdur/1000 + 5
                        new = k != self.tkey

                        if new or flick or looped:
                            self.tstart = time.time()
                            self.songs += 1

                            if new:
                                self.tkey = k
                                t, a = p["title"], p["artist"]
                                self.root.after(0, lambda tt=t, ta=a: self.log(f"Playing: {tt}  {ta}"))

                                meta = None
                                if self.qobuz_ok: meta = self.qobuz.search(t, a)
                                if not meta: meta = itunes_lookup(a, t)

                                if meta:
                                    self.tcover = meta.get("cover")
                                    self.talbum = meta.get("album", "")
                                    self.tqual = meta.get("quality", "")
                                    self.tdur = meta.get("duration_ms", 0)
                                    self.turls = {
                                        "track": meta.get("track_url", ""),
                                        "album": meta.get("album_url", ""),
                                        "artist": meta.get("artist_url", ""),
                                    }
                                    src = meta.get("src", "")
                                    self.root.after(0, lambda s=src, q=self.tqual:
                                        self.log(f"[{s}] {q}" if q else f"[{s}] loaded"))
                                    if self.tcover:
                                        threading.Thread(target=self._fetch_cover, args=(self.tcover,), daemon=True).start()
                                else:
                                    self.tcover = None; self.talbum = ""
                                    self.tqual = self.cfg.get("quality_label", ""); self.tdur = 0
                                    self.turls = {}

                                self.root.after(0, lambda: self._set_np(p["title"], p["artist"], self.talbum, self.tqual))

                            elif flick:
                                self.root.after(0, lambda: self.log("Restarted"))
                            elif looped:
                                self.root.after(0, lambda: self.log("Looped"))

                        self._push_rpc(p["title"], p["artist"], self.talbum or "", self.tcover, self.tqual)

                    else:
                        if self.playing:
                            self.root.after(0, lambda: self.log("Paused"))
                            self.playing = False; self.tstart = 0
                            self.root.after(0, lambda: self.np.config(text="PAUSED", fg=AMBER))
                        try:
                            if self.rpc and self.rpc_ok: self.rpc.clear()
                        except: pass

                    self.prev_raw = raw

            except Exception as e:
                self.root.after(0, lambda err=e: self.log(f"Error: {err}"))
                if not self.rpc_ok: self._connect_rpc()

            time.sleep(self.cfg.get("update_interval", 3))

    def _reset(self):
        self.tkey = self.tcover = self.talbum = self.prev_raw = None
        self.tqual = ""; self.tdur = 0; self.tstart = 0; self.playing = False

    # ui updates
    def _set_np(self, title, artist, album, quality):
        self.l_title.config(text=title); self.l_artist.config(text=artist)
        self.l_album.config(text=album or " "); self.l_qual.config(text=quality)
        self.np.config(text="NOW PLAYING", fg=GREEN)

    def _fetch_cover(self, url):
        d = get_img(url)
        if d:
            try:
                ph = mk_rounded(d, 150, 12)
                self.root.after(0, lambda p=ph: self._setcov(p))
            except: pass

    def _setcov(self, ph):
        self.cov_img = ph; self.cov.config(image=ph)

    def _clear_np(self):
        self.l_title.config(text=" "); self.l_artist.config(text=" ")
        self.l_album.config(text=" "); self.l_qual.config(text="")
        self.np.config(text="Nothing Playing", fg=MUTED)
        self.l_pos.config(text="0:00"); self.l_dur.config(text="0:00")
        self.bar.delete("all"); self.cov.config(image=self.ph); self.cov_img = None

    # connect / disconnect
    def _toggle(self, _e=None):
        if not self.monitoring: self._start()
        else: self._stop()

    def _start(self):
        self._read_fields(); save_cfg(self.cfg)
        set_autostart(self.cfg.get("start_with_windows", False))

        if not self.cfg.get("discord_app_id", "").strip():
            messagebox.showwarning("Missing App ID", "Enter your Discord Application ID first.")
            return

        self.sess_start = time.time(); self.songs = 0; self.listen_s = 0; self.ltick = 0
        self.log("Starting...")

        email = self.cfg.get("qobuz_email", "").strip()
        pw = self.cfg.get("qobuz_pw_hash", "").strip()
        if email and pw:
            self.log("Initializing Qobuz API...")
            if self.qobuz.init(log=self.log):
                self.qobuz_ok = self.qobuz.login(email, pw, log=self.log)
                if not self.qobuz_ok: self.log("Falling back to iTunes")
            else:
                self.log("API init failed, iTunes mode"); self.qobuz_ok = False
        else:
            self.log("No Qobuz creds, iTunes mode"); self.qobuz_ok = False

        if not self._connect_rpc():
            self._setstatus("Failed", RED); return

        self.monitoring = True
        self._setstatus("Connected", GREEN); self._setbtn("Disconnect", RED)
        threading.Thread(target=self._monitor, daemon=True).start()

    def _stop(self):
        self.monitoring = False; self._disconnect_rpc()
        self._clear_np(); self._reset(); self.qobuz_ok = False
        self._setstatus("Idle", MUTED); self._setbtn("Connect", BLUE)
        self.log("Disconnected")

    def _setstatus(self, t, c):
        self.st_lbl.config(text=t, fg=c); self._mkdot(c)

    # tray
    def _mktray(self):
        if not HAS_TRAY: return
        try:
            ico = Image.open(ICON_PNG) if os.path.exists(ICON_PNG) else Image.new("RGBA", (64,64), BLUE)
            menu = pystray.Menu(
                pystray.MenuItem("Show", lambda: self.root.after(0, self._show)),
                pystray.MenuItem("Quit", lambda: self.root.after(0, self._quit)),
            )
            self.tray = pystray.Icon("QobuzRPC", ico, "Qobuz RPC", menu)
            threading.Thread(target=self.tray.run, daemon=True).start()
        except: pass

    def _show(self):
        self.root.deiconify(); self.root.lift(); self.root.focus_force()

    def _quit(self):
        self.monitoring = False; self._disconnect_rpc()
        if self.tray:
            try: self.tray.stop()
            except: pass
        self.root.destroy()

    def _close(self):
        if self.cfg.get("minimize_to_tray") and HAS_TRAY:
            self.root.withdraw()
            if not self.tray: self._mktray()
        else:
            self._quit()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    App().run()
