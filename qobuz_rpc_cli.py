import base64, hashlib, json, re, signal, sys, os, time
try: import requests
except ImportError: print("[!] pip install requests"); sys.exit(1)
try: from pypresence import Presence
except ImportError: print("[!] pip install pypresence"); sys.exit(1)
try: import psutil
except ImportError: print("[!] pip install psutil"); sys.exit(1)
try: import win32gui, win32process
except ImportError: print("[!] pip install pywin32"); sys.exit(1)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "config.json")

DEFAULT_CFG = {
    "discord_app_id": "", "qobuz_email": "", "qobuz_pw_hash": "",
    "quality_label": "Hi-Res 24-Bit / 96 kHz", "update_interval": 3,
    "show_quality_badge": True, "fallback_cover": "",
}

def load_cfg():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH) as f: return {**DEFAULT_CFG, **json.load(f)}
    return dict(DEFAULT_CFG)

def save_cfg(c):
    with open(CONFIG_PATH, "w") as f: json.dump(c, f, indent=2)


class QobuzAPI:
    BASE = "https://www.qobuz.com/api.json/0.2"
    WEB = "https://play.qobuz.com"

    def __init__(self):
        self.app_id = None; self.app_secret = None; self.user_auth_token = None
        self.s = requests.Session()
        self.s.headers["User-Agent"] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/110.0"
        self._bun = None

    def init(self, log=print):
        try:
            log("[*] Fetching web player...")
            html = self.s.get(f"{self.WEB}/login", timeout=15).text
            m = re.search(r'<script src="(/resources/\d+\.\d+\.\d+-[a-z]\d+/bundle\.js)"', html)
            if not m: m = re.search(r'<script[^>]+src="(/resources/[^"]*bundle[^"]*\.js)"', html)
            if not m: log("[!] bundle.js not found"); return False
            log("[*] Downloading bundle.js...")
            self._bun = self.s.get(f"{self.WEB}{m.group(1)}", timeout=20).text

            m2 = re.search(r'production:\{api:\{appId:"([^"]+)",appSecret:', self._bun)
            if not m2: log("[!] App ID not found"); return False
            self.app_id = m2.group(1)
            self.s.headers["X-App-Id"] = self.app_id
            log(f"[+] App ID: {self.app_id}")

            sm = re.search(r'\):[a-z]\.initialSeed\("([^"]+)",window\.utimezone\.([a-z]+)\)', self._bun)
            if sm:
                seed, tz = sm.group(1), sm.group(2)
                im = re.search(r'timezones:\[.*?name:".*?/' + tz[0].upper()+tz[1:] + r'",info:"([^"]*)",extras:"([^"]*)"', self._bun)
                if im:
                    enc = seed + im.group(1) + im.group(2)
                    if len(enc) > 44:
                        try: self.app_secret = base64.b64decode(enc[:-44]).decode("utf-8"); log("[+] App secret extracted")
                        except: pass
            return True
        except Exception as e:
            log(f"[!] Init failed: {e}"); return False

    def login(self, email, pw_md5, log=print):
        if not self.app_id: return False
        try:
            r = self.s.get(f"{self.BASE}/user/login", params={"email": email, "password": pw_md5}, timeout=15)
            if not r.ok:
                try: log(f"[!] Login failed: {r.json().get('message', r.status_code)}")
                except: log(f"[!] Login failed: HTTP {r.status_code}")
                return False
            d = r.json(); tok = d.get("user_auth_token")
            if not tok: return False
            self.user_auth_token = tok; self.s.headers["X-User-Auth-Token"] = tok
            u = d.get("user", {}); log(f"[+] Logged in: {u.get('display_name') or u.get('login') or email}")
            sub = u.get("credential", {}).get("label", "")
            if sub: log(f"[+] Sub: {sub}")
            return True
        except Exception as e:
            log(f"[!] Login error: {e}"); return False

    def search(self, title, artist):
        if not self.app_id: return None
        try:
            r = self.s.get(f"{self.BASE}/track/search",
                params={"query": f"{artist} {title}", "limit": "5", "offset": "0"}, timeout=10)
            r.raise_for_status()
            items = r.json().get("tracks", {}).get("items", [])
            if not items: return None
            best = items[0]
            for t in items:
                if (t.get("performer") or {}).get("name", "").lower() == artist.lower(): best = t; break
            alb = best.get("album") or {}; img = alb.get("image") or {}
            cover = img.get("mega") or img.get("extralarge") or img.get("large") or img.get("small") or ""
            if cover and not cover.startswith("http"):
                cover = f"https:{cover}" if cover.startswith("//") else ""
            bd = best.get("maximum_bit_depth") or 0; sr = best.get("maximum_sampling_rate") or 0
            if sr > 1000: sr /= 1000
            ql = ""
            if bd and sr: ql = f"Hi-Res {int(bd)}-Bit / {sr:g} kHz" if bd >= 24 else f"CD {int(bd)}-Bit / {sr:g} kHz"
            return {"title": best.get("title") or title, "artist": (best.get("performer") or {}).get("name") or artist,
                "album": alb.get("title") or "", "cover": cover or None,
                "duration_ms": int((best.get("duration") or 0) * 1000), "quality": ql, "src": "Qobuz"}
        except: return None


_it = {}
def itunes(artist, track):
    k = f"{artist}||{track}".lower()
    if k in _it: return _it[k]
    try:
        r = requests.get("https://itunes.apple.com/search",
            params={"term": f"{artist} {track}", "entity": "song", "limit": "5"}, timeout=6)
        items = r.json().get("results", [])
        if not items: _it[k] = None; return None
        best = items[0]
        for i in items:
            if i.get("artistName", "").lower() == artist.lower(): best = i; break
        art = best.get("artworkUrl100", "").replace("100x100bb", "600x600bb")
        out = {"title": best.get("trackName", track), "artist": best.get("artistName", artist),
            "album": best.get("collectionName", ""), "cover": art or None,
            "duration_ms": best.get("trackTimeMillis", 0), "quality": "", "src": "iTunes"}
        _it[k] = out; return out
    except: _it[k] = None; return None


def get_title():
    pids = []
    for p in psutil.process_iter(["pid", "name"]):
        try:
            if p.info["name"] and p.info["name"].lower() == "qobuz.exe": pids.append(p.info["pid"])
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

def fmt(s):
    m, s = divmod(int(max(0, s)), 60); h, m = divmod(m, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m}:{s:02d}"


def setup():
    print("\n  Qobuz RPC Setup\n  " + "="*30 + "\n")
    cfg = load_cfg()
    v = input(f"  Discord App ID [{cfg.get('discord_app_id','')}]: ").strip()
    if v: cfg["discord_app_id"] = v
    v = input(f"  Qobuz Email [{cfg.get('qobuz_email','')}]: ").strip()
    if v: cfg["qobuz_email"] = v
    v = input("  Qobuz Password (blank = keep): ").strip()
    if v: cfg["qobuz_pw_hash"] = hashlib.md5(v.encode()).hexdigest(); print("  -> hashed")
    print("\n  Quality: 1) Hi-Res 192  2) Hi-Res 96  3) CD  4) MP3")
    v = input("  Pick [2]: ").strip()
    qm = {"1":"Hi-Res 24-Bit / 192 kHz","2":"Hi-Res 24-Bit / 96 kHz","3":"CD 16-Bit / 44.1 kHz","4":"MP3 320 kbps"}
    cfg["quality_label"] = qm.get(v, cfg.get("quality_label", "Hi-Res 24-Bit / 96 kHz"))
    save_cfg(cfg)
    print(f"\n  Saved to {CONFIG_PATH}\n")


def main():
    if "--setup" in sys.argv: setup(); return

    cfg = load_cfg()
    if not cfg.get("discord_app_id"):
        print("[!] No Discord App ID. Run with --setup"); sys.exit(1)

    print(f"\n  Qobuz RPC (CLI)\n  {'='*30}\n")

    running = True
    def stop(*_): nonlocal running; running = False; print("\n[*] Stopping...")
    signal.signal(signal.SIGINT, stop)
    signal.signal(signal.SIGTERM, stop)

    qz = QobuzAPI(); qz_ok = False
    email = cfg.get("qobuz_email","").strip(); pw = cfg.get("qobuz_pw_hash","").strip()
    if email and pw:
        if qz.init(): qz_ok = qz.login(email, pw)
        if not qz_ok: print("[*] iTunes fallback")
    else: print("[*] No creds, iTunes mode")

    print("[*] Connecting to Discord...")
    try: rpc = Presence(cfg["discord_app_id"]); rpc.connect(); print("[+] Connected")
    except Exception as e: print(f"[!] {e}"); sys.exit(1)

    tkey = None; tcover = None; talbum = None; tqual = ""; tdur = 0
    tstart = 0.0; prev = None; playing = False
    songs = 0; listen_s = 0.0; ltick = 0.0; t0 = time.time()
    iv = cfg.get("update_interval", 3)

    print(f"[*] Monitoring (every {iv}s)...\n")

    while running:
        now = time.time()
        if playing and ltick > 0:
            dt = now - ltick
            if 0 < dt < 10: listen_s += dt
        ltick = now if playing else 0

        try:
            raw = get_title()
            if raw is None:
                if playing or tkey:
                    print(f"  [{time.strftime('%H:%M:%S')}] Qobuz gone")
                    tkey = tcover = talbum = prev = None; tqual = ""; tdur = 0; tstart = 0; playing = False
                try: rpc.clear()
                except: pass
            else:
                p = parse(raw)
                if p:
                    k = f"{p['title']}|{p['artist']}"; playing = True
                    flick = prev and not parse(prev) and k == tkey
                    looped = k == tkey and tdur > 0 and tstart > 0 and now - tstart > tdur/1000 + 5
                    new = k != tkey

                    if new or flick or looped:
                        tstart = now; songs += 1
                        if new:
                            tkey = k
                            print(f"  [{time.strftime('%H:%M:%S')}] {p['title']}  |  {p['artist']}")
                            meta = qz.search(p["title"], p["artist"]) if qz_ok else None
                            if not meta: meta = itunes(p["artist"], p["title"])
                            if meta:
                                tcover = meta.get("cover"); talbum = meta.get("album","")
                                tqual = meta.get("quality",""); tdur = meta.get("duration_ms",0)
                                extra = f" | {tqual}" if tqual else ""
                                print(f"             [{meta.get('src','')}] {talbum}{extra}")
                            else:
                                tcover = None; talbum = ""; tqual = cfg.get("quality_label",""); tdur = 0
                        elif flick: print(f"  [{time.strftime('%H:%M:%S')}] Restarted")
                        elif looped: print(f"  [{time.strftime('%H:%M:%S')}] Looped")

                    state = f"{p['artist']} \u00b7 {tqual}" if tqual else p["artist"]
                    kw = {"details": p["title"][:128], "state": state[:128],
                        "large_image": tcover or cfg.get("fallback_cover") or "qobuz_icon",
                        "large_text": (talbum or "Qobuz")[:128]}
                    if tstart > 0: kw["start"] = int(tstart)
                    if cfg.get("show_quality_badge", True):
                        kw["small_image"] = "qobuz_icon"; kw["small_text"] = tqual or "Qobuz"
                    try: rpc.update(**kw)
                    except: pass
                else:
                    if playing:
                        print(f"  [{time.strftime('%H:%M:%S')}] Paused")
                        playing = False; tstart = 0
                    try: rpc.clear()
                    except: pass
                prev = raw
        except Exception as e:
            print(f"  [!] {e}")

        time.sleep(iv)

    try: rpc.clear(); rpc.close()
    except: pass
    print(f"\n  Session: {fmt(time.time()-t0)} | Listened: {fmt(listen_s)} | {songs} songs\n")


if __name__ == "__main__":
    main()
