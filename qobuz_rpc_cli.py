import hashlib, json, re, signal, sys, os, time
try: import requests
except ImportError: print("[!] pip install requests"); sys.exit(1)
try: from pypresence import Presence
except ImportError: print("[!] pip install pypresence"); sys.exit(1)
try: import psutil
except ImportError: print("[!] pip install psutil"); sys.exit(1)
try: import win32gui, win32process
except ImportError: print("[!] pip install pywin32"); sys.exit(1)
try: import win32cred
except Exception: win32cred = None

import asyncio
try:
    from winrt.windows.media.control import GlobalSystemMediaTransportControlsSessionManager as _SMTCMgr
    HAS_SMTC = True
except Exception:
    HAS_SMTC = False

if getattr(sys, 'frozen', False):
    SCRIPT_DIR = os.path.dirname(sys.executable)
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "config.json")

DEFAULT_CFG = {
    "discord_app_id": "", "qobuz_email": "", "qobuz_pw_hash": "",
    "quality_label": "Hi-Res 24-Bit / 96 kHz", "update_interval": 3,
    "show_quality_badge": True, "fallback_cover": "", "use_smtc": True,
}

def load_cfg():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH) as f: return {**DEFAULT_CFG, **json.load(f)}
    cfg = dict(DEFAULT_CFG)
    save_cfg(cfg)
    return cfg

def save_cfg(c):
    with open(CONFIG_PATH, "w") as f: json.dump(c, f, indent=2)


# qobuz password hash lives in the windows credential manager, not config.json
CRED_TARGET = "QobuzRPC"
def cred_get():
    if win32cred is None: return ""
    try:
        c = win32cred.CredRead(CRED_TARGET, win32cred.CRED_TYPE_GENERIC)
        b = c.get("CredentialBlob") or b""
        return b.decode("utf-16-le") if isinstance(b, (bytes, bytearray)) else str(b)
    except Exception:
        return ""

def cred_set(h):
    if win32cred is None or not h: return False
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


class QobuzAPI:
    BASE = "https://www.qobuz.com/api.json/0.2"
    WEB = "https://play.qobuz.com"

    def __init__(self):
        self.app_id = None; self.user_auth_token = None
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
            ql = quality_str(best.get("maximum_bit_depth"), best.get("maximum_sampling_rate"))
            return {"title": best.get("title") or title, "artist": (best.get("performer") or {}).get("name") or artist,
                "album": alb.get("title") or "", "cover": cover or None,
                "duration_ms": int((best.get("duration") or 0) * 1000), "quality": ql, "src": "Qobuz"}
        except: return None


_it = {}
def itunes(artist, track):
    k = f"{artist}||{track}".lower()
    if k in _it: return _it[k]
    if len(_it) > 200: _it.pop(next(iter(_it)))
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
    if len(p) == 2 and p[0].strip() and p[1].strip():
        return {"title": p[0].strip(), "artist": p[1].strip()}
    return None

def fmt(s):
    m, s = divmod(int(max(0, s)), 60); h, m = divmod(m, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m}:{s:02d}"

def quality_str(bd, sr):
    bd = bd or 0; sr = sr or 0
    if sr > 1000: sr /= 1000
    if not (bd and sr): return ""
    return f"{'Hi-Res' if bd >= 24 else 'CD'} {int(bd)}-Bit / {sr:g} kHz"


# windows media session (SMTC) - exact title/artist/album + real position & play/pause
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
    if status == 4 and tl.last_updated_time:
        pos += max(0.0, time.time() - tl.last_updated_time.timestamp())
    if dur > 0: pos = min(pos, dur)
    return {"status": "playing" if status == 4 else "paused", "title": title,
        "artist": props.artist or "", "album": props.album_title or "", "pos": pos, "dur": dur}

def smtc_now():
    if not HAS_SMTC: return None
    try: return _smtc_run(_smtc_read())
    except: return None

def decide(st, sample, now):
    # pure state machine, shared by the scraper and SMTC sources, so it can be tested.
    # st keys: tkey, tstart, tdur, playing, pause_pos, prev_status
    # sample: {status,title,artist,album,pos,dur} or None. returns (event, new_st).
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


def setup():
    print("\n  Qobuz RPC Setup\n  " + "="*30 + "\n")
    cfg = load_cfg()
    v = input(f"  Discord App ID [{cfg.get('discord_app_id','')}]: ").strip()
    if v: cfg["discord_app_id"] = v
    v = input(f"  Qobuz Email [{cfg.get('qobuz_email','')}]: ").strip()
    if v: cfg["qobuz_email"] = v
    v = input("  Qobuz Password (blank = keep): ").strip()
    if v: set_pw_hash(cfg, hashlib.md5(v.encode()).hexdigest()); print("  -> hashed")
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
    email = cfg.get("qobuz_email","").strip(); pw = get_pw_hash(cfg)
    if email and pw:
        if qz.init(): qz_ok = qz.login(email, pw)
        if not qz_ok: print("[*] iTunes fallback")
    else: print("[*] No creds, iTunes mode")

    print("[*] Connecting to Discord...")
    try: rpc = Presence(cfg["discord_app_id"]); rpc.connect(); print("[+] Connected")
    except Exception as e: print(f"[!] {e}"); sys.exit(1)

    tkey = None; tcover = None; talbum = None; tqual = ""; tdur = 0
    tstart = 0.0; playing = False; pause_pos = 0.0; last_sig = None; prev_status = None
    songs = 0; listen_s = 0.0; ltick = 0.0; t0 = time.time()
    iv = cfg.get("update_interval", 3); use_smtc = cfg.get("use_smtc", True)

    print(f"[*] Monitoring (every {iv}s)...\n")

    while running:
        now = time.time()
        if playing and ltick > 0:
            dt = now - ltick
            if 0 < dt < 10: listen_s += dt
        ltick = now if playing else 0

        try:
            sample = smtc_now() if use_smtc else None
            if sample is None:
                raw = get_title()
                if raw is not None:
                    p = parse(raw)
                    if p: sample = {"status":"playing","title":p["title"],"artist":p["artist"],"album":None,"pos":None,"dur":None}
                    else: sample = {"status":"paused","title":"","artist":"","album":None,"pos":None,"dur":None}

            st = {"tkey":tkey,"tstart":tstart,"tdur":tdur,"playing":playing,"pause_pos":pause_pos,"prev_status":prev_status}
            ev, st = decide(st, sample, now)
            tkey=st["tkey"]; tstart=st["tstart"]; playing=st["playing"]; pause_pos=st["pause_pos"]; prev_status=st["prev_status"]
            ts = time.strftime('%H:%M:%S')

            if ev == "gone":
                print(f"  [{ts}] Qobuz gone")
                tkey = None; tcover = talbum = None; tqual = ""; tdur = 0; last_sig = None
                try: rpc.clear()
                except: pass
            elif ev == "pause":
                print(f"  [{ts}] Paused"); last_sig = None
                try: rpc.clear()
                except: pass
            elif ev == "new":
                songs += 1
                t, a = sample["title"], sample["artist"]
                print(f"  [{ts}] {t}  |  {a}")
                meta = qz.search(t, a) if qz_ok else None
                if not meta: meta = itunes(a, t)
                if meta:
                    tcover = meta.get("cover"); talbum = meta.get("album","") or (sample.get("album") or "")
                    tqual = meta.get("quality",""); tdur = meta.get("duration_ms",0)
                    print(f"             [{meta.get('src','')}] {talbum}{f' | {tqual}' if tqual else ''}")
                else:
                    tcover = None; talbum = sample.get("album") or ""; tqual = cfg.get("quality_label",""); tdur = 0
                if not tdur and sample.get("dur"): tdur = int(sample["dur"]*1000)
            elif ev == "loop":
                songs += 1; print(f"  [{ts}] Looped")
            elif ev == "resume":
                print(f"  [{ts}] Resumed")

            if sample and sample["status"] == "playing":
                state = f"{sample['artist']} \u00b7 {tqual}" if tqual else sample["artist"]
                kw = {"details": sample["title"][:128], "state": state[:128],
                    "large_image": tcover or cfg.get("fallback_cover") or "qobuz_icon",
                    "large_text": (talbum or "Qobuz")[:128]}
                if tstart > 0: kw["start"] = int(tstart)
                if cfg.get("show_quality_badge", True):
                    kw["small_image"] = "qobuz_icon"; kw["small_text"] = tqual or "Qobuz"
                sig = (sample["title"], state, tcover, talbum, kw.get("start"))
                if sig != last_sig:
                    try: rpc.update(**kw); last_sig = sig
                    except: pass
        except Exception as e:
            print(f"  [!] {e}")

        time.sleep(iv)

    try: rpc.clear(); rpc.close()
    except: pass
    print(f"\n  Session: {fmt(time.time()-t0)} | Listened: {fmt(listen_s)} | {songs} songs\n")


if __name__ == "__main__":
    main()
