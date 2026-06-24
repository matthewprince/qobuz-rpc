# Glassmorphism pywebview front-end for Qobuz RPC. Reuses qobuz_core for all the
# real work (Qobuz API, now-playing sources, the state machine, zero-config auth)
# and pypresence for Discord, then pushes live state into the web UI.
import hashlib, json, os, sys, threading, time

import webview
from pypresence import Presence, ActivityType, StatusDisplayType
from qobuz_core import (
    QobuzAPI, authenticate_qobuz, now_playing, decide, itunes_lookup,
    load_cfg, save_cfg, get_pw_hash, set_pw_hash, set_autostart,
    DEFAULT_DISCORD_APP_ID, PAUSE_ICON_URL, SCRIPT_DIR,
)

PW_PLACEHOLDER = "••••••••"

# The window is kept OFF the js_api object: pywebview serializes the api object,
# and a reference to the native window handle made it recurse on the .NET
# accessibility tree ("maximum recursion depth exceeded").
WIN = None

def _web_dir():
    base = sys._MEIPASS if getattr(sys, "frozen", False) else SCRIPT_DIR
    return os.path.join(base, "web")


class Backend:
    def __init__(self):
        self.cfg = load_cfg()
        self.monitoring = False
        self.state = {
            "status": "idle", "connected": False, "nowplaying": False, "paused": False,
            "title": "", "artist": "", "album": "", "cover": "", "quality": "",
            "tstart": 0.0, "pos": 0.0, "dur": 0.0,
            "session_start": 0.0, "listened": 0.0, "songs": 0,
            "qmode": "none", "msg": "Idle",
        }

    # ---- called from JS (pywebview.api.*) ----
    def get_state(self):
        return {"state": self.state, "settings": self._settings()}

    def _settings(self):
        return {
            "discord_app_id": self.cfg.get("discord_app_id", ""),
            "qobuz_email": self.cfg.get("qobuz_email", ""),
            "has_password": bool(get_pw_hash(self.cfg)),
            "quality_label": self.cfg.get("quality_label", "Hi-Res 24-Bit / 96 kHz"),
            "update_interval": self.cfg.get("update_interval", 3),
            "show_quality_badge": self.cfg.get("show_quality_badge", True),
            "auto_connect": self.cfg.get("auto_connect", False),
            "start_with_windows": self.cfg.get("start_with_windows", False),
        }

    def save_settings(self, s):
        self.cfg["discord_app_id"] = (s.get("discord_app_id") or "").strip()
        self.cfg["qobuz_email"] = (s.get("qobuz_email") or "").strip()
        pw = (s.get("qobuz_password") or "").strip()
        if pw and pw != PW_PLACEHOLDER:
            set_pw_hash(self.cfg, hashlib.md5(pw.encode()).hexdigest())
        self.cfg["quality_label"] = s.get("quality_label") or self.cfg.get("quality_label")
        try: self.cfg["update_interval"] = max(1, int(float(s.get("update_interval") or 3)))
        except Exception: pass
        self.cfg["show_quality_badge"] = bool(s.get("show_quality_badge"))
        self.cfg["auto_connect"] = bool(s.get("auto_connect"))
        self.cfg["start_with_windows"] = bool(s.get("start_with_windows"))
        save_cfg(self.cfg)
        try: set_autostart(self.cfg.get("start_with_windows", False))
        except Exception: pass
        return {"ok": True, "settings": self._settings()}

    def toggle_connect(self):
        if self.monitoring:
            self.monitoring = False
            return {"connected": False}
        self.monitoring = True
        threading.Thread(target=self._monitor, daemon=True).start()
        return {"connected": True}

    def minimize(self):
        if WIN:
            try: WIN.minimize()
            except Exception: pass

    def close(self):
        self.monitoring = False
        if WIN:
            try: WIN.destroy()
            except Exception: pass

    # ---- internals ----
    def _push(self):
        if not WIN: return
        try:
            WIN.evaluate_js("window.qSetState(%s)" % json.dumps(self.state))
        except Exception:
            pass

    def _msg(self, m):
        self.state["msg"] = m; self._push()

    def _monitor(self):
        app_id = self.cfg.get("discord_app_id") or DEFAULT_DISCORD_APP_ID
        self.state.update(status="connecting", msg="Authenticating Qobuz...")
        self._push()

        qz = QobuzAPI()
        qmode = authenticate_qobuz(qz, self.cfg, log=self._msg)

        def rpc_connect():
            try: r = Presence(app_id); r.connect(); return r
            except Exception: return None
        def rpc_drop(r):
            try:
                if r: r.close()
            except Exception: pass
            return None

        rpc = rpc_connect()
        last_reconnect = time.time()
        self.state.update(connected=True, qmode=qmode, session_start=time.time(),
            songs=0, listened=0.0, status=("connected" if rpc else "connecting"),
            msg=("Connected" if rpc else "Waiting for Discord..."))
        self._push()

        tkey = tcover = talbum = None; tqual = ""; tdur = 0
        tstart = 0.0; playing = False; pause_pos = 0.0; prev_status = None; last_sig = None
        ltick = 0.0; songs = 0; listen_s = 0.0

        while self.monitoring:
            now = time.time()
            if playing and ltick > 0:
                dt = now - ltick
                if 0 < dt < 10: listen_s += dt
            ltick = now if playing else 0

            if rpc is None and now - last_reconnect > 5:
                last_reconnect = now; rpc = rpc_connect()
                if rpc: last_sig = None

            try:
                sample = now_playing(self.cfg)
                stm = {"tkey": tkey, "tstart": tstart, "tdur": tdur, "playing": playing,
                    "pause_pos": pause_pos, "prev_status": prev_status}
                ev, stm = decide(stm, sample, now)
                tkey = stm["tkey"]; tstart = stm["tstart"]; playing = stm["playing"]
                pause_pos = stm["pause_pos"]; prev_status = stm["prev_status"]

                if ev == "gone":
                    tkey = tcover = talbum = None; tqual = ""; tdur = 0; last_sig = None
                    if rpc:
                        try: rpc.clear()
                        except Exception: rpc = rpc_drop(rpc); last_reconnect = now
                    self.state.update(nowplaying=False, paused=False, title="", artist="",
                        album="", cover="", quality="", tstart=0, pos=0, dur=0, msg="Nothing playing")
                elif ev == "pause":
                    last_sig = None
                    name = tkey.split("|", 1)[0] if tkey else ""
                    if rpc and name:
                        kw = {"activity_type": ActivityType.LISTENING, "status_display_type": StatusDisplayType.DETAILS,
                            "details": name[:128], "state": "Paused", "small_image": PAUSE_ICON_URL, "small_text": "Paused"}
                        if tcover: kw["large_image"] = tcover
                        if talbum: kw["large_text"] = talbum[:128]
                        try: rpc.update(**kw)
                        except Exception: rpc = rpc_drop(rpc); last_reconnect = now
                    self.state.update(paused=True, pos=pause_pos, msg="Paused")
                elif ev == "new":
                    songs += 1
                    t, a = sample["title"], sample["artist"]
                    meta = qz.search(t, a) if qmode != "none" else None
                    if not meta: meta = itunes_lookup(a, t)
                    if meta:
                        tcover = meta.get("cover"); talbum = meta.get("album", "") or (sample.get("album") or "")
                        tqual = meta.get("quality", ""); tdur = meta.get("duration_ms", 0)
                    else:
                        tcover = None; talbum = sample.get("album") or ""; tqual = self.cfg.get("quality_label", ""); tdur = 0
                    if not tcover and sample.get("art"): tcover = sample["art"]
                    if not tdur and sample.get("dur"): tdur = int(sample["dur"] * 1000)
                    self.state.update(nowplaying=True, paused=False, title=t, artist=a,
                        album=talbum or "", cover=tcover or "",
                        quality=(tqual if self.cfg.get("show_quality_badge", True) else ""),
                        tstart=tstart, dur=tdur / 1000.0, msg="Playing: %s" % t)
                elif ev == "loop":
                    songs += 1
                elif ev == "resume":
                    self.state.update(paused=False, nowplaying=True, tstart=tstart)

                if sample and sample["status"] == "playing":
                    self.state.update(nowplaying=True, paused=False, tstart=tstart, dur=tdur / 1000.0)
                    if rpc is not None:
                        show_q = tqual and self.cfg.get("show_quality_badge", True)
                        st_txt = ("%s · %s" % (sample["artist"], tqual)) if show_q else sample["artist"]
                        kw = {"activity_type": ActivityType.LISTENING, "status_display_type": StatusDisplayType.DETAILS,
                            "details": sample["title"][:128], "state": st_txt[:128]}
                        if tcover: kw["large_image"] = tcover
                        if talbum: kw["large_text"] = talbum[:128]
                        # send end too, not just start, so Discord draws the scrubbing
                        # progress bar; start alone renders as a plain elapsed counter
                        if tdur and tstart > 0: kw["start"] = int(tstart); kw["end"] = int(tstart + tdur / 1000)
                        elif tstart > 0: kw["start"] = int(tstart)
                        sig = (sample["title"], st_txt, tcover, talbum, kw.get("start"), kw.get("end"))
                        if sig != last_sig:
                            try: rpc.update(**kw); last_sig = sig
                            except Exception: rpc = rpc_drop(rpc); last_sig = None; last_reconnect = now
            except Exception as e:
                self.state["msg"] = "Error: %s" % e

            self.state.update(songs=songs, listened=listen_s, status=("connected" if rpc else "connecting"))
            self._push()
            time.sleep(self.cfg.get("update_interval", 3))

        if rpc:
            try: rpc.clear(); rpc.close()
            except Exception: pass
        self.state.update(connected=False, status="idle", nowplaying=False, paused=False,
            title="", artist="", album="", cover="", quality="", tstart=0, pos=0, dur=0, msg="Disconnected")
        self._push()


def main():
    global WIN
    backend = Backend()
    index = os.path.join(_web_dir(), "index.html")
    WIN = webview.create_window(
        "Qobuz RPC", url=index, js_api=backend,
        width=460, height=772, min_size=(404, 560),
        background_color="#060A12", frameless=True, easy_drag=False, resizable=True,
    )
    if backend.cfg.get("auto_connect"):
        threading.Thread(target=lambda: (time.sleep(0.7), backend.toggle_connect()), daemon=True).start()
    webview.start()


if __name__ == "__main__":
    main()
