# Live MPRIS integration test (Linux/WSL only). Publishes real MPRIS players on
# a private session bus with dbus-next, then runs the real qobuz_core reader
# (jeepney) against them to validate the D-Bus path + player selection end to
# end. Run: dbus-run-session -- python _mpris_livetest.py    (not run by CI)
import asyncio, threading, time, sys, os

REPO = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, REPO)

from dbus_next.aio import MessageBus
from dbus_next.service import ServiceInterface, method, dbus_property, PropertyAccess
from dbus_next import Variant


def make_player(status, meta):
    class P(ServiceInterface):
        def __init__(self):
            super().__init__("org.mpris.MediaPlayer2.Player")
        @dbus_property(access=PropertyAccess.READ)
        def PlaybackStatus(self) -> "s":
            return status
        @dbus_property(access=PropertyAccess.READ)
        def Metadata(self) -> "a{sv}":
            return meta
        @dbus_property(access=PropertyAccess.READ)
        def Position(self) -> "x":
            return 30_000_000
        @method()
        def PlayPause(self):
            pass
    return P()


PLAYERS = [
    # a non-Qobuz browser player that is also Playing (a distractor)
    ("org.mpris.MediaPlayer2.firefox.instance1", "Playing", {
        "xesam:title": Variant("s", "Other Song"),
        "xesam:artist": Variant("as", ["Someone"]),
        "xesam:album": Variant("s", "Other Album"),
        "mpris:length": Variant("x", 180_000_000),
        "mpris:artUrl": Variant("s", "https://i.scdn.co/image/abc"),
        "xesam:url": Variant("s", "https://open.spotify.com/track/x"),
    }),
    # the Qobuz web player (cover art on the Qobuz CDN -> tier-1 match)
    ("org.mpris.MediaPlayer2.chromium.instance1", "Playing", {
        "xesam:title": Variant("s", "Qobuz Song"),
        "xesam:artist": Variant("as", ["A", "B"]),
        "xesam:album": Variant("s", "Q Album"),
        "mpris:length": Variant("x", 200_000_000),
        "mpris:artUrl": Variant("s", "https://static.qobuz.com/images/covers/x_600.jpg"),
        "xesam:url": Variant("s", "https://play.qobuz.com/track/123"),
    }),
]

ready = threading.Event()

async def _serve():
    for name, status, meta in PLAYERS:
        bus = await MessageBus().connect()
        bus.export("/org/mpris/MediaPlayer2", make_player(status, meta))
        await bus.request_name(name)
    ready.set()
    await asyncio.Future()

def _thr():
    try: asyncio.run(_serve())
    except Exception as e: print("PUBLISHER ERROR:", repr(e))

threading.Thread(target=_thr, daemon=True).start()
if not ready.wait(15):
    print("publisher not ready"); sys.exit(2)
time.sleep(0.6)

import qobuz_core as core
print("HAS_MPRIS:", core.HAS_MPRIS)
conn = core._mpris_conn()
print("players on bus:", sorted(core._mpris_list(conn)))

fails = []

# tier 1: among two Playing players, the Qobuz-art one is chosen, with the right sample
s = core.mpris_now({})
print("mpris_now({}) ->", s)
if not (s and s["title"] == "Qobuz Song" and s["artist"] == "A, B" and abs(s["dur"]-200) < 0.1):
    fails.append("tier-1 qobuz selection")
if not (s and s.get("art") and "qobuz" in s["art"]):
    fails.append("cover art passthrough")

# re-read the live players, then exercise tier-2 / tier-3 on real data (markers stripped)
players = []
for n in sorted(core._mpris_list(conn)):
    ap = {k: core._unwrap(v) for k, v in core._mpris_getall(conn, n).items()}
    meta = {k: core._unwrap(v) for k, v in (ap.get("Metadata") or {}).items()}
    players.append({"name": n, "status": ap.get("PlaybackStatus"), "metadata": meta, "pos": ap.get("Position")})
for p in players:
    p["metadata"].pop("mpris:artUrl", None); p["metadata"].pop("xesam:url", None)

ff = core._mpris_choose(players, "firefox")
if not (ff and "firefox" in ff["name"]):
    fails.append("tier-2 configured-substring")
first = core._mpris_choose(players, "")
if not (first and first["status"] == "Playing"):
    fails.append("tier-3 first-playing")

print("selection: tier1(qobuz-art), tier2(pref=firefox), tier3(first-playing) ->",
      "all OK" if not fails else f"FAILED: {fails}")
print("RESULT:", "PASS" if not fails else "FAIL")
sys.exit(0 if not fails else 1)
