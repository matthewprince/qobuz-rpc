import hashlib, signal, sys, time

try: from pypresence import Presence
except ImportError: print("[!] pip install pypresence"); sys.exit(1)

# all the shared + platform logic lives in qobuz_core
from qobuz_core import (
    QobuzAPI, authenticate_qobuz, itunes_lookup, parse, fmt, quality_str, decide,
    load_cfg, save_cfg, set_pw_hash, now_playing, CONFIG_PATH, DEFAULT_DISCORD_APP_ID,
)


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
    app_id = cfg.get("discord_app_id") or DEFAULT_DISCORD_APP_ID
    if not app_id:
        print("[!] No Discord App ID. Run with --setup"); sys.exit(1)

    print(f"\n  Qobuz RPC (CLI)\n  {'='*30}\n")

    running = True
    def stop(*_): nonlocal running; running = False; print("\n[*] Stopping...")
    signal.signal(signal.SIGINT, stop)
    signal.signal(signal.SIGTERM, stop)

    qz = QobuzAPI()
    mode = authenticate_qobuz(qz, cfg, log=lambda m: print(f"[*] {m}"))
    qz_ok = mode != "none"
    if mode == "none": print("[*] No Qobuz auth, iTunes mode")

    print("[*] Connecting to Discord...")
    try: rpc = Presence(app_id); rpc.connect(); print("[+] Connected")
    except Exception as e: print(f"[!] {e}"); sys.exit(1)

    tkey = None; tcover = None; talbum = None; tqual = ""; tdur = 0
    tstart = 0.0; playing = False; pause_pos = 0.0; last_sig = None; prev_status = None
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
            sample = now_playing(cfg)

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
                if not meta: meta = itunes_lookup(a, t)
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
                show_q = tqual and cfg.get("show_quality_badge", True)
                state = f"{sample['artist']} · {tqual}" if show_q else sample["artist"]
                kw = {"details": sample["title"][:128], "state": state[:128]}
                big = tcover or cfg.get("fallback_cover")
                if big: kw["large_image"] = big
                if talbum: kw["large_text"] = talbum[:128]
                if tstart > 0: kw["start"] = int(tstart)
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
