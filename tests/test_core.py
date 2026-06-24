# pure-function tests for both entry points. these are what catch parse() drift.
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest
import qobuz_core as core
import qobuz_rpc as g
import qobuz_rpc_cli as c

MODS = [pytest.param(g, id="gui"), pytest.param(c, id="cli")]


@pytest.mark.parametrize("m", MODS)
class TestParse:
    def test_track_artist(self, m):
        assert m.parse("Song - Artist") == {"title": "Song", "artist": "Artist"}

    def test_app_name_is_none(self, m):
        assert m.parse("Qobuz") is None
        assert m.parse("qobuz") is None

    def test_title_only_is_none(self, m):
        assert m.parse("JustTitle") is None

    def test_blank_is_none(self, m):
        assert m.parse("") is None
        assert m.parse("   ") is None
        assert m.parse(None) is None

    def test_empty_sides_is_none(self, m):
        assert m.parse(" - ") is None

    def test_extra_hyphen_splits_once(self, m):
        assert m.parse("A - B - C") == {"title": "A", "artist": "B - C"}

    def test_strips_whitespace(self, m):
        assert m.parse("  S  -  A  ") == {"title": "S", "artist": "A"}


@pytest.mark.parametrize("m", MODS)
class TestFmt:
    def test_zero(self, m): assert m.fmt(0) == "0:00"
    def test_seconds(self, m): assert m.fmt(5) == "0:05"
    def test_minutes(self, m): assert m.fmt(65) == "1:05"
    def test_hours(self, m): assert m.fmt(3661) == "1:01:01"
    def test_negative_clamped(self, m): assert m.fmt(-3) == "0:00"


@pytest.mark.parametrize("m", MODS)
class TestQualityStr:
    def test_hires(self, m): assert m.quality_str(24, 96) == "Hi-Res 24-Bit / 96 kHz"
    def test_cd(self, m): assert m.quality_str(16, 44.1) == "CD 16-Bit / 44.1 kHz"
    def test_hz_normalized_to_khz(self, m): assert m.quality_str(24, 96000) == "Hi-Res 24-Bit / 96 kHz"
    def test_192(self, m): assert m.quality_str(24, 192000) == "Hi-Res 24-Bit / 192 kHz"
    def test_missing(self, m):
        assert m.quality_str(0, 0) == ""
        assert m.quality_str(None, None) == ""
        assert m.quality_str(24, 0) == ""


def _idle():
    return {"tkey": None, "tstart": 0.0, "tdur": 0, "playing": False, "pause_pos": 0.0, "prev_status": None}

def _smp(status="playing", title="T", artist="A", album="Al", pos=None, dur=None):
    return {"status": status, "title": title, "artist": artist, "album": album, "pos": pos, "dur": dur}


@pytest.mark.parametrize("m", MODS)
class TestDecide:
    def test_idle_stays_quiet(self, m):
        ev, st = m.decide(_idle(), None, 100.0)
        assert ev is None and st["playing"] is False

    def test_new_with_position(self, m):
        ev, st = m.decide(_idle(), _smp(pos=30.0, dur=200.0), 1000.0)
        assert ev == "new" and st["tkey"] == "T|A" and abs(st["tstart"] - 970.0) < 0.01

    def test_new_without_position(self, m):
        ev, st = m.decide(_idle(), _smp(pos=None), 1000.0)
        assert ev == "new" and st["tstart"] == 1000.0

    def test_same_track_is_stable_tick(self, m):
        ev, st = m.decide(_idle(), _smp(pos=30.0, dur=200.0), 1000.0)
        ev, st = m.decide(st, _smp(pos=33.0, dur=200.0), 1003.0)
        assert ev == "tick" and abs(st["tstart"] - 970.0) < 0.01  # no drift -> debounce holds

    def test_pause_captures_position(self, m):
        ev, st = m.decide(_idle(), _smp(pos=30.0, dur=200.0), 1000.0)
        ev, st = m.decide(st, _smp(status="paused", pos=33.0), 1004.0)
        assert ev == "pause" and st["playing"] is False and abs(st["pause_pos"] - 33.0) < 0.01

    def test_resume_with_position_excludes_pause_gap(self, m):
        ev, st = m.decide(_idle(), _smp(pos=30.0, dur=200.0), 1000.0)
        ev, st = m.decide(st, _smp(status="paused", pos=33.0), 1004.0)
        ev, st = m.decide(st, _smp(pos=33.0, dur=200.0), 1010.0)
        assert ev == "resume" and abs(st["tstart"] - 977.0) < 0.01

    def test_resume_without_position_uses_pause_pos(self, m):
        st = _idle(); st.update(tkey="T|A", tstart=950.0, playing=True, prev_status="playing")
        ev, st = m.decide(st, _smp(status="paused", pos=None), 1000.0)
        assert abs(st["pause_pos"] - 50.0) < 0.01
        ev, st = m.decide(st, _smp(pos=None), 1100.0)
        assert ev == "resume" and abs(st["tstart"] - 1050.0) < 0.01

    def test_seek_jumps_tstart(self, m):
        ev, st = m.decide(_idle(), _smp(pos=30.0, dur=200.0), 1000.0)
        ev, st = m.decide(st, _smp(pos=120.0, dur=200.0), 1002.0)
        assert ev == "seek" and abs(st["tstart"] - 882.0) < 0.01

    def test_gone_after_playing(self, m):
        ev, st = m.decide(_idle(), _smp(pos=None), 1000.0)
        ev, st = m.decide(st, None, 1001.0)
        assert ev == "gone" and st["playing"] is False

    def test_loop_detected_without_position(self, m):
        st = _idle(); st.update(tkey="T|A", tstart=1000.0, tdur=10000, playing=True, prev_status="playing")
        ev, st = m.decide(st, _smp(pos=None), 1016.0)
        assert ev == "loop" and st["tstart"] == 1016.0

    def test_track_change_is_new(self, m):
        ev, st = m.decide(_idle(), _smp(title="A1", pos=None), 1000.0)
        ev, st = m.decide(st, _smp(title="B1", pos=None), 1005.0)
        assert ev == "new" and st["tkey"] == "B1|A"


# MPRIS metadata -> sample, the pure half of the Linux source (no D-Bus needed)
class TestMprisSample:
    def test_playing_basic(self):
        s = core.mpris_sample("Playing",
            {"xesam:title": "T", "xesam:artist": ["A1", "A2"], "xesam:album": "Al",
             "mpris:length": 200_000_000}, 30_000_000)
        assert s["status"] == "playing" and s["title"] == "T"
        assert s["artist"] == "A1, A2" and s["album"] == "Al"
        assert abs(s["dur"] - 200.0) < 0.01 and abs(s["pos"] - 30.0) < 0.01

    def test_paused(self):
        s = core.mpris_sample("Paused", {"xesam:title": "T", "xesam:artist": ["A"]}, 0)
        assert s["status"] == "paused" and s["artist"] == "A"

    def test_stopped_is_none(self):
        assert core.mpris_sample("Stopped", {"xesam:title": "T"}, 0) is None

    def test_unknown_status_is_none(self):
        assert core.mpris_sample(None, {"xesam:title": "T"}, 0) is None

    def test_no_title_is_none(self):
        assert core.mpris_sample("Playing", {"xesam:artist": ["A"]}, 0) is None
        assert core.mpris_sample("Playing", {"xesam:title": "   "}, 0) is None

    def test_artist_string_not_list(self):
        s = core.mpris_sample("Playing", {"xesam:title": "T", "xesam:artist": "Solo"}, 0)
        assert s["artist"] == "Solo"

    def test_position_clamped_to_duration(self):
        s = core.mpris_sample("Playing",
            {"xesam:title": "T", "mpris:length": 100_000_000}, 250_000_000)
        assert abs(s["dur"] - 100.0) < 0.01 and abs(s["pos"] - 100.0) < 0.01

    def test_missing_length_and_position(self):
        s = core.mpris_sample("Playing", {"xesam:title": "T"}, None)
        assert s["dur"] == 0 and s["pos"] == 0

    def test_http_art_passthrough(self):
        s = core.mpris_sample("Playing",
            {"xesam:title": "T", "mpris:artUrl": "https://static.qobuz.com/x.jpg"}, 0)
        assert s["art"] == "https://static.qobuz.com/x.jpg"

    def test_file_art_dropped(self):
        s = core.mpris_sample("Playing",
            {"xesam:title": "T", "mpris:artUrl": "file:///tmp/cover.png"}, 0)
        assert s["art"] is None

    def test_art_absent_is_none(self):
        assert core.mpris_sample("Playing", {"xesam:title": "T"}, 0)["art"] is None


# MPRIS player selection - the "which player is Qobuz" resolution order
class TestMprisChoose:
    @staticmethod
    def _p(name, status="Playing", url="", art=""):
        return {"name": f"org.mpris.MediaPlayer2.{name}", "status": status,
                "metadata": {"xesam:url": url, "mpris:artUrl": art}, "pos": 0}

    def test_prefers_qobuz_in_name(self):
        chosen = core._mpris_choose(
            [self._p("firefox", "Playing"), self._p("qobuz", "Paused")], "")
        assert chosen["name"].endswith("qobuz")

    def test_matches_qobuz_in_art_url(self):
        chosen = core._mpris_choose(
            [self._p("firefox", "Playing", art="https://static.qobuz.com/cover.jpg")], "")
        assert chosen["name"].endswith("firefox")

    def test_configured_substring_wins_over_playing(self):
        chosen = core._mpris_choose(
            [self._p("chromium", "Playing"), self._p("firefox", "Paused")], "firefox")
        assert chosen["name"].endswith("firefox")

    def test_falls_back_to_first_playing(self):
        chosen = core._mpris_choose(
            [self._p("vlc", "Paused"), self._p("spotify", "Playing")], "")
        assert chosen["name"].endswith("spotify")

    def test_none_when_nothing_relevant(self):
        assert core._mpris_choose([self._p("vlc", "Paused")], "") is None
