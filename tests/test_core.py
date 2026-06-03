# pure-function tests for both entry points. these are what catch parse() drift.
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest
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
