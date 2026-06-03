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
