"""Tests for the translator module (with mocked API calls)."""
import json
import sys
import tempfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

from cdf_helper.translator import Translator, _is_english, _cache_key
from cdf_helper.parser import Part


def make_translator(tmp, translate_map=None):
    """Create a Translator whose _translate_api is mocked."""
    provider = Translator(
        app_id="test-appid",
        apikey="test-apikey",
        cache_path=Path(tmp) / "translate_cache.json",
    )
    if translate_map:
        provider._translate_api = lambda text: translate_map.get(text)
    else:
        provider._translate_api = lambda text: f"[{text}]"
    return provider


def test_is_english_detection():
    assert _is_english("INTERMEDIATE RELAY") is True
    assert _is_english("PCE-123") is True
    assert _is_english("备件名称") is False
    assert _is_english("") is False
    assert _is_english("泵筒总成 Pump Assembly") is False  # has CJK -> not English
    assert _is_english("12345") is True  # no CJK -> treated as English
    print("_is_english detection OK")


def test_skips_chinese_names():
    with tempfile.TemporaryDirectory() as tmp:
        provider = make_translator(tmp)
        parts = [
            Part(name="备件名称", qty=2),
            Part(name="法兰", qty=3),
        ]
        stats = provider.translate_names(parts, on_status=lambda m: None)
        assert parts[0].name == "备件名称"
        assert parts[1].name == "法兰"
        assert stats["translated"] == 0
        assert stats["from_cache"] == 0
        assert stats["errors"] == 0
        print("Chinese names skipped OK")


def test_translates_english_names():
    with tempfile.TemporaryDirectory() as tmp:
        provider = make_translator(tmp, {
            "INTERMEDIATE RELAY": "中间继电器",
            "PLASTIC SHELL": "塑料壳体",
        })
        parts = [
            Part(name="INTERMEDIATE RELAY", qty=2),
            Part(name="PLASTIC SHELL", qty=3),
            Part(name="备件名称", qty=1),
        ]
        stats = provider.translate_names(parts, on_status=lambda m: None)
        assert parts[0].name == "中间继电器"
        assert parts[1].name == "塑料壳体"
        assert parts[2].name == "备件名称"  # untouched
        assert stats["translated"] == 2
        print("English names translated OK")


def test_cache_hit_no_api_call():
    with tempfile.TemporaryDirectory() as tmp:
        provider = make_translator(tmp, {
            "PUMP ASSEMBLY": "泵总成",
        })
        calls = []
        provider._translate_api = lambda text: calls.append(text) or "泵总成"

        parts1 = [Part(name="PUMP ASSEMBLY", qty=2)]
        provider.translate_names(parts1, on_status=lambda m: None)
        assert parts1[0].name == "泵总成"
        assert len(calls) == 1

        calls.clear()
        parts2 = [Part(name="PUMP ASSEMBLY", qty=2)]
        stats = provider.translate_names(parts2, on_status=lambda m: None)
        assert calls == [], "should not call API again when cached"
        assert parts2[0].name == "泵总成"
        assert stats["from_cache"] == 1
        assert stats["translated"] == 0
        print("Cache hit skips API OK")


def test_api_failure_keeps_original():
    with tempfile.TemporaryDirectory() as tmp:
        provider = make_translator(tmp)
        provider._translate_api = lambda text: None  # simulate API failure
        parts = [Part(name="UNKNOWN PART", qty=1)]
        stats = provider.translate_names(parts, on_status=lambda m: None)
        assert parts[0].name == "UNKNOWN PART"  # original kept
        assert stats["errors"] == 1
        assert stats["translated"] == 0
        print("API failure keeps original OK")


def test_cache_persisted_to_disk():
    with tempfile.TemporaryDirectory() as tmp:
        cache_path = Path(tmp) / "translate_cache.json"
        provider = make_translator(tmp, {
            "BEARING": "轴承",
        })
        parts = [Part(name="BEARING", qty=5)]
        provider.translate_names(parts, on_status=lambda m: None)
        assert cache_path.exists()
        cached = json.loads(cache_path.read_text(encoding="utf-8"))
        assert cached[_cache_key("BEARING")] == "轴承"
        print("Cache persisted OK")


print("Translator module tests PASSED")
