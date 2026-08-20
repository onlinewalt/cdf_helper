"""Tests for the DeepSeek enrichment module (with a mocked HTTP backend)."""
import json
import sys
import tempfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

from cdf_helper.ai import AIProvider
from cdf_helper.parser import Part


def make_provider(tmp):
    provider = AIProvider(api_key="sk-test", cache_path=Path(tmp) / "cache.json")
    provider._post_json = lambda body: (
        '{"items": [{"id": 1, "weight_kg": 12.5, "unit_price": 350},'
        ' {"id": 2, "weight_kg": null, "unit_price": 99.9},'
        ' {"id": 3, "weight_kg": 0.2, "unit_price": null}]}'
    )
    return provider


def test_fills_missing_only():
    with tempfile.TemporaryDirectory() as tmp:
        provider = make_provider(tmp)
        parts = [
            Part(name="泵筒总成", qty=2),            # missing both
            Part(name="法兰", qty=3, weight=5.0),     # has weight, missing price
            Part(name="螺栓", qty=10, price=2.0),     # has price, missing weight
            Part(name="完整件", qty=1, weight=1.0, price=100.0),  # complete, untouched
        ]
        stats = provider.enrich(parts)
        assert parts[0].weight == 12.5 and parts[0].price == 350, parts[0]
        assert parts[1].weight == 5.0 and parts[1].price == 99.9, parts[1]   # weight preserved
        assert parts[2].weight == 0.2 and parts[2].price == 2.0, parts[2]    # price preserved
        assert parts[3].weight == 1.0 and parts[3].price == 100.0, parts[3]  # untouched
        assert stats == {"requested": 3, "filled": 3, "from_cache": 0, "errors": 0}, stats


def test_cache_hit_no_api_call():
    with tempfile.TemporaryDirectory() as tmp:
        provider = make_provider(tmp)
        parts1 = [Part(name="泵筒总成", qty=2)]
        provider.enrich(parts1)
        assert parts1[0].weight == 12.5

        calls = []
        provider._post_json = lambda body: calls.append(1) or None
        parts2 = [Part(name="泵筒总成", qty=2)]
        stats = provider.enrich(parts2)
        assert calls == [], "should not call API again when cached"
        assert parts2[0].weight == 12.5
        assert stats["from_cache"] == 1


def test_api_failure_keeps_blanks():
    with tempfile.TemporaryDirectory() as tmp:
        provider = AIProvider(api_key="sk-test", cache_path=Path(tmp) / "cache.json")
        provider._post_json = lambda body: None
        parts = [Part(name="未知备件", qty=1)]
        stats = provider.enrich(parts)
        assert parts[0].weight is None and parts[0].price is None
        assert stats["errors"] == 1


def test_json_fenced_response():
    with tempfile.TemporaryDirectory() as tmp:
        provider = make_provider(tmp)
        provider._post_json = lambda body: '```json\n{"items":[{"id":1,"weight_kg":7,"unit_price":55}]}\n```'
        parts = [Part(name="A", qty=1)]
        provider.enrich(parts)
        assert parts[0].weight == 7.0 and parts[0].price == 55.0


print("AI module tests PASSED")