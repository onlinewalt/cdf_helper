"""Shared pytest fixtures/config for CDF_helper tests.

Centralizes UTF-8 stdout handling (so Chinese assertions render on any
platform / CI log) and provides an isolated Flask test client, so individual
test modules stay free of module-level side effects (import-safe under pytest).
"""
import sys

import pytest


def pytest_configure(config):
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


@pytest.fixture
def client(tmp_path, monkeypatch):
    """Flask test client with an isolated config.json and TESTING=True."""
    import webapp
    from cdf_helper import config as app_config

    monkeypatch.setattr(app_config, "CONFIG_PATH", tmp_path / "config.json")
    monkeypatch.setitem(webapp.app.config, "TESTING", True)
    with webapp.app.test_client() as c:
        yield c
