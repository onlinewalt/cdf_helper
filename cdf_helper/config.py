"""本地配置读写（AI API Key / base URL / model 等），支持 config.json 与环境变量。

Profiles model: config.json may store a ``profiles`` mapping of named presets
plus a ``default_profile`` name, so users with several AI providers (e.g.
DeepSeek, OpenAI, a local Ollama) can switch between them. The legacy
top-level ``api_key``/``api_url``/``model`` keys are still honoured as the
"_last" (most-recently-saved) profile for backward compatibility.
"""

import json
import os
from pathlib import Path
from typing import List, TypedDict

PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = PROJECT_ROOT / "config.json"

ENV_API_KEY = "AI_API_KEY"
ENV_API_URL = "AI_API_URL"
ENV_MODEL = "AI_MODEL"

DEFAULT_API_URL = "https://api.deepseek.com"
DEFAULT_MODEL = "deepseek-chat"

LAST_PROFILE = "_last"


class AiConfig(TypedDict):
    """Typed shape of the AI configuration (the config Module's Interface).

    Callers in webapp.py / main.py / ai.py previously hand-indexed a bare
    dict returned by get_ai_config(); a key rename broke them with no static
    signal (mypy --strict passed because the return type was ``dict``). This
    TypedDict localizes the contract here and keeps the existing ``cfg["x"]``
    access syntax working, gaining static field checking at the seam.
    """

    api_key: str
    api_url: str
    model: str


def load_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def save_config(cfg: dict) -> None:
    data = load_config()
    data.update(cfg)
    CONFIG_PATH.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _last_profile(cfg: dict) -> dict:
    """Backward-compat shim: expose the legacy top-level keys as a profile dict."""
    return {
        "api_key": cfg.get("api_key", ""),
        "api_url": cfg.get("api_url", ""),
        "model": cfg.get("model", ""),
    }


def get_profiles() -> List[str]:
    """Names of all saved profiles, with the default last.

    Order is stable: default first (if set), then named profiles, then the
    legacy "_last" (if any top-level keys exist). Used by the UI "switch
    provider" dropdown.
    """
    cfg = load_config()
    profiles = dict(cfg.get("profiles", {}))
    if profiles and profiles.get(cfg.get("default_profile")):
        default = cfg["default_profile"]
        ordered = [default] + [p for p in profiles if p != default]
    else:
        ordered = list(profiles)
    # legacy top-level keys surface as an unnamed "_last" profile
    last = _last_profile(cfg)
    if any(v for v in last.values()):
        if LAST_PROFILE not in profiles:
            ordered.append(LAST_PROFILE)
    return ordered


def save_profile(name: str, api_key: str, api_url: str, model: str, default: bool = False) -> None:
    """Persist a named profile (and optionally make it the default).

    An empty *name* is treated as the legacy "_last" (anonymous) profile.
    """
    cfg = load_config()
    profiles = dict(cfg.get("profiles", {}))
    key = name or LAST_PROFILE
    profiles[key] = {"api_key": api_key or "", "api_url": api_url or "", "model": model or ""}
    cfg["profiles"] = profiles
    if default or (not profiles.get(cfg.get("default_profile")) and key != LAST_PROFILE):
        cfg["default_profile"] = key
    # keep the legacy top-level view in sync with the default profile so old
    # env-var fallback paths and any hand-edited config.json still work.
    if key != LAST_PROFILE:
        default_cfg = profiles[key]
        cfg["api_key"] = default_cfg["api_key"]
        cfg["api_url"] = default_cfg["api_url"]
        cfg["model"] = default_cfg["model"]
    CONFIG_PATH.write_text(
        json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def get_api_key(profile: str = None) -> str:
    """返回 API Key：优先环境变量，其次指定/默认 profile。"""
    key = os.environ.get(ENV_API_KEY, "").strip()
    if key:
        return key
    cfg = load_config()
    if profile is None:
        profile = cfg.get("default_profile") or LAST_PROFILE
    if profile == LAST_PROFILE or profile is None:
        return _last_profile(cfg).get("api_key", "").strip()
    return cfg.get("profiles", {}).get(profile, {}).get("api_key", "").strip()


def get_ai_config(profile: str = None) -> AiConfig:
    """返回生效的 AI 配置，合并自环境变量与 config.json。

    Keys: api_key, api_url（不含 /chat/completions 的基础地址）, model。
    环境变量优先于 config.json；缺失值回退到默认（DeepSeek）值。

    If *profile* is given, that named profile is used; otherwise the configured
    default profile is. Legacy top-level keys still work as the "_last" profile.
    """
    cfg = load_config()
    if profile is None:
        profile = cfg.get("default_profile") or LAST_PROFILE
    if profile == LAST_PROFILE:
        prof = _last_profile(cfg)
    else:
        prof = cfg.get("profiles", {}).get(profile, {})
    api_url = os.environ.get(ENV_API_URL, "").strip() or prof.get("api_url", "").strip()
    model = os.environ.get(ENV_MODEL, "").strip() or prof.get("model", "").strip()
    return AiConfig(
        api_key=get_api_key(profile),
        api_url=api_url or DEFAULT_API_URL,
        model=model or DEFAULT_MODEL,
    )