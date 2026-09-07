"""本地配置读写（AI API Key / base URL / model 等），支持 config.json 与环境变量。"""

import json
import os
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = PROJECT_ROOT / "config.json"

ENV_API_KEY = "AI_API_KEY"
ENV_API_URL = "AI_API_URL"
ENV_MODEL = "AI_MODEL"

DEFAULT_API_URL = "https://api.deepseek.com"
DEFAULT_MODEL = "deepseek-chat"


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


def get_api_key() -> str:
    """返回 API Key：优先环境变量，其次 config.json。"""
    key = os.environ.get(ENV_API_KEY, "").strip()
    if key:
        return key
    return load_config().get("api_key", "").strip()


def get_ai_config() -> dict:
    """返回生效的 AI 配置，合并自环境变量与 config.json。

    Keys: api_key, api_url（不含 /chat/completions 的基础地址）, model。
    环境变量优先于 config.json；缺失值回退到默认（DeepSeek）值。
    """
    cfg = load_config()
    api_url = os.environ.get(ENV_API_URL, "").strip() or cfg.get("api_url", "").strip()
    model = os.environ.get(ENV_MODEL, "").strip() or cfg.get("model", "").strip()
    return {
        "api_key": get_api_key(),
        "api_url": api_url or DEFAULT_API_URL,
        "model": model or DEFAULT_MODEL,
    }