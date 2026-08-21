"""本地配置读写（DeepSeek API Key 等），支持 config.json 与环境变量。"""

import json
import os
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = PROJECT_ROOT / "config.json"

ENV_API_KEY = "DEEPSEEK_API_KEY"
ENV_TRANS_APP_ID = "TRANSLATE_APP_ID"
ENV_TRANS_APIKEY = "TRANSLATE_APIKEY"


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


def get_trans_credentials() -> tuple:
    """返回翻译 API 凭据 (app_id, apikey)：优先环境变量，其次 config.json。"""
    app_id = os.environ.get(ENV_TRANS_APP_ID, "").strip()
    apikey = os.environ.get(ENV_TRANS_APIKEY, "").strip()
    if not app_id or not apikey:
        cfg = load_config()
        app_id = app_id or cfg.get("trans_app_id", "").strip()
        apikey = apikey or cfg.get("trans_apikey", "").strip()
    return app_id, apikey