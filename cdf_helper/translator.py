"""Translate spare-part names from English to Chinese via the Niutrans API.

- **Always-on**: automatically translates part names detected as English (no CJK chars).
- Non-fatal: API failures or missing credentials skip translation gracefully.
- Results cached in ``translate_cache.json`` (keyed by ``sha1(name)``) to avoid repeat charges.
- Only translates names with no Chinese characters — already-Chinese names are left untouched.
"""

import hashlib
import json
import re
import time
from pathlib import Path

try:
    import requests
except ImportError:
    requests = None

from cdf_helper.parser import Part

API_URL = "https://api.niutrans.com/v2/text/translate"
TIMEOUT = 30
CACHE_PATH = Path(__file__).resolve().parent.parent / "translate_cache.json"
BATCH_SIZE = 20

_CJK = re.compile(r"[\u4e00-\u9fff]")


def _is_english(text: str) -> bool:
    """True when *text* has no CJK characters (likely English or code)."""
    return bool(text) and not _CJK.search(text)


def _cache_key(name: str) -> str:
    return hashlib.sha1(name.encode("utf-8")).hexdigest()[:16]


class Translator:
    def __init__(self, app_id: str, apikey: str, cache_path: Path = CACHE_PATH):
        if not app_id or not apikey:
            raise ValueError("缺少翻译 API 凭据（app_id / apikey）")
        if requests is None:
            raise RuntimeError("需要安装 requests 库：pip install requests")
        self.app_id = app_id
        self.apikey = apikey
        self.cache_path = Path(cache_path)
        self._cache = self._load_cache()

    # ---- cache ---------------------------------------------------------
    def _load_cache(self) -> dict:
        if self.cache_path.exists():
            try:
                return json.loads(self.cache_path.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, OSError):
                return {}
        return {}

    def _save_cache(self) -> None:
        try:
            self.cache_path.parent.mkdir(parents=True, exist_ok=True)
            self.cache_path.write_text(
                json.dumps(self._cache, ensure_ascii=False, indent=1), encoding="utf-8"
            )
        except OSError:
            pass

    # ---- API -----------------------------------------------------------
    def _generate_auth_str(self, params: dict) -> str:
        """Generate MD5 auth string: sorted(params + apikey) joined by &."""
        sorted_params = sorted(list(params.items()) + [("apikey", self.apikey)], key=lambda x: x[0])
        param_str = "&".join(f"{k}={v}" for k, v in sorted_params)
        return hashlib.md5(param_str.encode("utf-8")).hexdigest()

    def _translate_api(self, text: str):
        """Call the Niutrans API for one text. Returns translated string or None."""
        data = {
            "from": "en",
            "to": "zh",
            "appId": self.app_id,
            "timestamp": int(time.time()),
            "srcText": text,
        }
        data["authStr"] = self._generate_auth_str(data)

        for attempt in range(2):
            try:
                response = requests.post(API_URL, data=data, timeout=TIMEOUT)
                if response.status_code == 429 and attempt == 0:
                    time.sleep(5)
                    continue
                response.raise_for_status()
                resp = response.json()
                translated = self._extract_text(resp)
                if translated:
                    return translated.strip()
                return None
            except (requests.RequestException, json.JSONDecodeError, KeyError, TypeError):
                if attempt == 0:
                    time.sleep(5)
                    continue
                return None
        return None

    @staticmethod
    def _extract_text(resp: dict) -> str:
        """Extract translated text from various possible Niutrans response formats."""
        if not isinstance(resp, dict):
            return ""
        data = resp.get("data")
        if isinstance(data, dict):
            for key in ("translatedText", "result", "dst"):
                val = data.get(key)
                if val:
                    if isinstance(val, list):
                        return "".join(str(v) for v in val)
                    return str(val)
        for key in ("translatedText", "result", "dst"):
            val = resp.get(key)
            if val:
                if isinstance(val, list):
                    return "".join(str(v) for v in val)
                return str(val)
        return ""

    # ---- public API ----------------------------------------------------
    def translate_names(self, parts, on_status=None) -> dict:
        """Translate English part names to Chinese in-place.

        Returns stats dict: {translated, from_cache, errors}
        """
        stats = {"translated": 0, "from_cache": 0, "errors": 0}

        def _report(msg):
            if on_status:
                on_status(msg)

        targets = [(i, p) for i, p in enumerate(parts) if p.name and _is_english(p.name)]
        if not targets:
            _report("所有备件名称均为中文，无需翻译。")
            return stats

        _report(f"发现 {len(targets)} 条英文备件名称，开始翻译 ...")

        # 1) Cache hits
        pending = []
        for i, part in targets:
            key = _cache_key(part.name)
            cached = self._cache.get(key)
            if cached is not None:
                part.name = cached
                stats["from_cache"] += 1
            else:
                pending.append((i, part))

        _report(f"命中本地缓存 {stats['from_cache']} 条。")

        # 2) Batch API calls
        batches = [pending[k:k + BATCH_SIZE] for k in range(0, len(pending), BATCH_SIZE)]
        for bi, batch in enumerate(batches, 1):
            _report(f"正在翻译（批次 {bi}/{len(batches)}，{len(batch)} 条）...")
            for i, part in batch:
                original = part.name
                translated = self._translate_api(original)
                self._cache[_cache_key(original)] = translated
                if translated:
                    part.name = translated
                    stats["translated"] += 1
                else:
                    stats["errors"] += 1
            self._save_cache()

        _report(f"翻译完成：新增 {stats['translated']} 条，缓存命中 {stats['from_cache']} 条，失败 {stats['errors']} 条。")
        return stats


def translate_names(parts, app_id, apikey, cache_path=CACHE_PATH, on_status=None) -> dict:
    """Convenience entry: create Translator and translate English part names in-place."""
    provider = Translator(app_id=app_id, apikey=apikey, cache_path=cache_path)
    return provider.translate_names(parts, on_status=on_status)
