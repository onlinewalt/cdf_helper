"""DeepSeek 智能估算：为缺失 重量/单价 的备件调用 DeepSeek 补充数据。

- 仅填补来源文件中缺失的 重量(KG) / 单价(RMB)，不覆盖已有值。
- 分批（BATCH_SIZE 条/次）请求，降低调用次数与成本。
- 结果缓存在 ai_cache.json（按 名称+规格 哈希），重复生成不重复计费。
"""

import hashlib
import json
import time
import urllib.error
import urllib.request

from pathlib import Path

API_URL = "https://api.deepseek.com/chat/completions"
MODEL = "deepseek-chat"
BATCH_SIZE = 50
CACHE_PATH = Path(__file__).resolve().parent.parent / "ai_cache.json"
TIMEOUT = 120

_SYSTEM_PROMPT = (
    "你是一名船舶备件采购与海关申报助理。请根据备件的名称、规格和单位，"
    "估算其重量（单位：公斤 KG）与单价（单位：人民币元 RMB）。"
    "基于船舶行业常识给出合理、保守的估计；确实无法判断时，对应字段返回 null。"
)


def _cache_key(name: str, spec: str) -> str:
    return hashlib.sha1(f"{name}|{spec}".encode("utf-8")).hexdigest()[:16]


class AIProvider:
    def __init__(self, api_key: str, cache_path: Path = CACHE_PATH):
        if not api_key:
            raise ValueError("缺少 DeepSeek API Key")
        self.api_key = api_key
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

    # ---- enrichment ----------------------------------------------------
    def enrich(self, parts, on_status=None) -> dict:
        """估算缺失重量/单价的备件，就地写入 part.weight / part.price。

        Returns stats dict: {requested, filled, from_cache, errors}
        """
        stats = {"requested": 0, "filled": 0, "from_cache": 0, "errors": 0}

        def _report(msg):
            if on_status:
                on_status(msg)

        targets = [
            (i, p)
            for i, p in enumerate(parts)
            if p.weight is None or p.price is None
        ]
        if not targets:
            _report("所有备件已包含重量和单价，无需调用 DeepSeek。")
            return stats

        stats["requested"] = len(targets)
        _report(f"有 {len(targets)} 条备件缺少 重量/单价，开始用 DeepSeek 估算 ...")

        # 1) 先用本地缓存
        pending = []
        for i, part in targets:
            cached = self._cache.get(_cache_key(part.name, part.type or ""))
            if cached:
                if part.weight is None and cached.get("weight") is not None:
                    part.weight = cached["weight"]
                if part.price is None and cached.get("price") is not None:
                    part.price = cached["price"]
                stats["from_cache"] += 1
            else:
                pending.append((i, part))
        _report(f"命中本地缓存 {stats['from_cache']} 条。")

        # 2) 分批调用 API
        batches = [pending[k:k + BATCH_SIZE] for k in range(0, len(pending), BATCH_SIZE)]
        for bi, batch in enumerate(batches, 1):
            _report(f"正在请求 DeepSeek（批次 {bi}/{len(batches)}，{len(batch)} 条）...")
            results = self._call_batch(batch)
            if results is None:
                stats["errors"] += len(batch)
                _report("  该批次请求失败，相关条目标记跳过。")
                continue
            for (i, part) in batch:
                r = results.get(str(i))
                if not r:
                    stats["errors"] += 1
                    continue
                weight = r.get("weight_kg")
                price = r.get("unit_price")
                if part.weight is None and weight is not None:
                    part.weight = float(weight)
                if part.price is None and price is not None:
                    part.price = float(price)
                # 无论是否填成功都写缓存，避免重复计费
                self._cache[_cache_key(part.name, part.type or "")] = {
                    "weight": part.weight,
                    "price": part.price,
                }
                stats["filled"] += 1
            self._save_cache()

        _report(f"DeepSeek 估算完成：新增 {stats['filled']} 条，失败 {stats['errors']} 条。")
        return stats

    # ---- API -----------------------------------------------------------
    def _call_batch(self, batch):
        """Call DeepSeek for one batch; returns {id_str: {weight_kg, unit_price}} or None."""
        lines = []
        for idx, (i, part) in enumerate(batch, 1):
            spec = part.type or "无"
            lines.append(f'{idx}. 名称: {part.name}；规格: {spec}；单位: {part.unit or "个"}')
        prompt = (
            "请为以下备件估算重量与单价，以 JSON 格式返回，不要输出任何其他内容：\n"
            '{"items":[{"id":1,"weight_kg":0.5,"unit_price":100}]}\n'
            "备件列表：\n" + "\n".join(lines)
        )

        body = {
            "model": MODEL,
            "messages": [
                {"role": "system", "content": _SYSTEM_PROMPT},
                {"role": "user", "content": prompt},
            ],
            "temperature": 0.2,
            "response_format": {"type": "json_object"},
        }
        content = self._post_json(body)
        if content is None:
            return None

        data = self._parse_json(content)
        if not data:
            return None

        mapping = {}
        items = data.get("items") if isinstance(data, dict) else None
        if not isinstance(items, list):
            return None
        for item in items:
            if not isinstance(item, dict) or "id" not in item:
                continue
            try:
                pos = int(item["id"])
            except (TypeError, ValueError):
                continue
            if not 1 <= pos <= len(batch):
                continue
            i = batch[pos - 1][0]
            mapping[str(i)] = {
                "weight_kg": self._to_float(item.get("weight_kg")),
                "unit_price": self._to_float(item.get("unit_price")),
            }
        return mapping

    def _post_json(self, body: dict):
        """POST to DeepSeek and return the assistant message text (or None on failure)."""
        payload = json.dumps(body).encode("utf-8")
        req = urllib.request.Request(
            API_URL,
            data=payload,
            headers={
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json",
            },
        )
        for attempt in range(2):  # one retry on 429/5xx
            try:
                with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
                    raw = resp.read().decode("utf-8")
                obj = json.loads(raw)
                return obj["choices"][0]["message"]["content"]
            except urllib.error.HTTPError as e:
                if e.code in (429, 500, 502, 503, 504) and attempt == 0:
                    time.sleep(5)
                    continue
                return None
            except (urllib.error.URLError, TimeoutError, OSError, KeyError, ValueError):
                return None
        return None

    @staticmethod
    def _parse_json(text: str):
        text = text.strip()
        # 兼容模型偶尔输出 ```json 包裹
        if text.startswith("```"):
            text = text.strip("`")
            if text.startswith("json"):
                text = text[4:]
            text = text.strip()
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            return None

    @staticmethod
    def _to_float(value):
        if value is None:
            return None
        try:
            return round(float(value), 2)
        except (TypeError, ValueError):
            return None


def enrich_parts(parts, api_key, cache_path=CACHE_PATH, on_status=None) -> dict:
    """便捷入口：创建 provider 并 enrich。"""
    provider = AIProvider(api_key=api_key, cache_path=cache_path)
    return provider.enrich(parts, on_status=on_status)