# AGENTS.md — CDF Helper

Single-package Python tool (no virtualenv, no build step). Run directly with Python 3.9+.

## Commands

| Task | Command |
|---|---|
| Web UI (starts on port 5000) | `python3 main.py` |
| CLI generate | `python3 main.py generate --template T.xlsx --source S.xls --vessel "远怡湖 COSMERRY LAKE" --port 上海` |
| Run all tests | `python3 test_ai.py && python3 test_web.py && python3 test_packing.py` |
| Run packing parser tests only | `python3 test_packing.py` |
| Run AI tests only | `python3 test_ai.py` |

Install deps (already satisfied in most environments):
```
python3 -m pip install -r requirements.txt
```

## Architecture

```
main.py                  # entry: `python3 main.py` → Web UI; `generate` → CLI
webapp.py                # Flask app, routes: GET /, POST /generate, GET /download/<file>
cdf_helper/
├── parser.py    # Excel → List[Part] (Chinese headers + English packing lists + vessel detection)
├── generator.py # fills template.xlsx → output .xlsx
├── ai.py        # DeepSeek enrichment for missing weight/price (batching + local cache)
├── config.py    # API key from env DEEPSEEK_API_KEY > config.json
└── __init__.py  # version 0.2.0
```

Entry points:
- **Web**: `python3 main.py` (no args) → Flask on 127.0.0.1:5000
- **CLI**: `python3 main.py generate ...`
- **Tests**: `test_ai.py`, `test_web.py`, `test_packing.py` — each is a standalone script with `if __name__ == "__main__"`

## Key conventions

- **Templates**: sheet name is hardcoded to `"Sheet1"` in `generator.TEMPLATE_SHEET`. Template layout:
  - A1 (merged A1:G1): title `船名：<name>`
  - Row 2: headers (序号 / 备件名称 / 数量 / 单位 / 重量(KG) / 单价/RMB / 金额RMB)
  - Rows 3–6: sample data (cleared and reused for styling)
  - Row 7: total row (yellow fill)
- **Source files**: any `.xls`/`.xlsx` in the project root is a candidate. Files containing "船名" are treated as the vessel-name lookup table, not source data. Files containing "报关清单" are treated as templates.
- **Vessel name**: auto-detected from source files. Format: `中文 英文` (e.g., `远怡湖 COSMERRY LAKE`), resolved via `中英文船名25-5-14.xls` lookup workbook.
- **AI**: optional (`--ai` flag or "use_ai" checkbox). Key from `--api-key`, `config.json`, or `DEEPSEEK_API_KEY` env var. Results cached in `ai_cache.json` (keyed by `sha1(name|spec)`) to avoid repeat charges. API failures are non-fatal — missing fields stay empty.
- **Filenames**: sanitized via `generator.sanitize_filename` (strips `\/:*?"<>|`). Output pattern: `<vessel>-<port>-报关清单-<date>.xlsx`.
- **Temp files**: `uploads/` and `generated/` are gitignored. Old files (>7 days) are cleaned on webapp startup.

## Testing notes

- `test_web.py` and `test_packing.py` use `glob` to find real Excel files in the project root. Tests may print "SKIP" if expected files are absent — this is normal.
- `test_web.py` creates an isolated temp `config.json` to avoid clobbering real config.
- `test_ai.py` mocks `_post_json` — no real API calls.
- `test_packing.py:test_real_file_if_present` requires files in `uploads/` — skipped if absent.
- To run a single test function (not just a file): there's no pytest; run via `python3 -c "import test_packing; test_packing.test_synthetic()"`.

## Gotchas

1. **Port 5000 may be taken** — `main.py:_free_port` scans 5000–5019 for a free port automatically.
2. **Formulas use A1-style relative refs** (`=ROW()-2`, `=C{r}*F{r}`) — rows are inserted before the total row when parts exceed 4.
3. **xlrd 2.0+ only reads `.xls`** — `.xlsx` is handled by openpyxl. Do not install the old xlrd 1.x.
4. **Style copying is cell-level** — `generator._copy_style` uses the private `_style` attribute. This works on current openpyxl versions but may break on major upgrades.
5. **No CSRF protection on the web form** — do not expose to the public internet.
6. **`SECRET_KEY` is hardcoded** (`webapp.py` line 54) — fine for local dev only.
7. **Chinese locale assumed** — encoding errors may appear if system locale doesn't support UTF-8; `main.py:_ensure_utf8_stdio()` mitigates this.

## File references

- Template example (do not delete, used by tests): `veseel name-destination port-报关清单-date.xlsx`
- Vessel name lookup table: `中英文船名25-5-14.xls`
- Windows launcher: `启动CDF助手.bat` (ASCII-only to avoid encoding issues)
