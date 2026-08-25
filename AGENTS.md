# AGENTS.md — CDF Helper

Single-package Python tool (no virtualenv, no build step). Run directly with Python 3.9+.

## Commands

| Task | Command |
|---|---|
| Web UI (starts on port 5000) | `python3 main.py` |
| CLI generate | `python3 main.py generate --template T.xlsx --source S.xls --vessel "远怡湖 COSMERRY LAKE" --port 上海` |
| Run all tests (pytest) | `python3 -m pytest` |
| Run all tests (standalone) | `python3 test_ai.py && python3 test_web.py && python3 test_packing.py` |
| Run packing parser tests only | `python3 -m pytest test_packing.py` |
| Run AI tests only | `python3 -m pytest test_ai.py` |

Install runtime deps (already satisfied in most environments):
```
python3 -m pip install -r requirements.txt
```
Install dev/test deps (pytest, coverage, ruff):
```
python3 -m pip install -r requirements-dev.txt
```

## Architecture

```
main.py                  # entry: `python3 main.py` → Web UI; `generate` → CLI
webapp.py                # Flask app, routes: GET /, POST /generate, GET /download/<file>
cdf_helper/
├── parser.py    # Excel → List[Part] (Chinese headers + English packing lists + multi-line cells + vessel detection)
├── generator.py # fills template.xlsx → output .xlsx
├── ai.py        # DeepSeek enrichment for missing weight/price (batching + local cache)
├── config.py    # API key from env DEEPSEEK_API_KEY > config.json
└── __init__.py  # version 0.2.0
```

Entry points:
- **Web**: `python3 main.py` (no args) → Flask on 127.0.0.1:5000
- **CLI**: `python3 main.py generate ...`
- **Tests**: `test_ai.py`, `test_web.py`, `test_packing.py` — pytest-based. Each has an `if __name__ == "__main__": pytest.main(...)` guard, so it also runs standalone (`python3 test_*.py`). Shared fixtures live in `conftest.py`.

## Key conventions

- **Templates**: sheet name is hardcoded to `"Sheet1"` in `generator.TEMPLATE_SHEET`. Template layout:
  - A1 (merged A1:G1): title `船名：<name>`
  - Row 2: headers (序号 / 备件名称 / 数量 / 单位 / 重量(KG) / 单价/RMB / 金额RMB)
  - Rows 3–6: sample data (cleared and reused for styling)
  - Row 7: total row (yellow fill)
- **Source files**: any `.xls`/`.xlsx` in the project root is a candidate. Files containing "船名" are treated as the vessel-name lookup table, not source data. Files containing "报关清单" are treated as templates. `test_web.py` previously used `glob("*.xls*")` to pick up both `.xls` and `.xlsx` from root.
- **Packed single-column format**: when all header fields (序号/设备/名称/编号/单位/数量/备注) are in a single cell with 2+ space-separated headers, the parser detects this and uses `_parse_packed_sheet` to extract parts from space-separated data rows. Handles both multi-row (header and data in separate rows) and multi-line cell (header and data in same cell, newline-separated) variants.
- **Vessel name**: auto-detected from source files. Format: `中文 英文` (e.g., `远怡湖 COSMERRY LAKE`), resolved via `中英文船名25-5-14.xls` lookup workbook. Trailing hyphens in English names (e.g. `YINNIAN-`) are stripped.
- **AI**: optional (`--ai` flag or "use_ai" checkbox). Key from `--api-key`, `config.json`, or `DEEPSEEK_API_KEY` env var. Results cached in `ai_cache.json` (keyed by `sha1(name|spec)`) to avoid repeat charges. API failures are non-fatal — missing fields stay empty.
- **Filenames**: sanitized via `generator.sanitize_filename` (strips `\/:*?"<>|`). Output pattern: `<vessel>-<port>-报关清单-<date>.xlsx`.
- **Temp files**: `uploads/` and `generated/` are gitignored. Old files (>7 days) are cleaned on webapp startup.

## Testing notes

- Tests run under `pytest`. Each `test_*.py` also runs standalone via its `pytest.main` guard.
- `test_web.py` is **hermetic**: the template is read from the committed repo fixture and the source workbook is built in memory and uploaded through the Flask client, so it needs **no root data files** and is stable in CI (sample data files like `new file.xlsx`, `远棠湾-舟山.xlsx`, etc. are gitignored).
- `test_packing.py` and `test_ai.py` are hermetic except for a few explicit "real file" checks, which `pytest.skip()` when their fixture is absent. The committed template workbook (`veseel name-destination port-报关清单-date.xlsx`) and the vessel lookup (`中英文船名25-5-14.xls`) are always available.
- The AI tests mock the DeepSeek HTTP call (`AIProvider._post_json`) — no network.
- To run a single test: `python3 -m pytest test_packing.py::test_synthetic` (or filter with `-k test_synthetic`).

## CI

- `.github/workflows/tests.yml` runs on every push/PR: `ruff check .` (pyflakes gate) + `pytest --cov=cdf_helper` with a coverage report artifact.

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
- Source data example (new format with multi-line/merged cells): `new file.xlsx`
- Source data example (packed single-column format): `远怡湖missing sheets.xlsx`
- Windows launcher: `启动CDF助手.bat` (ASCII-only to avoid encoding issues)
