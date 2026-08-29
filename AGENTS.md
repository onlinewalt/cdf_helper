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
- Type-check: non-blocking `mypy --strict --ignore-missing-imports cdf_helper` (surfaces annotation gaps and real type errors; does not fail PRs).
- A separate `mutation` job (non-blocking, `continue-on-error`) runs `mutmut` over the covered lines of `cdf_helper/` — including `parser._find_packing_header` and the `AIProvider` mapping — and reports surviving mutants for follow-up. It never fails PRs.
- Local mutation testing: `mutmut run` (mutmut does **not** run on native Windows; use WSL or rely on the CI job).

## Gotchas

1. **Port 5000 may be taken** — `main.py:_free_port` scans 5000–5019 for a free port automatically.
2. **Formulas use A1-style relative refs** (`=ROW()-2`, `=C{r}*F{r}`) — rows are inserted before the total row when parts exceed 4.
3. **xlrd 2.0+ only reads `.xls`** — `.xlsx` is handled by openpyxl. Do not install the old xlrd 1.x.
4. **Style copying is cell-level** — `generator._copy_style` uses the private `_style` attribute. This works on current openpyxl versions but may break on major upgrades.
5. **No CSRF protection on the web form** — do not expose to the public internet.
6. **`SECRET_KEY` is hardcoded** (`webapp.py` line 54) — fine for local dev only.
7. **Chinese locale assumed** — encoding errors may appear if system locale doesn't support UTF-8; `main.py:_ensure_utf8_stdio()` mitigates this.

## Diagnosing unprocessable sheets

When a source workbook (or a single sheet) yields 0 parts, run this checklist before opening the file. Every step is tied to a real `parser.py` symbol so you can `grep`/`read` quickly instead of guessing.

### Step 1 — capture the parser's own diagnostics
`parse_source(path, warn=...)` calls back per sheet / per row with the reason it skipped or could not parse. Two messages to watch for:

- `f"{path.name}/{sheet.name}: 未找到 Item/Quantity 表头，跳过该表"` → "no Item/Quantity header → skip sheet" (the sheet was identified as an English packing list but had no recognisable header row).
- `f"{path.name} ...缺...1 列: {name}"` → a header row *was* found but an individual data row was missing a column and was dropped (recoverable).

If **all** sheets across the file resolve to 0 parts, `parse_source` raises `ValueError("在 {path.name} 中没有解析到备件数据")`. Always pass a `warn` callable — silent 0-result runs hide the cause.

Quick scan script (repo root):
```python
from pathlib import Path
from cdf_helper.parser import parse_source
warns = []
parts = parse_source(Path("SOME_FILE.xlsx"), warn=warns.append)
print("parts:", len(parts))
for w in warns: print(" -", w)
```

### Step 2 — root-cause checklist (symbol that owns it)

- **Workbook fails to open / every sheet missing.** `_open_sheets` (.xlsx branch) first calls `_make_openpyxl_lenient()`, which monkeypatches `openpyxl.worksheet._reader._cast_number` so a non-numeric value in a number-typed cell (e.g. a stray `.`) returns a string instead of raising `ValueError: could not convert string to float` (which would otherwise abort the whole load). `.xls` goes through `xlrd`. If the crash persists, the patch wasn't applied (`_LENIENT_OPENPYXL` guard) or openpyxl changed its internals.
- **File misrouted as the vessel table, not data.** A file whose text contains `船名` is consumed by `detect_vessel` / `_split_zh_en` (regex `r"船名\s*[:：]\s*(\S+)"`; also `船名/单位` / `船名／单位`) as the bilingual vessel-name lookup, **not** as parts data — this is a caller-level convention in `main.py`/`webapp.py` source selection. If a genuine data file literally contains "船名" it may be skipped; call `detect_vessel(paths)` first to confirm intent.
- **Chinese header sheet not recognised.** `_find_header_row` (Chinese path) keys off `HEADER_ALIASES` via the precomputed `_ALIAS_TO_FIELD` table, normalising each cell with `_norm` (collapses *all* whitespace incl. newlines via `_WHITESPACE_RE`). A wrapped label like `订\n订单数\n量\n(NUM)` is matched only because `_norm` strips the `\n`; if `_norm` is bypassed, `订...订单数...` never contains `数量`/`金额` and the header row is missed → the sheet falls through to the English packing path → usually 0 parts.
- **Packed single-column format not taken.** `_is_packed_format(mapping)` is True only when every mapped field points to the *same* column (`len(set(cols)) == 1`). Headers spread across ≥2 cells/rows → False → routes to `_parse_standard`/`_parse_multiline_row` (which may then yield 0 parts).
- **English packing list: no header found → sheet skipped.** `_find_packing_header` (docstring: rows 696-709) requires a row carrying **both** an Item-like label **and** a Quantity-like label; a footer/data row never carries both, so this is safe. No such row → no header → `_parse_packing_sheet` walks rows for the End-of-Listing sentinel; if none, 0 parts.
- **OCR-corrupted labels.** `_find_packing_header` is OCR-tolerant: exact substrings first, then `_close_enough(t, keywords)` (Levenshtein ≤ `max_dist=2` via `_dist`). `Quantity`→`Quanty`/`Q'ty`/`qtt`, `Item`→`ftem`/`tem`/`particulars`/`description`, a `Num` column = quantity. `Qantily` is dist 2 to `quantity` → passes. Tokens farther than 2 (e.g. `Quantitty`) do **not** — extend the keyword tuple or raise `max_dist`.
- **End-of-Listing boundary over/under-shoots.** `_END_RE = re.compile(r"end\s*of\s*listing", re.IGNORECASE)` tolerates OCR merges (`End ofListing`, `endoflisting`). `_find_end_of_listing(rows, start_idx, last_row)` bounds `_collect_packing_items`; a different sentinel (`End of Table`, `***`) bleeds rows — extend `_END_RE` or `_is_signature`.
- **Header found but rows malformed.** `_parse_multiline_row` / `_parse_standard` emit per-row `warn("...缺...1 列")` when a row lacks the expected column count; the row is skipped, not fatal. A high ratio ⇒ the column mapping (from `_find_header_row`) is wrong — inspect `mapping`.

### Step 3 — regression / invariant checks
Each `test_*.py` test below locks down one historical failure; if a real file regresses, reproduce it hermetically with an `openpyxl.Workbook()` helper (pattern: `_write_*_workbook` + `assert len(parse_source(...)) == N`):

- `test_bad_number_cell_xlsx_loads` — corrupt numeric cell no longer aborts the `.xlsx` load (`_make_openpyxl_lenient`).
- `test_multiline_wrapped_header` — wrapped `订 订单数 量 (NUM)` collapses under `_norm`.
- `test_deshanghai_no_item_header_and_headerless` — Item-less header + headerless End-of-Listing recovery (`_parse_packing_headerless`).
- `test_yuantong_ocr_header_and_merged_footer` — OCR `Qantily` header + merged `End ofListing` footer.
- `test_packed_synthetic` (and `test_packed_real_file_if_present`) — packed single-column extraction (`_parse_packed_sheet` / `_is_packed_format`).
- `test_end_of_listing_embedded_in_data` / `test_synthetic` — `_END_RE` boundary + End-of-Listing inside a data row.

Green baseline: `python3 -m pytest -q` → `19 passed`; `ruff check .` → `All checks passed!`.

### Step 4 — when the whole file raises
`parse_source` raises `ValueError("在 {path.name} 中没有解析到备件数据")` when **every** sheet yields 0 parts. In that case loop `_open_sheets(path)` and call `_find_header_row`, `_find_packing_header`, `_find_end_of_listing` per sheet manually — exactly one of the Step-2 branches above will reveal the offending sheet.

## File references

- Template example (do not delete, used by tests): `veseel name-destination port-报关清单-date.xlsx`
- Vessel name lookup table: `中英文船名25-5-14.xls`
- Source data example (new format with multi-line/merged cells): `new file.xlsx`
- Source data example (packed single-column format): `远怡湖missing sheets.xlsx`
- Windows launcher: `启动CDF助手.bat` (ASCII-only to avoid encoding issues)
