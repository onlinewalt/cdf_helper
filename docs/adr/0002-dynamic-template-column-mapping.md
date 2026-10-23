# ADR-0002: Dynamic column mapping for uploaded templates

Status: **Proposed** (not yet implemented).

## Context

The generator fills the customs-declaration (报关清单) template by writing parts
into a small set of hardcoded column indices:

```python
EXPECTED_HEADERS = ["序号", "备件名称", "数量", "单位", "重量(KG)", "单价/RMB", "金额RMB"]
SEQ_COL, NAME_COL, QTY_COL, UNIT_COL, WEIGHT_COL, PRICE_COL, AMOUNT_COL = 1, 2, 3, 4, 5, 6, 7
```

`validate_template` (called by `POST /generate`) enforces that an uploaded
template's Sheet1 row-2 headers match `EXPECTED_HEADERS` **in this exact
order**, and that the "合计" total row sits at row 7. The result:

- An uploaded template whose **headers are in a different order**, use **renamed
  labels** (e.g. `重量` instead of `重量(KG)`), add **extra columns**, or shift
  the **total row** is rejected with a flash message
  ("模板表头第 2 行与标准格式不符，无法确定列映射").
- A standard-layout template uploads and generates correctly (verified: 200,
  download link emitted).

This is a *strict-fail*, not a silent-corruption, gate — but it means the system
cannot yet "generate a file based on an uploaded template" in the general sense:
it only generates from templates that are byte-compatible with the one committed
workbook. The requirement, then, is to **fill an uploaded template dynamically
by reading its header row and mapping columns by label**, so non-standard but
structurally-similar templates (reordered/renamed columns, extra columns) are
accepted instead of rejected.

## Decision drivers

1. **User-reported**: "CDF helper is not able to generate file based on
   uploaded template" — the symptom is the strict header-order check, not a
   missing upload path (uploads/ is created at import, field names match
   `template_upload`, `validate_template` pre-flight runs).
2. **Correctness**: writing `Part.weight` (per-unit) or `price` to a fixed
   column index that the user has reordered would silently put weight in the
   price column — worse than rejecting. Any dynamic mapping must be
   header-driven, never index-assumed.
3. **The total row**: `generate` inserts rows before the "合计" total row when
   parts exceed `TEMPLATE_DATA_CAPACITY` (4), and writes SUM formulas anchored
   on `EXPECTED_TOTAL_ROW`. A dynamic template shifts the total row's position;
   the SUM anchors must be re-derived from the template, not hardcoded.
4. **Locality / deepening** (per repo architecture conventions): the column
   indices and header contract are currently a shallow cluster scattered across
   `generator.py` constants and `_validate_template_layout`. Extracting a
   `TemplateLayout` value (discovered once from the workbook, passed to
   `generate`) is a deepening seam with one real adapter — worth doing now
   because two callers (`do_generate` pre-flight validation, `generate`
   filling) already need the same layout facts.

## Options

### Option A — Dynamic mapping by header label (recommended)

Discover the column index of each logical field by scanning Sheet1 row 2 for
the canonical label (with a small alias set, e.g. `重量(KG)`/`重量`/`Weight`),
and locate the "合计" total row by *scanning* (first row whose A-cell is
`合计` or matches a total marker regex) instead of assuming row 7.

```
TemplateLayout(sheet_name, first_data_row=3, last_header_col,
               seq_col, name_col, qty_col, unit_col, weight_col,
               price_col, amount_col, total_row)
```

`validate_template` returns (or raises) a `TemplateLayout`; `generate` and
`_write_part_row`/`_write_total_row` consume the layout's column indices.

- **Pros**: genuinely supports reordered/relabelled columns + shifted total
  rows; the alias set is the only new maintenance surface; tests build a
  template with a *different* column order and assert the output still aligns.
- **Cons**: adds a layout-discovery step; must still reject truly malformed
  templates (missing required columns). Row-insertion logic (`_insert_rows`
  before total row) must use `layout.total_row` not the constant.

### Option B — Accept templates only if headers match exactly (status quo)

Keep the current `EXPECTED_HEADERS`-in-order check and the hardcoded indices.

- **Pros**: zero change; the fail-fast behavior is already user-friendly.
- **Cons**: does not satisfy the requirement — a reordered template still
  cannot be generated. The column indices remain a shallow cluster spread
  across the module, violating the deepening principle (the deletion test:
  removing the constants would scatter their meaning across `_write_part_row`,
  `_write_total_row`, and `_validate_template_layout`).

### Option C — Column-mapping UI on the upload form

Let the user map "which uploaded column is Quantity?" via the web form, then
persist a per-template mapping. (Heavy — not requested; deferred.)

## Decision

**Option A.** Map columns by header label with a small alias set, and detect
the total row by scanning. This is a deepening of the existing shallow
column-contract cluster into a `TemplateLayout` adapter behind a small
interface (`validate_template(path) -> TemplateLayout`), satisfying the
"generate from an uploaded template" requirement without silently accepting
garbage.

## Consequences

- `generator.py` gains a `TemplateLayout` (dataclass) built by a single
  `_discover_layout(wb)` reader, plus `_ALIASES` for the header variants
  accepted (e.g. `重量(KG)` ≡ `重量`). `_validate_template_layout` is renamed/refactored
  to return the layout (raising `ValueError` on missing required columns, as
  today).
- `generate`, `_write_part_row`, `_write_total_row`, `amount_formula` and the
  SUM-anchor logic consume `layout.<field>_col` / `layout.total_row` instead of
  the module constants `NAME_COL` etc. The constants become defaults or are
  removed if unused elsewhere.
- New hermetic tests in `test_generator.py`: a template with columns in a
  different order + a renamed header (e.g. `重量`) + total row at row 9 must
  still yield correctly-placed part rows and a working SUM.
- No AGENTS.md change needed (the "uploaded .xlsx → Sheet1" convention is
  unchanged); the existing template fixture workbook and the committed
  standard template are still the default.

## Open questions

- How tolerant should the alias set be? (Start conservative: exact-or-one-
  variant; expand only if a real uploaded template is rejected.)
- Should a *missing* optional column (e.g. 重量) leave the cell blank, as
  today's `Part.weight is None` guard already does? Yes — keep that behaviour.
- Does the total row scan need to handle templates whose "合计" sits in a
  non-A column (some users right-align it)? Defer to Option A's regex + linear
  scan; revisit if a real fixture misaligns.
