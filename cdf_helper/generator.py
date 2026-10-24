"""Generate the CDF (报关清单) workbook by filling the user's template.

The template layout (Sheet1):
    A1 (merged A1:G1) : title "船名：<vessel>"
    row 2             : headers 序号/备件名称/数量/单位/重量(KG)/单价/RMB/金额RMB
    rows 3.N         : one row per part (序号 and 金额 are formulas)
    last data row + 1 : 合计 total row (yellow fill, SUM formulas)
"""

import copy
import re
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

TEMPLATE_SHEET = "Sheet1"
FIRST_DATA_ROW = 3
HEADER_ROW = 2
LAST_HEADER_COL = 7
TEMPLATE_DATA_CAPACITY = 4  # rows 3..6 in the blank template

# --- Template layout contract (validated up-front, see validate_template) ---
EXPECTED_HEADERS = ["序号", "备件名称", "数量", "单位", "重量(KG)", "单价/RMB", "金额RMB"]
TOTAL_MARKER = "合计"
EXPECTED_TOTAL_ROW = FIRST_DATA_ROW + TEMPLATE_DATA_CAPACITY  # 7
# Columns O/P (15/16) in the blank template carry stray "套"/2 leftovers.
LEFTOVER_COLS = (15, 16)

# Header aliases accepted when discovering a column's role from an uploaded
# template. Conservative by default (exact-or-one-variant); extend only if a
# real template is rejected with a header-mismatch error.
_HEADER_ALIASES = {
    "序号":   {"序号"},
    "备件名称": {"备件名称", "名称", "品名", "Particulars", "Description", "Item"},
    "数量":   {"数量", "Qty", "Quantity"},
    "单位":   {"单位", "Unit"},
    "重量(KG)": {"重量(KG)", "重量", "Weight", "KG"},
    "单价/RMB": {"单价/RMB", "单价", "价格", "Unit Price", "Price"},
    "金额RMB": {"金额RMB", "金额", "Amount"},
}
# Required fields for a template to be fillable; 序号/金额 rows are written by
# the generator (formula / total), so a template missing those columns is still
# acceptable, but the five below must be present.
_REQUIRED_FIELDS = {"备件名称", "数量", "单位", "重量(KG)", "单价/RMB", "金额RMB"}

# --- Unit translation seam ---------------------------------------------------
# English packing-list units (PCE/SET/PC/EA) are translated to the Chinese unit
# the 报关清单's 单位 column expects; Chinese units pass through; unknown tokens
# pass through unchanged; None/empty -> None (the writer applies its "个"
# default). Applied at output time (_write_part_row) so the parser's Part.unit
# keeps source fidelity and existing parser tests are unaffected. Pure &
# stateless — testable directly without parsing a workbook.
_UNIT_ALIASES = {
    "pcs": "件", "pce": "件", "pc": "个", "pc.": "个",
    "ea": "个", "each": "个", "set": "套", "set.": "套",
    "件": "件", "套": "套", "个": "个", "只": "只", "包": "包",
    "kg": "KG", "kgs": "KG", "g": "g", "gs": "g",
}


def normalize_unit(u):
    """Translate an English unit word (PCE/SET/PC/EA) to its Chinese equivalent
    (件/套/个); pass Chinese units through; ``None``/empty -> ``None``."""
    if u is None:
        return None
    s = str(u).strip()
    if not s:
        return None
    return _UNIT_ALIASES.get(s.lower(), s)


@dataclass
class TemplateLayout:
    """Resolved column mapping discovered from a template's header row.

    The deep module's interface: `_discover_layout(wb)` is the single reader;
    every writer (`generate`, `_write_part_row`, `_write_total_row`, the formula
    seams) consumes `layout.<field>_col` instead of hardcoded indices/letters, so
    an uploaded template with reordered or renamed columns is filled correctly
    rather than rejected or silently misfiled. `first_data_row`/`total_row`
    locate where data begins and where the 合计 row sits (scanned, not assumed).
    """
    sheet_name: str = TEMPLATE_SHEET
    first_data_row: int = FIRST_DATA_ROW
    seq_col: int = 1
    name_col: int = 2
    qty_col: int = 3
    unit_col: int = 4
    weight_col: int = 5
    price_col: int = 6
    amount_col: int = 7
    total_row: int = EXPECTED_TOTAL_ROW
    last_header_col: int = LAST_HEADER_COL

    def col_letter(self, name: str) -> str:
        return get_column_letter(getattr(self, f"{name}_col"))

    def field_col(self, name: str) -> int:
        return getattr(self, f"{name}_col")

    @classmethod
    def default(cls) -> "TemplateLayout":
        """Standard-template layout (first_data_row=3, cols 1..7, total row 7)."""
        return cls()


def _discover_layout(wb, sheet_name=TEMPLATE_SHEET):
    """Read Sheet1's header row + total-marker row and return a TemplateLayout.

    Columns are mapped by **label** (via `_HEADER_ALIASES`), so reordered or
    renamed headers are accepted. The total row is *scanned* (first row at/after
    the header whose A-cell matches `TOTAL_MARKER` / a total-row regex), so a
    shifted 合计 row is handled. Raises `ValueError` with an actionable message
    when a required column is missing or the total row is not found.
    """
    if sheet_name not in wb.sheetnames:
        names = ", ".join(wb.sheetnames) if wb.sheetnames else "（空）"
        raise ValueError(
            f'模板缺少名为"{sheet_name}"的工作表，无法填充。'
            f'当前工作表：{names}。'
            '请使用标准报关清单模板（工作表名为 Sheet1）。'
        )
    ws = wb[sheet_name]
    raw_headers = [ws.cell(HEADER_ROW, c).value for c in range(1, LAST_HEADER_COL + 1)]
    col_by_field = {}
    last_header_col = LAST_HEADER_COL
    # extend scan beyond the nominal 7 columns in case an uploaded template
    # pads with extra columns before one of the required labels.
    scan_cols = max(LAST_HEADER_COL, len([h for h in raw_headers if h is not None]))
    for col in range(1, scan_cols + 1):
        raw = ws.cell(HEADER_ROW, col).value
        if raw is None:
            continue
        label = str(raw).strip()
        if label == TOTAL_MARKER:
            continue
        matched_field = None
        for fld, aliases in _HEADER_ALIASES.items():
            if label in aliases:
                matched_field = fld
                break
        if matched_field and matched_field not in col_by_field:
            col_by_field[matched_field] = col
            last_header_col = max(last_header_col, col)

    missing = [f for f in _REQUIRED_FIELDS if f not in col_by_field]
    if missing:
        found = {ws.cell(HEADER_ROW, c).value for c in range(1, last_header_col + 1)}
        raise ValueError(
            f'模板表头（第 {HEADER_ROW} 行）缺少必需列：{", ".join(missing)}。'
            f'已找到的表头：{sorted(str(h) for h in found if h)}。'
            '请使用标准报关清单模板，第 2 行为：' + " / ".join(EXPECTED_HEADERS) + '。'
        )

    # --- locate the 合计 total row by scanning from the header downward ---
    total_row = None
    for row in range(HEADER_ROW + 1, ws.max_row + 1):
        v = ws.cell(row, 1).value
        if v is None:
            continue
        s = str(v).strip()
        if s == TOTAL_MARKER or re.match(r"合计\s*[:：]*", s) or "合计" in s:
            total_row = row
            break
    if total_row is None:
        # fall back to the standard position if no marker is present
        total_row = FIRST_DATA_ROW + TEMPLATE_DATA_CAPACITY

    return TemplateLayout(
        sheet_name=sheet_name,
        first_data_row=FIRST_DATA_ROW,
        seq_col=col_by_field.get("序号", 1),
        name_col=col_by_field["备件名称"],
        qty_col=col_by_field["数量"],
        unit_col=col_by_field["单位"],
        weight_col=col_by_field["重量(KG)"],
        price_col=col_by_field["单价/RMB"],
        amount_col=col_by_field["金额RMB"],
        total_row=total_row,
        last_header_col=last_header_col,
    )


def seq_num_formula(layout: TemplateLayout, row: int) -> str:
    """序号 auto-increment: ``=ROW()-(first_data_row-1)``.

    Anchored to `layout.first_data_row` (row 3 -> `=ROW()-2`), so a shift in
    the data-start row updates both the seq column and the SUM anchor in one
    place. Centralised as a seam so it can't drift from the SUM anchor.
    """
    return f"=ROW()-{layout.first_data_row - 1}"


def amount_formula(layout: TemplateLayout, row: int) -> str:
    """金额: ``=Qty<row>*Price<row>`` (数量 * 单价), relative to the part's row.

    Column letters come from `layout`, so a reordered template still computes
    amount from the right cells.
    """
    qty_l = layout.col_letter("qty")
    price_l = layout.col_letter("price")
    return f"={qty_l}{row}*{price_l}{row}"


def sum_formula(layout: TemplateLayout, fld: str) -> str:
    """合计: ``=SUM(<field>$<first_data_row>:INDEX(<field>:<field>,ROW()-1))``.

    Sums the data column from `layout.first_data_row` up to the row above the
    total row. Column letter and `$N` anchor come from `layout`, so a reordered
    or shifted template still produces a correct, self-contained SUM.
    """
    col_letter = layout.col_letter(fld)
    return (
        f"=SUM({col_letter}${layout.first_data_row}:"
        f"INDEX({col_letter}:{col_letter},ROW()-1))"
    )


def _copy_style(src_cell, dst_cell):
    dst_cell._style = copy.copy(src_cell._style)


def _style_row_template(ws, row, last_col=LAST_HEADER_COL):
    """Capture the cell styles of one template row (data or total) as a list."""
    return [copy.copy(ws.cell(row, col)._style) for col in range(1, last_col + 1)]


def _apply_styles(ws, row, styles):
    for col, style in enumerate(styles, start=1):
        ws.cell(row, col)._style = copy.copy(style)


def _validate_template_layout(wb, sheet_name=TEMPLATE_SHEET) -> TemplateLayout:
    """Return the resolved TemplateLayout, raising ``ValueError`` if the
    template cannot be filled.

    Delegates to `_discover_layout` (which maps columns by header label and
    scans for the total row). A ValueError surfaces an actionable header/total
    mismatch message so the web layer can fail fast instead of producing a
    silently corrupt workbook.
    """
    return _discover_layout(wb, sheet_name)


def validate_template(template_path, sheet_name=TEMPLATE_SHEET):
    """Standalone pre-flight check (does **not** mutate the file).

    Returns ``None`` when the template is acceptable, or a human-readable
    error string.  Used by the web layer to give early feedback before
    source files are parsed.
    """
    template_path = Path(template_path)
    if not template_path.is_file():
        return f'模板文件不存在：{template_path}。'

    suffix = template_path.suffix.lower()
    if suffix == ".xls":
        return (
            f'模板文件“{template_path.name}”是旧版 .xls 格式，'
            '仅支持 .xlsx 模板。请转换为 .xlsx 后重试，'
            '或使用服务器提供的标准报关清单模板。'
        )
    if suffix != ".xlsx":
        return (
            f'模板文件“{template_path.name}”格式不受支持（{suffix}）。'
            '请上传 .xlsx 报关清单模板。'
        )

    try:
        wb = load_workbook(template_path, data_only=False, read_only=True)
    except Exception as exc:
        return f'无法读取模板文件：{exc}。'

    try:
        _validate_template_layout(wb, sheet_name)
    except ValueError as exc:
        return str(exc)
    finally:
        wb.close()

    return None


def _write_part_row(ws, row, part, data_styles, layout, include_spec):
    """Write one part's row (序号/name/qty/unit/weight/price/amount) and style it.

    Column indices come from `layout`, so reordered headers are filled to the
    correct cells. 重量(KG) is emitted as per-ROW total (per-unit kg * qty); the
    sheet's 合计 total row SUMs it. A part with no weight leaves the cell blank.
    """
    ws.cell(row, layout.seq_col).value = seq_num_formula(layout, row)
    ws.cell(row, layout.name_col).value = _display_name(part, include_spec)
    ws.cell(row, layout.qty_col).value = part.qty
    ws.cell(row, layout.unit_col).value = normalize_unit(part.unit) or "个"
    ws.cell(row, layout.weight_col).value = (part.weight * part.qty) if part.weight is not None else None
    ws.cell(row, layout.price_col).value = part.price
    ws.cell(row, layout.amount_col).value = amount_formula(layout, row)
    _apply_styles(ws, row, data_styles)
    ws.row_dimensions[row].height = 20


def _write_total_row(ws, row, total_styles, layout):
    """Write the '合计' total row: marker in col A, SUM over qty/weight/amount.

    All columns other than the total-row marker, the SUM columns (qty / weight
    / amount) are blanked — matching the standard template's total row, which
    carries only 合计 + SUM(C) + SUM(E) + SUM(G).
    """
    ws.cell(row, 1).value = TOTAL_MARKER
    ws.cell(row, layout.qty_col).value = sum_formula(layout, "qty")
    ws.cell(row, layout.weight_col).value = sum_formula(layout, "weight")
    ws.cell(row, layout.amount_col).value = sum_formula(layout, "amount")
    # blank the remaining header-span columns (name/unit/price) for tidiness
    for col in range(1, layout.last_header_col + 1):
        v = ws.cell(row, col).value
        if v is None:
            continue
        if col in (1, layout.qty_col, layout.weight_col, layout.amount_col):
            continue
        ws.cell(row, col).value = None
    _apply_styles(ws, row, total_styles)
    ws.row_dimensions[row].height = 20


def generate(template_path, parts, vessel_name, output_dir, output_name,
             include_spec=True, sheet_name=TEMPLATE_SHEET):
    """Fill the template with parts and save to output_dir/output_name.

    Returns the absolute path of the generated file. The template's column
    mapping is discovered by header label (see `_discover_layout`), so
    reordered/renamed-column templates are filled to the right cells.
    """
    template_path = Path(template_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    wb = load_workbook(template_path, data_only=False)
    layout = _validate_template_layout(wb, sheet_name)
    ws = wb[sheet_name]

    # Title cell
    ws.cell(1, 1).value = f"船名：{vessel_name}"

    data_styles = _style_row_template(ws, layout.first_data_row, layout.last_header_col)
    total_styles = _style_row_template(ws, layout.total_row, layout.last_header_col)

    # Clear the sample data rows (their styles are reused via data_styles)
    for r in range(layout.first_data_row, layout.total_row):
        for col in range(1, layout.last_header_col + 1):
            ws.cell(r, col).value = None

    # Insert extra rows before the total row if we have more parts than capacity
    n = len(parts)
    capacity = layout.total_row - layout.first_data_row
    if n > capacity:
        ws.insert_rows(layout.total_row, n - capacity)
        total_row = layout.first_data_row + n  # insert pushed the original total down
    else:
        # ≤ capacity: keep the template's original total-row position (it was
        # NOT cleared by the range above, since range excludes total_row).
        total_row = layout.total_row
    for i, part in enumerate(parts):
        r = layout.first_data_row + i
        _write_part_row(ws, r, part, data_styles, layout, include_spec)

    _write_total_row(ws, total_row, total_styles, layout)

    # Tidy up template leftovers outside the table (cols O/P on the standard
    # 7-wide template) if they fall inside the header span.
    for col in LEFTOVER_COLS:
        if col > layout.last_header_col:
            continue
        cell = ws.cell(total_row, col)
        if cell.value is not None:
            cell.value = None

    out_path = output_dir / output_name
    wb.save(out_path)
    wb.close()
    return str(out_path)


def _display_name(part, include_spec):
    name = part.name.strip()
    if include_spec and part.type:
        spec = part.type.strip()
        if spec and spec != name:
            return f"{name} {spec}"
    return name


def sanitize_filename(name, fallback="file"):
    """Remove characters that are invalid in Windows filenames."""
    cleaned = re.sub(r'[\\/:*?"<>|\r\n\t]', "-", str(name)).strip()
    cleaned = cleaned.strip(" .")
    return cleaned or fallback