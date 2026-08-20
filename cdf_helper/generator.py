"""Generate the CDF (报关清单) workbook by filling the user's template.

The template layout (Sheet1):
    A1 (merged A1:G1) : title "船名：<vessel>"
    row 2             : headers 序号/备件名称/数量/单位/重量(KG)/单价/RMB/金额RMB
    rows 3..N         : one row per part (序号 and 金额 are formulas)
    last data row + 1 : 合计 total row (yellow fill, SUM formulas)
"""

import copy
import re
from pathlib import Path

from openpyxl import load_workbook

TEMPLATE_SHEET = "Sheet1"
FIRST_DATA_ROW = 3
HEADER_ROW = 2
LAST_HEADER_COL = 7
TEMPLATE_DATA_CAPACITY = 4  # rows 3..6 in the blank template


def _copy_style(src_cell, dst_cell):
    dst_cell._style = copy.copy(src_cell._style)


def _style_row_template(ws, row):
    """Capture the cell styles of one template row (data or total) as a list."""
    return [copy.copy(ws.cell(row, col)._style) for col in range(1, LAST_HEADER_COL + 1)]


def _apply_styles(ws, row, styles):
    for col, style in enumerate(styles, start=1):
        ws.cell(row, col)._style = copy.copy(style)


def generate(template_path, parts, vessel_name, output_dir, output_name,
             include_spec=True, sheet_name=TEMPLATE_SHEET):
    """Fill the template with parts and save to output_dir/output_name.

    Returns the absolute path of the generated file.
    """
    template_path = Path(template_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    wb = load_workbook(template_path, data_only=False)
    ws = wb[sheet_name]

    # Title cell
    ws.cell(1, 1).value = f"船名：{vessel_name}"

    data_styles = _style_row_template(ws, FIRST_DATA_ROW)
    total_row_in_template = FIRST_DATA_ROW + TEMPLATE_DATA_CAPACITY  # 7
    total_styles = _style_row_template(ws, total_row_in_template)

    # Clear the sample rows 3..6 (their styles are reused via data_styles)
    for r in range(FIRST_DATA_ROW, FIRST_DATA_ROW + TEMPLATE_DATA_CAPACITY):
        for col in range(1, LAST_HEADER_COL + 1):
            ws.cell(r, col).value = None

    # Insert extra rows before the total row if we have more parts than capacity
    n = len(parts)
    if n > TEMPLATE_DATA_CAPACITY:
        ws.insert_rows(total_row_in_template, n - TEMPLATE_DATA_CAPACITY)

    total_row = FIRST_DATA_ROW + n

    for i, part in enumerate(parts):
        r = FIRST_DATA_ROW + i
        ws.cell(r, 1).value = "=ROW()-2"
        ws.cell(r, 2).value = _display_name(part, include_spec)
        ws.cell(r, 3).value = part.qty
        ws.cell(r, 4).value = part.unit
        ws.cell(r, 5).value = part.weight
        ws.cell(r, 6).value = part.price
        ws.cell(r, 7).value = f"=C{r}*F{r}"
        _apply_styles(ws, r, data_styles)
        ws.row_dimensions[r].height = 20

    # Total row
    ws.cell(total_row, 1).value = "合计"
    ws.cell(total_row, 2).value = None
    ws.cell(total_row, 3).value = "=SUM(C$3:INDEX(C:C,ROW()-1))"
    ws.cell(total_row, 4).value = None
    ws.cell(total_row, 5).value = "=SUM(E$3:INDEX(E:E,ROW()-1))"
    ws.cell(total_row, 6).value = None
    ws.cell(total_row, 7).value = "=SUM(G$3:INDEX(G:G,ROW()-1))"
    _apply_styles(ws, total_row, total_styles)
    ws.row_dimensions[total_row].height = 20

    # Tidy up template leftovers outside the table (row 7 cols O/P)
    for col in (15, 16):
        cell = ws.cell(total_row, col)
        if cell.value is not None:
            cell.value = None

    out_path = output_dir / output_name
    wb.save(out_path)
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