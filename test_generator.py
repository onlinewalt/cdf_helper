"""White-box tests for the generator Module.

Covers the narrowed interface exposed by Candidate 7:
- seq_num_formula / amount_formula / sum_formula (the A1 formula seam)
- _style_row_template / _apply_styles (the openpyxl private-_style adapter seam)
- generate end-to-end writes the right cell values + formulas into the template
"""
import sys
import tempfile
from pathlib import Path

import pytest
from openpyxl import load_workbook

sys.stdout.reconfigure(encoding="utf-8")

from cdf_helper.generator import (
    FIRST_DATA_ROW, TEMPLATE_DATA_CAPACITY,
    seq_num_formula, amount_formula, sum_formula,
    _style_row_template, _apply_styles, _write_part_row, _write_total_row,
    generate, validate_template,
)
from cdf_helper.parser import Part


def _template_path():
    """The committed template fixture (glob is locale-safe per test_web.py)."""
    for f in Path(".").glob("*报关清单*.xlsx"):
        return Path(f)
    pytest.skip("no committed template workbook found in repo root")


def test_formula_helpers_match_template_contract():
    """The formula seam reproduces the template's documented A1 formulas so
    a future FIRST_DATA_ROW shift updates seq + SUM anchors in one place."""
    assert seq_num_formula(3) == "=ROW()-2"      # FIRST_DATA_ROW=3 -> ROW()-2
    assert seq_num_formula(7) == "=ROW()-2"      # anchor is invariant to row
    assert amount_formula(5) == "=C5*F5"
    assert amount_formula(6) == "=C6*F6"
    assert sum_formula("C") == "=SUM(C$3:INDEX(C:C,ROW()-1))"
    assert sum_formula("E") == "=SUM(E$3:INDEX(E:E,ROW()-1))"
    assert sum_formula("G") == "=SUM(G$3:INDEX(G:G,ROW()-1))"


def test_style_capture_and_apply_seam():
    """_style_row_template / _apply_styles are the seam over openpyxl's private
    _style attribute; style must survive a capture -> clear -> apply round-trip."""
    tpl = _template_path()
    wb = load_workbook(tpl, data_only=False)
    ws = wb["Sheet1"]
    styles = _style_row_template(ws, FIRST_DATA_ROW)
    assert len(styles) == 7 and all(s is not None for s in styles)
    # clear a row then re-apply -> style present, value None
    row = FIRST_DATA_ROW + TEMPLATE_DATA_CAPACITY
    for col in range(1, 8):
        ws.cell(row, col).value = None
    _apply_styles(ws, row, styles)
    assert ws.cell(row, 3)._style is not None
    assert ws.cell(row, 3).value is None


def test_write_part_row_writes_values_and_formulas():
    """_write_part_row exercises the row-writing logic in isolation."""
    tpl = _template_path()
    wb = load_workbook(tpl, data_only=False)
    ws = wb["Sheet1"]
    r = FIRST_DATA_ROW
    # clear first so leftover template data doesn't leak
    for col in range(1, 8):
        ws.cell(r, col).value = None
    part = Part(name="阀门", qty=3, unit="只", type="DN50", weight=1.5, price=200.0)
    styles = _style_row_template(ws, FIRST_DATA_ROW)
    _write_part_row(ws, r, part, styles, include_spec=True)
    assert ws.cell(r, 1).value == seq_num_formula(r)
    assert ws.cell(r, 2).value == "阀门 DN50"           # include_spec True -> name + spec
    assert ws.cell(r, 3).value == 3                     # qty
    assert ws.cell(r, 4).value == "只"
    assert ws.cell(r, 5).value == 1.5
    assert ws.cell(r, 6).value == 200.0
    assert ws.cell(r, 7).value == amount_formula(r)


def test_write_total_row_writes_marker_and_sums():
    """_write_total_row exercises the total-row logic in isolation."""
    tpl = _template_path()
    wb = load_workbook(tpl, data_only=False)
    ws = wb["Sheet1"]
    total_row = FIRST_DATA_ROW + TEMPLATE_DATA_CAPACITY
    for col in range(1, 8):
        ws.cell(total_row, col).value = None
    total_styles = _style_row_template(ws, total_row)
    _write_total_row(ws, total_row, total_styles)
    assert ws.cell(total_row, 1).value == "合计"
    assert ws.cell(total_row, 3).value == sum_formula("C")
    assert ws.cell(total_row, 5).value == sum_formula("E")
    assert ws.cell(total_row, 7).value == sum_formula("G")


def test_generate_end_to_end_over_more_than_capacity():
    """generate with >4 parts must insert rows before the total row and the
    SUM formulas must still anchor to C$3 (data start), total row lands at
    FIRST_DATA_ROW + N."""
    tpl = _template_path()
    assert validate_template(tpl) is None  # pre-flight
    parts = [
        Part(name=f"备件{i}", qty=i, unit="个", type=f"Type-{i}", weight=1.0, price=float(i * 10))
        for i in range(1, 8)  # 7 parts -> exceeds TEMPLATE_DATA_CAPACITY (4)
    ]
    with tempfile.TemporaryDirectory() as td:
        out = generate(tpl, parts, vessel_name="远怡湖", output_dir=Path(td),
                       output_name="远怡湖-报关清单.xlsx", include_spec=True)
        out_path = Path(out)
        assert out_path.is_file()
        wb = load_workbook(out_path, data_only=False)
        ws = wb["Sheet1"]
        total_row = FIRST_DATA_ROW + len(parts)  # 3 + 7 = 10
        assert ws.cell(1, 1).value == "船名：远怡湖"
        assert ws.cell(total_row, 1).value == "合计"
        assert ws.cell(total_row, 3).value == sum_formula("C")
        # last data row: seq formula + amount formula use that row's index
        last = FIRST_DATA_ROW + len(parts) - 1
        assert ws.cell(last, 1).value == seq_num_formula(last)
        assert ws.cell(last, 7).value == amount_formula(last)
        # row capacity matched: rows 3..(3+N-1) are data, total at row 3+N
        assert ws.cell(total_row, 1).value == "合计"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
