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
    FIRST_DATA_ROW, LAST_HEADER_COL, TEMPLATE_DATA_CAPACITY,
    TemplateLayout, seq_num_formula, amount_formula, sum_formula,
    _style_row_template, _apply_styles, _write_part_row, _write_total_row,
    generate, validate_template,
)
from cdf_helper.parser import Part


def _std_layout() -> TemplateLayout:
    return TemplateLayout.default()


def _template_path():
    """The committed template fixture (glob is locale-safe per test_web.py)."""
    for f in Path(".").glob("*报关清单*.xlsx"):
        return Path(f)
    pytest.skip("no committed template workbook found in repo root")


def test_formula_helpers_match_template_contract():
    """The formula seam reproduces the template's documented A1 formulas so
    a future FIRST_DATA_ROW shift updates seq + SUM anchors in one place.

    With the standard layout (first_data_row=3, cols A-G), the layout-aware
    seams must produce the exact strings the standard template expects.
    """
    L = _std_layout()
    assert seq_num_formula(L, 3) == "=ROW()-2"      # FIRST_DATA_ROW=3 -> ROW()-2
    assert seq_num_formula(L, 7) == "=ROW()-2"      # anchor is invariant to row
    assert amount_formula(L, 5) == "=C5*F5"
    assert amount_formula(L, 6) == "=C6*F6"
    assert sum_formula(L, "qty") == "=SUM(C$3:INDEX(C:C,ROW()-1))"
    assert sum_formula(L, "weight") == "=SUM(E$3:INDEX(E:E,ROW()-1))"
    assert sum_formula(L, "amount") == "=SUM(G$3:INDEX(G:G,ROW()-1))"


def test_style_capture_and_apply_seam():
    """_style_row_template / _apply_styles are the seam over openpyxl's private
    _style attribute; style must survive a capture -> clear -> apply round-trip."""
    tpl = _template_path()
    wb = load_workbook(tpl, data_only=False)
    ws = wb["Sheet1"]
    styles = _style_row_template(ws, FIRST_DATA_ROW, LAST_HEADER_COL)
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
    L = _std_layout()
    # clear first so leftover template data doesn't leak
    for col in range(1, 8):
        ws.cell(r, col).value = None
    part = Part(name="阀门", qty=3, unit="只", type="DN50", weight=1.5, price=200.0)
    styles = _style_row_template(ws, FIRST_DATA_ROW)
    _write_part_row(ws, r, part, styles, layout=L, include_spec=True)
    assert ws.cell(r, 1).value == seq_num_formula(L, r)
    assert ws.cell(r, 2).value == "阀门 DN50"           # include_spec True -> name + spec
    assert ws.cell(r, 3).value == 3                     # qty
    assert ws.cell(r, 4).value == "只"
    # weight: template column is 重量(KG) per-ROW total (its total row does
    # =SUM(E$3...)), and Part.weight is per-unit — so the writer emits qty*weight.
    assert ws.cell(r, 5).value == 1.5 * 3
    assert ws.cell(r, 6).value == 200.0                 # unit price (per-unit)
    assert ws.cell(r, 7).value == amount_formula(L, r)  # amount formula

    # None weight -> blank cell (must not raise), preserving pre-multiplication
    # behaviour for parts the AI left unfilled.
    part_no_w = Part(name="阀门", qty=3, unit="只", weight=None)
    _write_part_row(ws, r + 1, part_no_w, styles, layout=L, include_spec=True)
    assert ws.cell(r + 1, 5).value is None


def test_write_total_row_writes_marker_and_sums():
    """_write_total_row exercises the total-row logic in isolation."""
    tpl = _template_path()
    wb = load_workbook(tpl, data_only=False)
    ws = wb["Sheet1"]
    L = _std_layout()
    total_row = FIRST_DATA_ROW + TEMPLATE_DATA_CAPACITY
    for col in range(1, 8):
        ws.cell(total_row, col).value = None
    total_styles = _style_row_template(ws, total_row)
    _write_total_row(ws, total_row, total_styles, L)
    assert ws.cell(total_row, 1).value == "合计"
    assert ws.cell(total_row, 3).value == sum_formula(L, "qty")
    assert ws.cell(total_row, 5).value == sum_formula(L, "weight")
    assert ws.cell(total_row, 7).value == sum_formula(L, "amount")


def test_generate_end_to_end_over_more_than_capacity():
    """generate with >4 parts must insert rows before the total row and the
    SUM formulas must still anchor to the layout's first_data_row (C$3 for the
    standard template); total row lands at FIRST_DATA_ROW + N."""
    tpl = _template_path()
    assert validate_template(tpl) is None  # pre-flight
    L = _std_layout()
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
        assert ws.cell(total_row, 3).value == sum_formula(L, "qty")
        # last data row: seq formula + amount formula use that row's index
        last = FIRST_DATA_ROW + len(parts) - 1
        assert ws.cell(last, 1).value == seq_num_formula(L, last)
        assert ws.cell(last, 7).value == amount_formula(L, last)
        # row capacity matched: rows 3..(3+N-1) are data, total at row 3+N
        assert ws.cell(total_row, 1).value == "合计"


def _write_reordered_template(path):
    """A standard-layout template whose columns are REORDERED (quantity before
    name, spec last) plus a renamed weight header ('重量' vs '重量(KG)') —
    exercises _discover_layout's label-based mapping (Option A of ADR-0002)."""
    from openpyxl import Workbook
    wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws.merge_cells("A1:G1"); ws["A1"] = "船名："
    # column order: 数量 / 备件名称 / 单位 / 重量 / 单价/RMB / 金额RMB / 序号
    ws["A2"], ws["B2"], ws["C2"] = "数量", "备件名称", "单位"
    ws["D2"], ws["E2"], ws["F2"], ws["G2"] = "重量", "单价/RMB", "金额RMB", "序号"
    ws["A3"], ws["B3"], ws["C3"], ws["D3"], ws["E3"], ws["F3"], ws["G3"] = (
        3, "阀门", "只", 1.5, 200, "=E3*C3", 1)
    ws["A7"] = "合计"
    wb.save(path); wb.close()


def test_reorderd_template_columns_are_mapped_by_label():
    """Dynamic template column mapping: an uploaded template with reordered
    AND relabelled columns (qty before name; '重量' not '重量(KG)') must still
    emit parts to the correctly-mapped cells and SUM to the right column letters."""
    from cdf_helper.generator import _discover_layout
    with tempfile.TemporaryDirectory() as td:
        tpl = Path(td) / "reordered.xlsx"
        _write_reordered_template(tpl)
        assert validate_template(tpl) is None  # accepted (labels match aliases)
        disc = load_workbook(tpl)
        lay = _discover_layout(disc); disc.close()
        assert lay.qty_col == 1 and lay.name_col == 2 and lay.weight_col == 4
        assert lay.seq_col == 7 and lay.total_row == 7
        parts = [Part(name="阀门", qty=3, unit="只", weight=1.5, price=200.0)]
        out = generate(tpl, parts, vessel_name="船", output_dir=Path(td),
                       output_name="out.xlsx", include_spec=False)
        wb = load_workbook(out); ws = wb["Sheet1"]
        # qty lands in col A (1), name in col B (2), weight = 1.5*3 in col D (4)
        assert ws.cell(3, lay.qty_col).value == 3
        assert ws.cell(3, lay.name_col).value == "阀门"
        assert ws.cell(3, lay.weight_col).value == 4.5
        # SUM in total row uses the *discovered* column letters, not C/E/G
        assert ws.cell(lay.total_row, lay.qty_col).value == sum_formula(lay, "qty")
        assert ws.cell(lay.total_row, lay.weight_col).value == sum_formula(lay, "weight")
        assert ws.cell(lay.total_row, lay.amount_col).value == sum_formula(lay, "amount")
        wb.close()


def test_total_row_position_for_under_capacity_parts():
    """Regression for a pre-existing latent bug: with <= capacity parts the total
    row must stay at the template's original 合计 position (row 7 on the standard
    template), not FIRST_DATA_ROW + N — otherwise a stale duplicate 合计 survives
    and SUM anchors drift."""
    tpl = _template_path()
    parts = [Part(name=f"备件{i}", qty=i, unit="个", weight=1.0, price=float(i))
             for i in range(1, 3)]  # 2 parts -> under capacity (4)
    L = _std_layout()
    with tempfile.TemporaryDirectory() as td:
        out = generate(tpl, parts, vessel_name="船", output_dir=Path(td),
                       output_name="船-港-报关清单-2026-08-20.xlsx")
        wb = load_workbook(out); ws = wb["Sheet1"]
        # parts occupy rows 3..4; total 合计 sits at the template's row 7
        assert ws.cell(3, 2).value == "备件1"
        assert ws.cell(4, 2).value == "备件2"
        for blank_row in range(5, 7):  # rows 5,6 cleared
            assert ws.cell(blank_row, 2).value is None
        assert ws.cell(L.total_row, 1).value == "合计"
        assert ws.cell(L.total_row, 3).value == sum_formula(L, "qty")
        assert ws.cell(L.total_row, 5).value == sum_formula(L, "weight")
        assert ws.cell(L.total_row, 7).value == sum_formula(L, "amount")
        # no duplicate 合计 anywhere else above the total row
        assert not any(ws.cell(r, 1).value == "合计" for r in range(3, L.total_row))


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
