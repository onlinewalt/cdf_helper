"""Parse spare-parts source Excel files into a normalized list of parts.

A Part is a dataclass with:
    name      - part name (名称 / 物资名称)
    qty       - quantity (数量)
    unit      - unit (单位)
    type      - model / spec (规格 / 物资规格/品牌)
    weight    - weight in KG (重量), optional
    price     - unit price in RMB (单价), optional

Both .xlsx (openpyxl) and legacy .xls (xlrd) files are supported.
The header row is located automatically by recognizable column names.
"""

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from openpyxl import load_workbook

try:
    import xlrd
except ImportError:
    xlrd = None


@dataclass
class Part:
    name: str
    qty: float = 1.0
    unit: str = "个"
    type: Optional[str] = None
    weight: Optional[float] = None
    price: Optional[float] = None


HEADER_ALIASES = {
    "name": ("备件名称", "名称", "备件", "品名", "物料名称", "物资名称", "名称及规格", "名称/规格", "part name", "name", "item"),
    "type": ("型号", "规格", "型号规格", "规格型号", "物资规格", "物资规格/品牌", "规格/品牌", "type", "spec", "model", "品牌"),
    "qty": ("数量", "件数", "qty", "quantity", "数量/套", "数量  "),
    "unit": ("单位", "unit", "计量单位", "单位  "),
    "weight": ("重量", "重量(kg)", "weight", "单重", "毛重"),
    "price": ("单价", "单价/rmb", "unit price", "price", "价格"),
}

_CONTAINS_RULES = (
    # (needle, field) checked in order; a header cell that contains any needle maps to the field
    ("名称及规格", "name"),
    ("名称", "name"),
    ("品名", "name"),
    ("规格", "type"),
    ("型号", "type"),
    ("品牌", "type"),
    ("数量", "qty"),
    ("件数", "qty"),
    ("单位", "unit"),
    ("重量", "weight"),
    ("单价", "price"),
    ("价格", "price"),
)


def _norm(s: str) -> str:
    return str(s).strip().replace(" ", "").replace("\u3000", "").lower()


class _Cell:
    """Minimal unified cell wrapper over xlrd / openpyxl cells."""

    __slots__ = ("value", "column", "row")

    def __init__(self, value, column, row):
        self.value = value
        self.column = column
        self.row = row


class _Sheet:
    """Minimal unified sheet wrapper supporting iter_rows() and cell()."""

    def __init__(self, source):
        self._source = source
        if source.__class__.__module__.startswith("xlrd"):
            self._kind = "xls"
        else:
            self._kind = "xlsx"

    def iter_rows(self):
        if self._kind == "xls":
            for r in range(self._source.nrows):
                yield [
                    _Cell(self._source.cell_value(r, c), c + 1, r + 1)
                    for c in range(self._source.ncols)
                ]
        else:
            for row in self._source.iter_rows():
                yield [_Cell(c.value, c.column, c.row) for c in row]

    def cell(self, row, column):
        if self._kind == "xls":
            return _Cell(self._source.cell_value(row - 1, column - 1), column, row)
        return self._source.cell(row, column)


def _open_sheet(path: Path):
    path = Path(path)
    if path.suffix.lower() == ".xls":
        if xlrd is None:
            raise RuntimeError("需要 xlrd 库来读取 .xls 文件：pip install xlrd")
        wb = xlrd.open_workbook(str(path))
        return _Sheet(wb.sheet_by_index(0))
    wb = load_workbook(path, data_only=True)
    return _Sheet(wb.active)


def _find_header_row(sheet):
    """Return (header_row_idx, {field: col_idx}) for the first row that looks like a parts header."""
    for cells in sheet.iter_rows():
        mapping = {}
        for cell in cells:
            if cell.value is None:
                continue
            text = _norm(cell.value)
            if not text:
                continue
            matched = False
            for field, aliases in HEADER_ALIASES.items():
                for alias in aliases:
                    if _norm(alias) == text:
                        if field not in mapping:
                            mapping[field] = cell.column
                        matched = True
                        break
                if matched:
                    break
            if matched:
                continue  # exact alias wins; do not also contains-match this cell
            for needle, field in _CONTAINS_RULES:
                if field in mapping:
                    continue
                if needle in text:
                    mapping[field] = cell.column
                    break
        if "name" in mapping and "qty" in mapping:
            return cells[0].row, mapping
    return None, None


def _to_number(value) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "").replace("，", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        pass
    m = re.search(r"-?\d+(?:\.\d+)?", text)
    if m:
        try:
            return float(m.group(0))
        except ValueError:
            return None
    return None


def _to_text(value) -> Optional[str]:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def parse_source(path, warn=None) -> list:
    """Parse a source Excel file into a list of Part objects.

    Optional callable warn(msg) receives messages about rows that could not be
    fully parsed (e.g. missing qty) so the caller can report them.
    """
    path = Path(path)
    sheet = _open_sheet(path)
    header_row, mapping = _find_header_row(sheet)
    if header_row is None:
        raise ValueError(f"无法在 {path.name} 中找到表头行（需要含 名称 与 数量 列）")

    name_col = mapping["name"]
    qty_col = mapping["qty"]
    type_col = mapping.get("type")
    if type_col == name_col:
        type_col = None
    unit_col = mapping.get("unit")
    weight_col = mapping.get("weight")
    price_col = mapping.get("price")

    parts = []
    for cells in sheet.iter_rows():
        if cells[0].row <= header_row:
            continue
        name = _to_text(cells[name_col - 1].value)
        if name is None:
            continue
        qty = _to_number(cells[qty_col - 1].value)
        if qty is None:
            if warn:
                warn(f"{path.name} 第 {cells[0].row} 行缺少数量，按 1 处理: {name}")
            qty = 1.0
        unit = _to_text(cells[unit_col - 1].value) if unit_col else None
        parts.append(
            Part(
                name=name,
                qty=qty,
                unit=unit or "个",
                type=_to_text(cells[type_col - 1].value) if type_col else None,
                weight=_to_number(cells[weight_col - 1].value) if weight_col else None,
                price=_to_number(cells[price_col - 1].value) if price_col else None,
            )
        )

    if not parts:
        raise ValueError(f"在 {path.name} 中没有解析到备件数据")

    return parts


def parse_sources(paths, warn=None) -> list:
    """Parse one or more source files, merging all parts in order."""
    all_parts = []
    for p in paths:
        all_parts.extend(parse_source(p, warn=warn))
    return all_parts


def _iter_cells(sheet):
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is not None:
                yield cell


def detect_vessel(paths) -> Optional[str]:
    """Try to guess the vessel name from the source files (船名 / 船名/单位 fields)."""
    for p in paths:
        p = Path(p)
        try:
            sheet = _open_sheet(p)
        except Exception:
            continue
        for cell in _iter_cells(sheet):
            text = str(cell.value)
            m = re.search(r"船名\s*[:：]\s*(\S+)", text)
            if m:
                return m.group(1)
            if "船名/单位" in text or "船名／单位" in text:
                nxt = sheet.cell(cell.row, cell.column + 1).value
                if nxt is not None and str(nxt).strip():
                    return str(nxt).strip()
    return None