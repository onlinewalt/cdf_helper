"""Parse spare-parts source Excel files into a normalized list of parts.

A Part is a dataclass with:
    name      - part name (名称 / 物资名称 / Particulars)
    qty       - quantity (数量)
    unit      - unit (单位)
    type      - model / spec (规格 / Type / 物资规格/品牌)
    weight    - weight in KG (重量), optional
    price     - unit price in RMB (单价), optional

Supported source layouts:
1. 中文签收单/物料清单：表头含 名称 与 数量（自动识别）。
2. 英文 Receipt/Packing List（如远通海事 Yuantong）：表头含 Item/Quantity(Unit)/Particulars，
   数量形如 "N PCE"，以 "** End of Listing **" 结尾；多 sheet 各自为一张装箱单。
Both .xlsx (openpyxl) and legacy .xls (xlrd) files are supported.
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

# ---- packing-list (English Receipt/Packing List) patterns -------------
_QTY_RE = re.compile(
    r"(?<!\d)(\d+(?:\.\d+)?)\s*"
    r"(PCE|PCS|PC|SETS|SET|SHEET|PKT|MTR|PAIR|CASE|CAN|EA|BAG|ROLL|BIL)(?!\w)",
    re.IGNORECASE,
)
_END_RE = re.compile(r"end\s+of\s+listing", re.IGNORECASE)

_IRRELEVANT_RE = re.compile(
    r"^\s*(?:"
    r"\*\*|\(\d+\)|（\d+）|Equipment|Item|Quantity|Particulars|Part\s*No|Serial\s*No|"
    r"Cus\s*Ref|Status|Gross\s*Weight|Order\s*No|Order\s*Date|Msg\s*No|Vessel|Owner\s*Ref|"
    r"Department|Delivery|Package|Coordinator|Email|Tel|Date|Signature|Issued|Page|"
    r"RI#|Shipment#|http"
    r")",
    re.IGNORECASE,
)

_UNIT_WORDS = frozenset(
    w.upper()
    for w in (
        "PCE", "PCS", "PC", "SET", "SETS", "SHEET", "PKT", "MTR", "PAIR",
        "CASE", "CAN", "EA", "BAG", "ROLL", "个", "件", "套", "台", "支",
        "根", "张", "块", "盒", "米", "只", "片",
    )
)

_FOOTER_KEYWORDS = (
    "signature",
    "盖章", "蓋章", "蓝章", "藍章",
    "签收", "簽收", "簸收",
    "issued by",
    "date 日期",
    "htt",
)


def _trim_footer(text: str) -> str:
    """Cut trailing footer fragments (signatures, URLs) glued onto a name."""
    low = text.lower()
    idx = len(text)
    for kw in _FOOTER_KEYWORDS:
        i = low.find(kw)
        if i != -1:
            idx = min(idx, i)
    if idx < len(text):
        text = text[:idx].strip(" -–—")
    return text


def _norm(s: str) -> str:
    return str(s).strip().replace(" ", "").replace("\u3000", "").lower()


def _clean(value) -> str:
    """Collapse whitespace/newlines of a cell value into a single line."""
    if value is None:
        return ""
    text = str(value).replace("\r", " ").replace("\n", " ")
    return " ".join(text.split())


def _clean_name(text: str) -> str:
    """Drop pure-number / stray-unit / separator tokens from a name fragment."""
    tokens = []
    for t in text.split():
        t = t.strip(":：")
        if re.fullmatch(r"[\d.,\-]+", t):
            continue
        if re.fullmatch(r"[=\-*]+", t):
            continue
        if t.upper() in _UNIT_WORDS:
            continue
        tokens.append(t)
    return " ".join(tokens).strip(" -–—")


def _split_type(text: str):
    """Return (text_without_type, type_part). Splits a "Type:xxx" tail out of a cell."""
    m = re.search(r"Type\s*[:：]", text, re.IGNORECASE)
    if not m:
        return text, None
    before = text[: m.start()].strip(" -–—")
    after = text[m.end():].strip()
    return before, after or None


def _is_irrelevant(text: str) -> bool:
    if text.startswith(":") or text.endswith(":"):
        return True
    if re.match(r"dwg\.?", text, re.IGNORECASE):
        return True
    if re.fullmatch(r"[\d.,\-()（）\s]+", text):
        return True
    return bool(_IRRELEVANT_RE.match(text))


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

    @property
    def name(self) -> str:
        if self._kind == "xls":
            return str(self._source.name)
        return str(self._source.title)

    @property
    def last_row(self) -> int:
        if self._kind == "xls":
            return int(self._source.nrows)
        return int(self._source.max_row)

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


def _open_sheets(path: Path):
    path = Path(path)
    if path.suffix.lower() == ".xls":
        if xlrd is None:
            raise RuntimeError("需要 xlrd 库来读取 .xls 文件：pip install xlrd")
        wb = xlrd.open_workbook(str(path))
        return [_Sheet(s) for s in wb.sheets()]
    wb = load_workbook(path, data_only=True)
    return [_Sheet(ws) for ws in wb.worksheets]


# ---- standard (Chinese header) parsing --------------------------------

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


_SIGNATURE_KEYWORDS = ("签字", "签收", "盖章", "供船", "签收日期")


def _is_signature(text: str) -> bool:
    return any(k in text for k in _SIGNATURE_KEYWORDS)


def _parse_standard(sheet, mapping, header_row, path, warn) -> list:
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
        if _is_signature(name):
            continue  # signature / footer rows must not become parts
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
    return parts


# ---- packing-list (English Receipt/Packing List) parsing --------------

def _find_packing_header(sheet):
    """Return (header_row, exclude_cols).

    A packing header is a row containing both "item" and "quantity" (the
    particulars/description label may be missing or named "Description").
    Columns labelled Part No / Serial No / Dwg.* are metadata and excluded
    from name extraction.
    """
    for cells in sheet.iter_rows():
        texts = [_clean(c.value) for c in cells if c.value is not None]
        joined = " ".join(texts).lower()
        if "item" not in joined or "quantity" not in joined:
            continue
        exclude = set()
        for c in cells:
            t = _clean(c.value).lower()
            if not t:
                continue
            if any(k in t for k in ("item", "quantity", "particulars", "description")):
                continue  # merged main header cell; not a metadata column
            if "part no" in t or "serial no" in t or "dwg" in t:
                exclude.add(c.column)
        return cells[0].row, exclude
    return None, set()


def _row_parts(cells, exclude_cols=()):
    """Extract (qty, unit, name_parts, type_parts) from one row of cells."""
    qty, unit = None, None
    name_parts, type_parts = [], []
    for cell in cells:
        if cell.column in exclude_cols:
            continue
        text = _clean(cell.value)
        if not text:
            continue
        base, type_part = _split_type(text)
        if type_part:
            type_parts.append(type_part)
        if not base:
            continue
        m = _QTY_RE.search(base)
        if m:
            if qty is None:
                qty = float(m.group(1))
                unit = m.group(2)
            rest = base[: m.start()] + base[m.end():]
            rest = _clean_name(rest)
            if rest and not _is_irrelevant(rest):
                name_parts.append(rest)
            continue
        if _is_irrelevant(base):
            continue
        name_parts.append(_clean_name(base))
    return qty, unit, name_parts, type_parts


def _parse_packing_sheet(sheet, path, warn) -> list:
    header_row, exclude_cols = _find_packing_header(sheet)
    if header_row is None:
        if warn:
            warn(f"{path.name}/{sheet.name}: 未找到 Item/Quantity 表头，跳过该表")
        return []

    rows = list(sheet.iter_rows())
    end = sheet.last_row + 1
    for idx in range(header_row, sheet.last_row + 1):
        text = " ".join(_clean(c.value) for c in rows[idx - 1] if c.value is not None)
        if _END_RE.search(text):
            end = idx
            break

    items = []
    i = header_row
    while i < end:
        cells = rows[i - 1]
        qty, unit, name_parts, type_parts = _row_parts(cells, exclude_cols)
        if qty is None:
            i += 1
            continue

        # look ahead: continuation / Type lines until the next qty row or end
        j = i + 1
        while j < end:
            nqty, _, nname, ntype = _row_parts(rows[j - 1], exclude_cols)
            if nqty is not None:
                break
            text = " ".join(_clean(c.value) for c in rows[j - 1] if c.value is not None)
            if _END_RE.search(text):
                break
            name_parts.extend(nname)
            type_parts.extend(ntype)
            j += 1

        name = " ".join(dict.fromkeys(p for p in name_parts if p)).strip(" -–—")
        type_str = "; ".join(dict.fromkeys(t for t in type_parts if t)).strip(" -–—")
        name = _trim_footer(name)
        type_str = _trim_footer(type_str)
        if not name and type_str and len(type_str) <= 60:
            name, type_str = type_str, ""
        if not name:
            if warn:
                warn(f"{path.name}/{sheet.name} 第 {i} 行备件缺少名称（数量 {qty:g} {unit or '个'}）")
            name = "(未填写名称)"

        items.append(Part(name=name, qty=qty, unit=unit or "个", type=type_str or None))
        i = j
    return items


# ---- number / text helpers --------------------------------------------

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


# ---- public API -------------------------------------------------------

def parse_source(path, warn=None) -> list:
    """Parse a source Excel file (all sheets) into a list of Part objects.

    Optional callable warn(msg) receives messages about skipped/unparseable rows.
    """
    path = Path(path)
    parts = []
    for sheet in _open_sheets(path):
        header_row, mapping = _find_header_row(sheet)
        if header_row is not None:
            parts.extend(_parse_standard(sheet, mapping, header_row, path, warn))
            continue
        parts.extend(_parse_packing_sheet(sheet, path, warn))

    if not parts:
        raise ValueError(f"在 {path.name} 中没有解析到备件数据")
    return parts


def parse_sources(paths, warn=None) -> list:
    """Parse one or more source files, merging all parts in order."""
    all_parts = []
    for p in paths:
        all_parts.extend(parse_source(p, warn=warn))
    return all_parts


def detect_vessel(paths) -> Optional[str]:
    """Try to guess the vessel name from the source files (船名 / 船名/单位 / Vessel Name)."""
    for p in paths:
        p = Path(p)
        try:
            sheets = _open_sheets(p)
        except Exception:
            continue
        for sheet in sheets:
            for row in sheet.iter_rows():
                texts = [str(c.value).strip() for c in row if c.value is not None]
                joined = " ".join(texts)

                m = re.search(r"船名\s*[:：]\s*(\S+)", joined)
                if m:
                    return m.group(1)

                for i, t in enumerate(texts):
                    if "船名/单位" in t or "船名／单位" in t:
                        for other in texts[i + 1:]:
                            if other:
                                return other
                        break

                for t in texts:
                    if "vessel name" not in t.lower():
                        continue
                    m = re.search(r"[Vv]essel\s+[Nn]ame\s*[:：]?\s*(.*)", t)
                    val = (m.group(1) if m else "").strip()
                    if not val:
                        # label cell only; the value sits in a sibling cell on this row
                        for other in texts:
                            if other != t and other.lstrip(":"):
                                candidate = other.lstrip(":")
                                pm = re.search(r"\(([^)]*)\)", candidate)
                                if pm and re.search(r"[\u4e00-\u9fff]", pm.group(1)):
                                    return pm.group(1).strip()
                                clean = candidate.split("(")[0].strip()
                                if clean:
                                    return clean
                        continue
                    pm = re.search(r"\(([^)]*)\)", val)
                    if pm and re.search(r"[\u4e00-\u9fff]", pm.group(1)):
                        return pm.group(1).strip()
                    clean = val.split("(")[0].strip()
                    if clean:
                        return clean
    return None