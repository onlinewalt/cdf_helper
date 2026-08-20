"""Tests for the packing-list (English Receipt/Packing List) parser, incl. a
best-effort check against the real uploaded file when present."""
import os
import sys
import tempfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

from openpyxl import Workbook

from cdf_helper.parser import (
    parse_source,
    parse_sources,
    detect_vessel,
    detect_vessel_pair,
    load_vessel_names,
    bilingual_vessel,
)


def _write_workbook(path: Path):
    wb = Workbook()

    # P1: Sheet1-like, merged header in A, item/qty/name in separate columns + Type line
    ws = wb.active
    ws.title = "P1"
    ws["A1"] = "Item      Quantity(Unit)         Particulars"
    ws["A2"] = "1"
    ws["B2"] = "1  PCE"
    ws["C2"] = "消防泵出海阀(蝶阀)"
    ws["B3"] = "Type:5k250"
    ws["A4"] = "2"
    ws["B4"] = "2 PCE"
    ws["C4"] = "液压法兰式蝶阀"
    ws["B6"] = "** End of Listing **"

    # P2: Sheet3-like, header split across A/B, qty+name in same cell, Type line after
    ws = wb.create_sheet("P2")
    ws["A1"] = "Item"
    ws["B1"] = "Quantity(Unit)          Particulars"
    ws["A2"] = "1"
    ws["B2"] = "3  PC  时间继电器"
    ws["B3"] = "Type:OMRON H3CR-H8L AC220V 50/60HZ"
    ws["A4"] = "2"
    ws["B4"] = "3 PCE"
    ws["C4"] = "时间继电器"
    ws["B5"] = "Type:OMRON H3CR-A8E"
    ws["B7"] = "**End of Listing**"

    # P3: Sheet8-like, qty+name in one merged cell, Type line is the only real name source
    ws = wb.create_sheet("P3")
    ws["A1"] = "Item       Quantity(Unit)          Particulars"
    ws["A2"] = "  1              14   PCE                         plate"
    ws["B2"] = "51506-04H-160"
    ws["D3"] = "Type: 辅机滑油冷却器换热片"
    ws["A4"] = "**   End   of   Listing**"

    # P4: footer text glued onto a name must be trimmed
    ws = wb.create_sheet("P4")
    ws["A1"] = "Item Quantity(Unit) Particulars"
    ws["A2"] = "1"
    ws["B2"] = "4 PCE"
    ws["C2"] = "Sealing Ring 蓋章及簸收 httn://10.13.30.15:8080/x"
    ws["B3"] = "Type:07.02.01.063"
    ws["A5"] = "** End of Listing **"

    # NoHeader: no Item header -> skipped, but carries a Vessel Name
    ws = wb.create_sheet("NoHeader")
    ws["A1"] = "Vessel Name :VESSEL-A(测试船)"
    ws["A2"] = "1  2 PCE  某某"

    wb.save(path)


def test_synthetic():
    with tempfile.TemporaryDirectory() as td:
        path = Path(td) / "packing.xlsx"
        _write_workbook(path)

        warns = []
        parts = parse_source(path, warn=warns.append)

        # P1 2 items + P2 2 items + P3 1 item + P4 1 item; NoHeader skipped
        assert len(parts) == 6, [p.name for p in parts]
        assert any("未找到 Item/Quantity" in w for w in warns), warns

        p1, p2, p3, p4, p5, p6 = parts
        assert p1.name == "消防泵出海阀(蝶阀)" and p1.qty == 1 and p1.unit == "PCE"
        assert p1.type == "5k250", p1.type
        assert p2.name == "液压法兰式蝶阀" and p2.qty == 2
        assert p3.name == "时间继电器" and p3.qty == 3 and p3.unit == "PC"
        assert p3.type == "OMRON H3CR-H8L AC220V 50/60HZ", p3.type
        assert p4.name == "时间继电器" and p4.qty == 3 and p4.type == "OMRON H3CR-A8E"
        assert p5.qty == 14 and p5.unit == "PCE"
        assert p5.type == "辅机滑油冷却器换热片", p5.type
        assert p6.name == "Sealing Ring", p6.name  # footer trimmed
        assert p6.qty == 4 and p6.type == "07.02.01.063"

        assert detect_vessel([path]) == "测试船"
        assert detect_vessel_pair([path]) == ("测试船", "VESSEL-A")


def test_vessel_lookup():
    lookup = Path("中英文船名25-5-14.xls")
    if not lookup.exists():
        print("SKIP: 中英文船名 lookup not present")
        return
    zh2en, en2zh = load_vessel_names(lookup)
    assert len(zh2en) == len(en2zh) > 100
    assert zh2en["远怡湖"] == "COSMERRY LAKE"
    assert zh2en["中远安特卫普"] == "COSCO ANTWERP"
    assert en2zh["COSMERRYLAKE"] == "远怡湖"
    assert en2zh["COSCOANTWERP"] == "中远安特卫普"

    # English match handles traditional/simplified differences in the source
    assert bilingual_vessel("中遠安特偉普", zh2en, en2zh, english="coscO ANTWERP", chinese="中遠安特偉普") == "中远安特卫普 COSCO ANTWERP"
    assert bilingual_vessel("遠怡湖", zh2en, en2zh, english="cOSMERRY LAKE", chinese="遠怡湖") == "远怡湖 COSMERRY LAKE"
    # Chinese-only match through the lookup (typed forms)
    assert bilingual_vessel("远怡湖", zh2en, en2zh) == "远怡湖 COSMERRY LAKE"
    assert bilingual_vessel("遠怡湖", zh2en, en2zh) == "远怡湖 COSMERRY LAKE"
    # unresolvable -> unchanged
    assert bilingual_vessel("测试船", zh2en, en2zh) == "测试船"


def test_real_file_if_present():
    real = Path("uploads/中远安特伟普-大连26-8-20.xlsx")
    if not real.exists():
        print("SKIP: 中远安特伟普 file not present")
        return
    parts = parse_sources([real])
    assert len(parts) > 0
    assert all(p.name for p in parts), "every parsed part must have a name"
    assert any(p.name == "(未填写名称)" for p in parts), "Sheet1 unnamed items expected"
    vessel = detect_vessel([real])
    print(f"real file: {len(parts)} items, vessel={vessel}")

    # second real sample: 9 mixed-format sheets (Chinese 签收单 + Yuantong packing lists)
    real2 = Path("uploads/远怡湖-湛江26-8-20.xlsx")
    if not real2.exists():
        print("SKIP: 远怡湖 file not present")
        return
    parts2 = parse_sources([real2])
    assert len(parts2) == 29, [p.name for p in parts2]
    assert detect_vessel([real2]) == "遠怡湖"
    print(f"real file2: {len(parts2)} items, vessel=遠怡湖")


if __name__ == "__main__":
    test_synthetic()
    test_real_file_if_present()
    print("\nALL PARSER TESTS PASSED")