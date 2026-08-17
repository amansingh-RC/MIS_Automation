import io

import openpyxl

from factory_inward import process, karat_from_variant
from collections import OrderedDict


def test_variant_karat():
    assert karat_from_variant("G-NA-75.00-RG") == "18KT"
    assert karat_from_variant("G-NA-91.80-YG") == "22KT"
    assert karat_from_variant("PG-NA-99.50-YG") == "24KT"
    assert karat_from_variant("PG-NA-24KT-YG") == "24KT"     # literal fallback
    assert karat_from_variant("G-NA-80.00-YG") is None       # out of band
    assert karat_from_variant("G-NA-80.00-YG", other="OT") == "OT"
    assert karat_from_variant("NO-KARAT-HERE", other="OT") is None  # no value
    print("OK  karat from variant (decimal % + literal NNKT + OT)")


def make_input() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["B6"] = "All Transaction Summary"
    ws["B9"] = "From Date :- 24/06/2026\nTo Date :- 24/06/2026\n"
    hdr = {2: "Trans Date", 3: "Party Name", 4: "Style No", 5: "Variant Name",
           6: "Net Wt", 8: "Pg Wt", 10: "Line Remark"}
    for c, label in hdr.items():
        ws.cell(row=12, column=c, value=label)
    # party, variant, net, pg
    data = [
        ("Royal Chain Private Limited", "PG-NA-24KT-YG", 100.0, 95.0),
        ("Royal Chain Private Limited", "G-NA-91.80-YG", 20.0, 18.0),
        # heavy enough to rank 5th, but refinishing is excluded from the TOP 5
        ("RC-REFINISHING", "G-NA-75.00-RG", 10.0, 11.0),
        ("RC-REFINISHING", "G-NA-80.00-YG", 3.0, 2.0),   # out of band -> OT
        ("UNKNOWN PARTY", "G-NA-91.80-YG", 5.0, 4.0),   # not in position table
        ("ARK CHAINS", "PG-NA-24KT-YG", 80.0, 70.0),
        ("TITAN", "G-NA-91.80-YG", 60.0, 55.0),
        ("Aalishaan Jewels LLP", "G-NA-75.00-RG", 50.0, 45.0),
        # big enough for 2nd place, but "reparing" parties never enter the TOP 5
        ("REPARING-ARK CHAINS", "G-NA-91.80-YG", 90.0, 85.0),
        # would rank 5th, but hallmarking is excluded too
        ("QUALTIY ASSEY & HALLMARKS LLP", "G-NA-91.80-YG", 40.0, 35.0),
        ("SHREE GANESH JEWELS", "PG-NA-24KT-YG", 12.0, 10.0),
    ]
    r = 13
    for party, var, net, pg in data:
        ws.cell(row=r, column=3, value=party)
        ws.cell(row=r, column=5, value=var)
        ws.cell(row=r, column=6, value=net)
        ws.cell(row=r, column=8, value=pg)
        r += 1
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def main():
    test_variant_karat()
    # explicit position table: Royal Chain=1, RC-REFINISHING=2
    positions = OrderedDict([("royal chain private limited", 1),
                             ("rc-refinishing", 2)])
    out_bytes, result = process(make_input(), positions=positions,
                                gen_date="25-06-2026")

    assert result.date_from == "24-06-2026", result.date_from
    assert result.skipped == 0, result.skipped        # OT row is kept, not skipped
    assert result.unlisted == [
        "Aalishaan Jewels LLP", "ARK CHAINS", "QUALTIY ASSEY & HALLMARKS LLP",
        "REPARING-ARK CHAINS", "SHREE GANESH JEWELS", "TITAN",
        "UNKNOWN PARTY"], result.unlisted
    # grand: net 138+80+60+50+90+40+12 = 470 ; pg 130+70+55+45+85+35+10 = 430
    assert abs(result.grand_net - 470.0) < 1e-9, result.grand_net
    assert abs(result.grand_pg - 430.0) < 1e-9, result.grand_pg
    print("OK  dates, skipped, unlisted, grand totals")

    wb = openpyxl.load_workbook(io.BytesIO(out_bytes))
    ws = wb.active
    assert ws.title == "Factory Inward"
    assert ws["A1"].value == "Factory Inward ( GOODS RECEIPT NOTE ) Report Summary"
    assert ws["A2"].value == "Date From - 24-06-2026 To 24-06-2026"
    assert ws["D2"].value == "Date: 25-06-2026"
    assert [ws.cell(row=3, column=c).value for c in range(1, 6)] == \
        ["Sr. No.", "Party Name", "KARAT", "Net Wt", "Pg Wt"]

    # Royal Chain (Sr.No 1) first, karat order 24 then 22
    assert ws["A4"].value == 1 and ws["B4"].value == "Royal Chain Private Limited"
    assert ws["C4"].value == "24KT" and ws["C5"].value == "22KT"
    assert ws["B6"].value == "Royal Chain Private Limited Total"
    assert abs(ws["D6"].value - 120.0) < 1e-9
    # RC-REFINISHING Sr.No 2
    assert ws["A7"].value == 2 and ws["B7"].value == "RC-REFINISHING"
    # unlisted parties follow, alphabetically; UNKNOWN PARTY is last of 9
    rows = {ws.cell(row=r, column=2).value: ws.cell(row=r, column=1).value
            for r in range(4, ws.max_row + 1)}
    assert rows.get("Aalishaan Jewels LLP") == 3, rows
    assert rows.get("UNKNOWN PARTY") == 9, rows
    print("OK  position order, Sr. No., karat order, totals")

    # Grand total + karat summary block present
    flat = [tuple(ws.cell(row=r, column=c).value for c in range(1, 6))
            for r in range(4, ws.max_row + 1)]
    assert any(t[1] == "Grand Total" for t in flat)
    assert any(t[2] == "KARAT" for t in flat)         # block header
    # out-of-band karat kept as "OT"
    assert "OT" in {t[2] for t in flat if t[2]}, flat
    print("OK  grand total + karat summary block + OT")

    # --- TOP 5 block on the right (G:I), ranked by Pg Wt --------------------
    assert ws["G1"].value == "TOP 5"
    assert ws["G1"].fill.fgColor.rgb == "00FFFF00", ws["G1"].fill.fgColor.rgb
    assert [ws.cell(row=3, column=c).value for c in (7, 8, 9)] == \
        ["Party Name", "Net Wt", "Pg Wt"]
    top = [ws.cell(row=r, column=7).value for r in range(4, 9)]
    assert top == ["Royal Chain Private Limited", "ARK CHAINS", "TITAN",
                   "Aalishaan Jewels LLP", "SHREE GANESH JEWELS"], top
    # reparing / refinishing / hallmarking parties are skipped even though
    # each of them outranks SHREE GANESH JEWELS on Pg Wt
    block = {ws.cell(row=r, column=7).value for r in range(4, 20)}
    assert "REPARING-ARK CHAINS" not in block, block
    assert "RC-REFINISHING" not in block, block
    assert "QUALTIY ASSEY & HALLMARKS LLP" not in block, block
    # the block's own total covers the five listed parties only
    assert ws["G9"].value == "Grand Total"
    assert abs(ws["H9"].value - 322.0) < 1e-9, ws["H9"].value
    assert abs(ws["I9"].value - 293.0) < 1e-9, ws["I9"].value
    # karat summary below it still covers everything, excluded parties included
    assert [ws.cell(row=11, column=c).value for c in (7, 8, 9)] == \
        ["KARAT", "Net Wt", "Pg Wt"]
    karats = {ws.cell(row=r, column=7).value: ws.cell(row=r, column=8).value
              for r in range(12, 16)}
    assert list(karats) == ["24KT", "22KT", "18KT", "OT"], list(karats)
    assert abs(karats["22KT"] - 215.0) < 1e-9, karats      # includes REPARING
    assert ws["G16"].value == "Grand Total"
    assert abs(ws["H16"].value - 470.0) < 1e-9, ws["H16"].value
    assert abs(ws["I16"].value - 430.0) < 1e-9, ws["I16"].value
    print("OK  TOP 5 block: Pg-Wt order, reparing/refinishing/hallmark "
          "excluded, totals")

    # pivot only — no detail sheet
    assert wb.sheetnames == ["Factory Inward"], wb.sheetnames
    print("OK  single pivot sheet (no detail sheet)")
    print("\nALL TESTS PASSED")


if __name__ == "__main__":
    main()
