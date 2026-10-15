import io

import openpyxl

from groupsales import process, karat_from_melting


def test_karat_bands():
    assert karat_from_melting(99.99) == "24KT"
    assert karat_from_melting(91.8) == "22KT"
    assert karat_from_melting(75.0) == "18KT"
    assert karat_from_melting(58.5) == "14KT"
    assert karat_from_melting(37.5) == "9KT"
    assert karat_from_melting(37.0) == "9KT" and karat_from_melting(38.0) == "9KT"
    assert karat_from_melting(36.9) is None and karat_from_melting(38.1) is None
    assert karat_from_melting(80.0) is None
    print("OK  karat bands")


def make_input() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["B6"] = "All Transaction Summary"
    ws["B9"] = "From Date :- 24/06/2026\nTo Date :- 24/06/2026\n"
    # header at row 12 (cols B..), with the columns we need
    hdr = {2: "Trans Date", 3: "Party Name", 11: "Metal Fineness",
           14: "Variant Name", 15: "Net Wt", 16: "Pg Wt", 23: "Groupsales"}
    for c, label in hdr.items():
        ws.cell(row=12, column=c, value=label)
    # group, party, fineness, variant, net, pg
    data = [
        ("CUBAN", "ROYAL CHAIN", 0.918, "", 100.0, 90.0),   # 22KT via fineness
        ("CUBAN", "ARK CHAINS", 0.75, "", 20.0, 15.0),      # 18KT
        ("ALPHA", "TITAN", 0.75, "", 10.0, 8.0),            # 18KT
        ("ALPHA", "TITAN", 0.80, "", 5.0, 4.0),   # out of band (fineness) -> OT
        ("ZETA", "ARK CHAINS", None, "PG-NA-24KT-YG", 30.0, 28.0),  # 24KT
        ("", "AALISHAAN", 0.999, "", 40.0, 38.0),  # blank group -> "(blank)"
        # Variant Name decides karat: 99.00 -> 24KT even though fineness=0.75(18KT)
        ("GAMMA", "NOVEL JEWELS", 0.75, "G-NA-99.00-YG", 10.0, 9.0),
        ("ZETA", "TITAN", None, "G-NA-80.00-YG", 6.0, 5.0),  # out of band -> OT
        ("DELTA", "SKIPME", None, "", 7.0, 6.0),   # no karat value -> skipped
        ("EPS", "TINY CO", 0.918, "", 2.0, 1.0),   # 6th party, below the TOP 5
    ]
    r = 13
    for group, party, fin, var, net, pg in data:
        ws.cell(row=r, column=23, value=group)
        ws.cell(row=r, column=3, value=party)
        if fin is not None:
            ws.cell(row=r, column=11, value=fin)
        if var:
            ws.cell(row=r, column=14, value=var)
        ws.cell(row=r, column=15, value=net)
        ws.cell(row=r, column=16, value=pg)
        r += 1
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_top5_without_party_column():
    """No Party Name column -> the TOP 5 ranks the Groupsales groups instead."""
    wb = openpyxl.load_workbook(io.BytesIO(make_input()))
    wb.active.cell(row=12, column=3).value = None       # drop the header
    buf = io.BytesIO()
    wb.save(buf)
    out_bytes, result = process(buf.getvalue())
    assert result.top_label == "Groupsales", result.top_label
    ws = openpyxl.load_workbook(io.BytesIO(out_bytes)).active
    assert ws.cell(row=3, column=6).value == "Groupsales"
    # group totals ranked by Pg Wt: CUBAN 105 > (blank) 38 > ZETA 33 > ALPHA 12
    assert [ws.cell(row=r, column=6).value for r in range(4, 9)] == \
        ["CUBAN", "(blank)", "ZETA", "ALPHA", "GAMMA"]
    print("OK  TOP 5 falls back to Groupsales when Party Name is absent")


def test_9kt_band():
    """A 9KT row (melting 37–38) gets its own line, first in ascending order,
    and counts towards the totals."""
    wb = openpyxl.load_workbook(io.BytesIO(make_input()))
    ws = wb.active
    r = ws.max_row + 1
    ws.cell(row=r, column=23, value="CUBAN")
    ws.cell(row=r, column=3, value="ROYAL CHAIN")
    ws.cell(row=r, column=14, value="G-NA-37.50-YG")
    ws.cell(row=r, column=15, value=4.0)
    ws.cell(row=r, column=16, value=3.0)
    buf = io.BytesIO()
    wb.save(buf)
    out_bytes, result = process(buf.getvalue())
    assert abs(result.grand_net - 227.0) < 1e-9, result.grand_net   # 223 + 4
    ws = openpyxl.load_workbook(io.BytesIO(out_bytes)).active
    assert [ws.cell(row=r, column=2).value for r in (4, 5, 6)] == \
        ["9KT", "18KT", "22KT"]
    assert ws["A7"].value == "CUBAN Total"
    assert abs(ws["C7"].value - 124.0) < 1e-9, ws["C7"].value
    # karat summary under the TOP 5 block starts with 9KT
    assert ws["F12"].value == "9KT", ws["F12"].value
    print("OK  9KT band: own line, ascending order, counted in totals")


def main():
    test_karat_bands()
    out_bytes, result = process(make_input())

    assert result.date == "24.06.2026", result.date
    assert result.skipped == 1, result.skipped        # DELTA: no karat value
    assert result.other == 2, result.other            # the 0.80 and 80.00 rows
    # grand: net 100+20+10+5+30+40+10+6+2 = 223 ; pg 90+15+8+4+28+38+9+5+1 = 198
    assert abs(result.grand_net - 223.0) < 1e-9, result.grand_net
    assert abs(result.grand_pg - 198.0) < 1e-9, result.grand_pg
    print("OK  date, skipped, other, grand totals")

    wb = openpyxl.load_workbook(io.BytesIO(out_bytes))
    ws = wb.active
    assert ws.title == "Groupsales Reports"
    assert ws["A1"].value == "Groupsales Reports"
    assert ws["A2"].value == "Date: 24.06.2026"
    assert [ws.cell(row=3, column=c).value for c in range(1, 5)] == \
        ["Groupsales", "Karat", "Net Wt", "Pg Wt"]

    # CUBAN first (data / first-appearance order), karat ascending: 18KT then 22KT
    assert ws["A4"].value == "CUBAN"
    assert ws["B4"].value == "18KT" and abs(ws["C4"].value - 20.0) < 1e-9
    assert ws["B5"].value == "22KT" and abs(ws["C5"].value - 100.0) < 1e-9
    assert ws["A6"].value == "CUBAN Total"
    assert abs(ws["C6"].value - 120.0) < 1e-9

    # ZETA (24KT) came from the Variant Name even though Metal Fineness is blank
    zeta = next(r for r in range(4, ws.max_row + 1)
                if ws.cell(row=r, column=1).value == "ZETA")
    assert ws.cell(row=zeta, column=2).value == "24KT"
    assert abs(ws.cell(row=zeta, column=3).value - 30.0) < 1e-9

    # blank-Groupsales rows are kept under "(blank)"
    blank = next(r for r in range(4, ws.max_row + 1)
                 if ws.cell(row=r, column=1).value == "(blank)")
    assert ws.cell(row=blank, column=2).value == "24KT"
    assert abs(ws.cell(row=blank, column=3).value - 40.0) < 1e-9

    # Variant Name decides karat: GAMMA's "G-NA-99.00-YG" -> 24KT (NOT 18KT
    # from its Metal Fineness 0.75)
    gamma = next(r for r in range(4, ws.max_row + 1)
                 if ws.cell(row=r, column=1).value == "GAMMA")
    assert ws.cell(row=gamma, column=2).value == "24KT", \
        ws.cell(row=gamma, column=2).value

    # out-of-band rows land in an "OT" row, last within the group
    alpha = next(r for r in range(4, ws.max_row + 1)
                 if ws.cell(row=r, column=1).value == "ALPHA")
    assert ws.cell(row=alpha, column=2).value == "18KT"
    assert ws.cell(row=alpha + 1, column=2).value == "OT"
    assert abs(ws.cell(row=alpha + 1, column=3).value - 5.0) < 1e-9
    assert ws.cell(row=alpha + 2, column=1).value == "ALPHA Total"
    assert abs(ws.cell(row=alpha + 2, column=3).value - 15.0) < 1e-9

    # an out-of-band melting in the Variant Name is OT too (ZETA: 24KT + OT)
    assert ws.cell(row=zeta + 1, column=2).value == "OT"
    assert abs(ws.cell(row=zeta + 1, column=3).value - 6.0) < 1e-9

    # a row with no karat value at all is still dropped
    assert not any(ws.cell(row=r, column=1).value == "DELTA"
                   for r in range(4, ws.max_row + 1))

    # last row of the pivot (column A) is the grand total
    gt = max(r for r in range(4, ws.max_row + 1)
             if ws.cell(row=r, column=1).value == "Grand Total")
    assert abs(ws.cell(row=gt, column=3).value - 223.0) < 1e-9
    print("OK  layout, order, ascending karat, OT band, (blank) group, "
          "variant-decides")

    # --- TOP 5 block on the right (F:H), ranked by Pg Wt --------------------
    assert result.top_label == "Party Name", result.top_label
    assert ws["F1"].value == "TOP 5"
    assert ws["F1"].fill.fgColor.rgb == "00FFFF00", ws["F1"].fill.fgColor.rgb
    assert [ws.cell(row=3, column=c).value for c in (6, 7, 8)] == \
        ["Party Name", "Net Wt", "Pg Wt"]
    # ROYAL 90 > ARK 43 > AALISHAAN 38 > TITAN 17 > NOVEL 9 ; TINY CO (1) is out
    top = [(ws.cell(row=r, column=6).value, ws.cell(row=r, column=7).value,
            ws.cell(row=r, column=8).value) for r in range(4, 9)]
    assert [t[0] for t in top] == ["ROYAL CHAIN", "ARK CHAINS", "AALISHAAN",
                                   "TITAN", "NOVEL JEWELS"], top
    assert abs(top[1][1] - 50.0) < 1e-9 and abs(top[1][2] - 43.0) < 1e-9, top[1]
    assert abs(top[3][1] - 21.0) < 1e-9 and abs(top[3][2] - 17.0) < 1e-9, top[3]
    assert not any(ws.cell(row=r, column=6).value == "TINY CO"
                   for r in range(4, ws.max_row + 1))
    assert not any(ws.cell(row=r, column=6).value == "SKIPME"
                   for r in range(4, ws.max_row + 1))
    # TOP 5 total = the five shown only (221/197), not the grand total (223/198)
    assert ws["F9"].value == "Grand Total"
    assert abs(ws["G9"].value - 221.0) < 1e-9, ws["G9"].value
    assert abs(ws["H9"].value - 197.0) < 1e-9, ws["H9"].value
    print("OK  TOP 5 block: title, Pg-Wt order, cut at 5, own grand total")

    # --- karat summary under the TOP 5 block --------------------------------
    assert [ws.cell(row=11, column=c).value for c in (6, 7, 8)] == \
        ["KARAT", "Net Wt", "Pg Wt"]
    karats = {ws.cell(row=r, column=6).value:
              (ws.cell(row=r, column=7).value, ws.cell(row=r, column=8).value)
              for r in range(12, 16)}
    assert list(karats) == ["18KT", "22KT", "24KT", "OT"], list(karats)
    assert abs(karats["18KT"][0] - 30.0) < 1e-9, karats["18KT"]
    assert abs(karats["22KT"][0] - 102.0) < 1e-9, karats["22KT"]
    assert abs(karats["24KT"][0] - 80.0) < 1e-9, karats["24KT"]
    assert abs(karats["OT"][0] - 11.0) < 1e-9, karats["OT"]
    assert ws["F16"].value == "Grand Total"
    assert abs(ws["G16"].value - 223.0) < 1e-9, ws["G16"].value
    assert abs(ws["H16"].value - 198.0) < 1e-9, ws["H16"].value
    print("OK  karat summary block (all data) + grand total")

    # pivot only — no detail sheet
    assert wb.sheetnames == ["Groupsales Reports"], wb.sheetnames
    print("OK  single pivot sheet (no detail sheet)")

    test_top5_without_party_column()
    test_9kt_band()
    print("\nALL TESTS PASSED")


if __name__ == "__main__":
    main()
