from __future__ import annotations

import io
from datetime import date, datetime, timedelta

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter

_EXCEL_EPOCH = datetime(1899, 12, 30)


def new_workbook() -> openpyxl.Workbook:
    """A fresh workbook. Its default empty 'Sheet' is removed by workbook_bytes."""
    return openpyxl.Workbook()


def workbook_bytes(wb: openpyxl.Workbook) -> bytes:
    """Drop the leftover empty default sheet, then return the .xlsx bytes."""
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        default = wb["Sheet"]
        empty = (default.max_row <= 1 and default.max_column <= 1
                 and default["A1"].value is None)
        if empty:
            del wb["Sheet"]
    if wb.sheetnames:
        wb.active = 0
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# --- shared "TOP 5" block ---------------------------------------------------
# A right-hand section (title bar + top parties by Pg Wt + a karat summary)
# used by Groupsales and Factory Inward/Outward.
TOP_N = 5
_TOP_THIN = Side(style="thin", color="000000")
_TOP_BORDER = Border(left=_TOP_THIN, right=_TOP_THIN,
                     top=_TOP_THIN, bottom=_TOP_THIN)
_TOP_TITLE_FILL = PatternFill("solid", fgColor="FFFF00")
_TOP_HEAD_FILL = PatternFill("solid", fgColor=Color(theme=4,
                                                    tint=0.7999816888943144))
_TOP_NUM_FMT = "#,##0.000"
_TOP_FONT = "Cambria"
_TOP_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
# Rework / service accounts — real work, but not sales, so they never take a
# TOP 5 slot. Matched as substrings ("reparing" is how the export spells it;
# "refinish" covers RC-REFINISHING, AJ-REFINISHING, …).
_TOP_EXCLUDE = ("repar", "repair", "refinish", "hallmark")


def is_top_excluded(name) -> bool:
    """True for rework / hallmarking accounts that must skip the TOP 5."""
    lowered = str(name).casefold()
    return any(word in lowered for word in _TOP_EXCLUDE)


def rank_top(totals: dict, n: int = TOP_N, exclude: bool = False) -> list:
    """``{name: [net, pg]}`` -> the ``n`` biggest by Pg Wt (desc).

    Ties break on Net Wt, then name. With ``exclude`` the rework / hallmarking
    accounts are dropped before ranking (their weights still count towards the
    report's own grand total — only the ranking ignores them).
    """
    items = [kv for kv in totals.items()
             if not (exclude and is_top_excluded(kv[0]))]
    items.sort(key=lambda kv: (-kv[1][1], -kv[1][0], str(kv[0]).casefold()))
    return items[:n]


def _top_row(ws, row: int, col: int, label, net, pg, *, bold: bool, fill=None):
    """One label + Net Wt + Pg Wt row of the block, fully bordered."""
    for i, value in enumerate((label, net, pg)):
        cell = ws.cell(row=row, column=col + i, value=value)
        cell.font = Font(name=_TOP_FONT, bold=bold, size=10)
        cell.alignment = _TOP_CENTER
        cell.border = _TOP_BORDER
        if fill is not None:
            cell.fill = fill
        if i and isinstance(value, (int, float)):
            cell.number_format = _TOP_NUM_FMT


def add_top_block(ws, col: int, label: str, ranked: list, karat_totals: dict,
                  karat_order: list, grand_net: float, grand_pg: float,
                  widths=(34, 15, 15), header_row: int = 3) -> list[dict]:
    """Write the TOP 5 block in the three columns starting at ``col``.

    ``ranked`` is the output of :func:`rank_top`. Its Grand Total covers only
    the listed names; the karat summary below it covers the whole report.
    Returns the block's rows for the app preview.
    """
    first, last = get_column_letter(col), get_column_letter(col + 2)
    ws.merge_cells(f"{first}1:{last}1")
    ws[f"{first}1"].value = f"TOP {TOP_N}"
    ws[f"{first}1"].font = Font(name=_TOP_FONT, bold=True, size=14)
    ws[f"{first}1"].alignment = _TOP_CENTER
    for c in range(col, col + 3):
        ws.cell(row=1, column=c).fill = _TOP_TITLE_FILL
        ws.cell(row=1, column=c).border = _TOP_BORDER

    rows: list[dict] = []
    _top_row(ws, header_row, col, label, "Net Wt", "Pg Wt",
             bold=True, fill=_TOP_HEAD_FILL)
    r = header_row + 1
    t_net = t_pg = 0.0
    for name, (net, pg) in ranked:
        _top_row(ws, r, col, name, net, pg, bold=False)
        rows.append({label: name, "Net Wt": round(net, 3),
                     "Pg Wt": round(pg, 3)})
        t_net += net
        t_pg += pg
        r += 1
    _top_row(ws, r, col, "Grand Total", t_net, t_pg,
             bold=True, fill=_TOP_HEAD_FILL)
    rows.append({label: "Grand Total", "Net Wt": round(t_net, 3),
                 "Pg Wt": round(t_pg, 3)})

    # Karat summary — the whole report, not just the listed names.
    r += 2
    _top_row(ws, r, col, "KARAT", "Net Wt", "Pg Wt",
             bold=True, fill=_TOP_HEAD_FILL)
    r += 1
    for karat in [k for k in karat_order if k in karat_totals]:
        k_net, k_pg = karat_totals[karat]
        _top_row(ws, r, col, karat, k_net, k_pg, bold=False)
        r += 1
    _top_row(ws, r, col, "Grand Total", grand_net, grand_pg,
             bold=True, fill=_TOP_HEAD_FILL)

    for i, width in enumerate(widths):
        ws.column_dimensions[get_column_letter(col + i)].width = width
    return rows


def excel_serial_to_date(value: float) -> datetime:
    return _EXCEL_EPOCH + timedelta(days=float(value))


def fmt_date(value, pattern: str = "%d/%m/%Y") -> str:
    """Format a cell that may be a datetime, an Excel serial, or text."""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime(pattern)
    if isinstance(value, date):
        return value.strftime(pattern)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        # Plausible Excel date serials (skip tiny numbers that aren't dates).
        if value > 59:
            return excel_serial_to_date(value).strftime(pattern)
        return str(value)
    return str(value).strip()
