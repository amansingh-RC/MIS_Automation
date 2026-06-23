# RCL MIS Report Automation

A web app with two tools (tabs):

1. **Loss Report** — raw export → formatted **Loss Report Summary** sheet.
2. **Scrap + Stock Report** — one daily export → one workbook with **two
   sheets** (`SCRAP REPORT` and `STOCK REPORT`).

---

## Tool 1 — Loss Report

Takes the raw **Monthly Loss Report Summary** export (`.xls` or `.xlsx`) and
returns the clean **Loss Report Summary** sheet — the formatted single sheet
you build by hand every day.

### What it produces

A new `.xlsx` with one sheet (`Loss Report Summary`) that matches your target:

- Title row + a date band (**dates read automatically** from the export's
  `From Date :- / To Date :-` preamble).
- Columns reordered to the report layout (**Scrap before Sample**).
- The totals row (blank Wc Name) removed.
- A yellow spacer column, then two new columns:

| New column | Formula (written as a live Excel formula) |
|------------|-------------------------------------------|
| **FINAL LOSS** | `Loss Quantity Pg + Gain Pg`  → `=G{row}+H{row}` |
| **LOSS %** | `FINAL LOSS / (Process + Unutilized + Scrap + Sample)`  → `=IFERROR(K{row}/(C{row}+D{row}+E{row}+F{row}),"")` |

Columns are detected **by header name**, so it tolerates minor naming/order
changes. Gains stored as `(4.89)` or as negative numbers are handled correctly.

---

## Tool 2 — Scrap + Stock Report

Upload **one** daily export; both reports run and you download **one workbook
with two sheets**.

**SCRAP REPORT** (sheet 1): finds the header row with `Stock Status`, keeps
rows where status is `SCRAP` or `HL-SCRAP`, groups by *WC Group + WC Name*, and
sums Gross / Metal weight (group totals + grand total).

**STOCK REPORT** (sheet 2): finds the header row with `Wcgroup Name`, then:
- fills blank *WC Group* / *WC Name* from *Party Name*;
- removes Item Groups: ALLOY, Beads (gms), Color Stone, CZ, Synthetic Stone,
  Pearl;
- under *OTHER METAL* keeps only the `OM` variant;
- skips Total / Grand Total / Subtotal rows;
- merges duplicate *WC Group + WC Name* (weights summed once).

These are faithful Python ports of `RCL_SCRAP_REPORT_TOOL.html` and
`RCL_STOCK_REPORT_DAILY_TOOL.html`.

---

## Setup (one time)

```powershell
cd C:\Users\ShaGai\Desktop\MIS_Automation
pip install -r requirements.txt
```

## Run

```powershell
streamlit run app.py
```

The browser opens automatically. Then:

1. Drag-drop one or more raw report files (`.xls` or `.xlsx`).
2. Check the preview (and the dates it detected).
3. Click **Download** to get the formatted `Loss Report Summary` sheet.

## Files

| File | Purpose |
|------|---------|
| `app.py` | Streamlit web interface (two tabs) |
| `loss_report.py` | Loss Report logic: read, detect columns, build formatted output |
| `rcl_reports.py` | Scrap + Stock logic and combined 2-sheet workbook |
| `test_loss_report.py` | Self-test of the Loss Report (`python test_loss_report.py`) |
| `test_rcl_reports.py` | Self-test of Scrap + Stock (`python test_rcl_reports.py`) |
| `requirements.txt` | Python dependencies |

## Notes

- The output is always a modern `.xlsx`, even if you upload an old `.xls`.
- FINAL LOSS / LOSS % are real formulas, so they recalculate if you edit a value
  in Excel.
- Only the Loss Report sheet is produced; other sub-reports are ignored.
