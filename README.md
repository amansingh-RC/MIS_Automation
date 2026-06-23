# Loss Report Automation

A web app that takes the raw **Monthly Loss Report Summary** export
(`.xls` or `.xlsx`) and returns the clean **Loss Report Summary** sheet — the
formatted single sheet you build by hand every day.

## What it produces

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
| `app.py` | Streamlit web interface (upload / preview / download) |
| `loss_report.py` | Core logic: read, detect columns, build the formatted output |
| `test_loss_report.py` | Self-test of the formulas + template (`python test_loss_report.py`) |
| `requirements.txt` | Python dependencies |

## Notes

- The output is always a modern `.xlsx`, even if you upload an old `.xls`.
- FINAL LOSS / LOSS % are real formulas, so they recalculate if you edit a value
  in Excel.
- Only the Loss Report sheet is produced; other sub-reports are ignored.
