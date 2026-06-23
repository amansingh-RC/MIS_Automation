"""
Streamlit web app for Royal Chain MIS report automation.

Run with:
    streamlit run app.py

Tools:
  1. Loss Report   -> upload raw export, get the formatted Loss Report sheet.
  2. Scrap + Stock -> upload one daily export, get one workbook with two sheets.

Adding a new tool later: write a render_<name>_tab() function and add one line
to the TABS list at the bottom.
"""

from __future__ import annotations

from datetime import date

import pandas as pd
import streamlit as st

import ui
from loss_report import LossReportError, process
from rcl_reports import process_scrap_and_stock

st.set_page_config(
    page_title="RCL MIS Automation",
    page_icon="https://royalchaingroup.com/wp-content/uploads/2025/09/favicon.png",
    layout="wide")
ui.inject_css()
ui.render_header()

XLSX_MIME = ("application/vnd.openxmlformats-officedocument"
             ".spreadsheetml.sheet")


# ===========================================================================
# Tab — Loss Report
# ===========================================================================
def render_loss_tab() -> None:
    ui.section_title(
        "Loss Report",
        "Upload the raw export → get the formatted Loss Report Summary sheet "
        "with Final Loss and Loss %.")

    with st.expander("How the new columns are calculated"):
        st.markdown(
            "- **Final Loss** = Loss Quantity Pg + Gain Pg\n"
            "- **Loss %** = Final Loss ÷ (Process + Unutilized + "
            "Unutilized Scrap + Unutilized Sample)")

    files = st.file_uploader(
        "Drop the Loss Report here  ·  .xls / .xlsx",
        type=["xlsx", "xlsm", "xls"], accept_multiple_files=True,
        key="loss_upl")

    if not files:
        st.info("Upload one or more files to begin.")
        return

    for file in files:
        st.divider()
        st.subheader(f"📄 {file.name}")
        try:
            out_bytes, report = process(file.getvalue())
        except LossReportError as exc:
            st.error(f"Could not process this file: {exc}")
            continue
        except Exception as exc:  # noqa: BLE001
            st.error(f"Unexpected error reading the file: {exc}")
            continue

        note = ""
        if report.date_from or report.date_to:
            note = f"  ·  Dates: {report.date_from} → {report.date_to}"
        st.success(f"Processed {len(report.rows)} work-center rows.{note}")

        preview = pd.DataFrame(report.rows)
        if "Loss %" in preview:
            preview["Loss %"] = preview["Loss %"].apply(
                lambda v: f"{v:.3%}" if v is not None else "")
        st.dataframe(preview, use_container_width=True, hide_index=True)

        out_name = file.name.rsplit(".", 1)[0] + "_processed.xlsx"
        st.download_button(
            f"⬇  Download {out_name}", data=out_bytes, file_name=out_name,
            mime=XLSX_MIME, key="loss_dl_" + file.name)


# ===========================================================================
# Tab — Scrap + Stock (one input -> one workbook with two sheets)
# ===========================================================================
def render_scrap_stock_tab() -> None:
    ui.section_title(
        "Scrap + Stock Report",
        "Upload one daily export → both reports run → one workbook with two "
        "sheets (SCRAP REPORT and STOCK REPORT).")

    with st.expander("What each report does"):
        st.markdown(
            "**SCRAP REPORT** — keeps rows where *Stock Status* is `SCRAP` or "
            "`HL-SCRAP`, grouped by WC Group + WC Name (Gross / Metal summed)."
            "\n\n**STOCK REPORT** — fills blank WC from Party Name; removes "
            "ALLOY, Beads, Color Stone, CZ, Synthetic Stone, Pearl; keeps only "
            "`OM` under OTHER METAL; skips Total rows; merges duplicates.")

    file = st.file_uploader(
        "Drop the daily report here  ·  .xls / .xlsx",
        type=["xlsx", "xlsm", "xls"], accept_multiple_files=False,
        key="ss_upl")

    if file is None:
        st.info("Upload one file to run both reports.")
        return

    st.divider()
    st.subheader(f"📄 {file.name}")
    try:
        out_bytes, summary = process_scrap_and_stock(
            file.getvalue(), date.today().isoformat())
    except LossReportError as exc:
        st.error(f"Could not process this file: {exc}")
        return
    except Exception as exc:  # noqa: BLE001
        st.error(f"Unexpected error reading the file: {exc}")
        return

    st.markdown("##### ♻️ Scrap Report")
    a, b, c = st.columns(3)
    a.metric("Scrap rows", f"{summary['scrap_rows']:,}")
    b.metric("Gross Weight", f"{summary['scrap_gross']:,.3f}")
    c.metric("Metal Weight", f"{summary['scrap_metal']:,.3f}")

    st.markdown("##### 📦 Stock Report")
    d, e, f = st.columns(3)
    d.metric("Rows after filtering", f"{summary['stock_filtered']:,}")
    e.metric("WC Groups", f"{summary['stock_groups']:,}")
    f.metric("Gross Weight", f"{summary['stock_gross']:,.3f}")

    st.divider()
    out_name = "RCL_Scrap_Stock_" + date.today().isoformat() + ".xlsx"
    st.download_button(
        f"⬇  Download {out_name}  (2 sheets)", data=out_bytes,
        file_name=out_name, mime=XLSX_MIME, key="ss_dl")


# ===========================================================================
# Tab registry — add a (label, render_fn) tuple here to add a new tool.
# ===========================================================================
TABS = [
    ("📉  Loss Report", render_loss_tab),
    ("♻️  Scrap + Stock", render_scrap_stock_tab),
]

for tab, (_label, render_fn) in zip(st.tabs([t[0] for t in TABS]), TABS):
    with tab:
        render_fn()

ui.render_footer()
