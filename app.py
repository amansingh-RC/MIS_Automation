"""
Streamlit web app for Royal Chain MIS report automation.

Run with:
    streamlit run app.py

Each tab is one report tool: upload its file, preview, download just that
report. The sidebar offers a single "Combined Workbook" download that merges
every uploaded tab's output sheets into one .xlsx.

Adding a new tool later: write an `add_<name>_sheet(wb, file_bytes)` function in
its own module, then add one entry to the TOOLS list below.
"""

from __future__ import annotations

from datetime import date

import pandas as pd
import streamlit as st

import ui
from loss_report import LossReportError, add_loss_sheet
from lot_rejection import add_lot_rejection_sheet
from rcl_reports import add_scrap_stock_sheets
from report_common import new_workbook, workbook_bytes

st.set_page_config(
    page_title="RCL MIS Automation",
    page_icon="https://royalchaingroup.com/wp-content/uploads/2025/09/favicon.png",
    layout="wide")
ui.inject_css()
ui.render_header()

XLSX_MIME = ("application/vnd.openxmlformats-officedocument"
             ".spreadsheetml.sheet")
TODAY = date.today().isoformat()


# ===========================================================================
# Per-tool sheet adders (signature: (workbook, file_bytes) -> result)
# ===========================================================================
def _add_scrap_stock(wb, file_bytes):
    return add_scrap_stock_sheets(wb, file_bytes, TODAY)


# ===========================================================================
# Per-tool preview renderers (result -> Streamlit output)
# ===========================================================================
def _preview_loss(result):
    note = (f"  ·  Dates: {result.date_from} → {result.date_to}"
            if (result.date_from or result.date_to) else "")
    st.success(f"Processed {len(result.rows)} work-center rows.{note}")
    df = pd.DataFrame(result.rows)
    if "Loss %" in df:
        df["Loss %"] = df["Loss %"].apply(
            lambda v: f"{v:.3%}" if v is not None else "")
    st.dataframe(df, use_container_width=True, hide_index=True)


def _preview_scrap_stock(result):
    st.success("Both reports generated (2 sheets).")
    st.markdown("##### ♻️ Scrap Report")
    a, b, c = st.columns(3)
    a.metric("Scrap rows", f"{result['scrap_rows']:,}")
    b.metric("Gross Weight", f"{result['scrap_gross']:,.3f}")
    c.metric("Metal Weight", f"{result['scrap_metal']:,.3f}")
    st.markdown("##### 📦 Stock Report")
    d, e, f = st.columns(3)
    d.metric("Rows after filtering", f"{result['stock_filtered']:,}")
    e.metric("WC Groups", f"{result['stock_groups']:,}")
    f.metric("Gross Weight", f"{result['stock_gross']:,.3f}")


def _preview_lot(result):
    note = (f"  ·  {result.date_from} → {result.date_to}"
            if (result.date_from or result.date_to) else "")
    st.success(f"Processed {len(result.rows)} rejection rows.  "
               f"Total Wt: {result.total_wt:,.2f}{note}")
    st.dataframe(pd.DataFrame(result.rows), use_container_width=True,
                 hide_index=True)


# ===========================================================================
# Tool registry — add a dict here to add a new tab.
# ===========================================================================
TOOLS = [
    {
        "key": "loss", "label": "📉  Loss Report", "title": "Loss Report",
        "subtitle": "Raw export → formatted Loss Report Summary with "
                    "Final Loss and Loss %.",
        "help": "**Final Loss** = Loss Quantity + Gain Pg.  \n"
                "**Loss %** = Final Loss ÷ (Process + Unutilized + Scrap + "
                "Sample).",
        "add": add_loss_sheet, "preview": _preview_loss,
    },
    {
        "key": "scrap_stock", "label": "♻️  Scrap + Stock",
        "title": "Scrap + Stock Report",
        "subtitle": "One daily export → one workbook with SCRAP REPORT and "
                    "STOCK REPORT sheets.",
        "help": "**Scrap**: keeps SCRAP / HL-SCRAP rows grouped by WC.  \n"
                "**Stock**: removes ALLOY/Beads/stones/Pearl, keeps OM under "
                "OTHER METAL, merges duplicates.",
        "add": _add_scrap_stock, "preview": _preview_scrap_stock,
    },
    {
        "key": "lot", "label": "🧾  Lot Rejection",
        "title": "Lot Rejection Report",
        "subtitle": "Raw export → clean Lot Rejection Report in the standard "
                    "layout with a Grand Total of Wt.",
        "help": "Rebuilds the report with the title, date band, merged "
                "Operation / Wc / User columns, and a recomputed Wt total.",
        "add": add_lot_rejection_sheet, "preview": _preview_lot,
    },
]

# Bytes of every successfully-readable upload, keyed by tool — drives the
# combined workbook in the sidebar.
uploaded_bytes: dict[str, bytes] = {}


def render_tool_tab(spec: dict) -> None:
    ui.section_title(spec["title"], spec["subtitle"])
    with st.expander("How it works"):
        st.markdown(spec["help"])

    file = st.file_uploader(
        "Drop the report here  ·  .xls / .xlsx",
        type=["xlsx", "xlsm", "xls"], accept_multiple_files=False,
        key=spec["key"] + "_upl")
    if file is None:
        st.info("Upload a file to process this report.")
        return

    data = file.getvalue()
    st.divider()
    st.subheader(f"📄 {file.name}")
    try:
        wb = new_workbook()
        result = spec["add"](wb, data)
        out_bytes = workbook_bytes(wb)
    except LossReportError as exc:
        st.error(f"Could not process this file: {exc}")
        return
    except Exception as exc:  # noqa: BLE001
        st.error(f"Unexpected error reading the file: {exc}")
        return

    uploaded_bytes[spec["key"]] = data          # include in combined workbook
    spec["preview"](result)

    out_name = f"{spec['title'].replace(' ', '_')}_{TODAY}.xlsx"
    st.download_button(
        f"⬇  Download {spec['title']}", data=out_bytes, file_name=out_name,
        mime=XLSX_MIME, key=spec["key"] + "_dl")


for tab, spec in zip(st.tabs([t["label"] for t in TOOLS]), TOOLS):
    with tab:
        render_tool_tab(spec)


# ===========================================================================
# Sidebar — universal combined workbook (all uploaded tabs as sub-sheets)
# ===========================================================================
def render_combined_sidebar() -> None:
    with st.sidebar:
        st.markdown("### 📚 Combined Workbook")
        st.caption("One file with every uploaded tab's output as sub-sheets.")

        if not uploaded_bytes:
            st.info("Upload a file in any tab to enable this.")
            return

        wb = new_workbook()
        included, errors = [], []
        for spec in TOOLS:
            if spec["key"] not in uploaded_bytes:
                continue
            try:
                spec["add"](wb, uploaded_bytes[spec["key"]])
                included.append(spec["title"])
            except Exception as exc:  # noqa: BLE001
                errors.append(f"{spec['title']}: {exc}")

        if included:
            st.success("Included: " + ", ".join(included))
            st.download_button(
                "⬇  Download Combined Workbook",
                data=workbook_bytes(wb),
                file_name=f"RCL_MIS_Combined_{TODAY}.xlsx",
                mime=XLSX_MIME, key="combined_dl")
        for msg in errors:
            st.warning(msg)


render_combined_sidebar()
ui.render_footer()
