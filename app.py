from __future__ import annotations

from datetime import date

import pandas as pd
import streamlit as st

import ui
from factory_inward import (add_factory_inward_sheet, default_positions,
                            load_positions)
from factory_outward import add_factory_outward_sheet
from filling_loss import add_filling_loss_sheet
from groupsales import add_groupsales_sheet
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
TODAY_DMY = date.today().strftime("%d-%m-%Y")

def _add_scrap_stock(wb, file_bytes):
    return add_scrap_stock_sheets(wb, file_bytes, TODAY)


def _get_positions(tool_key):
    """Per-tab session override position table, else the bundled default."""
    return st.session_state.get(tool_key + "_positions") or default_positions()


def _add_factory_inward(wb, file_bytes):
    return add_factory_inward_sheet(
        wb, file_bytes, _get_positions("factory_inward"), TODAY_DMY)


def _add_factory_outward(wb, file_bytes):
    return add_factory_outward_sheet(
        wb, file_bytes, _get_positions("factory_outward"), TODAY_DMY)


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
    st.markdown("#####  Scrap Report")
    a, b, c = st.columns(3)
    a.metric("Scrap rows", f"{result['scrap_rows']:,}")
    b.metric("Gross Weight", f"{result['scrap_gross']:,.3f}")
    c.metric("Metal Weight", f"{result['scrap_metal']:,.3f}")
    st.markdown("#####  Stock Report")
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


def _preview_factory(result):
    note = (f"  ·  {result.date_from} → {result.date_to}"
            if (result.date_from or result.date_to) else "")
    msg = (f"Grand Total — Net Wt: {result.grand_net:,.3f}  ·  "
           f"Pg Wt: {result.grand_pg:,.3f}{note}")
    if result.unlisted:
        msg += (f"  ·  {len(result.unlisted)} party(ies) not in position "
                "table (placed at end)")
    if result.skipped:
        msg += f"  ·  {result.skipped} row(s) skipped (variant matched no karat)"
    st.success(msg)
    st.dataframe(pd.DataFrame(result.rows), use_container_width=True,
                 hide_index=True)


def _factory_extra(spec):
    """Optional position-table override for a Factory tab (namespaced per tab)."""
    sess_key = spec["key"] + "_positions"
    with st.expander("Party position table (optional override)"):
        st.caption("By default the bundled position table is used. Upload a "
                   "new one (columns: Name, Position) to override it.")
        pf = st.file_uploader("Position table (.xlsx / .xls)",
                              type=["xlsx", "xlsm", "xls"],
                              key=spec["key"] + "_pos_upl")
        if pf is not None:
            try:
                st.session_state[sess_key] = load_positions(pf.getvalue())
                st.success(f"Loaded {len(st.session_state[sess_key])} "
                           "party positions from the uploaded table.")
            except Exception as exc:  # noqa: BLE001
                st.error(f"Position table error: {exc}")
        else:
            st.session_state.pop(sess_key, None)
            st.caption(f"Using bundled table "
                       f"({len(default_positions())} parties).")


def _preview_filling_loss(result):
    note = f"  ·  {result.date_range}" if result.date_range else ""
    counts = "  ·  ".join(f"{k}: {v}" for k, v in result.remark_counts.items())
    msg = (f"Grand Total — Issue Pg: {result.grand['issue_pg']:,.3f}  ·  "
           f"Return Pg: {result.grand['ret_pg']:,.3f}  ·  "
           f"Loss Pg: {result.grand['loss_pg']:,.3f}{note}")
    if counts:
        msg += f"  ·  {counts}"
    if result.skipped:
        msg += f"  ·  {result.skipped} row(s) with no karat"
    st.success(msg)
    st.dataframe(pd.DataFrame(result.rows), use_container_width=True,
                 hide_index=True)


def _preview_groupsales(result):
    note = f"  ·  Date: {result.date}" if result.date else ""
    msg = (f"Grand Total — Net Wt: {result.grand_net:,.3f}  ·  "
           f"Pg Wt: {result.grand_pg:,.3f}{note}")
    if result.skipped:
        msg += f"  ·  {result.skipped} row(s) skipped (melting matched no karat)"
    st.success(msg)
    st.dataframe(pd.DataFrame(result.rows), use_container_width=True,
                 hide_index=True)

TOOLS = [
    {
        "key": "loss", "label": "  Loss Report", "title": "Loss Report",
        "subtitle": "Raw export → formatted Loss Report Summary with "
                    "Final Loss and Loss %.",
        "help": "**Final Loss** = Loss Quantity + Gain Pg.  \n"
                "**Loss %** = Final Loss ÷ (Process + Unutilized + Scrap + "
                "Sample).",
        "add": add_loss_sheet, "preview": _preview_loss,
    },
    {
        "key": "scrap_stock", "label": "  Scrap + Stock",
        "title": "Scrap + Stock Report",
        "subtitle": "One daily export → one workbook with SCRAP REPORT and "
                    "STOCK REPORT sheets.",
        "help": "**Scrap**: keeps SCRAP / HL-SCRAP rows grouped by WC.  \n"
                "**Stock**: removes ALLOY/Beads/stones/Pearl, keeps OM under "
                "OTHER METAL, merges duplicates.",
        "add": _add_scrap_stock, "preview": _preview_scrap_stock,
    },
    {
        "key": "lot", "label": "  Lot Rejection",
        "title": "Lot Rejection Report",
        "subtitle": "Raw export → clean Lot Rejection Report in the standard "
                    "layout with a Grand Total of Wt.",
        "help": "Rebuilds the report with the title, date band, merged "
                "Operation / Wc / User columns, and a recomputed Wt total.",
        "add": add_lot_rejection_sheet, "preview": _preview_lot,
    },
    {
        "key": "groupsales", "label": "  Groupsales",
        "title": "Groupsales Reports",
        "subtitle": "Raw export → pivot of Groupsales → Karat → Net Wt / Pg Wt "
                    "with per-group and grand totals.",
        "help": "Karat is derived from Metal Fineness × 100 (fallback: Variant "
                "Name): 24KT 99–100, 22KT 91–92.5, 18KT 74.5–76, 14KT "
                "57.5–59.8. Rows with a blank Groupsales are grouped under "
                "\"(blank)\".",
        "add": add_groupsales_sheet, "preview": _preview_groupsales,
    },
    {
        "key": "filling_loss", "label": "  Filling Loss & Recovery",
        "title": "Filling Loss & Recovery Report",
        "subtitle": "Raw GOLD-FILLING loss export → remark × karat pivot of "
                    "Issue / Return / Unutilized / Loss weights.",
        "help": "Continuation rows inherit the batch/operation/WC above. Karat "
                "comes from the Variant Name. **Roundup** = ROUNDUP(Return Pg, "
                "3); **Roundup2** = Roundup − MINIFS(Roundup by Batch). "
                "**Remark**: REP (REP-GOLD-FILLING-WK), else DOUBLE ENTRY for "
                "non B-GOLD-FILING / CAST-FILLING operations or NONE batches, "
                "else FINISH when Roundup2 = 0 (one FINISH kept per batch), "
                "else DOUBLE ENTRY. The pivot sums the 12 weight columns by "
                "remark and karat.",
        "add": add_filling_loss_sheet, "preview": _preview_filling_loss,
    },
    {
        "key": "factory_inward", "label": "  Factory Inward",
        "title": "Factory Inward (GRN)",
        "subtitle": "Raw GRN export → pivot of Sr. No. / Party Name / Karat / "
                    "Net Wt / Pg Wt, ordered by the party position table.",
        "help": "Karat comes from the Variant Name (decimal melting %, else a "
                "literal NNKT). Parties are ordered by the position table and "
                "numbered serially (1, 2, 3…); parties not in the table go to "
                "the end. Includes a karat-only summary block.",
        "add": _add_factory_inward, "preview": _preview_factory,
        "extra": _factory_extra,
    },
    {
        "key": "factory_outward", "label": "  Factory Outward",
        "title": "Factory Outward",
        "subtitle": "Raw export → pivot of Sr. No. / Party Name / Karat / "
                    "Net Wt / Pg Wt, ordered by the party position table.",
        "help": "Same processing as Factory Inward: Karat from the Variant "
                "Name, parties ordered by the position table and numbered "
                "serially, with per-party totals, a Grand Total, and a "
                "karat-only summary block.",
        "add": _add_factory_outward, "preview": _preview_factory,
        "extra": _factory_extra,
    },
]

uploaded_bytes: dict[str, bytes] = {}


def render_tool_tab(spec: dict) -> None:
    ui.section_title(spec["title"], spec["subtitle"])
    with st.expander("How it works"):
        st.markdown(spec["help"])
    if spec.get("extra"):
        spec["extra"](spec)

    file = st.file_uploader(
        "Drop the report here  ·  .xls / .xlsx",
        type=["xlsx", "xlsm", "xls"], accept_multiple_files=False,
        key=spec["key"] + "_upl")
    if file is None:
        st.info("Upload a file to process this report.")
        return

    data = file.getvalue()
    st.divider()
    st.subheader(f" {file.name}")
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


def render_combined_sidebar() -> None:
    with st.sidebar:
        st.markdown("###  Combined Workbook")
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
