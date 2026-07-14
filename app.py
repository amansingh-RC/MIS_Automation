from __future__ import annotations

from datetime import date, datetime

import pandas as pd
import streamlit as st

import ui
from factory_inward import (add_factory_inward_sheet, default_positions,
                            load_positions)
from factory_outward import add_factory_outward_sheet
from filling_loss import (add_cast_gold_buffing_sheet, add_filling_loss_sheet,
                          add_gold_buffing_sheet)
from groupsales import add_groupsales_sheet
from inter_wc_transfer import add_inter_wc_transfer_sheet
import html as _html

import mailer
from loss_report import LossReportError, add_loss_sheet
from lot_rejection import add_lot_rejection_sheet
from rcl_reports import (add_scrap_stock_manish_sheets,
                         add_scrap_stock_sheets)
from report_common import new_workbook, workbook_bytes
from return_summary import add_return_summary_sheet

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

# Default email subject / body shown (and editable) in both mail sections.
DEFAULT_MAIL_SUBJECT = "MIS | Daily Report"
DEFAULT_MAIL_MESSAGE = ("Hey Sir,\nPls find the MIS daily report.\n\n"
                        "Thanks & Regards")


MANISH_LOSS_WC = [
    "CAST GOLD BUFFING WK", "CAST RHODIUM PLATING-WK", "CAST-MELTING-WK",
    "ENAMEL WK", "GOLD METAL SETTING WK", "GOLD-ASSEMBLY-WK",
    "GOLD-BUFFING-WK", "GOLD-CASTING-WK", "GOLD-EP-WK", "GOLD-FILLING-WK",
    "GOLD-REPAIR-ASSEMBLY-WK", "GOLD-REPAIR-WK", "GOLD-SHORT-WK",
    "GOLD-TREE-CUTTING-WK", "MEDIA-POLISH-WK", "REFINING-WK",
    "REPAIR CAST GOLD BUFFING WK", "REPAIR-GOLD-BUFFING-WK",
    "RHODIUM-PLATING-WK", "ROLLING-WK",
]


def _add_loss_manish(wb, file_bytes):
    return add_loss_sheet(wb, file_bytes, wc_filter=MANISH_LOSS_WC)


# Manish Report: Wcgroup Names to keep in the Scrap / Stock sheets.
MANISH_SCRAP_WCG = [
    "CAST GOLD BUFFING HOD", "CENTRAL-OFFICE-1F", "ELECTROPLATING",
    "GOLD-BUFFING-HOD", "GOLD-CASTING", "GOLD-FILLING", "GOLD-REPAIR-HOD",
    "GOLD-SHORT-HOD", "HAMMERING", "HAND-CUTTING", "MEDIA-POLISH",
]
MANISH_STOCK_WCG = [
    "CAD", "CAM", "CAST GOLD BUFFING HOD", "CAST-MAGNATIC-HOD",
    "CENTRAL-OFFICE-1F", "ELECTROPLATING", "ENAMEL HOD", "GOLD METALSETTING",
    "GOLD-ASSEMBLY", "GOLD-BUFFING-HOD", "GOLD-CASTING", "GOLD-FILLING",
    "GOLD-REPAIR-HOD", "GOLD-SEPERATION", "GOLD-SHORT-HOD", "HALLMARK",
    "HAMMERING", "HAND-CUTTING", "MEDIA-POLISH", "PACKING & QA",
    "PHOTOGRAPHY HOD", "REFINING-HOD", "WAX",
]


def _add_scrap_stock_manish(wb, file_bytes):
    now = datetime.now()
    return add_scrap_stock_manish_sheets(
        wb, file_bytes, TODAY_DMY, f"{now.hour}:{now.minute:02d}",
        MANISH_SCRAP_WCG, MANISH_STOCK_WCG)


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


def _preview_return_summary(result):
    g = result.grand

    def gv(k, i):
        return g.get(k, [0.0, 0.0])[i]

    msg = (f"{result.parties} parties  ·  {result.date_from} to "
           f"{result.date_to}  ·  Receipt Metal Net: {gv('METAL', 0):,.3f}  ·  "
           f"Issue Net: {gv('ISSUE', 0):,.3f}  ·  "
           f"Return Net: {gv('RETURN', 0):,.3f}")
    if result.unlisted:
        msg += (f"  ·  {len(result.unlisted)} party(ies) not in the position "
                "table (placed at end)")
    st.success(msg)
    st.dataframe(pd.DataFrame(result.rows), use_container_width=True,
                 hide_index=True)


def _preview_inter_wc(result):
    note = f"  ·  DATE : {result.date}" if result.date else ""
    st.success(f"{result.source_count} source workers  ·  {result.kept_rows} "
               f"rows kept  ·  Grand Total — Weight: {result.grand_weight:,.3f}"
               f"  ·  Pg Metal: {result.grand_pg:,.3f}{note}")
    st.dataframe(pd.DataFrame(result.rows), use_container_width=True,
                 hide_index=True)


def render_return_summary(spec: dict) -> None:
    """Custom tab renderer: this report needs TWO uploads (inward + outward)."""
    ui.section_title(spec["title"], spec["subtitle"])
    with st.expander("How it works"):
        st.markdown(spec["help"])
    _factory_extra(spec)                       # optional position-table override

    col_in, col_out = st.columns(2)
    with col_in:
        f_in = st.file_uploader("Factory Inward (GRN)  ·  .xls / .xlsx",
                                type=["xlsx", "xlsm", "xls"],
                                key=spec["key"] + "_in")
    with col_out:
        f_out = st.file_uploader("Factory Outward (INV)  ·  .xls / .xlsx",
                                 type=["xlsx", "xlsm", "xls"],
                                 key=spec["key"] + "_out")
    if f_in is None or f_out is None:
        st.info("Upload BOTH the Factory Inward (GRN) and Factory Outward "
                "(INV) exports to build this report.")
        return

    st.divider()
    try:
        wb = new_workbook()
        result = add_return_summary_sheet(
            wb, f_in.getvalue(), f_out.getvalue(),
            _get_positions(spec["key"]))
        out_bytes = workbook_bytes(wb)
    except LossReportError as exc:
        st.error(f"Could not process these files: {exc}")
        return
    except Exception as exc:  # noqa: BLE001
        st.error(f"Unexpected error reading the files: {exc}")
        return

    spec["preview"](result)
    out_name = f"{spec['title'].replace(' ', '_')}_{TODAY}.xlsx"
    st.download_button(
        f"⬇  Download {spec['title']}", data=out_bytes, file_name=out_name,
        mime=XLSX_MIME, key=spec["key"] + "_dl")


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
        "manish": _add_loss_manish,
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
        "manish": _add_scrap_stock_manish,
    },
    {
        "key": "lot", "label": "  Lot Rejection",
        "title": "Lot Rejection Report",
        "subtitle": "Raw export → clean Lot Rejection Report in the standard "
                    "layout with a Grand Total of Wt.",
        "help": "Rebuilds the report with the title, date band, merged "
                "Operation / Wc / User columns, and a recomputed Wt total.",
        "add": add_lot_rejection_sheet, "preview": _preview_lot,
        "manish": add_lot_rejection_sheet,      # included as-is in Manish Report
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
        "key": "gold_buffing", "label": "  Gold Buffing",
        "title": "Gold Buffing Loss & Recovery Report",
        "subtitle": "Raw GOLD-BUFFING loss export → remark × karat pivot of "
                    "Issue / Return / Unutilized / Loss weights.",
        "help": "Same process as Filling Loss & Recovery, on the buffing data. "
                "Continuation rows inherit the batch/operation/WC above. Karat "
                "comes from the Variant Name. **Roundup** = ROUNDUP(Return Pg, "
                "3); **Roundup2** = Roundup − MINIFS(Roundup by Batch). "
                "**Remark**: REP (REPAIR-GOLD-BUFFING-WK), else DOUBLE ENTRY "
                "for operations other than B-GOLD-BUFFING / Buffing-2 or NONE "
                "batches, else FINISH when Roundup2 = 0 (one FINISH kept per "
                "batch), else DOUBLE ENTRY.",
        "add": add_gold_buffing_sheet, "preview": _preview_filling_loss,
    },
    {
        "key": "cast_gold_buffing", "label": "  Cast Gold Buffing",
        "title": "Cast Gold Buffing Loss & Recovery Report",
        "subtitle": "Raw CAST GOLD BUFFING loss export → remark × karat pivot "
                    "of Issue / Return / Unutilized / Loss weights.",
        "help": "Same process as Filling Loss & Recovery, on the cast gold "
                "buffing data. Continuation rows inherit the batch/operation/"
                "WC above. Karat comes from the Variant Name. **Roundup** = "
                "ROUNDUP(Return Pg, 3); **Roundup2** = Roundup − MINIFS(Roundup "
                "by Batch). **Remark**: REP (REPAIR CAST GOLD BUFFING WK), else "
                "DOUBLE ENTRY for operations other than B-GOLD-BUFFING / "
                "Buffing-2 or NONE batches, else FINISH when Roundup2 = 0 (one "
                "FINISH kept per batch), else DOUBLE ENTRY.",
        "add": add_cast_gold_buffing_sheet, "preview": _preview_filling_loss,
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
    {
        "key": "return_summary", "label": "Return % Summary",
        "title": "Return % Summary Report",
        "subtitle": "Factory Inward (GRN) + Factory Outward (INV) exports → "
                    "per-party × karat pivot of Receipt Metal / Issue Metal / "
                    "Issue / Return with Net Goods and Return %.",
        "help": "Both files get the Factory Inward/Outward treatment (Karat "
                "from the Variant Name, out-of-band = OT) without the pivot, "
                "then are merged. **24KT → 24KT-METAL**. **Remark**: inward + "
                "24KT = Metal (Receipt Metal), outward + 24KT = Issue Metal, "
                "outward + other = Issue, inward + other = Return. Parties are "
                "ordered by the position table. **Net Goods** = Issue − "
                "Return; **Return %** = Return ÷ Issue (Net and Pg).",
        "render": render_return_summary, "preview": _preview_return_summary,
    },
    {
        "key": "inter_wc", "label": " Inter WC Transfer",
        "title": "Inter WC Group Transfer Outward",
        "subtitle": "Raw Inter WC Group Transfer Outward export → Source Worker "
                    "→ Dest Worker pivot of Weight and Pg Metal Weight.",
        "help": "Keeps only the configured Source Workers (CAM, CAST GOLD "
                "BUFFING HOD, … WAX), then pivots by Source Worker → Dest "
                "Worker, summing **Weight** and **Pg Metal Weight**, with a "
                "per-source total and a Grand Total. Rows are in alphabetical "
                "(pivot) order.",
        "add": add_inter_wc_transfer_sheet, "preview": _preview_inter_wc,
        "manish": add_inter_wc_transfer_sheet,   # goes into Manish, not Combined
        "skip_combined": True,
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
        spec.get("render", render_tool_tab)(spec)


def _smtp_config() -> "mailer.SmtpConfig":
    try:
        smtp = dict(st.secrets.get("smtp", {}))
    except Exception:  # noqa: BLE001 (no secrets.toml present)
        smtp = {}
    return mailer.SmtpConfig.from_dict(smtp)


def _mail_default(key: str) -> str:
    try:
        return str(st.secrets.get("mail", {}).get(key, "") or "")
    except Exception:  # noqa: BLE001
        return ""


def _email_section(key: str, default_subject: str, default_message: str,
                   attach_name: str, attach_bytes: bytes,
                   default_to: str) -> None:
    """Multi-recipient email with an editable Subject and Message and the
    workbook attached. The body is exactly the message (plain, clean)."""
    cfg = _smtp_config()
    with st.expander("Email this report"):
        if not cfg.configured:
            st.caption("Email not configured. Copy "
                       "`.streamlit/secrets.toml.example` to "
                       "`.streamlit/secrets.toml` and add your Gmail SMTP "
                       "settings + App Password.")
        to = st.text_area(
            "Send to  ·  add multiple (one per line or comma-separated)",
            value=st.session_state.get(key + "_to", default_to),
            key=key + "_to", height=90,
            placeholder="name@royalchains.com\nanother@royalchains.com")
        subject = st.text_input(
            "Subject", value=st.session_state.get(key + "_subj",
                                                  default_subject),
            key=key + "_subj")
        message = st.text_area(
            "Message  ·  your custom email text",
            value=st.session_state.get(key + "_msg", default_message),
            key=key + "_msg", height=140)

        recipients = mailer.parse_addrs(to)
        invalid = [a for a in recipients if not mailer.valid_addr(a)]
        if recipients:
            note = f"{len(recipients)} recipient(s): " + ", ".join(recipients)
            if invalid:
                st.warning(f"{note}\n\n⚠ Invalid: {', '.join(invalid)}")
            else:
                st.caption(note)
        if st.button("Send Email", key=key + "_send",
                     use_container_width=True, disabled=not cfg.configured):
            try:
                sent = mailer.send_email(
                    cfg, to, subject or default_subject, _email_html(message),
                    attachments=[(attach_name, attach_bytes)])
                st.success(f"Email successfully sent to {len(sent)} "
                           f"recipient(s): " + ", ".join(sent))
            except mailer.MailError as exc:
                st.error(str(exc))


def _email_html(message: str) -> str:
    """Render the plain-text message as a clean, simple email body."""
    safe_msg = _html.escape(message or "").replace("\n", "<br>")
    return (f'<div style="font-family:Arial,Helvetica,sans-serif;'
            f'font-size:14px;color:#202124;line-height:1.5">{safe_msg}</div>')


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
            if (not spec.get("add") or spec.get("skip_combined")
                    or spec["key"] not in uploaded_bytes):
                continue
            try:
                spec["add"](wb, uploaded_bytes[spec["key"]])
                included.append(spec["title"])
            except Exception as exc:  # noqa: BLE001
                errors.append(f"{spec['title']}: {exc}")

        if included:
            data = workbook_bytes(wb)
            fname = f"RCL_MIS_Combined_{TODAY}.xlsx"
            st.success("Included: " + ", ".join(included))
            st.download_button(
                "⬇  Download Combined Workbook", data=data,
                file_name=fname, mime=XLSX_MIME, key="combined_dl")
            _email_section(
                "combined", default_subject=DEFAULT_MAIL_SUBJECT,
                default_message=DEFAULT_MAIL_MESSAGE,
                attach_name=fname, attach_bytes=data,
                default_to=_mail_default("combined_to"))
        for msg in errors:
            st.warning(msg)


def render_manish_sidebar() -> None:
    """Second combined workbook: each configured tab contributes a customized
    (column/row-limited) version of its report."""
    with st.sidebar:
        st.divider()
        st.markdown("###  Manish Report")
        st.caption("A combined workbook with a customized version of each "
                   "report (e.g. Loss Report limited to specific work centres).")

        specs = [s for s in TOOLS
                 if s.get("manish") and s["key"] in uploaded_bytes]
        if not specs:
            st.info("Upload a file in a configured tab (Loss Report) to "
                    "enable this.")
            return

        wb = new_workbook()
        included, errors = [], []
        for spec in specs:
            try:
                spec["manish"](wb, uploaded_bytes[spec["key"]])
                included.append(spec["title"])
            except Exception as exc:  # noqa: BLE001
                errors.append(f"{spec['title']}: {exc}")

        if included:
            data = workbook_bytes(wb)
            fname = f"Manish_Report_{TODAY}.xlsx"
            st.success("Included: " + ", ".join(included))
            st.download_button(
                "⬇  Download Manish Report", data=data,
                file_name=fname, mime=XLSX_MIME, key="manish_dl")
            _email_section(
                "manish", default_subject=DEFAULT_MAIL_SUBJECT,
                default_message=DEFAULT_MAIL_MESSAGE,
                attach_name=fname, attach_bytes=data,
                default_to=_mail_default("manish_to"))
        for msg in errors:
            st.warning(msg)


render_combined_sidebar()
render_manish_sidebar()
ui.render_footer()
