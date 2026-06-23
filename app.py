"""
Streamlit web app for the Monthly Loss Report automation.

Run with:
    streamlit run app.py

Drag-drop your Loss Report Excel file and download it back with two extra
columns -- Final Loss and Loss %. Everything else in the sheet is left exactly
as it was.
"""

from __future__ import annotations

import pandas as pd
import streamlit as st

from loss_report import LossReportError, process

st.set_page_config(page_title="Loss Report Automation", page_icon="📊",
                   layout="wide")

st.title("📊 Monthly Loss Report Automation")
st.caption(
    "Upload the Loss Report sheet. The app keeps every original column and "
    "adds **Final Loss** and **Loss %** on the right."
)

with st.expander("How the new columns are calculated", expanded=False):
    st.markdown(
        "- **Final Loss** = Loss Quantity Pg + Gain Pg  \n"
        "  *(Gain shown like `(4.89)` counts as negative.)*\n"
        "- **Loss %** = Final Loss ÷ (Process Qty + Unutilized Qty + "
        "Unutilized Scrap Qty + Unutilized Sample Qty)"
    )

uploaded = st.file_uploader(
    "Upload Loss Report (.xlsx / .xls)", type=["xlsx", "xlsm", "xls"],
    accept_multiple_files=True,
    help=("You can drop more than one file; each is processed separately. "
          "Note: old .xls files keep all data but not original cell colours "
          "(the output is always a modern .xlsx)."),
)

if uploaded:
    for file in uploaded:
        st.divider()
        st.subheader(f"📄 {file.name}")
        try:
            out_bytes, report = process(file.getvalue())
        except LossReportError as exc:
            st.error(f"Could not process this file: {exc}")
            continue
        except Exception as exc:  # noqa: BLE001 - surface any read error
            st.error(f"Unexpected error reading the file: {exc}")
            continue

        date_note = ""
        if report.date_from or report.date_to:
            date_note = f" · Dates: {report.date_from} → {report.date_to}"
        st.success(f"Processed {len(report.rows)} work-center rows.{date_note}")
        if not (report.date_from and report.date_to):
            st.warning("Could not read the From/To dates from the sheet's "
                       "preamble — the date band may be blank. Check the file.")

        preview = pd.DataFrame(report.rows)
        if "Loss %" in preview:
            preview["Loss %"] = preview["Loss %"].apply(
                lambda v: f"{v:.3%}" if v is not None else ""
            )
        st.dataframe(preview, use_container_width=True, hide_index=True)

        out_name = file.name.rsplit(".", 1)[0] + "_processed.xlsx"
        st.download_button(
            label=f"⬇️ Download {out_name}",
            data=out_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=file.name,
        )
else:
    st.info("Upload one or more .xlsx files to begin.")
