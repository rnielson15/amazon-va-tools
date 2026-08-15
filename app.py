"""
Streamlit front end for the Amazon VA tools.

Run locally:
    streamlit run app.py

Drag your Amazon "Bulk Operations" export onto the page, review the bleeder
counts, and download the finished report. No command line needed after launch.

Tools 2 and 3 can be added later as extra entries in TOOLS below, each pointing
at its own engine's generate() function -- the UI code does not change.
"""

import streamlit as st
import pandas as pd

from bleeders import generate as generate_bleeders, workbook_to_bytes

# Registry of available tools. Add future tools here.
TOOLS = {
    "Bleeders Report": {
        "generate": generate_bleeders,
        "blurb": "Enabled keywords / targets / search terms with >10 clicks and $0 sales.",
        "output_name": "Bleeders_Report.xlsx",
    },
}

st.set_page_config(page_title="Amazon VA Tools", page_icon="📊", layout="centered")
st.title("Amazon VA Tools")

tool_name = st.selectbox("Report to build", list(TOOLS.keys()))
tool = TOOLS[tool_name]
st.caption(tool["blurb"])

uploaded = st.file_uploader(
    "Drop your Amazon Bulk Operations export (.xlsx)",
    type=["xlsx"],
    help="Campaign Manager → Bulk operations → 60-day range → Create spreadsheet for download",
)

if uploaded is not None:
    try:
        with st.spinner("Building report…"):
            summary, wb = tool["generate"](uploaded)
            data = workbook_to_bytes(wb)
    except Exception as exc:                     # noqa: BLE001
        st.error(f"Could not process that file: {exc}")
        st.stop()

    total = sum(n for _, n in summary)

    st.subheader(f"{total} bleeders found")
    st.dataframe(
        pd.DataFrame(summary, columns=["Section", "Bleeders"]),
        hide_index=True,
        use_container_width=True,
    )

    st.download_button(
        label="⬇️  Download report",
        data=data,
        file_name=tool["output_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if total == 0:
        st.info("No bleeders found — the relevant source tabs may be empty in this export.")
else:
    st.info("Upload a file to begin.")
