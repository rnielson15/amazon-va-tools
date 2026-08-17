"""
Streamlit front end for the Amazon VA tools.

Run locally:
    streamlit run app.py

Each tool declares its own inputs (which files it needs and any settings like
gross margin). The UI renders the right uploaders and fields automatically, so
adding tool #3 later is just another entry in TOOLS -- no UI changes.
"""

import streamlit as st
import pandas as pd

from bleeders import generate as gen_v1
from bleeders_v2 import generate as gen_v2, describe as describe_v2
from common import workbook_to_bytes

TOOLS = {
    "Bleeders Report — clicks, no sales (v1)": {
        "blurb": "Enabled keywords / targets / search terms with more than 10 clicks and $0 sales.",
        "inputs": [{"key": "bulk", "label": "Amazon Bulk Operations export (.xlsx)"}],
        "params": [],
        "run": lambda files, params: gen_v1(files["bulk"]),
        "output_name": "Bleeders_Report.xlsx",
    },
    "Bleeders 2.0 — high ACOS, low orders": {
        "blurb": "Sponsored Brands & Display targets with 5 or fewer orders and ACOS above your break-even threshold.",
        "inputs": [
            {"key": "sp_search", "label": "Sponsored Products Search-term report (.xlsx)", "optional": True},
            {"key": "sp_target", "label": "Sponsored Products Targeting report (.xlsx)", "optional": True},
            {"key": "sb_kw", "label": "Sponsored Brands Keyword report (.xlsx)", "optional": True},
            {"key": "sd_tgt", "label": "Sponsored Display Targeting report (.xlsx)", "optional": True},
        ],
        "params": [{"key": "gross_margin", "label": "Break-even ACOS % (= your gross margin)",
                    "default": 35.0, "min": 0.0, "max": 100.0, "step": 1.0}],
        "run": lambda files, params: gen_v2(files, params),
        "describe": describe_v2,
        "output_name": "Bleeders_Report_2.xlsx",
    },
}

st.set_page_config(page_title="Amazon VA Tools", page_icon="📊", layout="centered")
st.title("Amazon VA Tools")

tool_name = st.selectbox("Report to build", list(TOOLS.keys()))
tool = TOOLS[tool_name]
st.caption(tool["blurb"])

# --- settings (params) ---
params = {}
for p in tool["params"]:
    params[p["key"]] = st.number_input(
        p["label"], value=p.get("default", 0.0),
        min_value=p.get("min"), max_value=p.get("max"), step=p.get("step", 1.0),
    )

# Live description of what the current settings will flag.
if tool.get("describe") and params:
    st.info(tool["describe"](params))

# --- file inputs ---
files = {}
for inp in tool["inputs"]:
    files[inp["key"]] = st.file_uploader(inp["label"], type=["xlsx"], key=inp["key"])

required = [i for i in tool["inputs"] if not i.get("optional")]
have_required = all(files[i["key"]] is not None for i in required)
any_file = any(files[i["key"]] is not None for i in tool["inputs"])
ready = have_required and any_file

if ready:
    try:
        with st.spinner("Building report…"):
            summary, wb = tool["run"](files, params)
            data = workbook_to_bytes(wb)
    except Exception as exc:                     # noqa: BLE001
        st.error(f"Could not process the file(s): {exc}")
        st.stop()

    total = sum(n for _, n in summary)
    st.subheader(f"{total} bleeders found")
    st.dataframe(pd.DataFrame(summary, columns=["Section", "Bleeders"]),
                 hide_index=True, use_container_width=True)
    st.download_button(
        "⬇️  Download report", data=data, file_name=tool["output_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    if total == 0:
        st.info("No bleeders found — the source report(s) may have no rows that meet the rule.")
else:
    st.info("Upload at least one report above to begin.")
