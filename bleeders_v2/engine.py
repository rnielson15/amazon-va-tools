"""
bleeders_v2.engine  —  "Bleeders 2.0"
-------------------------------------
Flags targets/search terms that are *unprofitable*: they have a small number of
orders (1-5) and an ACOS above a break-even-derived threshold. Zero-order rows
are intentionally excluded -- those belong to the original Bleeders Report.

The threshold comes from your break-even ACOS (= gross margin), typed into the
app each run, plus a per-section markup taken from the SOP:

    Brands / Display / SP Search Terms : threshold = break-even + 10 points
    SP Targeting                       : threshold = break-even + 20 points

Example at 35% break-even: SP search/Brands/Display flag ACOS >= 45%,
SP targeting flags ACOS >= 55%.

Columns are matched by header name, whitespace/case-insensitive (Amazon leaves
stray trailing spaces on some headers).
"""

import pandas as pd
from common import build_report, workbook_to_bytes  # noqa: F401 (re-exported)

MIN_ORDERS = 1
MAX_ORDERS = 5

# Amazon attribution columns differ by report family
ACOS = "Total Advertising Cost of Sales (ACOS)"
ROAS = "Total Return on Advertising Spend (ROAS)"
ORDERS_SP = "7 Day Total Orders (#)"          # Sponsored Products reports
ORDERS_SBSD = "14 Day Total Orders (#) - (Click)"   # Brands / Display reports
SALES_SP = "7 Day Total Sales"
SALES_SBSD = "14 Day Total Sales"


def _norm(s):
    return str(s).strip().lower()


def _get(row, name, default=""):
    target = _norm(name)
    for k in row.index:
        if _norm(k) == target:
            v = row[k]
            return default if pd.isna(v) else v
    return default


def _col(df, name):
    target = _norm(name)
    for k in df.columns:
        if _norm(k) == target:
            return k
    return None


def _num(v):
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return 0
        return float(v)
    except (TypeError, ValueError):
        return v if v not in (None, "") else ""


TEXT_HEADERS = {"campaign name", "ad group name", "targeting", "match type",
                "customer search term", "customer search terms"}

# ---------------------------------------------------------------------------
# Block definitions, in template order. Each maps one source report to one
# output section.
# ---------------------------------------------------------------------------
BLOCKS = [
    {
        "file_key": "sp_search", "section": "Sponsored Products", "sub": "Search Term",
        "markup": 0.10, "exclude_exact": True,
        "orders_col": ORDERS_SP, "acos_col": ACOS,
        "headers": ["Campaign Name", "Ad Group Name", "Targeting", "Match Type", "",
                    "Customer Search Terms", "Impressions", "Clicks",
                    "Click-Thru Rate (CTR)", "Cost Per Click (CPC)", "Spend",
                    "7 Day Total Sales", "Total Advertising Cost of Sales (ACOS)",
                    "Total Return on Advertising Spend (ROAS)", "7 Day Total Orders (#)"],
        "map": {"Campaign Name": "Campaign Name", "Ad Group Name": "Ad Group Name",
                "Targeting": "Targeting", "Match Type": "Match Type",
                "Customer Search Terms": "Customer Search Term",
                "Impressions": "Impressions", "Clicks": "Clicks",
                "Click-Thru Rate (CTR)": "Click-Thru Rate (CTR)",
                "Cost Per Click (CPC)": "Cost Per Click (CPC)", "Spend": "Spend",
                "7 Day Total Sales": SALES_SP,
                "Total Advertising Cost of Sales (ACOS)": ACOS,
                "Total Return on Advertising Spend (ROAS)": ROAS,
                "7 Day Total Orders (#)": ORDERS_SP},
    },
    {
        "file_key": "sp_target", "section": "Sponsored Products", "sub": "Targeting",
        "markup": 0.20, "exclude_exact": False,
        "orders_col": ORDERS_SP, "acos_col": ACOS,
        "headers": ["Campaign Name", "Ad Group Name", "Targeting", "Match Type", "",
                    "Clicks", "Click-Thru Rate (CTR)", "Cost Per Click (CPC)", "Spend",
                    "7 Day Total Sales", "ACOS", "Orders"],
        "map": {"Campaign Name": "Campaign Name", "Ad Group Name": "Ad Group Name",
                "Targeting": "Targeting", "Match Type": "Match Type", "Clicks": "Clicks",
                "Click-Thru Rate (CTR)": "Click-Thru Rate (CTR)",
                "Cost Per Click (CPC)": "Cost Per Click (CPC)", "Spend": "Spend",
                "7 Day Total Sales": SALES_SP, "ACOS": ACOS, "Orders": ORDERS_SP},
    },
    {
        "file_key": "sb_kw", "section": "Sponsored Brands", "sub": "Keywords",
        "markup": 0.10, "exclude_exact": False,
        "orders_col": ORDERS_SBSD, "acos_col": ACOS,
        "headers": ["Campaign Name", "Targeting", "Match Type", "Clicks", "",
                    "Cost Per Click (CPC)", "Spend", "14 Day Total Sales", "ACOS", "Orders"],
        "map": {"Campaign Name": "Campaign Name", "Targeting": "Targeting",
                "Match Type": "Match Type", "Clicks": "Clicks",
                "Cost Per Click (CPC)": "Cost Per Click (CPC)", "Spend": "Spend",
                "14 Day Total Sales": SALES_SBSD, "ACOS": ACOS, "Orders": ORDERS_SBSD},
    },
    {
        "file_key": "sd_tgt", "section": "Sponsored Display", "sub": "Targeting",
        "markup": 0.10, "exclude_exact": False,
        "orders_col": ORDERS_SBSD, "acos_col": ACOS,
        "headers": ["Campaign Name", "Ad Group Name", "Targeting", "Clicks", "",
                    "Spend", "Cost Per Click (CPC)", "14 Day Total Sales", "ACOS", "Orders"],
        "map": {"Campaign Name": "Campaign Name", "Ad Group Name": "Ad Group Name",
                "Targeting": "Targeting", "Clicks": "Clicks", "Spend": "Spend",
                "Cost Per Click (CPC)": "Cost Per Click (CPC)",
                "14 Day Total Sales": SALES_SBSD, "ACOS": ACOS, "Orders": ORDERS_SBSD},
    },
]


def _filter(df, block, threshold):
    if df is None or df.empty:
        return df.iloc[0:0] if df is not None else pd.DataFrame()
    oc, ac = _col(df, block["orders_col"]), _col(df, block["acos_col"])
    if oc is None or ac is None:
        return df.iloc[0:0]
    d = df.copy()
    orders = pd.to_numeric(d[oc], errors="coerce").fillna(0)
    acos = pd.to_numeric(d[ac], errors="coerce")
    mask = (orders >= MIN_ORDERS) & (orders <= MAX_ORDERS) & (acos >= threshold)
    if block.get("exclude_exact"):
        mt = _col(df, "Match Type")
        if mt is not None:
            mask &= d[mt].astype(str).str.strip().str.upper() != "EXACT"
    return d[mask]


def _rows(df, block, threshold):
    hits = _filter(df, block, threshold)
    out = []
    for _, srow in hits.iterrows():
        row = {}
        for header in block["headers"]:
            if header == "":
                continue
            src = block["map"].get(header)
            val = _get(srow, src) if src else ""
            row[header] = val if header.lower() in TEXT_HEADERS else _num(val)
        out.append(row)
    return out


def generate(files, params):
    """files: dict of file-like objects keyed sp_search, sp_target, sb_kw, sd_tgt
    (any subset may be present/None). params: {'gross_margin': percent}.
    Returns (summary, workbook). Only blocks whose file was provided appear.
    """
    break_even = float(params.get("gross_margin", 35)) / 100.0

    sections, summary = [], []
    for block in BLOCKS:
        f = files.get(block["file_key"])
        if f is None:
            continue
        df = pd.read_excel(f)
        rows = _rows(df, block, break_even + block["markup"])
        sections.append({"section": block["section"], "sub": block["sub"],
                         "headers": block["headers"], "rows": rows})
        summary.append((f'{block["section"]} - {block["sub"]}', len(rows)))

    wb = build_report(sections)
    return summary, wb


def describe(params):
    be = float(params.get("gross_margin", 35))
    return (f"1–{MAX_ORDERS} orders. ACOS ≥ {be + 10:.0f}% "
            f"(Brands/Display/SP search) or ≥ {be + 20:.0f}% (SP targeting), "
            f"from your {be:.0f}% break-even.")
