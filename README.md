# Amazon VA Tools

Automates the reports currently produced by hand from Amazon Seller Central
bulk exports. First tool: the **Bleeders Report**. Two more to be added.

A *bleeder* is an enabled keyword, target, or customer search term that got
**more than 10 clicks but generated $0 in sales** — money spent with nothing to
show for it. The tool finds every bleeder across Sponsored Products, Brands, and
Display and drops them into the report layout, with an empty blue column left
for the operator to mark **pause** vs. **negative-exact**.

## Why this is more reliable than doing it by hand

Columns are matched by **header name** ("Clicks", "Sales", "State"), not by
column letter. Amazon reorders and adds columns between exports — the manual SOP
already drifted out of sync (it says Clicks is column AM; in current exports it
is AS). Header-name matching keeps working regardless. Same input always
produces the same output.

## Setup (one time)

Requires Python 3.10+.

```bash
git clone <your-repo-url>
cd amazon-va-tools
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Run the app (the normal way)

```bash
streamlit run app.py
```

A browser tab opens. Drag your bulk export on, review the counts, click
**Download report**. No command line after launch.

### Getting the bulk export from Amazon

Campaign Manager → **Bulk operations** → date range **60 days** → check
Sponsored Products / Brands / Display + the two Search-term boxes →
**Create spreadsheet for download**.

## Run from the command line (optional)

```bash
python cli.py path/to/bulk-export.xlsx -o Bleeders_Report.xlsx
```

## The rules, in one place

All logic lives in `bleeders/engine.py`:

- `MIN_CLICKS = 10` and `MAX_SALES = 0.0` — the bleeder thresholds.
- `BLOCKS` — one entry per report section, with the exact template headers and
  which rows qualify. Change a section by editing its entry; no other code moves.

## Project layout

```
amazon-va-tools/
├── app.py              # Streamlit UI (registry of tools at the top)
├── cli.py              # command-line entry point
├── bleeders/
│   ├── __init__.py
│   └── engine.py       # all Bleeders Report logic + config
├── requirements.txt
└── sample_data/        # local test files (git-ignored)
```

## Adding tools 2 and 3 later

Each tool becomes its own package next to `bleeders/` (e.g. `tool_two/engine.py`)
exposing a `generate(source)` that returns `(summary, workbook)`. Register it in
the `TOOLS` dict at the top of `app.py` and it appears in the dropdown — the UI
code does not change. Build the abstraction only when the second tool's spec is
known; don't generalize ahead of it.

## Known items to confirm against live data

- **Sponsored Brands & Display** blocks were built from the SOP but not yet
  verified against an export where those campaigns had spend (they were empty in
  the first file). Verify row-for-row once such an export exists.
- **Sponsored Display** currently keeps every entity except the ad-group row, per
  the SOP's "uncheck Ad group." Confirm that matches intent when SD has bleeders.

## Note on data

Real exports and generated reports are git-ignored. Do not commit client data.
