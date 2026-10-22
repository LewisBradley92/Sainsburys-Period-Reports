# Sainsbury's / Argos Periodic Reporting App

## What it does
- Generates Argos, Supermarket and Convenience LIVE workbooks from the uploaded templates.
- Uses order `order_info` to allocate orders to Argos, Supermarket or Convenience.
- Uses audit visit dates to select the reporting period.
- Excludes Compliance Audit, Date Check, Mystery Shop and Petrol/Diesel.
- Refreshes `Period Input`, merges `YTD Input` by audit `internal_id`, and refreshes the `Regions` tab from the master store databases.
- Produces three LIVE workbooks and a ZIP containing those three files.

## Run
Python 3.10+ recommended.

```bash
pip install -r requirements.txt
streamlit run app.py
```

The app defaults to Fast mode, which skips cloud formula recalculation. Open the generated LIVE workbooks in a compatible spreadsheet engine and confirm the Summary and Performance tabs calculate correctly before relying on the figures. Fast mode is not a guarantee that cached formula results are up to date.

## Important
- Confirm period start/end dates against the uploaded Sainsbury's calendar.
- The Q26 helper value is populated from the audit export's `primary_result`; confirm that this field is the intended business definition for Q26 before using the output for reporting.
- This version preserves array-formula prototypes when filling input rows and retains historical Q26 values when merging YTD.
