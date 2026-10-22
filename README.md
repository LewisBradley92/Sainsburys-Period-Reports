# Sainsbury's / Argos Periodic Reporting App

## What it does
- Generates Argos, Supermarket and Convenience reports from the three LIVE templates.
- Uses `order_export.csv` `order_info` prefixes such as `P8 - Argos - Knife`, `P8 - Supermarket - Tobacco`, and `P8 - Convenience - E-cig` to allocate orders.
- Excludes Compliance Audit, Date Check, Mystery Shop and Petrol/Diesel.
- Refreshes `Period Input`, appends/updates `YTD Input` by audit `internal_id`, keeping the latest date of visit for duplicate audit IDs.
- Produces LIVE and client-facing versions of each report, plus a ZIP containing all six.

## Run locally
Python 3.10+ recommended.

```bash
pip install -r requirements.txt
streamlit run app.py
```

LibreOffice Calc is required for reliable formula recalculation before converting formulas to values in client-facing workbooks. The `packages.txt` file installs `libreoffice-calc` on Streamlit Community Cloud. Commit `packages.txt` and redeploy; allow the first build to finish installing the system package.

## Important
The reporting period dates are entered manually in the UI for each run. Confirm the dates against the uploaded Sainsbury's calendar. This avoids assuming a calendar structure if the calendar layout changes.
