# Sainsbury's / Argos / Convenience Period Reporting App

Streamlit app that updates the three LIVE reporting workbooks for Argos, Supermarket and Convenience.

## Important Q26 pass/fail fix
The `Basic No ID logic (Q26)` column in Period Input and YTD Input is now populated per audit from the source audit export's `primary_result` (`pass`, `fail`, or `abort`) instead of copying the template's first-row literal `pass` marker to every row. This prevents failed audits from being incorrectly counted as passes.

## Run
Deploy `app.py` on Streamlit Community Cloud with `requirements.txt` and `packages.txt`. Fast mode skips cloud formula recalculation; open LIVE workbooks in Excel to recalculate formulas before using report figures.
