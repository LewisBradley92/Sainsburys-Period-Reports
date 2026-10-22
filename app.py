import io, os, re, shutil, subprocess, tempfile, zipfile
from datetime import datetime, date
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.formula.translate import Translator
from copy import copy

st.set_page_config(page_title="Sainsbury's / Argos Reporting", layout="wide")
st.title("Sainsbury's / Argos Periodic Reporting")
st.caption("Build Argos, Supermarket and Convenience LIVE and client-ready workbooks.")

REPORTS = {
    "Argos": {"template": "argos", "category": "Argos", "suffix": "Argos"},
    "Supermarket": {"template": "super", "category": "Supermarket", "suffix": "Super"},
    "Convenience": {"template": "conv", "category": "Convenience", "suffix": "Conv"},
}
EXCLUDED_ITO = {"compliance audit", "date check", "mystery shop", "petrol/diesel"}


def sync_regions_tab(wb, report_name, master_upload, argos_master_upload):
    """Refresh the report's Regions tab A:D from the appropriate master store database."""
    if "Regions" not in wb.sheetnames:
        raise ValueError(f"{report_name} workbook has no 'Regions' tab.")
    if report_name == "Argos":
        # Argos must use its own store database, not the Sainsbury's master.
        argos_master_upload.seek(0)
        src_wb = openpyxl.load_workbook(argos_master_upload, data_only=True, read_only=True)
        expected_sheet = "Argos Store Listing - Argos & S"
        if expected_sheet in src_wb.sheetnames:
            src_ws = src_wb[expected_sheet]
        else:
            # Tolerate minor tab-name differences between database versions.
            candidates = [name for name in src_wb.sheetnames
                          if "argos" in name.casefold() and "store" in name.casefold()]
            if not candidates:
                available = ", ".join(src_wb.sheetnames)
                src_wb.close()
                raise ValueError(
                    f"Could not find the Argos store-listing tab in the Argos database. "
                    f"Available tabs: {available}"
                )
            src_ws = src_wb[candidates[0]]
        # Regions A:D = Store Code, Region, District, Store Name
        records = []
        for row in src_ws.iter_rows(min_row=2, values_only=True):
            store_code = row[0]   # A
            store_name = row[2]   # C
            region = row[6]       # G
            district = row[7]     # H
            if store_code is not None:
                records.append((store_code, region, district, store_name))
    else:
        master_upload.seek(0)
        src_wb = openpyxl.load_workbook(master_upload, data_only=True, read_only=True)
        sheet_name = "Supermarket" if report_name == "Supermarket" else "Convenience"
        src_ws = src_wb[sheet_name]
        records = []
        for row in src_ws.iter_rows(min_row=2, values_only=True):
            if report_name == "Supermarket":
                # Regions A:D = Region, Store Code, Store Name, Zone
                store_code, store_name, zone, region = row[2], row[3], row[9], row[10]
                rec = (region, store_code, store_name, zone)
            else:
                # Regions A:D = Zone, Region, Store Code, Store Name
                store_code, store_name, zone, region = row[2], row[3], row[9], row[10]
                rec = (zone, region, store_code, store_name)
            if store_code is not None:
                records.append(rec)
    src_wb.close()

    # Keep first row headers and replace the Regions data area with the current master listing.
    ws = wb["Regions"]
    for r in range(2, ws.max_row + 1):
        for c in range(1, 5):
            ws.cell(r, c).value = None
    # De-duplicate on report store code while preserving source order.
    deduped, seen = [], set()
    for rec in records:
        code = rec[0] if report_name == "Argos" else (rec[1] if report_name == "Supermarket" else rec[2])
        key = norm(code)
        if key and key not in seen:
            seen.add(key)
            deduped.append(rec)
    for r, rec in enumerate(deduped, start=2):
        for c, val in enumerate(rec, start=1):
            code_col = 1 if report_name == "Argos" else (2 if report_name == "Supermarket" else 3)
            ws.cell(r, c).value = numeric_site_code(val) if c == code_col else val
        if r > 2:
            for c in range(1, 5):
                ws.cell(r,c)._style = copy(ws.cell(2,c)._style)
    ws.auto_filter.ref = f"A1:D{max(1, len(deduped)+1)}"
    return len(deduped)

def read_csv_upload(upload):
    upload.seek(0)
    return pd.read_csv(upload, low_memory=False, dtype={"order_internal_id": str, "internal_id": str})

def norm(v):
    return re.sub(r"\s+", " ", str(v or "")).strip().casefold()

def numeric_site_code(value):
    """Store numeric-looking site codes as numbers for Excel lookups."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return int(value) if isinstance(value, float) and value.is_integer() else value
    text = str(value).strip()
    if re.fullmatch(r"\d+", text) and (text == "0" or not text.startswith("0")):
        return int(text)
    if re.fullmatch(r"\d+\.0+", text) and not text.startswith("0"):
        return int(float(text))
    return value

def clean_ito(value):
    v = norm(value)
    if "alcohol" in v: return "Alcohol"
    if "e-cig" in v or "ecig" in v: return "E-Cig"
    if "tobacco" in v or "cigarette" in v: return "Tobacco"
    if "knife" in v or "bladed" in v: return "Knife"
    if "online" in v or "click and collect" in v: return "Online"
    return str(value or "").strip()

def order_period_and_category(order_info):
    s = str(order_info or "")
    m = re.match(r"^\s*P\s*(\d+)\s*-\s*(Argos|Supermarket|Convenience)\b", s, flags=re.I)
    if not m:
        return None, None
    return int(m.group(1)), m.group(2).title()

def date_series(df, col):
    if col not in df.columns:
        return pd.Series(pd.NaT, index=df.index)
    return pd.to_datetime(df[col], errors="coerce", dayfirst=True)

def find_formula_row(ws, start_col=15):
    for r in range(2, min(ws.max_row, 15) + 1):
        if any(isinstance(ws.cell(r,c).value, str) and ws.cell(r,c).value.startswith("=") for c in range(start_col, ws.max_column+1)):
            return r
    return 2

def clear_data_rows(ws, start_row=2):
    # Clear cell contents but retain formatting and column positions.
    for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.value = None

def write_input_rows(ws, df, header_width=14):
    # Capture formula prototypes before clearing existing data.
    formula_row = find_formula_row(ws, header_width+1)
    formula_prototypes = {}
    for col in range(header_width+1, ws.max_column+1):
        value = ws.cell(formula_row, col).value
        if isinstance(value, str) and value.startswith("="):
            formula_prototypes[col] = value
    # Clear existing data and retain template styles.
    clear_data_rows(ws, 2)
    for idx, (_, rec) in enumerate(df.iterrows(), start=2):
        for col in range(1, header_width+1):
            value = rec.iloc[col-1] if col-1 < len(rec) else None
            if pd.isna(value) if not isinstance(value, (list, dict)) else False:
                value = None
            if isinstance(value, pd.Timestamp):
                value = value.to_pydatetime()
            ws.cell(idx, col).value = value
        # Extend existing helper formulas to the new row without changing column layout.
        if idx != formula_row:
            for col in range(header_width+1, ws.max_column+1):
                src_formula = formula_prototypes.get(col)
                if src_formula:
                    try:
                        ws.cell(idx, col).value = Translator(src_formula, origin=ws.cell(formula_row,col).coordinate).translate_formula(ws.cell(idx,col).coordinate)
                    except Exception:
                        ws.cell(idx, col).value = src_formula
                else:
                    ws.cell(idx,col)._style = copy(ws.cell(formula_row,col)._style)
        # Copy row style from template row 2 across the input row.
        if idx > 2:
            for col in range(1, ws.max_column+1):
                ws.cell(idx,col)._style = copy(ws.cell(2,col)._style)
                if ws.cell(idx,col).value is None:
                    pass
    # remove stale rows below final record; cells cleared above already.
    ws.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(ws.max_column)}{max(2, len(df)+1)}"

def rows_from_source(audit_df, order_df, period_num, report_name, start_date, end_date):
    audits = audit_df.copy()
    if "order_internal_id" not in audits or "internal_id" not in audits:
        raise ValueError("Audit export must contain order_internal_id and internal_id.")
    if "internal_id" not in order_df or "order_info" not in order_df:
        raise ValueError("Order export must contain internal_id and order_info.")
    orders = order_df.copy()
    orders["__period_category"] = orders["order_info"].apply(order_period_and_category)
    orders["__period"] = orders["__period_category"].apply(lambda x: x[0])
    orders["__category"] = orders["__period_category"].apply(lambda x: x[1])
    orders["__ito_norm"] = orders.get("item_to_order", pd.Series("", index=orders.index)).map(norm)
    orders = orders[~orders["__ito_norm"].apply(lambda x: any(ex in x for ex in EXCLUDED_ITO))]
    # order_info determines report allocation only; audit visit date determines reporting period.
    orders = orders[orders["__category"] == report_name]
    # Use the order export as the allocation authority.
    order_cols = ["internal_id", "order_info", "item_to_order"]
    orders = orders[order_cols].drop_duplicates("internal_id", keep="last").rename(
        columns={"internal_id": "__matched_order_id", "item_to_order": "__order_ito"})
    audits = audits.merge(orders, left_on="order_internal_id", right_on="__matched_order_id", how="inner")
    if "__order_ito" in audits:
        if "item_to_order" in audits:
            audits["item_to_order"] = audits["__order_ito"].combine_first(audits["item_to_order"])
        else:
            audits["item_to_order"] = audits["__order_ito"]
    audits = audits.drop(columns=[c for c in ["__matched_order_id", "order_info", "__order_ito"] if c in audits.columns], errors="ignore")
    audits["__visit"] = date_series(audits, "date_of_visit")
    start = pd.Timestamp(start_date)
    end = pd.Timestamp(end_date)
    audits = audits[(audits["__visit"] >= start) & (audits["__visit"] <= end)]
    # Global exclusions based on both order ITO and audit category.
    if "item_to_order" in audits:
        audits = audits[~audits["item_to_order"].map(norm).apply(lambda x: any(ex in x for ex in EXCLUDED_ITO))]
    # Map the first 14 fixed template columns, preserving their positions.
    columns = ["order_internal_id","client_name","internal_id","site_internal_id","site_name",
               "site_address_1","site_address_2","site_address_3","site_post_code",
               "submitted_date","approval_date","item_to_order","date_of_visit","site_code"]
    out = pd.DataFrame()
    for c in columns:
        out[c] = audits[c] if c in audits.columns else None
    for c in ["submitted_date","approval_date","date_of_visit"]:
        out[c] = pd.to_datetime(out[c], errors="coerce", dayfirst=True)
    out["site_code"] = out["site_code"].apply(numeric_site_code)
    return out.drop_duplicates("internal_id", keep="last").reset_index(drop=True)

def merge_ytd(existing_df, incoming_df):
    if existing_df.empty:
        return incoming_df.copy()
    # Match on audit internal_id. If the same audit appears again, keep the row with latest visit date.
    all_rows = pd.concat([existing_df, incoming_df], ignore_index=True)
    all_rows["__visit_sort"] = pd.to_datetime(all_rows.get("date_of_visit"), errors="coerce", dayfirst=True)
    all_rows["__order"] = range(len(all_rows))
    all_rows = all_rows.sort_values(["__visit_sort","__order"], na_position="first")
    all_rows = all_rows.drop_duplicates(subset=["internal_id"], keep="last")
    return all_rows.drop(columns=["__visit_sort","__order"], errors="ignore").reset_index(drop=True)

def extract_existing_ytd(ws):
    headers = [ws.cell(1,c).value for c in range(1,15)]
    rows=[]
    for r in range(2, ws.max_row+1):
        vals=[ws.cell(r,c).value for c in range(1,15)]
        if any(v is not None for v in vals) and vals[2] is not None:
            rows.append(vals)
    return pd.DataFrame(rows, columns=headers)


def update_period_headers(wb, period_num):
    """Update static period labels on report/output tabs without touching formulas or input tabs."""
    pattern = re.compile(r"(?i)\bPeriod\s+\d+\b")
    input_tabs = {"Input", "Regions", "Period Input", "YTD Input"}
    updated = 0
    for ws in wb.worksheets:
        if ws.title in input_tabs:
            continue
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and not value.startswith("=") and pattern.search(value):
                    new_value = pattern.sub(f"Period {int(period_num)}", value)
                    if new_value != value:
                        cell.value = new_value
                        updated += 1
    return updated

def recalc_with_libreoffice(path):
    """Recalculate an XLSX in-place via a temporary LibreOffice round-trip."""
    soffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not soffice:
        return False, "LibreOffice is not installed. Add packages.txt with libreoffice-calc and redeploy."
    try:
        with tempfile.TemporaryDirectory() as tmp:
            profile = os.path.join(tmp, "lo-profile")
            outdir = os.path.join(tmp, "converted")
            os.makedirs(outdir, exist_ok=True)
            # Isolate LibreOffice profile to avoid locked-profile errors in hosted environments.
            cmd = [
                soffice, "-env:UserInstallation=file://" + profile,
                "--headless", "--convert-to", "xlsx",
                "--outdir", outdir, path
            ]
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=240)
            converted = os.path.join(outdir, os.path.basename(path))
            if proc.returncode == 0 and os.path.exists(converted) and os.path.getsize(converted) > 0:
                # Ensure output is a readable workbook before replacing the original.
                test_wb = openpyxl.load_workbook(converted, data_only=False, read_only=True)
                test_wb.close()
                shutil.copy2(converted, path)
                # Verify calculation produced cached values for formulas on Summary onwards.
                value_wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
                formula_wb = openpyxl.load_workbook(path, data_only=False, read_only=True)
                checked = 0
                missing = 0
                for ws in formula_wb.worksheets:
                    if ws.title in {"Input", "Regions", "Period Input", "YTD Input"}:
                        continue
                    wsv = value_wb[ws.title]
                    for row in ws.iter_rows():
                        for cell in row:
                            if isinstance(cell.value, str) and cell.value.startswith("="):
                                checked += 1
                                cached = wsv[cell.coordinate].value
                                if cached is None or (isinstance(cached, str) and cached.startswith("#")):
                                    missing += 1
                formula_wb.close(); value_wb.close()
                if checked and missing > max(3, int(checked * 0.02)):
                    return False, f"LibreOffice converted the workbook but {missing} of {checked} formula cells still have blank/error cached results."
                return True, proc.stdout + proc.stderr
            return False, (proc.stdout + proc.stderr or f"LibreOffice returned code {proc.returncode} without an output workbook.")
    except Exception as e:
        return False, str(e)

def make_client_copy(live_path, client_path, keep_sheets, freeze_values=True):
    # Only freeze formulas after a calculation engine has successfully recalculated the LIVE file.
    wb_formula = openpyxl.load_workbook(live_path, data_only=False)
    if freeze_values:
        wb_values = openpyxl.load_workbook(live_path, data_only=True)
    else:
        wb_values = None
    for sheet in list(wb_formula.sheetnames):
        if sheet not in keep_sheets:
            del wb_formula[sheet]
    if wb_values is not None:
        for sn in keep_sheets:
            if sn not in wb_formula.sheetnames or sn not in wb_values.sheetnames:
                continue
            wsf, wsv = wb_formula[sn], wb_values[sn]
            for row in wsf.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        cached = wsv[cell.coordinate].value
                        # Never replace a formula with an empty/error cache.
                        if cached is not None and not (isinstance(cached, str) and cached.startswith("#")):
                            cell.value = cached
    wb_formula.calculation.fullCalcOnLoad = True
    wb_formula.calculation.forceFullCalc = True
    wb_formula.calculation.calcMode = "auto"
    wb_formula.save(client_path)

def update_input_orders(ws, orders):
    """Append new order records to columns A:C without overwriting existing Input data."""
    marker_row = next((r for r in range(1, ws.max_row + 1)
                       if norm(ws.cell(r, 1).value) == "input"), None)
    if marker_row is None:
        raise ValueError(f"Could not find the 'Input' marker in column A of {ws.title}.")
    start_row = marker_row + 1
    records = orders.copy()
    if "internal_id" not in records.columns:
        raise ValueError("Order export is missing internal_id.")
    records = records.dropna(subset=["internal_id"]).drop_duplicates("internal_id", keep="last")
    records["__type"] = records.get("schedule_type", pd.Series("", index=records.index)).map(
        lambda x: "Revisit" if "revisit" in norm(x) else "Standard")
    records["__ito"] = records.get("item_to_order", pd.Series("", index=records.index)).map(clean_ito)

    # Preserve existing order rows. Add only order IDs that are not already listed.
    existing_ids = set()
    last_used = start_row - 1
    for r in range(start_row, ws.max_row + 1):
        current = ws.cell(r, 1).value
        if current is not None and str(current).strip():
            existing_ids.add(str(current).strip().casefold())
            last_used = r
    new_records = records[~records["internal_id"].astype(str).str.strip().str.casefold().isin(existing_ids)]
    write_row = max(start_row, last_used + 1)
    template_row = start_row
    for i, (_, rec) in enumerate(new_records.iterrows(), start=write_row):
        for c in range(1, ws.max_column + 1):
            ws.cell(i, c)._style = copy(ws.cell(template_row, c)._style)
            src = ws.cell(template_row, c)
            if c >= 4 and isinstance(src.value, str) and src.value.startswith("="):
                try:
                    ws.cell(i, c).value = Translator(src.value, origin=src.coordinate).translate_formula(ws.cell(i, c).coordinate)
                except Exception:
                    ws.cell(i, c).value = src.value
        ws.cell(i, 1).value = str(rec["internal_id"]).strip()
        ws.cell(i, 2).value = rec["__type"]
        ws.cell(i, 3).value = rec["__ito"]
    return len(new_records)

st.markdown("### Upload the source files and master store databases")
col1,col2,col3=st.columns(3)
with col1:
    argos_file=st.file_uploader("Argos LIVE workbook", type=["xlsx"], key="argos")
    super_file=st.file_uploader("Supermarket LIVE workbook", type=["xlsx"], key="super")
with col2:
    conv_file=st.file_uploader("Convenience LIVE workbook", type=["xlsx"], key="conv")
    audit_file=st.file_uploader("Audit export CSV", type=["csv"], key="audit")
with col3:
    order_file=st.file_uploader("Order export CSV", type=["csv"], key="orders")
    calendar_file=st.file_uploader("Sainsbury's calendar", type=["xlsx"], key="calendar")
col4, col5 = st.columns(2)
with col4:
    sains_master_file = st.file_uploader("Sainsbury's Master Store Database", type=["xlsx"], key="sains_master")
with col5:
    argos_master_file = st.file_uploader("Argos Store Database", type=["xlsx"], key="argos_master")

st.markdown("### Reporting period")
pcol1,pcol2,pcol3=st.columns(3)
with pcol1:
    period_num=st.number_input("Period number", min_value=1, max_value=13, value=8, step=1)
with pcol2:
    start_date=st.date_input("Period start date", value=date(2026,10,1))
with pcol3:
    end_date=st.date_input("Period end date", value=date(2026,10,31))

if start_date > end_date:
    st.error("Period start date must be on or before the end date.")

if st.button("Build all 3 reports (6 files)", type="primary", disabled=not all([argos_file,super_file,conv_file,audit_file,order_file,calendar_file,sains_master_file,argos_master_file]) or start_date>end_date):
    try:
        audit_df=read_csv_upload(audit_file)
        order_df=read_csv_upload(order_file)
        order_df["__allocation"] = order_df["order_info"].apply(order_period_and_category)
        order_df["__period"] = order_df["__allocation"].apply(lambda x:x[0])
        order_df["__category"] = order_df["__allocation"].apply(lambda x:x[1])
        order_df["__ito_norm"] = order_df.get("item_to_order", pd.Series("", index=order_df.index)).map(norm)
        valid_orders=order_df[(~order_df["__ito_norm"].apply(lambda x: any(ex in x for ex in EXCLUDED_ITO))) & order_df["__category"].notna()].copy()
        # Build the current-period order list from audits whose visit dates fall in the selected period.
        audit_dates = date_series(audit_df, "date_of_visit")
        in_period_audits = audit_df[(audit_dates >= pd.Timestamp(start_date)) & (audit_dates <= pd.Timestamp(end_date))]
        current_period_order_ids = set(in_period_audits["order_internal_id"].dropna().astype(str).str.strip())
        valid_orders = valid_orders[valid_orders["internal_id"].astype(str).str.strip().isin(current_period_order_ids)]
        st.info(f"Source audit rows: {len(audit_df):,}. Orders matched to audits visited in P{int(period_num)}: {len(valid_orders):,}.")
        templates={"Argos":argos_file,"Supermarket":super_file,"Convenience":conv_file}
        keep_sheets={
            "Argos":["Period Input","YTD Input","Summary","Zone Performance","Region 1 Performance","Region 2 Performance","Region 3 Performance","Region 4 Performance"],
            "Supermarket":["Period Input","YTD Input","Summary","Zone Performance","Zone North Performance","Zone Central Performance","Zone South Performance"],
            "Convenience":["Period Input","YTD Input","Summary","Zone Performance","Zone Central Performance","Zone North Performance","Zone South Performance"],
        }
        zipbuf=io.BytesIO()
        results=[]
        with tempfile.TemporaryDirectory() as td:
            for report_name, upload in templates.items():
                period_rows=rows_from_source(audit_df, order_df, int(period_num), report_name, start_date, end_date)
                wb=openpyxl.load_workbook(upload)
                region_count = sync_regions_tab(wb, report_name, sains_master_file, argos_master_file)
                period_ws=wb["Period Input"]; ytd_ws=wb["YTD Input"]
                existing=extract_existing_ytd(ytd_ws)
                new_ytd=merge_ytd(existing,period_rows)
                write_input_rows(period_ws,period_rows,14)
                write_input_rows(ytd_ws,new_ytd,14)
                # Refresh order lists if a compatible three-column list exists.
                report_orders=valid_orders[valid_orders["__category"]==report_name]
                input_order_count = update_input_orders(wb["Input"], report_orders)
                header_count = update_period_headers(wb, int(period_num))
                live_name=f"Measure 25 Executive Report {REPORTS[report_name]['suffix']} - Period {int(period_num)} - LIVE.xlsx"
                client_name=f"Measure 25 Executive Report {REPORTS[report_name]['suffix']} - Period {int(period_num)}.xlsx"
                live_path=os.path.join(td,live_name); client_path=os.path.join(td,client_name)
                wb.calculation.fullCalcOnLoad=True
                wb.calculation.forceFullCalc=True
                wb.calculation.calcMode="auto"
                wb.save(live_path)
                ok, msg=recalc_with_libreoffice(live_path)
                if not ok:
                    st.warning(f"{report_name}: Excel formulas were not recalculated by LibreOffice. Client copy will use cached values if available. {msg}")
                make_client_copy(live_path,client_path,keep_sheets[report_name],freeze_values=ok)
                if not ok:
                    st.warning(f"{report_name}: client copy retains formulas because recalculation was unavailable. This avoids exporting blank or #CALC! values. Open the workbook in Excel to calculate it; install/use LibreOffice in the deployment environment to produce a fully value-only client copy.")
                results.extend([(live_name,live_path),(client_name,client_path)])
                st.write(f"**{report_name}:** {len(period_rows):,} period audits; {len(new_ytd):,} YTD records; {region_count:,} store records refreshed on Regions tab; {input_order_count:,} orders written to Input; {header_count:,} period headers updated.")
            with zipfile.ZipFile(zipbuf,"w",zipfile.ZIP_DEFLATED) as zf:
                for name,path in results: zf.write(path,arcname=name)
            zipbuf.seek(0)
            st.success("All six workbooks have been generated.")
            st.download_button("Download all 6 workbooks (ZIP)", data=zipbuf.getvalue(),
                               file_name=f"Sainsburys_Argos_Period_{int(period_num)}_Reports.zip",
                               mime="application/zip")
            for name,path in results:
                with open(path,"rb") as f:
                    st.download_button(f"Download {name}", data=f.read(), file_name=name,
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       key="download_"+name)
    except Exception as e:
        st.exception(e)
