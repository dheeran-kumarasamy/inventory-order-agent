import io
import re
from datetime import date

import pandas as pd
from openpyxl import Workbook

from sheets_loader import load_leadtime_from_sheet, load_sales_from_sheet, load_stock_from_sheet

TYPE_ABBREV = {
    "CENTRE BASS": "CB",
    "SOLID": "SOLID",
    "HOLLOW": "HOLLOW",
    "HALF SOLID": "HALF SOLID",
    "LIGHT": "LIGHT",
}


def calc_avg_sales(sales: pd.DataFrame) -> pd.DataFrame:
    grp = sales.groupby("ProductName")
    avg = (grp["Total"].sum() / grp["Month"].nunique()).reset_index()
    avg.columns = ["ProductName", "AvgMonthlySales"]
    avg["AvgDailySales"] = (avg["AvgMonthlySales"] / 30).round(3)
    return avg


def extract_parts(name: str):
    parts = re.split(r"\s*-\s*", name, maxsplit=1)
    prefix = parts[0].strip()
    suffix = (parts[1] if len(parts) > 1 else "").upper()
    tokens = [t.strip() for t in re.split(r"\s*[Xx]\s*", prefix)]
    od = tokens[0] if len(tokens) > 0 else ""
    grooves = tokens[1] if len(tokens) > 1 else ""
    ptype = next((k for k in TYPE_ABBREV if k in suffix), None)
    return od, grooves, ptype


def find_rc(od, grooves, ptype, rc_lookup):
    abbrev = TYPE_ABBREV.get(ptype, "")
    candidate = f"{od} X {grooves} - {abbrev} - RC".upper() if abbrev else None
    if candidate and candidate in rc_lookup:
        return rc_lookup[candidate]
    prefix = f"{od} X {grooves}".upper()
    matches = [v for k, v in rc_lookup.items() if k.startswith(prefix) and "RC" in k]
    return matches[0] if matches else None


def _build_excel(mach_out: pd.DataFrame, mfg_out: pd.DataFrame) -> io.BytesIO:
    wb = Workbook()
    wb.remove(wb.active)

    ws_sum = wb.create_sheet("Summary")
    ws_sum.append(["Metric", "Value"])
    ws_sum.append(["Machining Items", len(mach_out)])
    ws_sum.append(["Manufacturing Items", len(mfg_out)])
    ws_sum.append(["Machining Units", int(mach_out["Suggested Order Qty"].sum()) if len(mach_out) else 0])
    ws_sum.append(["Manufacturing Units", int(mfg_out["Suggested Order Qty"].sum()) if len(mfg_out) else 0])

    ws_mach = wb.create_sheet("Machining Orders")
    ws_mach.append(list(mach_out.columns))
    for _, row in mach_out.iterrows():
        ws_mach.append(list(row.values))

    ws_mfg = wb.create_sheet("Manufacturing Orders")
    ws_mfg.append(list(mfg_out.columns))
    for _, row in mfg_out.iterrows():
        ws_mfg.append(list(row.values))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def generate_report(master_sheet_url, lt_url, mach_lead, mfg_lead):
    today = date.today()

    stock = load_stock_from_sheet(master_sheet_url)
    sales = load_sales_from_sheet(master_sheet_url)
    _ = load_leadtime_from_sheet(lt_url) if lt_url else None
    avg = calc_avg_sales(sales)

    below = stock[stock["StockLevel"] < stock["ReorderLevel"]].copy()
    rc_mask = below["ProductName"].str.contains(r"\bRC\b", case=False, na=False)
    mfg_df = below[rc_mask].copy()
    mach_df = below[~rc_mask].copy()

    all_rc = stock[stock["ProductName"].str.contains(r"\bRC\b", case=False, na=False)]
    rc_lookup = {n.upper(): n for n in all_rc["ProductName"]}

    def enrich(df, lead):
        rows = []
        for _, row in df.iterrows():
            pname = row["ProductName"]
            stk = int(row["StockLevel"])
            reorder = int(row["ReorderLevel"])
            shortage = max(reorder - stk, 0)
            sm = avg[avg["ProductName"].str.upper() == pname.upper()]
            adaily = round(sm["AvgDailySales"].values[0], 3) if len(sm) else 0.0
            amonthly = round(sm["AvgMonthlySales"].values[0], 1) if len(sm) else 0.0
            suggested = max(int(adaily * lead) + shortage, 1)
            rows.append(
                {
                    "pname": pname,
                    "stk": stk,
                    "reorder": reorder,
                    "shortage": shortage,
                    "adaily": adaily,
                    "amonthly": amonthly,
                    "suggested": suggested,
                }
            )
        return rows

    mach_rows = []
    for r in enrich(mach_df, mach_lead):
        pname = r["pname"]
        is_vp = "V-PULLEY" in pname.upper()
        rc = None
        if is_vp:
            od, grooves, ptype = extract_parts(pname)
            rc = find_rc(od, grooves, ptype, rc_lookup)
        mach_rows.append(
            {
                "Product Name": pname,
                "RC Required": rc if rc else ("No RC Found" if is_vp else "N/A"),
                "Current Stock": r["stk"],
                "Reorder Level": r["reorder"],
                "Shortage": r["shortage"],
                "Avg Monthly Sales": r["amonthly"],
                "Avg Daily Sales": r["adaily"],
                "Machining Lead Time (Days)": mach_lead,
                "Suggested Order Qty": r["suggested"],
            }
        )

    mfg_rows = []
    for r in enrich(mfg_df, mfg_lead):
        mfg_rows.append(
            {
                "RC Product Name": r["pname"],
                "Current Stock": r["stk"],
                "Reorder Level": r["reorder"],
                "Shortage": r["shortage"],
                "Avg Monthly Sales": r["amonthly"],
                "Avg Daily Sales": r["adaily"],
                "Manufacturing Lead Time (Days)": mfg_lead,
                "Suggested Order Qty": r["suggested"],
            }
        )

    mach_out = pd.DataFrame(mach_rows).sort_values("Shortage", ascending=False)
    mfg_out = pd.DataFrame(mfg_rows).sort_values("Shortage", ascending=False)
    buf = _build_excel(mach_out, mfg_out)
    return buf, mach_out, mfg_out, today
