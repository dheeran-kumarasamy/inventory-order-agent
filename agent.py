import io
import math
import re
from datetime import date

import pandas as pd
from fpdf import FPDF
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

HIGHLIGHT_FILL = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=10)
ALT_ROW_FILL = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
CENTER = Alignment(horizontal="center", vertical="center")
MACH_HIGHLIGHT_COLS = set()
NAME_COL_W = 40    # Excel column width for product name columns
DATA_COL_W = 15    # Excel column width for stock/qty/sales columns
ROW_H_PER_LINE = 15  # Excel row height (pts) per wrapped line
NAME_COLS = {"Product Name", "RC Required", "RC Product Name", "Product name"}
DATA_COLS = {
    "Product Stock",
    "Avg Monthly Sales",
    "RC Stock",
    "Order",
    "Current Stock",
    "Stock",
    "M/C Order",
    "Manufacturing Order",
    "Avg Sales",
}

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


def _apply_excel_sheet(ws, df: pd.DataFrame, highlight_cols: set[str]):
    cols = list(df.columns)
    col_idx = {c: i + 1 for i, c in enumerate(cols)}
    numeric_cols = {c for c in cols if pd.api.types.is_numeric_dtype(df[c])}

    for c in cols:
        letter = ws.cell(row=1, column=col_idx[c]).column_letter
        ws.column_dimensions[letter].width = NAME_COL_W if c in NAME_COLS else DATA_COL_W

    ws.append(cols)
    ws.row_dimensions[1].height = ROW_H_PER_LINE * 1.2

    for ri, (_, row) in enumerate(df.iterrows(), start=2):
        ws.append(list(row.values))
        max_lines = 1
        for c in cols:
            if c in NAME_COLS:
                val = str(row[c]) if pd.notna(row[c]) else ""
                lines = math.ceil(len(val) / NAME_COL_W) if val else 1
                max_lines = max(max_lines, lines)
        ws.row_dimensions[ri].height = max(ROW_H_PER_LINE, max_lines * ROW_H_PER_LINE)

    for rnum in range(1, len(df) + 2):
        is_header = rnum == 1
        is_even_data = not is_header and (rnum % 2 == 0)
        for c in cols:
            cidx = col_idx[c]
            cell = ws.cell(row=rnum, column=cidx)
            if is_header:
                cell.font = HEADER_FONT
                cell.fill = HEADER_FILL
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            else:
                if c in highlight_cols:
                    cell.fill = HIGHLIGHT_FILL
                elif is_even_data:
                    cell.fill = ALT_ROW_FILL
                else:
                    cell.fill = WHITE_FILL
                if c in numeric_cols:
                    cell.alignment = CENTER
                else:
                    cell.alignment = Alignment(horizontal="left", wrap_text=True, vertical="top")


def _build_excel(mach_out: pd.DataFrame, mfg_out: pd.DataFrame) -> io.BytesIO:
    wb = Workbook()
    wb.remove(wb.active)

    ws_sum = wb.create_sheet("Summary")
    ws_sum.append(["Metric", "Value"])
    ws_sum.append(["Machining Items", len(mach_out)])
    ws_sum.append(["GP Items", len(mfg_out)])
    ws_sum.append(["Machining Units", int(mach_out["Order"].sum()) if len(mach_out) else 0])
    ws_sum.append(["GP Units", int(mfg_out["Order"].sum()) if len(mfg_out) else 0])

    _apply_excel_sheet(wb.create_sheet("Machining Orders"), mach_out, MACH_HIGHLIGHT_COLS)
    _apply_excel_sheet(wb.create_sheet("GP Orders"), mfg_out, set())

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def _build_consolidated_excel(consolidated_out: pd.DataFrame) -> io.BytesIO:
    wb = Workbook()
    wb.remove(wb.active)
    _apply_excel_sheet(wb.create_sheet("Consolidated order"), consolidated_out, set())

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def _add_pdf_table(pdf: FPDF, title: str, df: pd.DataFrame, report_date: date, highlight_cols=None):
    highlight_cols = highlight_cols or set()
    pdf.add_page()
    if df.empty:
        pdf.set_font("Helvetica", "B", 18)
        pdf.cell(0, 12, "Inventory Order Report", new_x="LMARGIN", new_y="NEXT", align="C")
        pdf.set_font("Helvetica", "B", 13)
        pdf.cell(0, 10, title, new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 8, "No items.", new_x="LMARGIN", new_y="NEXT")
        return

    FONT_SIZE = 7
    LINE_H = 4
    PAD = 2
    cols = list(df.columns)
    avail = pdf.w - pdf.l_margin - pdf.r_margin
    numeric_cols = {c for c in cols if pd.api.types.is_numeric_dtype(df[c])}

    def _lines_needed(text, width, font_style=""):
        pdf.set_font("Helvetica", font_style, FONT_SIZE)
        if not text:
            return 1
        words = text.split()
        if not words:
            return 1
        lines, line = 1, ""
        for w in words:
            test = f"{line} {w}".strip()
            if pdf.get_string_width(test) > width - PAD:
                lines += 1
                line = w
            else:
                line = test
        return lines

    pdf.set_font("Helvetica", "", FONT_SIZE)
    col_widths = {}
    name_w = pdf.get_string_width("W" * 40) + PAD * 2
    data_w = pdf.get_string_width("0" * 15) + PAD * 2
    for c in cols:
        if c in DATA_COLS:
            col_widths[c] = data_w
        elif c in NAME_COLS:
            col_widths[c] = name_w

    fixed_total = sum(col_widths.get(c, 0) for c in cols)
    other_cols = [c for c in cols if c not in col_widths]
    remaining = avail - fixed_total
    if other_cols:
        per_other = max(remaining / len(other_cols), 25)
        for c in other_cols:
            col_widths[c] = per_other

    total_w = sum(col_widths.values())
    if total_w > avail and total_w > 0:
        scale = avail / total_w
        for c in cols:
            col_widths[c] = col_widths[c] * scale

    row_counter = [0]

    def _draw_page_header():
        pdf.set_font("Helvetica", "B", 18)
        pdf.cell(0, 12, "Inventory Order Report", new_x="LMARGIN", new_y="NEXT", align="C")
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 6, f"Generated: {report_date.strftime('%d %B %Y')}", new_x="LMARGIN", new_y="NEXT", align="C")
        pdf.ln(2)
        pdf.set_font("Helvetica", "B", 13)
        pdf.cell(0, 10, title, new_x="LMARGIN", new_y="NEXT")

    def _draw_row(values, is_header=False):
        x_start = pdf.l_margin
        y_start = pdf.get_y()
        style = "B" if is_header else ""
        pdf.set_font("Helvetica", style, FONT_SIZE)

        max_lines = 1
        for i, c in enumerate(cols):
            txt = str(values[i]) if values[i] is not None else ""
            w = col_widths[c]
            nl = _lines_needed(txt, w, style)
            if nl > max_lines:
                max_lines = nl
        row_h = max(max_lines * LINE_H, LINE_H)

        if y_start + row_h > pdf.h - 15:
            pdf.add_page()
            _draw_page_header()
            _draw_row(cols, is_header=True)
            y_start = pdf.get_y()

        if is_header:
            bg = (68, 114, 196)
        elif row_counter[0] % 2 == 0:
            bg = (217, 226, 243)
        else:
            bg = (255, 255, 255)

        if not is_header:
            row_counter[0] += 1

        for i, c in enumerate(cols):
            txt = str(values[i]) if values[i] is not None else ""
            w = col_widths[c]
            x = x_start + sum(col_widths[cols[j]] for j in range(i))
            is_highlighted = c in highlight_cols and not is_header

            if is_highlighted:
                pdf.set_fill_color(255, 230, 153)
            else:
                pdf.set_fill_color(*bg)
            pdf.rect(x, y_start, w, row_h, style="F")

            if is_header:
                pdf.set_xy(x, y_start)
                pdf.set_font("Helvetica", "B", FONT_SIZE)
                pdf.set_text_color(255, 255, 255)
                pdf.multi_cell(w, LINE_H, txt, border=0, align="C")
                pdf.set_text_color(0, 0, 0)
            elif c in DATA_COLS:
                y_off = (row_h - LINE_H) / 2
                pdf.set_xy(x, y_start + y_off)
                pdf.set_font("Helvetica", "", FONT_SIZE)
                pdf.cell(w, LINE_H, txt, border=0, align="C")
            else:
                pdf.set_xy(x, y_start)
                pdf.set_font("Helvetica", "", FONT_SIZE)
                pdf.multi_cell(w, LINE_H, txt, border=0, align="L")

            pdf.rect(x, y_start, w, row_h)

        pdf.set_y(y_start + row_h)

    _draw_page_header()
    _draw_row(cols, is_header=True)
    for _, row in df.iterrows():
        vals = [row[c] if pd.notna(row[c]) else "" for c in cols]
        _draw_row(vals)


def _build_pdf(mach_out: pd.DataFrame, mfg_out: pd.DataFrame, report_date: date) -> io.BytesIO:
    pdf = FPDF(orientation="L", format="A4")
    pdf.set_auto_page_break(auto=False)

    FONT_SIZE = 7
    LINE_H = 4
    PAD = 2

    # --- Summary Page ---
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 18)
    pdf.cell(0, 12, "Inventory Order Report", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 8, f"Generated: {report_date.strftime('%d %B %Y')}", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(8)

    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 10, "Summary", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 10)
    total_mach = int(mach_out["Order"].sum()) if len(mach_out) else 0
    total_mfg = int(mfg_out["Order"].sum()) if len(mfg_out) else 0
    for label, val in [
        ("Machining Items", len(mach_out)),
        ("GP Items", len(mfg_out)),
        ("Machining Units", total_mach),
        ("GP Units", total_mfg),
    ]:
        pdf.cell(80, 7, label, border=1)
        pdf.cell(40, 7, str(val), border=1, new_x="LMARGIN", new_y="NEXT")

    _add_pdf_table(pdf, "Machining Orders", mach_out, report_date, highlight_cols=MACH_HIGHLIGHT_COLS)
    _add_pdf_table(pdf, "GP Orders", mfg_out, report_date)

    buf = io.BytesIO()
    buf.write(pdf.output())
    buf.seek(0)
    return buf


def _build_consolidated_pdf(consolidated_out: pd.DataFrame, report_date: date) -> io.BytesIO:
    pdf = FPDF(orientation="L", format="A4")
    pdf.set_auto_page_break(auto=False)
    _add_pdf_table(pdf, "Consolidated order", consolidated_out, report_date)

    buf = io.BytesIO()
    buf.write(pdf.output())
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
    rc_stock_lookup = {
        row["ProductName"].upper(): int(row["StockLevel"])
        for _, row in all_rc.iterrows()
    }
    remaining_rc_stock = dict(rc_stock_lookup)

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
    machining_shortfall_by_rc = {}
    for r in enrich(mach_df, mach_lead):
        pname = r["pname"]
        is_vp = "V-PULLEY" in pname.upper()
        rc = None
        rc_available = None
        if is_vp:
            od, grooves, ptype = extract_parts(pname)
            rc = find_rc(od, grooves, ptype, rc_lookup)
        if rc:
            rc_key = rc.upper()
            rc_available = remaining_rc_stock.get(rc_key, 0)
            suggested = min(r["suggested"], rc_available)
            remaining_rc_stock[rc_key] = max(rc_available - suggested, 0)
            machining_shortfall = max(r["suggested"] - suggested, 0)
            if machining_shortfall:
                machining_shortfall_by_rc[rc_key] = machining_shortfall_by_rc.get(rc_key, 0) + machining_shortfall
        else:
            suggested = r["suggested"]
        mach_rows.append(
            {
                "Product Name": pname,
                "Product Stock": r["stk"],
                "Avg Monthly Sales": r["amonthly"],
                "RC Required": rc if rc else ("No RC Found" if is_vp else "N/A"),
                "RC Stock": rc_available if rc_available is not None else "N/A",
                "Order": suggested,
            }
        )

    mfg_order_map = {}
    for r in enrich(mfg_df, mfg_lead):
        rc_key = r["pname"].upper()
        mfg_order_map[rc_key] = {
            "RC Product Name": r["pname"],
            "Current Stock": r["stk"],
            "Order": r["suggested"],
            "Avg Monthly Sales": r["amonthly"],
        }

    for rc_key, shortfall in machining_shortfall_by_rc.items():
        if rc_key in mfg_order_map:
            mfg_order_map[rc_key]["Order"] += shortfall
        else:
            mfg_order_map[rc_key] = {
                "RC Product Name": rc_lookup.get(rc_key, rc_key),
                "Current Stock": rc_stock_lookup.get(rc_key, 0),
                "Order": shortfall,
                "Avg Monthly Sales": 0.0,
            }

    mfg_rows = list(mfg_order_map.values())

    mach_out = pd.DataFrame(mach_rows).sort_values("Order", ascending=False)
    mfg_out = pd.DataFrame(mfg_rows)
    if not mfg_out.empty:
        mfg_out = mfg_out[mfg_out["Order"] > 0].sort_values("Order", ascending=False)
    consolidated_rows = []
    for _, row in mach_out.iterrows():
        consolidated_rows.append(
            {
                "Product name": row["Product Name"],
                "Stock": row["Product Stock"],
                "RC Stock": row["RC Stock"],
                "M/C Order": row["Order"],
                "Manufacturing Order": "",
                "Avg Sales": row["Avg Monthly Sales"],
            }
        )
    for _, row in mfg_out.iterrows():
        consolidated_rows.append(
            {
                "Product name": row["RC Product Name"],
                "Stock": row["Current Stock"],
                "RC Stock": "",
                "M/C Order": "",
                "Manufacturing Order": row["Order"],
                "Avg Sales": row["Avg Monthly Sales"] if "Avg Monthly Sales" in row else "",
            }
        )

    consolidated_out = pd.DataFrame(consolidated_rows)
    if not consolidated_out.empty:
        consolidated_out["_sort"] = pd.to_numeric(consolidated_out["M/C Order"], errors="coerce").fillna(0) + pd.to_numeric(consolidated_out["Manufacturing Order"], errors="coerce").fillna(0)
        consolidated_out = consolidated_out.sort_values(["_sort", "Product name"], ascending=[False, True]).drop(columns=["_sort"])

    excel_buf = _build_excel(mach_out, mfg_out)
    pdf_buf = _build_pdf(mach_out, mfg_out, today)
    consolidated_excel_buf = _build_consolidated_excel(consolidated_out)
    consolidated_pdf_buf = _build_consolidated_pdf(consolidated_out, today)
    return excel_buf, pdf_buf, consolidated_excel_buf, consolidated_pdf_buf, mach_out, mfg_out, consolidated_out, today
