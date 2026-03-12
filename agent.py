import io
import math
import re
from datetime import date

import pandas as pd
from fpdf import FPDF
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill

HIGHLIGHT_FILL = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
CENTER = Alignment(horizontal="center")
MACH_HIGHLIGHT_COLS = {"Product Name", "RC Required", "Suggested Order Qty"}
TEXT_COL_W = 40    # Excel column width (chars) for text columns
NUM_COL_W  = 14   # Excel column width (chars) for numeric columns
ROW_H_PER_LINE = 15  # Excel row height (pts) per wrapped line

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

        def _apply_sheet(ws, df, highlight_cols):
            cols = list(df.columns)
            col_idx = {c: i + 1 for i, c in enumerate(cols)}
            numeric_cols = {c for c in cols if pd.api.types.is_numeric_dtype(df[c])}

            # Set column widths
            for c in cols:
                letter = ws.cell(row=1, column=col_idx[c]).column_letter
                ws.column_dimensions[letter].width = NUM_COL_W if c in numeric_cols else TEXT_COL_W

            # Write header
            ws.append(cols)
            ws.row_dimensions[1].height = ROW_H_PER_LINE

            # Write data rows
            for ri, (_, row) in enumerate(df.iterrows(), start=2):
                ws.append(list(row.values))
                max_lines = 1
                for c in cols:
                    if c not in numeric_cols:
                        val = str(row[c]) if pd.notna(row[c]) else ""
                        lines = math.ceil(len(val) / TEXT_COL_W) if val else 1
                        max_lines = max(max_lines, lines)
                ws.row_dimensions[ri].height = max(ROW_H_PER_LINE, max_lines * ROW_H_PER_LINE)

            # Apply styles
            for rnum in range(1, len(df) + 2):
                for c in cols:
                    cidx = col_idx[c]
                    cell = ws.cell(row=rnum, column=cidx)
                    if c in highlight_cols:
                        cell.fill = HIGHLIGHT_FILL
                    if c in numeric_cols:
                        cell.alignment = CENTER
                    else:
                        cell.alignment = Alignment(horizontal="left", wrap_text=True, vertical="top")

        _apply_sheet(wb.create_sheet("Machining Orders"), mach_out, MACH_HIGHLIGHT_COLS)
        _apply_sheet(wb.create_sheet("Manufacturing Orders"), mfg_out, set())

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


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
    total_mach = int(mach_out["Suggested Order Qty"].sum()) if len(mach_out) else 0
    total_mfg = int(mfg_out["Suggested Order Qty"].sum()) if len(mfg_out) else 0
    for label, val in [
        ("Machining Items", len(mach_out)),
        ("Manufacturing Items", len(mfg_out)),
        ("Machining Units", total_mach),
        ("Manufacturing Units", total_mfg),
    ]:
        pdf.cell(80, 7, label, border=1)
        pdf.cell(40, 7, str(val), border=1, new_x="LMARGIN", new_y="NEXT")

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

    def _add_table(title, df, highlight_cols=None):
        highlight_cols = highlight_cols or set()
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 13)
        pdf.cell(0, 10, title, new_x="LMARGIN", new_y="NEXT")
        if df.empty:
            pdf.set_font("Helvetica", "", 10)
            pdf.cell(0, 8, "No items.", new_x="LMARGIN", new_y="NEXT")
            return

        cols = list(df.columns)
        avail = pdf.w - pdf.l_margin - pdf.r_margin
        numeric_cols = {c for c in cols if pd.api.types.is_numeric_dtype(df[c])}

        # Numeric col widths: sized to the max data value width (not header)
        pdf.set_font("Helvetica", "", FONT_SIZE)
        col_widths = {}
        for c in cols:
            if c in numeric_cols:
                max_w = max(
                    (pdf.get_string_width(str(v)) for v in df[c] if pd.notna(v)),
                    default=pdf.get_string_width("0"),
                )
                col_widths[c] = max_w + PAD * 2

        # Distribute remaining width to text columns
        numeric_total = sum(col_widths.get(c, 0) for c in cols)
        text_cols = [c for c in cols if c not in numeric_cols]
        remaining = avail - numeric_total
        if text_cols:
            per_text = max(remaining / len(text_cols), 25)
            for c in text_cols:
                col_widths[c] = per_text

        def _draw_row(values, is_header=False):
            x_start = pdf.l_margin
            y_start = pdf.get_y()
            style = "B" if is_header else ""
            pdf.set_font("Helvetica", style, FONT_SIZE)

            # Pre-calculate row height from tallest cell
            max_lines = 1
            for i, c in enumerate(cols):
                txt = str(values[i]) if values[i] is not None else ""
                w = col_widths[c]
                nl = _lines_needed(txt, w, style)
                if nl > max_lines:
                    max_lines = nl
            row_h = max(max_lines * LINE_H, LINE_H)

            # Page break if row won't fit
            if y_start + row_h > pdf.h - 15:
                pdf.add_page()
                y_start = pdf.get_y()

            # Draw each cell
            for i, c in enumerate(cols):
                txt = str(values[i]) if values[i] is not None else ""
                w = col_widths[c]
                x = x_start + sum(col_widths[cols[j]] for j in range(i))
                is_highlighted = c in highlight_cols

                # Fill highlight background
                if is_highlighted:
                    pdf.set_fill_color(255, 230, 153)  # light yellow
                    pdf.rect(x, y_start, w, row_h, style="F")
                else:
                    pdf.set_fill_color(255, 255, 255)

                if c in numeric_cols and not is_header:
                    # Numeric data: center-aligned, vertically centered, no wrap
                    y_off = (row_h - LINE_H) / 2
                    pdf.set_xy(x, y_start + y_off)
                    pdf.set_font("Helvetica", style, FONT_SIZE)
                    pdf.cell(w, LINE_H, txt, border=0, align="C")
                else:
                    # Text / header: left-aligned with word wrap
                    pdf.set_xy(x, y_start)
                    pdf.set_font("Helvetica", style, FONT_SIZE)
                    pdf.multi_cell(w, LINE_H, txt, border=0, align="L")

                # Draw cell border rectangle
                pdf.rect(x, y_start, w, row_h)

            pdf.set_y(y_start + row_h)

        # Header row
        _draw_row(cols, is_header=True)
        # Data rows
        for _, row in df.iterrows():
            vals = [row[c] if pd.notna(row[c]) else "" for c in cols]
            _draw_row(vals)

    _add_table("Machining Orders", mach_out, highlight_cols=MACH_HIGHLIGHT_COLS)
    _add_table("Manufacturing Orders", mfg_out)

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
        rc_available = None
        if is_vp:
            od, grooves, ptype = extract_parts(pname)
            rc = find_rc(od, grooves, ptype, rc_lookup)
        if rc:
            rc_available = rc_stock_lookup.get(rc.upper(), 0)
            suggested = min(r["suggested"], rc_available)
        else:
            suggested = r["suggested"]
        mach_rows.append(
            {
                    "Product Name": pname,
                "Product Stock": r["stk"],
                "Avg Monthly Sales": r["amonthly"],
                    "RC Required": rc if rc else ("No RC Found" if is_vp else "N/A"),
                "RC Stock": rc_available if rc_available is not None else "N/A",
                "Suggested Order Qty": suggested,
            }
        )

    mfg_rows = []
    for r in enrich(mfg_df, mfg_lead):
        mfg_rows.append(
            {
                    "RC Product Name": r["pname"],
                "Current Stock": r["stk"],
                "Suggested Order Qty": r["suggested"],
            }
        )

    mach_out = pd.DataFrame(mach_rows).sort_values("Suggested Order Qty", ascending=False)
    mfg_out = pd.DataFrame(mfg_rows).sort_values("Suggested Order Qty", ascending=False)
    excel_buf = _build_excel(mach_out, mfg_out)
    pdf_buf = _build_pdf(mach_out, mfg_out, today)
    return excel_buf, pdf_buf, mach_out, mfg_out, today
