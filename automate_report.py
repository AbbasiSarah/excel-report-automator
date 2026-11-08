"""Excel Report Automator

Turns a folder of messy Excel files (one per branch per month) into ONE clean,
formatted Excel report with KPIs, summary tables, charts and a log of every fix.

Usage:
    python automate_report.py --input input --output output/sales_report.xlsx
"""
import argparse
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Different branches use different column names. Map them all to one standard name.
COLUMN_NAMES = {
    "date": ["date", "order date", "order_date", "sale date"],
    "product": ["product", "item", "product name", "product_name"],
    "qty": ["qty", "quantity", "units", "unit sold"],
    "price": ["unit price", "price", "unit_price", "rate"],
    "region": ["region", "area", "zone"],
}
BRANCH_REGION = {"lahore": "Punjab", "karachi": "Sindh", "islamabad": "Capital"}
DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%d %b %Y", "%d-%m-%Y", "%m/%d/%Y"]

NAVY, TEAL, GOLD, LIGHT = "0B1E4A", "2FE0BF", "F5B70A", "EEF3FF"


def standard_columns(df):
    rename = {}
    for col in df.columns:
        key = str(col).strip().lower()
        for std, options in COLUMN_NAMES.items():
            if key in options:
                rename[col] = std
    df = df.rename(columns=rename)
    missing = set(COLUMN_NAMES) - set(df.columns)
    if missing:
        raise ValueError(f"missing columns {sorted(missing)}")
    return df[list(COLUMN_NAMES)]


def parse_date(value):
    if isinstance(value, (datetime, pd.Timestamp)):
        return pd.Timestamp(value)
    text = str(value).strip()
    for fmt in DATE_FORMATS:
        try:
            return pd.Timestamp(datetime.strptime(text, fmt))
        except ValueError:
            continue
    return pd.NaT


def load_and_clean(folder):
    frames, issues = [], []
    files = sorted(Path(folder).glob("*.xlsx"))
    if not files:
        raise SystemExit(f"No .xlsx files found in {folder}")
    for f in files:
        branch = re.split(r"[_\-\s]", f.stem)[0].lower()
        try:
            df = standard_columns(pd.read_excel(f))
        except ValueError as e:
            issues.append((f.name, "-", "Skipped file", str(e)))
            continue
        df["branch"] = branch.title()
        df["source_file"] = f.name
        df["source_row"] = df.index + 2          # +2 = Excel row number (header is row 1)
        frames.append(df)

    data = pd.concat(frames, ignore_index=True)
    n_start = len(data)

    # 1. tidy text
    data["product"] = data["product"].astype(str).str.strip().str.title()
    data["region"] = data["region"].fillna("").astype(str).str.strip().str.title()
    blank = data["region"] == ""
    for _, r in data[blank].iterrows():
        issues.append((r.source_file, r.source_row, "Blank region", f"filled from branch: {BRANCH_REGION.get(r.branch.lower(), 'Unknown')}"))
    data.loc[blank, "region"] = data.loc[blank, "branch"].str.lower().map(BRANCH_REGION).fillna("Unknown")

    # 2. dates in any format -> real dates
    data["date"] = data["date"].apply(parse_date)
    for _, r in data[data["date"].isna()].iterrows():
        issues.append((r.source_file, r.source_row, "Unreadable date", "row removed"))
    data = data[data["date"].notna()]

    # 3. exact duplicates
    dup = data.duplicated(subset=["date", "product", "qty", "price", "branch"], keep="first")
    for _, r in data[dup].iterrows():
        issues.append((r.source_file, r.source_row, "Duplicate row", "removed"))
    data = data[~dup]

    # 4. missing or impossible numbers
    bad_qty = data["qty"].isna() | (data["qty"] <= 0)
    for _, r in data[bad_qty].iterrows():
        issues.append((r.source_file, r.source_row, "Missing quantity" if pd.isna(r.qty) else "Negative quantity",
                       "row removed" if pd.isna(r.qty) else f"value {int(r.qty)}; row removed"))
    data = data[~bad_qty]
    no_price = data["price"].isna()
    medians = data.groupby("product")["price"].median()
    for _, r in data[no_price].iterrows():
        issues.append((r.source_file, r.source_row, "Missing price", f"filled with {r['product']} median: {medians[r['product']]:.2f}"))
    data.loc[no_price, "price"] = data.loc[no_price, "product"].map(medians)

    data["qty"] = data["qty"].astype(int)
    data["revenue"] = (data["qty"] * data["price"]).round(2)
    data["month"] = data["date"].dt.to_period("M").astype(str)
    data = data.sort_values(["date", "branch"]).reset_index(drop=True)
    stats = {"files": len(files), "rows_in": n_start, "rows_out": len(data), "fixes": len(issues)}
    return data, pd.DataFrame(issues, columns=["File", "Row", "Problem", "Action"]), stats


# ---------------- Excel writing ----------------
THIN = Side(style="thin", color="C9D3EA")
def header(ws, row, cols, start_col=1):
    for i, name in enumerate(cols):
        c = ws.cell(row=row, column=start_col + i, value=name)
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center"); c.border = Border(bottom=THIN)


def write_table(ws, df, row, col=1, money_cols=(), int_cols=()):
    header(ws, row, list(df.columns), col)
    for r_i, rec in enumerate(df.itertuples(index=False), start=row + 1):
        for c_i, v in enumerate(rec, start=col):
            cell = ws.cell(row=r_i, column=c_i, value=v)
            name = df.columns[c_i - col]
            if name in money_cols: cell.number_format = '#,##0.00'
            if name in int_cols: cell.number_format = '#,##0'
            if (r_i - row) % 2 == 0: cell.fill = PatternFill("solid", fgColor=LIGHT)
            cell.border = Border(bottom=THIN)
    return row + len(df) + 1


def autosize(ws, min_w=10, max_w=45):
    for col in ws.columns:
        width = max((len(str(c.value)) for c in col if c.value is not None), default=min_w)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(min_w, min(max_w, width + 2))


def build_report(data, issues, stats, out_path):
    wb = Workbook()
    ws = wb.active; ws.title = "Summary"
    ws["A1"] = "Sales Report"; ws["A1"].font = Font(size=22, bold=True, color=NAVY)
    ws["A2"] = f"{data['date'].min():%d %b %Y} to {data['date'].max():%d %b %Y}  ·  built automatically from {stats['files']} files"
    ws["A2"].font = Font(italic=True, color="5A6785")

    by_branch = data.groupby("branch").agg(Orders=("qty", "size"), Units=("qty", "sum"), Revenue=("revenue", "sum")).reset_index().rename(columns={"branch": "Branch"})
    by_product = data.groupby("product").agg(Units=("qty", "sum"), Revenue=("revenue", "sum")).sort_values("Revenue", ascending=False).reset_index().rename(columns={"product": "Product"})
    by_month = data.pivot_table(index="month", columns="branch", values="revenue", aggfunc="sum").round(2).reset_index().rename(columns={"month": "Month"})

    kpis = [("Total revenue", data["revenue"].sum(), '#,##0'), ("Orders", len(data), '#,##0'), ("Units sold", data["qty"].sum(), '#,##0'),
            ("Best branch", by_branch.sort_values("Revenue").iloc[-1]["Branch"], None), ("Best product", by_product.iloc[0]["Product"], None),
            ("Data fixes logged", stats["fixes"], '#,##0')]
    for i, (label, value, fmt) in enumerate(kpis):
        col = 1 + (i % 3) * 2; row = 4 + (i // 3) * 3
        ws.cell(row=row, column=col, value=label).font = Font(color="5A6785", bold=True)
        c = ws.cell(row=row + 1, column=col, value=value); c.font = Font(size=18, bold=True, color=NAVY)
        if fmt: c.number_format = fmt
        for r in (row, row + 1):
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=LIGHT)

    r = 11
    ws.cell(row=r - 1, column=1, value="Revenue by branch").font = Font(bold=True, size=13, color=NAVY)
    end_b = write_table(ws, by_branch, r, money_cols=("Revenue",), int_cols=("Units", "Orders"))
    r2 = end_b + 2
    ws.cell(row=r2 - 1, column=1, value="Revenue by product").font = Font(bold=True, size=13, color=NAVY)
    end_p = write_table(ws, by_product, r2, money_cols=("Revenue",), int_cols=("Units",))
    r3 = end_p + 2
    ws.cell(row=r3 - 1, column=1, value="Monthly revenue by branch").font = Font(bold=True, size=13, color=NAVY)
    end_m = write_table(ws, by_month, r3, money_cols=tuple(by_month.columns[1:]))

    ch = BarChart(); ch.title = "Revenue by branch"; ch.style = 10; ch.height, ch.width = 7, 13
    ch.add_data(Reference(ws, min_col=4, min_row=r, max_row=end_b - 1), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=r + 1, max_row=end_b - 1)); ch.legend = None
    ch.y_axis.scaling.min = 0
    ch.series[0].graphicalProperties.solidFill = NAVY
    ws.add_chart(ch, "H4")
    ch2 = BarChart(); ch2.type = "bar"; ch2.title = "Revenue by product"; ch2.style = 10; ch2.height, ch2.width = 7, 13
    ch2.add_data(Reference(ws, min_col=3, min_row=r2, max_row=end_p - 1), titles_from_data=True)
    ch2.set_categories(Reference(ws, min_col=1, min_row=r2 + 1, max_row=end_p - 1)); ch2.legend = None
    ch2.x_axis.scaling.min = 0
    ch2.series[0].graphicalProperties.solidFill = TEAL
    ws.add_chart(ch2, "H19")
    ch3 = LineChart(); ch3.title = "Monthly revenue"; ch3.style = 12; ch3.height, ch3.width = 7, 13
    ch3.add_data(Reference(ws, min_col=2, max_col=by_month.shape[1], min_row=r3, max_row=end_m - 1), titles_from_data=True)
    ch3.set_categories(Reference(ws, min_col=1, min_row=r3 + 1, max_row=end_m - 1))
    for s_, color in zip(ch3.series, [NAVY, TEAL, GOLD, "8E7CFF", "FF6B6B"]):
        s_.smooth = False
        s_.graphicalProperties.line.solidFill = color
        s_.graphicalProperties.line.width = 28000
    ch3.y_axis.scaling.min = 0
    ws.add_chart(ch3, "H34")
    autosize(ws, 14, 22); ws.column_dimensions["A"].width = 22
    for sheet in (ws,):
        sheet.page_setup.orientation = "landscape"
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.fitToWidth, sheet.page_setup.fitToHeight = 1, 1

    wd = wb.create_sheet("Clean data")
    clean = data[["date", "branch", "region", "product", "qty", "price", "revenue"]].copy()
    clean["date"] = clean["date"].dt.date
    clean.columns = ["Date", "Branch", "Region", "Product", "Qty", "Unit price", "Revenue"]
    write_table(wd, clean, 1, money_cols=("Unit price", "Revenue"), int_cols=("Qty",))
    for row in wd.iter_rows(min_row=2, max_col=1):
        row[0].number_format = "DD MMM YYYY"
    wd.freeze_panes = "A2"; wd.auto_filter.ref = wd.dimensions; autosize(wd)

    wi = wb.create_sheet("Data fixes")
    wi["A1"] = f"{len(issues)} problems were found and fixed. Nothing was changed silently."
    wi["A1"].font = Font(bold=True, color=NAVY)
    write_table(wi, issues, 3); wi.freeze_panes = "A4"; autosize(wi)

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return by_branch, by_product


def main():
    ap = argparse.ArgumentParser(description="Merge and clean messy Excel files into one report.")
    ap.add_argument("--input", default="input", help="folder with the .xlsx files")
    ap.add_argument("--output", default="output/sales_report.xlsx", help="report file to create")
    a = ap.parse_args()
    data, issues, stats = load_and_clean(a.input)
    by_branch, by_product = build_report(data, issues, stats, a.output)
    print(f"Read {stats['files']} files, {stats['rows_in']} rows -> {stats['rows_out']} clean rows, {stats['fixes']} fixes logged")
    print(f"Total revenue: {data['revenue'].sum():,.2f}")
    print(by_branch.to_string(index=False))
    print(f"Report saved: {a.output}")


if __name__ == "__main__":
    main()
