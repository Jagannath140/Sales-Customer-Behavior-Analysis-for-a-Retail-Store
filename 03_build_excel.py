"""
03_build_excel.py
Purpose : Build excel/Retail_Sales_Analysis.xlsx with LIVE formulas (SUMIFS / INDEX-MATCH / RANK ...),
          conditional formatting, charts, a dashboard and a reconciliation sheet.
Run     : python python/03_build_excel.py   (after 01 and 02, from the project root)
Note    : Summary tables are formula-based "pivot-style" tables. Real PivotTables cannot be authored by
          Python libraries; excel/PIVOT_TABLE_STEPS in the README sheet explains the 2-minute manual build.
"""
import json
import os
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

ROOT = Path(__file__).resolve().parents[1]
OUT_FILE = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "excel/Retail_Sales_Analysis.xlsx"
CACHE = json.loads(Path(os.environ["XL_CACHE"]).read_text()) if os.environ.get("XL_CACHE") else {}
(ROOT / "excel").mkdir(exist_ok=True)

cust = pd.read_csv(ROOT / "data/cleaned/Customers.csv", parse_dates=["JoinDate"])
prod = pd.read_csv(ROOT / "data/cleaned/Products.csv")
sales = pd.read_csv(ROOT / "data/cleaned/Sales.csv", parse_dates=["SaleDate"])
master = pd.read_csv(ROOT / "data/cleaned/Sales_Master.csv")
cust_clv = pd.read_csv(ROOT / "outputs/customer_clv.csv")          # sorted by revenue (order only)
prod_perf = pd.read_csv(ROOT / "outputs/product_performance.csv")  # sorted by revenue (order only)
city_perf = pd.read_csv(ROOT / "outputs/city_performance.csv")
cat_perf = pd.read_csv(ROOT / "outputs/category_performance.csv")
monthly = pd.read_csv(ROOT / "outputs/monthly_revenue.csv")
fc = pd.read_csv(ROOT / "outputs/forecast_6_months.csv")
mysql = pd.read_csv(ROOT / "outputs/mysql_validation_metrics.tsv", sep="\t").set_index("Metric")["Value"]

NC, NP, NS = len(cust), len(prod), len(sales)           # 50, 30, 300
wb = xlsxwriter.Workbook(OUT_FILE, {"use_future_functions": True})
NAVY, ORANGE, GREY = "#1f4e79", "#e07b39", "#f2f2f2"
f = lambda **k: wb.add_format({"font_name": "Calibri", "font_size": 10, **k})
F = dict(
    title=f(bold=True, font_size=16, font_color=NAVY), sub=f(italic=True, font_color="#595959"),
    hdr=f(bold=True, font_color="white", bg_color=NAVY, border=1, align="center", valign="vcenter", text_wrap=True),
    txt=f(border=1), num=f(border=1, num_format="#,##0"), num2=f(border=1, num_format="#,##0.00"),
    pct=f(border=1, num_format="0.0%"), date=f(border=1, num_format="yyyy-mm-dd"),
    tot=f(bold=True, border=1, bg_color=GREY, num_format="#,##0"), totp=f(bold=True, border=1, bg_color=GREY, num_format="0.0%"),
    tott=f(bold=True, border=1, bg_color=GREY), inp=f(border=1, bg_color="#fff2cc", bold=True),
    inpp=f(border=1, bg_color="#fff2cc", bold=True, num_format="0%"), wrap=f(text_wrap=True, valign="top"),
    kl=f(bold=True, font_color="white", bg_color=NAVY, align="center", valign="vcenter", text_wrap=True),
    kv=f(bold=True, font_size=18, font_color=NAVY, bg_color="#deebf7", align="center", valign="vcenter", num_format="#,##0"),
    kvp=f(bold=True, font_size=18, font_color=NAVY, bg_color="#deebf7", align="center", valign="vcenter", num_format="0.0%"),
    ins=f(text_wrap=True, valign="top", bg_color="#fff8e5", border=1), h2=f(bold=True, font_size=12, font_color=NAVY),
    red=f(bg_color="#f8cbad", font_color="#9c0006"), orange=f(bg_color="#ffe699", font_color="#7f6000"),
    green=f(bg_color="#c6efce", font_color="#006100"), yellow=f(bg_color="#fff2cc"),
)


def wf(ws, cell, formula, fmt=None):
    """write a formula together with its cached value (from the LibreOffice recalculation pass)."""
    ws.write_formula(cell, formula, fmt, CACHE.get(f"{ws.get_name()}!{cell}", 0))


def header(ws, row, labels, col=0):
    for i, h in enumerate(labels):
        ws.write(row, col + i, h, F["hdr"])
    ws.set_row(row, 30)


README = wb.add_worksheet("README")
DASH = wb.add_worksheet("Dashboard")
PIV = wb.add_worksheet("Pivot_Summaries")
PRD = wb.add_worksheet("Product_Analysis")
CUS = wb.add_worksheet("Customer_Analysis")
FCS = wb.add_worksheet("Forecast")
REC = wb.add_worksheet("Reconciliation")
DQ = wb.add_worksheet("Data_Quality")
CR = wb.add_worksheet("Combined_Report")
WC = wb.add_worksheet("Customers_Clean")
WP = wb.add_worksheet("Products_Clean")
WS = wb.add_worksheet("Sales_Clean")
for ws, colr in [(README, NAVY), (DASH, ORANGE), (PIV, "#2e75b6"), (PRD, "#2e75b6"), (CUS, "#2e75b6"), (FCS, "#2e75b6"),
                 (REC, "#548235"), (DQ, "#548235"), (CR, "#7f7f7f"), (WC, "#bfbfbf"), (WP, "#bfbfbf"), (WS, "#bfbfbf")]:
    ws.set_tab_color(colr)

# ---------------------------------------------------------------- CLEAN DATA SHEETS (values)
header(WC, 0, ["CustomerID", "Name", "Age", "Gender", "City", "JoinDate"])
for r, row in enumerate(cust.itertuples(index=False), 1):
    WC.write(r, 0, row.CustomerID); WC.write(r, 1, row.Name); WC.write_number(r, 2, row.Age)
    WC.write(r, 3, row.Gender); WC.write(r, 4, row.City); WC.write_datetime(r, 5, row.JoinDate.to_pydatetime(), F["date"])
header(WP, 0, ["ProductID", "ProductName", "Category", "UnitPrice", "CostPrice"])
for r, row in enumerate(prod.itertuples(index=False), 1):
    WP.write(r, 0, row.ProductID); WP.write(r, 1, row.ProductName); WP.write(r, 2, row.Category)
    WP.write_number(r, 3, row.UnitPrice); WP.write_number(r, 4, row.CostPrice)
header(WS, 0, ["SalesID", "CustomerID", "ProductID", "Quantity", "SaleDate"])
for r, row in enumerate(sales.itertuples(index=False), 1):
    WS.write(r, 0, row.SalesID); WS.write(r, 1, row.CustomerID); WS.write(r, 2, row.ProductID)
    WS.write_number(r, 3, row.Quantity); WS.write_datetime(r, 4, row.SaleDate.to_pydatetime(), F["date"])
for ws, widths in [(WC, [11, 18, 6, 9, 12, 12]), (WP, [10, 18, 13, 10, 10]), (WS, [9, 11, 10, 9, 12])]:
    for i, w in enumerate(widths): ws.set_column(i, i, w)
    ws.freeze_panes(1, 0)
WC.autofilter(0, 0, NC, 5); WP.autofilter(0, 0, NP, 4); WS.autofilter(0, 0, NS, 4)

# ---------------------------------------------------------------- COMBINED SALES REPORT (formulas)
cols = ["SalesID", "SaleDate", "YearMonth", "Year", "CustomerID", "CustomerName", "City", "Gender", "ProductID", "ProductName",
        "Category", "Quantity", "UnitPrice", "CostPrice", "Revenue", "Cost", "Profit", "Margin %", "JoinDate", "SaleBeforeJoin"]
header(CR, 0, cols)
C_ = lambda col: f"Customers_Clean!${col}$2:${col}${NC + 1}"
P_ = lambda col: f"Products_Clean!${col}$2:${col}${NP + 1}"
for r in range(2, NS + 2):
    wf(CR, f"A{r}", f"=Sales_Clean!A{r}", F["txt"])
    wf(CR, f"B{r}", f"=Sales_Clean!E{r}", F["date"])
    wf(CR, f"C{r}", f'=TEXT(B{r},"yyyy-mm")', F["txt"])
    wf(CR, f"D{r}", f"=YEAR(B{r})", F["txt"])
    wf(CR, f"E{r}", f"=Sales_Clean!B{r}", F["txt"])
    for col, src in [("F", "B"), ("G", "E"), ("H", "D")]:
        wf(CR, f"{col}{r}", f"=INDEX({C_(src)},MATCH($E{r},{C_('A')},0))", F["txt"])
    wf(CR, f"I{r}", f"=Sales_Clean!C{r}", F["txt"])
    for col, src in [("J", "B"), ("K", "C")]:
        wf(CR, f"{col}{r}", f"=INDEX({P_(src)},MATCH($I{r},{P_('A')},0))", F["txt"])
    wf(CR, f"L{r}", f"=Sales_Clean!D{r}", F["num"])
    wf(CR, f"M{r}", f"=INDEX({P_('D')},MATCH($I{r},{P_('A')},0))", F["num"])
    wf(CR, f"N{r}", f"=INDEX({P_('E')},MATCH($I{r},{P_('A')},0))", F["num"])
    wf(CR, f"O{r}", f"=L{r}*M{r}", F["num"])
    wf(CR, f"P{r}", f"=L{r}*N{r}", F["num"])
    wf(CR, f"Q{r}", f"=O{r}-P{r}", F["num"])
    wf(CR, f"R{r}", f"=Q{r}/O{r}", F["pct"])
    wf(CR, f"S{r}", f"=INDEX({C_('F')},MATCH($E{r},{C_('A')},0))", F["date"])
    wf(CR, f"T{r}", f'=IF(B{r}<S{r},"Yes","No")', F["txt"])
for i, w in enumerate([9, 11, 10, 6, 11, 17, 11, 8, 10, 16, 12, 8, 9, 9, 10, 10, 10, 9, 11, 13]): CR.set_column(i, i, w)
CR.freeze_panes(1, 2); CR.autofilter(0, 0, NS, len(cols) - 1)
CR.conditional_format(f"R2:R{NS + 1}", {"type": "3_color_scale", "min_color": "#f8696b", "mid_color": "#ffeb84", "max_color": "#63be7b"})
CR.conditional_format(f"T2:T{NS + 1}", {"type": "cell", "criteria": "==", "value": '"Yes"', "format": F["orange"]})
R = lambda col: f"Combined_Report!${col}$2:${col}${NS + 1}"     # shorthand for a Combined_Report column range

# ---------------------------------------------------------------- PIVOT-STYLE SUMMARIES
PIV.write("A1", "Pivot-style summaries (live SUMIFS formulas on Combined_Report)", F["title"])
PIV.write("A2", "Business question each table answers is written above it. Amounts in QAR (currency per project document).", F["sub"])
PIV.set_column("A:A", 14); PIV.set_column("B:F", 13)
# Pivot 1: City
PIV.write("A4", "Pivot 1: Where does revenue come from? (Total Sales by City)", F["h2"])
header(PIV, 4, ["City", "Customers", "Revenue", "Profit", "Margin %"])
for i, city in enumerate(city_perf.City):
    r = 6 + i
    PIV.write(f"A{r}", city, F["txt"])
    wf(PIV, f"B{r}", f"=COUNTIF({C_('E')},A{r})", F["num"])
    wf(PIV, f"C{r}", f"=SUMIFS({R('O')},{R('G')},A{r})", F["num"])
    wf(PIV, f"D{r}", f"=SUMIFS({R('Q')},{R('G')},A{r})", F["num"])
    wf(PIV, f"E{r}", f"=D{r}/C{r}", F["pct"])
PIV.write("A11", "Total", F["tott"])
for c in "BCD": wf(PIV, f"{c}11", f"=SUM({c}6:{c}10)", F["tot"])
wf(PIV, "E11", "=D11/C11", F["totp"])
# Pivot 2: Category
PIV.write("A14", "Pivot 2: Which category performs best? (Category Performance)", F["h2"])
header(PIV, 14, ["Category", "Units Sold", "Revenue", "Profit", "Margin %", "Revenue Share"])
for i, cat in enumerate(cat_perf.Category):
    r = 16 + i
    PIV.write(f"A{r}", cat, F["txt"])
    wf(PIV, f"B{r}", f"=SUMIFS({R('L')},{R('K')},A{r})", F["num"])
    wf(PIV, f"C{r}", f"=SUMIFS({R('O')},{R('K')},A{r})", F["num"])
    wf(PIV, f"D{r}", f"=SUMIFS({R('Q')},{R('K')},A{r})", F["num"])
    wf(PIV, f"E{r}", f"=D{r}/C{r}", F["pct"])
    wf(PIV, f"F{r}", f"=C{r}/$C$18", F["pct"])
PIV.write("A18", "Total", F["tott"])
for c in "BCD": wf(PIV, f"{c}18", f"=SUM({c}16:{c}17)", F["tot"])
wf(PIV, "E18", "=D18/C18", F["totp"]); wf(PIV, "F18", "=SUM(F16:F17)", F["totp"])
# Pivot 3: Monthly
PIV.write("A21", "Pivot 3: How does revenue move through the year? (Monthly Revenue)", F["h2"])
header(PIV, 21, ["Month", "Revenue", "Profit", "Units", "Margin %", "MoM Growth"])
for i, m in enumerate(monthly.Month):
    r = 23 + i
    PIV.write(f"A{r}", m, F["txt"])
    wf(PIV, f"B{r}", f"=SUMIFS({R('O')},{R('C')},A{r})", F["num"])
    wf(PIV, f"C{r}", f"=SUMIFS({R('Q')},{R('C')},A{r})", F["num"])
    wf(PIV, f"D{r}", f"=SUMIFS({R('L')},{R('C')},A{r})", F["num"])
    wf(PIV, f"E{r}", f"=C{r}/B{r}", F["pct"])
    if i == 0: PIV.write(f"F{r}", "n/a", F["txt"])
    else: wf(PIV, f"F{r}", f"=B{r}/B{r-1}-1", F["pct"])
PIV.write("A35", "Total", F["tott"])
for c in "BCD": wf(PIV, f"{c}35", f"=SUM({c}23:{c}34)", F["tot"])
wf(PIV, "E35", "=C35/B35", F["totp"])
PIV.conditional_format("B23:B34", {"type": "data_bar", "bar_color": "#9dc3e6", "bar_solid": True})
PIV.conditional_format("F24:F34", {"type": "cell", "criteria": "<", "value": -0.25, "format": F["red"]})
PIV.conditional_format("F24:F34", {"type": "cell", "criteria": ">", "value": 0.25, "format": F["green"]})
# Pivot 4: Gender
PIV.write("A38", "Pivot 4: Who buys? (Gender split: indicative, see Data_Quality note)", F["h2"])
header(PIV, 38, ["Gender", "Customers", "Revenue", "Share"])
for i, g in enumerate(["Female", "Male"]):
    r = 40 + i
    PIV.write(f"A{r}", g, F["txt"])
    wf(PIV, f"B{r}", f"=COUNTIF({C_('D')},A{r})", F["num"])
    wf(PIV, f"C{r}", f"=SUMIFS({R('O')},{R('H')},A{r})", F["num"])
    wf(PIV, f"D{r}", f"=C{r}/$C$42", F["pct"])
PIV.write("A42", "Total", F["tott"]); wf(PIV, "B42", "=SUM(B40:B41)", F["tot"]); wf(PIV, "C42", "=SUM(C40:C41)", F["tot"]); wf(PIV, "D42", "=SUM(D40:D41)", F["totp"])

# ---------------------------------------------------------------- PRODUCT ANALYSIS
PRD.write("A1", "Product analysis: which products earn the money, and at what margin?", F["title"])
PRD.write("A2", "Margin alert: document rule (yellow cell)", F["txt"]); PRD.write("B2", 0.10, F["inpp"])
PRD.write("A3", "Margin alert: business target (yellow cell)", F["txt"]); PRD.write("B3", 0.20, F["inpp"])
PRD.write("D2", "Red = list margin below the document's 10% rule (no product triggers it). Orange = below the 20% target. Change the yellow cells to re-flag.", F["sub"])
header(PRD, 5, ["ProductID", "ProductName", "Category", "UnitPrice", "CostPrice", "List Margin %", "Units Sold", "Revenue", "Profit",
                "Realised Margin %", "Revenue Rank", "Margin Flag"])
p0, p1 = 7, 6 + NP
for i, pid in enumerate(prod_perf.ProductID):
    r = 7 + i
    PRD.write(f"A{r}", pid, F["txt"])
    wf(PRD, f"B{r}", f"=INDEX({P_('B')},MATCH($A{r},{P_('A')},0))", F["txt"])
    wf(PRD, f"C{r}", f"=INDEX({P_('C')},MATCH($A{r},{P_('A')},0))", F["txt"])
    wf(PRD, f"D{r}", f"=INDEX({P_('D')},MATCH($A{r},{P_('A')},0))", F["num"])
    wf(PRD, f"E{r}", f"=INDEX({P_('E')},MATCH($A{r},{P_('A')},0))", F["num"])
    wf(PRD, f"F{r}", f"=(D{r}-E{r})/D{r}", F["pct"])
    wf(PRD, f"G{r}", f"=SUMIFS({R('L')},{R('I')},$A{r})", F["num"])
    wf(PRD, f"H{r}", f"=SUMIFS({R('O')},{R('I')},$A{r})", F["num"])
    wf(PRD, f"I{r}", f"=SUMIFS({R('Q')},{R('I')},$A{r})", F["num"])
    wf(PRD, f"J{r}", f"=I{r}/H{r}", F["pct"])
    wf(PRD, f"K{r}", f"=RANK(H{r},$H${p0}:$H${p1})", F["num"])
    wf(PRD, f"L{r}", f'=IF(F{r}<$B$2,"Critical",IF(F{r}<$B$3,"Watch","OK"))', F["txt"])
PRD.write(f"A{p1+1}", "Total", F["tott"])
for c in "GHI": wf(PRD, f"{c}{p1+1}", f"=SUM({c}{p0}:{c}{p1})", F["tot"])
wf(PRD, f"J{p1+1}", f"=I{p1+1}/H{p1+1}", F["totp"])
PRD.conditional_format(f"F{p0}:F{p1}", {"type": "cell", "criteria": "<", "value": "$B$2", "format": F["red"]})
PRD.conditional_format(f"F{p0}:F{p1}", {"type": "cell", "criteria": "<", "value": "$B$3", "format": F["orange"]})
PRD.conditional_format(f"H{p0}:H{p1}", {"type": "data_bar", "bar_color": "#9dc3e6", "bar_solid": True})
PRD.conditional_format(f"L{p0}:L{p1}", {"type": "cell", "criteria": "==", "value": '"Watch"', "format": F["orange"]})
PRD.conditional_format(f"L{p0}:L{p1}", {"type": "cell", "criteria": "==", "value": '"Critical"', "format": F["red"]})
for i, w in enumerate([24, 16, 13, 10, 10, 12, 10, 11, 10, 12, 10, 11]): PRD.set_column(i, i, w)
PRD.freeze_panes(6, 2)

# ---------------------------------------------------------------- CUSTOMER ANALYSIS (+ CLV)
CUS.write("A1", "Customer analysis: who are the high-value customers? (historical CLV = spend and profit in 2023)", F["title"])
CUS.write("A2", "Spend highlight threshold (document: > 5,000)", F["txt"]); CUS.write("B2", 5000, F["inp"])
CUS.write("A3", "High-value cut-off = top 20% spend (formula)", F["txt"])
c0, c1 = 7, 6 + NC
wf(CUS, "B3", f"=PERCENTILE($G${c0}:$G${c1},0.8)", F["num"])
CUS.write("D2", "Yellow = spent more than the threshold. NOTE: 39 of 50 customers exceed 5,000, so the document's rule does not separate customers here; "
                "use the Segment column (top/bottom 20%) for real prioritisation.", F["sub"])
header(CUS, 5, ["CustomerID", "Name", "City", "Gender", "Orders", "Units", "Revenue (CLV)", "Profit (CLV)", "Avg Order Value",
                "First Purchase", "Last Purchase", "Segment", "Rank"])
for i, cid in enumerate(cust_clv.CustomerID):
    r = 7 + i
    CUS.write(f"A{r}", cid, F["txt"])
    for col, src in [("B", "B"), ("C", "E"), ("D", "D")]:
        wf(CUS, f"{col}{r}", f"=INDEX({C_(src)},MATCH($A{r},{C_('A')},0))", F["txt"])
    wf(CUS, f"E{r}", f"=COUNTIF({R('E')},$A{r})", F["num"])
    wf(CUS, f"F{r}", f"=SUMIFS({R('L')},{R('E')},$A{r})", F["num"])
    wf(CUS, f"G{r}", f"=SUMIFS({R('O')},{R('E')},$A{r})", F["num"])
    wf(CUS, f"H{r}", f"=SUMIFS({R('Q')},{R('E')},$A{r})", F["num"])
    wf(CUS, f"I{r}", f"=G{r}/E{r}", F["num"])
    wf(CUS, f"J{r}", f"=MINIFS({R('B')},{R('E')},$A{r})", F["date"])
    wf(CUS, f"K{r}", f"=MAXIFS({R('B')},{R('E')},$A{r})", F["date"])
    wf(CUS, f"L{r}", f'=IF(G{r}>=$B$3,"High value",IF(G{r}<=PERCENTILE($G${c0}:$G${c1},0.2),"Low value","Mid value"))', F["txt"])
    wf(CUS, f"M{r}", f"=RANK(G{r},$G${c0}:$G${c1})", F["num"])
CUS.write(f"A{c1+1}", "Total", F["tott"])
for c in "EFGH": wf(CUS, f"{c}{c1+1}", f"=SUM({c}{c0}:{c}{c1})", F["tot"])
CUS.conditional_format(f"G{c0}:G{c1}", {"type": "cell", "criteria": ">", "value": "$B$2", "format": F["yellow"]})
CUS.conditional_format(f"G{c0}:G{c1}", {"type": "data_bar", "bar_color": "#9dc3e6", "bar_solid": False})
CUS.conditional_format(f"L{c0}:L{c1}", {"type": "cell", "criteria": "==", "value": '"High value"', "format": F["green"]})
CUS.conditional_format(f"L{c0}:L{c1}", {"type": "cell", "criteria": "==", "value": '"Low value"', "format": F["red"]})
for i, w in enumerate([24, 17, 11, 9, 8, 8, 13, 12, 13, 12, 12, 11, 7]): CUS.set_column(i, i, w)
CUS.freeze_panes(6, 2)

# ---------------------------------------------------------------- FORECAST
FCS.write("A1", "Revenue forecast: next 6 months (from Python, python/02_analysis.py)", F["title"])
FCS.merge_range("A2:F3", "INDICATIVE ONLY. The data covers 12 months, too short for a reliable ARIMA model. In a 3-month back-test the model "
                "(MAE 37,031) did NOT beat a plain average (MAE 18,375). Read the forecast as 'revenue probably stays in the 30k-70k/month range', "
                "not as a precise prediction.", F["ins"])
header(FCS, 4, ["Month", "Actual Revenue", "Forecast", "Lower 80%", "Upper 80%"])
for i, m in enumerate(monthly.Month):
    r = 6 + i
    FCS.write(f"A{r}", m, F["txt"]); wf(FCS, f"B{r}", f"=Pivot_Summaries!B{23+i}", F["num"])
    for c in "CDE": FCS.write_blank(f"{c}{r}", None, F["num"])
for i, row in enumerate(fc.itertuples(index=False)):
    r = 18 + i
    FCS.write(f"A{r}", row.Month, F["txt"]); FCS.write_blank(f"B{r}", None, F["num"])
    FCS.write_number(f"C{r}", row.ForecastRevenue, F["num"]); FCS.write_number(f"D{r}", row.Lower80, F["num"]); FCS.write_number(f"E{r}", row.Upper80, F["num"])
FCS.set_column("A:E", 14)
FCS.write("A25", f"Model used: {fc.Model.iloc[0]} (chosen by back-test among small ARIMA candidates; the document's ARIMA(2,1,2) scored worse: MAE 53,729).", F["sub"])


def forecast_chart():
    ch = wb.add_chart({"type": "line"})
    ch.add_series({"name": "Actual 2023", "categories": ["Forecast", 5, 0, 22, 0], "values": ["Forecast", 5, 1, 22, 1],
                   "line": {"color": NAVY, "width": 2.25}, "marker": {"type": "circle", "size": 5, "fill": {"color": NAVY}}})
    ch.add_series({"name": "Forecast", "categories": ["Forecast", 5, 0, 22, 0], "values": ["Forecast", 5, 2, 22, 2],
                   "line": {"color": ORANGE, "width": 2.25, "dash_type": "dash"}, "marker": {"type": "circle", "size": 5, "fill": {"color": ORANGE}}})
    ch.add_series({"name": "Lower 80%", "categories": ["Forecast", 5, 0, 22, 0], "values": ["Forecast", 5, 3, 22, 3], "line": {"color": "#bfbfbf", "width": 1, "dash_type": "round_dot"}})
    ch.add_series({"name": "Upper 80%", "categories": ["Forecast", 5, 0, 22, 0], "values": ["Forecast", 5, 4, 22, 4], "line": {"color": "#bfbfbf", "width": 1, "dash_type": "round_dot"}})
    ch.set_title({"name": "Revenue forecast (indicative: 12 months of history)", "name_font": {"size": 11}})
    ch.set_y_axis({"num_format": "#,##0", "major_gridlines": {"visible": True, "line": {"color": "#e0e0e0"}}})
    ch.set_x_axis({"num_font": {"rotation": -45}}); ch.set_legend({"position": "bottom"}); ch.show_blanks_as("gap")
    return ch


FCS.insert_chart("G4", forecast_chart(), {"x_scale": 1.0, "y_scale": 1.0})

# ---------------------------------------------------------------- RECONCILIATION
REC.write("A1", "Cross-tool reconciliation: Excel vs MySQL vs Python", F["title"])
REC.write("A2", "Excel = live formulas here. MySQL = sql/04_validation_metrics.sql output. Python = recomputed with pandas from the cleaned master table.", F["sub"])
header(REC, 3, ["Metric", "Excel (formula)", "MySQL", "Python", "Status"])
mm = master
py = {"Total Revenue": mm.Revenue.sum(), "Total Profit": mm.Profit.sum(), "Total Units": mm.Quantity.sum(), "Transactions": len(mm),
      "Active Customers": mm.CustomerID.nunique(), "Products Sold": mm.ProductID.nunique(),
      "Electronics Revenue": mm[mm.Category == "Electronics"].Revenue.sum(), "Accessories Revenue": mm[mm.Category == "Accessories"].Revenue.sum(),
      "Dubai Revenue": mm[mm.City == "Dubai"].Revenue.sum(), "Sep-2023 Revenue": mm[mm.YearMonth == "2023-09"].Revenue.sum(),
      "Customer C046 Revenue": mm[mm.CustomerID == "C046"].Revenue.sum(), "Camera DSLR Revenue": mm[mm.ProductName == "Camera DSLR"].Revenue.sum()}
xl = {"Total Revenue": f"=SUM({R('O')})", "Total Profit": f"=SUM({R('Q')})", "Total Units": f"=SUM({R('L')})",
      "Transactions": f"=COUNTA({R('A')})", "Active Customers": f"=SUMPRODUCT(--(Customer_Analysis!$E${c0}:$E${c1}>0))",
      "Products Sold": f"=SUMPRODUCT(--(Product_Analysis!$G${p0}:$G${p1}>0))",
      "Electronics Revenue": f'=SUMIFS({R("O")},{R("K")},"Electronics")', "Accessories Revenue": f'=SUMIFS({R("O")},{R("K")},"Accessories")',
      "Dubai Revenue": f'=SUMIFS({R("O")},{R("G")},"Dubai")', "Sep-2023 Revenue": f'=SUMIFS({R("O")},{R("C")},"2023-09")',
      "Customer C046 Revenue": f'=SUMIFS({R("O")},{R("E")},"C046")', "Camera DSLR Revenue": f'=SUMIFS({R("O")},{R("J")},"Camera DSLR")'}
for i, met in enumerate(py):
    r = 5 + i
    REC.write(f"A{r}", met, F["txt"]); wf(REC, f"B{r}", xl[met], F["num"])
    REC.write_number(f"C{r}", float(mysql[met]), F["num"]); REC.write_number(f"D{r}", float(py[met]), F["num"])
    wf(REC, f"E{r}", f'=IF(AND(ROUND(B{r},2)=ROUND(C{r},2),ROUND(B{r},2)=ROUND(D{r},2)),"MATCH","CHECK")', F["txt"])
last = 4 + len(py)
REC.write(f"A{last+2}", "Overall", F["tott"]); wf(REC, f"B{last+2}", f'=IF(COUNTIF(E5:E{last},"MATCH")={len(py)},"ALL {len(py)} CHECKS MATCH","DISCREPANCY FOUND")', F["tott"])
REC.conditional_format(f"E5:E{last}", {"type": "cell", "criteria": "==", "value": '"MATCH"', "format": F["green"]})
REC.conditional_format(f"E5:E{last}", {"type": "cell", "criteria": "==", "value": '"CHECK"', "format": F["red"]})
REC.write(f"A{last+4}", "Power BI: after building the model, the same KPI cards must show these values (see powerbi/POWER_BI_GUIDE.md, validation table).", F["sub"])
REC.set_column("A:A", 26); REC.set_column("B:E", 16)

# ---------------------------------------------------------------- DATA QUALITY (live checks + cleaning log)
DQ.write("A1", "Data quality checks (live formulas) and cleaning log", F["title"])
header(DQ, 2, ["Table", "Rows", "Missing cells", "Duplicate IDs", "Result"])
dq = [("Customers_Clean", NC, "A", "F"), ("Products_Clean", NP, "A", "E"), ("Sales_Clean", NS, "A", "E")]
for i, (t, n, c0_, c1_) in enumerate(dq):
    r = 4 + i
    DQ.write(f"A{r}", t, F["txt"])
    wf(DQ, f"B{r}", f"=COUNTA({t}!A2:A{n+1})", F["num"])
    wf(DQ, f"C{r}", f"=COUNTBLANK({t}!{c0_}2:{c1_}{n+1})", F["num"])
    wf(DQ, f"D{r}", f"=ROWS({t}!A2:A{n+1})-SUMPRODUCT(1/COUNTIF({t}!A2:A{n+1},{t}!A2:A{n+1}))", F["num"])
    wf(DQ, f"E{r}", f'=IF(AND(C{r}=0,ROUND(D{r},0)=0),"Clean","Review")', F["txt"])
DQ.write("A8", "Sales dated before the customer's JoinDate (flagged, kept)", F["txt"]); wf(DQ, "D8", f'=COUNTIF({R("T")},"Yes")', F["num"])
DQ.write("A9", "Sales rows with no matching customer / product", F["txt"])
wf(DQ, "D9", f"=SUMPRODUCT(--ISNA(MATCH(Sales_Clean!B2:B{NS+1},{C_('A')},0)))+SUMPRODUCT(--ISNA(MATCH(Sales_Clean!C2:C{NS+1},{P_('A')},0)))", F["num"])
log = [
    "CLEANING LOG (python/01_clean_data.py is the source of truth)",
    "1. Dates: raw M/D/YYYY text converted to real dates (Excel: DATEVALUE). 0 unparseable values.",
    "2. Text: TRIM + PROPER applied to Name, City, Gender, ProductName, Category. 0 values needed changing.",
    "3. Duplicates: checked on full rows and on ID columns (Data > Remove Duplicates equivalent). 0 found.",
    "4. Missing values: none. The document's rules (Age = median, UnitPrice = check last 10 values) were therefore NOT triggered. Nothing was imputed.",
    "5. Validation: ages 25-45, quantities 1-5, every UnitPrice > CostPrice > 0, no orphan CustomerID/ProductID in Sales.",
    "6. ANOMALY kept and flagged: 24 of 300 sales are dated before the customer's JoinDate (Combined_Report column T).",
    "7. DATA NOTE: some Name/Gender pairs look inconsistent (e.g. C001 'Ahmed Khan' = Female). Kept as supplied; treat gender insights as indicative.",
    "8. Only one year of sales (2023-01-01 to 2023-12-31): limits forecasting and any year-over-year analysis.",
]
for i, line in enumerate(log): DQ.write(11 + i, 0, line, F["h2"] if i == 0 else None)
DQ.set_column("A:A", 58); DQ.set_column("B:E", 14)
DQ.conditional_format("E4:E6", {"type": "cell", "criteria": "==", "value": '"Clean"', "format": F["green"]})

# ---------------------------------------------------------------- DASHBOARD
DASH.hide_gridlines(2); DASH.set_column("A:L", 12.5)
DASH.merge_range("A1:L1", "Sales & Customer Behavior Analysis: Electronics Retail Store (2023)", F["title"])
DASH.set_row(0, 28)
kpis = [("Total Revenue (QAR)", f"=SUM({R('O')})", "kv"), ("Total Profit (QAR)", f"=SUM({R('Q')})", "kv"),
        ("Profit Margin", f"=SUM({R('Q')})/SUM({R('O')})", "kvp"), ("Units Sold", f"=SUM({R('L')})", "kv"),
        ("Active Customers", f"=SUMPRODUCT(--(Customer_Analysis!$E${c0}:$E${c1}>0))", "kv"),
        ("Avg Transaction (QAR)", f"=SUM({R('O')})/COUNTA({R('A')})", "kv")]
for i, (lab, frm, st) in enumerate(kpis):
    c = i * 2
    DASH.merge_range(2, c, 2, c + 1, lab, F["kl"]); DASH.merge_range(3, c, 3, c + 1, "", F[st])
    wf(DASH, xl_rowcol_to_cell(3, c), frm, F[st])
DASH.set_row(2, 22); DASH.set_row(3, 34)
DASH.merge_range("A6:L6", "What the data says", F["h2"])
ins = ["1. Electronics is 92% of revenue (571k of 620k) but earns only a 19% margin; Accessories earn 34% on just 8% of revenue.",
       "2. Revenue peaks in Aug-Sep (75.7k, 84.8k), then drops 47% in October and bottoms in November (26.5k): a strong seasonal swing, not a steady trend.",
       "3. The top 10 customers generate 37.5% of revenue; Camera DSLR alone is 19% (117k), so a few big-ticket items drive results.",
       "4. All five cities are close (90k-154k); Dubai leads on revenue, Muscat has the most customers (13) but the lowest revenue per customer."]
for i, t in enumerate(ins): DASH.merge_range(6 + i, 0, 6 + i, 11, t, F["ins"]); DASH.set_row(6 + i, 30)


def mk(kind, title, cats, vals, name, color, sub=None, **extra):
    ch = wb.add_chart({"type": kind, **({"subtype": sub} if sub else {})})
    ch.add_series({"name": name, "categories": cats, "values": vals, "fill": {"color": color}, "line": {"color": color}, **extra})
    ch.set_title({"name": title, "name_font": {"size": 11, "bold": True}}); ch.set_legend({"none": True})
    ch.set_size({"width": 560, "height": 300}); return ch


# 1 monthly revenue & profit line
c1_ = wb.add_chart({"type": "line"})
c1_.add_series({"name": "Revenue", "categories": ["Pivot_Summaries", 22, 0, 33, 0], "values": ["Pivot_Summaries", 22, 1, 33, 1],
                "line": {"color": NAVY, "width": 2.5}, "marker": {"type": "circle", "size": 6, "fill": {"color": NAVY}}})
c1_.add_series({"name": "Profit", "categories": ["Pivot_Summaries", 22, 0, 33, 0], "values": ["Pivot_Summaries", 22, 2, 33, 2],
                "line": {"color": ORANGE, "width": 2.5}, "marker": {"type": "circle", "size": 6, "fill": {"color": ORANGE}}})
c1_.set_title({"name": "When do we sell? Monthly revenue & profit", "name_font": {"size": 11}}); c1_.set_legend({"position": "bottom"})
c1_.set_y_axis({"num_format": "#,##0", "major_gridlines": {"visible": True, "line": {"color": "#e0e0e0"}}}); c1_.set_x_axis({"num_font": {"rotation": -45}})
c1_.set_size({"width": 560, "height": 300})
# 2 category revenue
c2_ = mk("column", "Which category earns the revenue?", ["Pivot_Summaries", 15, 0, 16, 0], ["Pivot_Summaries", 15, 2, 16, 2], "Revenue", NAVY,
         data_labels={"value": True, "num_format": "#,##0"})
c2_.set_y_axis({"num_format": "#,##0", "major_gridlines": {"visible": True, "line": {"color": "#e0e0e0"}}})
# 3 revenue by city
c3_ = mk("bar", "Which city sells the most? Revenue by city", ["Pivot_Summaries", 5, 0, 9, 0], ["Pivot_Summaries", 5, 2, 9, 2], "Revenue", "#2e75b6",
         data_labels={"value": True, "num_format": "#,##0"})
c3_.set_y_axis({"reverse": True}); c3_.set_x_axis({"num_format": "#,##0"})
# 4 customer distribution pie
c4_ = wb.add_chart({"type": "pie"})
c4_.add_series({"name": "Customers", "categories": ["Pivot_Summaries", 5, 0, 9, 0], "values": ["Pivot_Summaries", 5, 1, 9, 1],
                "data_labels": {"category": True, "percentage": True, "leader_lines": True, "position": "outside_end"},
                "points": [{"fill": {"color": c}} for c in [NAVY, ORANGE, "#2e75b6", "#a5a5a5", "#70ad47"]]})
c4_.set_title({"name": "Where do our customers live?", "name_font": {"size": 11}}); c4_.set_legend({"none": True}); c4_.set_size({"width": 560, "height": 300})
# 5 top 10 customers
c5_ = mk("bar", "Who are the top 10 customers? (revenue)", ["Customer_Analysis", 6, 1, 15, 1], ["Customer_Analysis", 6, 6, 15, 6], "Revenue", ORANGE,
         data_labels={"value": True, "num_format": "#,##0"})
c5_.set_y_axis({"reverse": True}); c5_.set_x_axis({"num_format": "#,##0"})
# 6 forecast
c6_ = forecast_chart(); c6_.set_size({"width": 560, "height": 300})
for i, ch in enumerate([c1_, c2_, c3_, c4_, c5_, c6_]):
    DASH.insert_chart(11 + (i // 2) * 16, 0 if i % 2 == 0 else 6, ch, {"x_offset": 4, "y_offset": 4})
DASH.write(11 + 3 * 16, 0, "Sources: Pivot_Summaries, Customer_Analysis, Forecast. Forecast is indicative only (12 months of history).", F["sub"])

# ---------------------------------------------------------------- README
README.set_column("A:A", 24); README.set_column("B:B", 110)
README.write("A1", "Sales & Customer Behavior Analysis: Excel workbook", F["title"])
rows = [
    ("Purpose", "Excel part of the 4-tool project (Excel, MySQL, Python, Power BI). Built from the supplied Customers/Products/Sales CSVs. No data was invented."),
    ("Dashboard", "KPI cards, 6 charts (each answers one business question) and 4 headline insights."),
    ("Pivot_Summaries", "Pivot-style tables: Pivot 1 Sales by City, Pivot 2 Category Performance, Pivot 3 Monthly Revenue, Pivot 4 Gender. Live SUMIFS formulas."),
    ("Product_Analysis", "Per-product units, revenue, profit, margin, rank. Conditional formatting flags thin margins (thresholds in yellow cells)."),
    ("Customer_Analysis", "Per-customer orders, revenue, profit (historical CLV), segment and rank. Highlights customers above the spend threshold."),
    ("Forecast", "6-month revenue forecast from Python (ARIMA) with the honest limitation note."),
    ("Reconciliation", "Excel vs MySQL vs Python control totals with MATCH/CHECK status."),
    ("Data_Quality", "Live data-quality checks and the cleaning log."),
    ("Combined_Report", "The combined sales report: one row per sale with customer + product fields and Revenue/Cost/Profit/Margin (formulas)."),
    ("*_Clean sheets", "Cleaned input data (dates as real dates, text standardised). Raw files are untouched in data/raw."),
    ("Colour code", "Navy tab = report, blue = analysis, green = checks, grey = data. Yellow cells = inputs you can change."),
    ("Real PivotTables", "Python cannot author native PivotTables, so the summaries are formula-based. To add a native one in 2 minutes: Combined_Report > "
                         "Insert > PivotTable. Pivot 1: Rows=City, Values=Sum of Revenue. Pivot 2: Rows=Category, Values=Sum of Quantity (document's version) and Sum of Revenue. "
                         "Add slicers for Year, City, Gender. The totals must equal the Pivot_Summaries tables."),
    ("Refreshing", "If you add sales rows, extend Sales_Clean and copy the Combined_Report formulas down; ranges use fixed rows (2-301)."),
    ("Currency", "The project document uses QAR; the data has no currency field, so amounts are labelled QAR."),
]
for i, (a, b) in enumerate(rows):
    README.write(2 + i, 0, a, f(bold=True, valign="top")); README.write(2 + i, 1, b, F["wrap"])
    README.set_row(2 + i, 32 if len(b) > 105 else 18)
README.set_row(13, 62)
for ws in (README, DASH, PIV, PRD, CUS, FCS, REC, DQ):       # print-friendly: landscape, one page wide
    ws.set_landscape(); ws.fit_to_pages(1, 0); ws.set_paper(9)
wb.close()
print("Saved", OUT_FILE)
