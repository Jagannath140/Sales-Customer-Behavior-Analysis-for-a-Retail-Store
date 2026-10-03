# Project Documentation

## 1. Business problem
An electronics retailer wants to understand sales performance, customer purchasing behaviour, profitability and future demand, using Excel (cleaning and reporting), MySQL (storage and querying), Python (analysis and forecasting) and Power BI (dashboard).

## 2. Data
| File | Rows | Columns |
|---|---|---|
| Customers.csv | 50 | CustomerID, Name, Age, Gender, City, JoinDate |
| Products.csv | 30 | ProductID, ProductName, Category, UnitPrice, CostPrice |
| Sales.csv | 300 | SalesID, CustomerID, ProductID, Quantity, SaleDate |

Relationships: `Sales.CustomerID > Customers.CustomerID`, `Sales.ProductID > Products.ProductID` (many-to-one). The sales table has no price, so **Revenue = Quantity x UnitPrice** and **Profit = Quantity x (UnitPrice - CostPrice)**, exactly as the project document defines them. Currency is not in the data; the document uses QAR, so amounts are labelled QAR.

### What the data supports, and what it does not
| Requirement in the document | Supported? | What I did |
|---|---|---|
| Revenue, profit, category/product/customer/city analysis | Yes | Done in all four tools |
| Customer Lifetime Value | Partly | Only 12 months and no churn/retention data, so a forward-looking CLV would need invented assumptions. I report **historical CLV** (revenue and profit per customer in the window). |
| Forecast with ARIMA/Prophet | Weakly | 12 monthly points is far below what ARIMA needs. I fitted ARIMA anyway (as required), back-tested it, and labelled the result **indicative** (section 6). |
| Fix missing Age (median) / UnitPrice (last 10 values) | Not needed | There are **no missing values**; nothing was imputed. |
| "Highlight customers > 5,000 QAR" | Runs, but weak | 39 of 50 customers exceed it, so it doesn't discriminate. I kept it (yellow cell, editable) and added top/bottom-20% segments. |
| "Highlight products with margin < 10%" | Runs, flags nothing | The lowest margin is 16.7%. I kept the rule and added a 20% "Watch" target that flags 5 products. |
| Excel native PivotTables | No (tooling) | Built formula-based pivot-style tables (SUMIFS) that update live; the README sheet gives the 2-minute steps for native PivotTables. |
| Power BI report | Not buildable here | Delivered data files, model, all DAX, page-by-page spec and a validation table. |

## 3. Data quality and cleaning (`python/01_clean_data.py`)
| Check | Result |
|---|---|
| Duplicate rows / duplicate IDs | 0 |
| Missing values | 0 |
| Date format | M/D/YYYY text > real dates (ISO in cleaned files); 0 unparseable |
| Text (trim, consistent case) | 0 values needed changing |
| Numeric validity | ages 25-45, quantity 1-5, every UnitPrice > CostPrice > 0 |
| Referential integrity | 0 orphan CustomerID/ProductID in Sales |
| Coverage | all 50 customers bought; all 30 products sold; 2023-01-01 to 2023-12-31 |
| **Anomaly (flagged, kept)** | **24 of 300 sales are dated before the customer's JoinDate** (looks like independently generated tables). Flag column `SaleBeforeJoin`. |
| **Data note** | Some Name/Gender pairs look inconsistent (e.g. C001 "Ahmed Khan" = Female). Kept as supplied; gender insights are indicative. |

Raw files are preserved in `data/raw`. Decisions are logged in `documentation/data_quality_report.txt`.

## 4. Methods
- **Python** (`02_analysis.py`): load > merge (`validate="m:1"` so a bad join can't silently duplicate rows) > EDA > revenue, product, category, city, customer and profit analysis > historical CLV and quintile-style segments > monthly trend > ARIMA forecast. 8 charts in `outputs/charts`.
- **MySQL** (`sql/`): 3 tables with primary/foreign keys and CHECK constraints, a reusable view `vw_sales_detail` (revenue/profit logic in one place), 13 business-question queries (joins, `GROUP BY`, `LAG`, `RANK`, `NTILE`), 6 stored procedures.
- **Excel**: clean-data sheets > `Combined_Report` (INDEX/MATCH lookups + Revenue/Profit formulas) > pivot-style summaries, product and customer analysis (SUMIFS, RANK, PERCENTILE, MINIFS/MAXIFS) > dashboard (6 charts), conditional formatting, forecast sheet, data-quality and reconciliation sheets.
- **Power BI**: star schema, SUMX-based measures, 4-page spec (`powerbi/`).

## 5. Results
| Metric | Value |
|---|---|
| Revenue / Profit / Margin | 619,615 / 125,915 / 20.3% |
| Units / Transactions / Avg transaction | 920 / 300 / 2,065 |
| Electronics | 571,070 revenue (92.2%), 109,230 profit, 19.1% margin, 410 units |
| Accessories | 48,545 revenue (7.8%), 16,685 profit, 34.4% margin, 510 units |
| Best month / worst month | Sep 84,830 / Nov 26,520 |
| Top product by revenue | Camera DSLR 117,000 (18.9% of revenue, 39 units) |
| Top sellers by units | Mouse 60, USB Cable 58, SSD 512GB 50 |
| City revenue | Dubai 154,160; Doha 138,475; Muscat 127,775; Riyadh 109,175; Abu Dhabi 90,030 |
| Revenue per customer | Doha 13,848; Riyadh 13,647; Abu Dhabi 12,861; Dubai 12,847; **Muscat 9,829** |
| Top 10 customers | 232,080 = 37.5% of revenue (top customer Rohit Sharma, 29,060) |
| High-value cut-off (top 20%) | 19,288 revenue per customer |
| Segments | 10 high (37.5% of revenue) / 30 mid (57.6%) / 10 low (4.9%) |
| Average historical CLV (profit) | 2,518 per customer |

## 6. Forecast and its limits
12 monthly points, strong seasonality (Aug-Sep spike, Nov trough), no repeating cycle to learn from. Back-test on the last 3 months (train on 9):

| Model | Back-test MAE |
|---|---|
| Plain average of training months (baseline) | **18,375** |
| ARIMA(1,0,0) (selected among ARIMA candidates) | 37,031 |
| ARIMA(0,1,1) | 46,148 |
| ARIMA(1,1,1) | 46,647 |
| ARIMA(2,1,2) (the document's order) | 53,729 |

**No ARIMA model beat the simple average.** The published forecast (ARIMA(1,0,0), 6 months) is ~47k rising to ~52k per month with an 80% interval of roughly 29k-73k. Read it as "revenue likely stays in the 30k-70k range", not as a prediction. With at least 2-3 years of monthly data, seasonal models (SARIMA/Prophet) become appropriate. Prophet was not used: it needs more history and adds a dependency for no gain here.

## 7. Business insights
1. **Concentration in electronics:** 92% of revenue, but only a 19% margin. Five products (Camera DSLR, Drone, Laptop, Projector, Graphic Card) are **54% of all revenue** at margins of 16.7-18.2%.
2. **Accessories are the margin engine but tiny:** 34% margin, 55% of units (510 of 920), yet 8% of revenue.
3. **Strong seasonality:** Aug-Sep = 26% of annual revenue; October falls 47% and November is the year's low (26.5k).
4. **Customer concentration:** top 10 customers = 37.5% of revenue; bottom 20% = 4.9%.
5. **Muscat underperforms per customer:** most customers (13) but 9.8k per customer versus 12.8-13.8k elsewhere.
6. **Possible churn risk:** 3 of the 10 high-value customers (Monica Singh, Amy Parker, Alisha Verma; 72k revenue together) made no purchase after September (Alisha Verma after March).

## 8. Recommendations
1. **Fix electronics margin:** renegotiate supplier cost or adjust price on the five low-margin big-ticket items. A 1-point margin gain on electronics is about 5.7k profit.
2. **Bundle accessories with electronics** (laptop + mouse/stand, camera + tripod): accessories carry ~34% margin and are bought in volume.
3. **Plan for seasonality:** stock and staff for Aug-Sep; run promotions or B2B pushes in Apr and Nov, the weakest months.
4. **Protect top customers:** loyalty terms or an account contact for the top 20%; run a win-back contact for the three customers who went quiet.
5. **Grow Muscat revenue per customer** with upsell offers (it has the customer base but the smallest baskets).
6. **Improve data:** capture 2-3 years of history, returns, inventory and a correct gender field, so forecasting and CLV can be modelled properly.

## 9. Testing and validation (what was actually run)
| Area | Evidence |
|---|---|
| Python | All four scripts run end to end; a **clean-copy rebuild from raw data reproduced identical cleaned files, KPIs and forecast**. A bug found during that test (missing `documentation/` folder on a fresh clone) was fixed. |
| MySQL | Executed on MariaDB 10.11 (MySQL-compatible): schema, load (50/30/300 rows), all 13 queries and all 6 stored procedures ran without errors and gave the expected results. Not run on Oracle MySQL 8 itself; nothing vendor-specific is used except standard window functions (need MySQL 8+). |
| Excel | 7,107 formula cells recalculated in LibreOffice: **0 errors**; segment counts, flags and totals equal Python. Dashboard rendered and visually checked. Not opened in desktop Excel. |
| Cross-check | 12 control totals (revenue, profit, units, transactions, customers, products, both categories, Dubai, Sep-2023, top customer, top product) **match across Excel, MySQL and Python** (`Reconciliation` sheet). |
| Power BI | **Not run** (needs Power BI Desktop). DAX mirrors verified calculations; expected values are in the guide's validation table. |

## 10. Limitations
- One year of data: no year-over-year view, weak forecast, historical-only CLV.
- Synthetic-looking data: 24 sales before the join date; inconsistent Name/Gender pairs; quantities limited to 1-5 and near-identical margins across cities (19.9-20.8%).
- Native Excel PivotTables and the `.pbix` file were not produced (tooling); both are replaced by documented equivalents.
- No inventory, returns, discounts or costs beyond `CostPrice`, so "profit" is gross profit.

## 11. File guide
| File | What it does | Where | How to use |
|---|---|---|---|
| `data/raw/*.csv` | Original data, untouched | data/raw | Source of truth; never edit |
| `data/cleaned/Customers.csv, Products.csv, Sales.csv` | Cleaned tables (ISO dates) | data/cleaned | Import into MySQL / Power BI |
| `data/cleaned/Sales_Master.csv` | One row per sale with customer, product, Revenue, Cost, Profit, flags | data/cleaned | Quick analysis in any tool |
| `python/01_clean_data.py` | Validates and cleans, builds master table | python | `python python/01_clean_data.py` |
| `python/02_analysis.py` | EDA, analysis, CLV, forecast, charts | python | run after 01 |
| `python/03_build_excel.py` | Builds the Excel workbook with formulas and charts | python | run after 01, 02 and the SQL validation export |
| `python/04_export_powerbi.py` | Copies Power BI input files and builds the forecast feed | python | run after 02 |
| `sql/01_schema_and_load.sql` | Creates DB, tables, keys, loads CSVs, creates view | sql | run first, from project root |
| `sql/02_analysis_queries.sql` | 13 business queries | sql | run or copy queries individually |
| `sql/03_stored_procedures.sql` | 6 summary-report procedures | sql | then `CALL sp_...` |
| `sql/04_validation_metrics.sql` | Control totals for reconciliation | sql | export to `outputs/mysql_validation_metrics.tsv` |
| `excel/Retail_Sales_Analysis.xlsx` | Workbook: dashboard, summaries, analysis, checks, data | excel | open in Excel; start at README sheet |
| `powerbi/POWER_BI_GUIDE.md`, `measures.dax`, `data/*` | Build guide, DAX, import files | powerbi | follow the guide |
| `outputs/*.csv, charts/*.png, summary_kpis.json` | Result tables and charts | outputs | use in reports/slides |
| `documentation/data_quality_report.txt` | Auto-generated cleaning log | documentation | reference |
| `documentation/INTERVIEW_GUIDE.md` | Resume text, pitch, Q&A | documentation | study before interviews |
| `requirements.txt`, `run_all.sh` | Dependencies and one-command rebuild | root | `pip install -r requirements.txt; bash run_all.sh` |
