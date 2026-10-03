# Power BI Guide: Sales & Customer Behavior Dashboard

**Honest note up front:** a `.pbix` file can only be created inside Power BI Desktop (Windows). I can't author or run it in this environment, so I've prepared everything you need to build it in about 30-40 minutes, and every number you should see is listed in the validation table below. The DAX was written to match calculations that *were* run and verified in MySQL, Python and Excel.

## Files you need (all in this folder)
| File | What it is |
|---|---|
| `data/Customers.csv`, `Products.csv`, `Sales.csv` | Cleaned tables (ISO dates) to import |
| `data/Forecast_Trend.csv` | Monthly actual revenue + 6-month forecast (from Python) for the forecast chart |
| `measures.dax` | All DAX measures + the Calendar table, ready to copy/paste |

## Step 1: Import
Home > Get data > Text/CSV, load the four files, then in Power Query check the types:
`SaleDate`, `JoinDate`, `MonthStart` = **Date**; `Quantity`, `Age` = Whole number; `UnitPrice`, `CostPrice`, `ActualRevenue`, `ForecastRevenue`, `Lower80`, `Upper80` = Decimal. Close & Apply.

## Step 2: Model (Model view)
Create the Calendar table (code at the bottom of `measures.dax`), mark it as a date table (`Calendar[Date]`), then set these relationships. All are **many-to-one, single direction** (dimension filters fact):

| From (many) | To (one) |
|---|---|
| `Sales[CustomerID]` | `Customers[CustomerID]` |
| `Sales[ProductID]` | `Products[ProductID]` |
| `Sales[SaleDate]` | `Calendar[Date]` |

`Forecast_Trend` stays **unrelated** (it already has its own months). This is a star schema: one fact table (Sales), three dimensions.

## Step 3: Measures
Create an empty table `_Measures`, then add every measure from `measures.dax`. Format Revenue/Profit/Cost as whole numbers with thousands separators, `Profit Margin %` as percentage.

**Important correction to the project document:** its DAX `Total Revenue = SUM(Sales[Quantity] * Products[UnitPrice])` will not work, because `SUM` takes a single column. The row-by-row multiplication needs an iterator, hence `SUMX(Sales, Sales[Quantity] * RELATED(Products[UnitPrice]))`. This is a classic interview question.

## Step 4: Report pages

**Page 1: Executive Overview** (story: *how are we doing, when, where, what*)
| Visual | Fields | Business question |
|---|---|---|
| 6 KPI cards | Total Revenue, Total Profit, Profit Margin %, Total Units, Total Customers, Avg Order Value | How are we doing? |
| Line chart | Axis `Calendar[Year-Month]`; Values Total Revenue, Total Profit | When do we sell? |
| Column chart | Axis `Products[Category]`; Values Total Revenue | Which category earns the revenue? |
| Bar chart (stacked) | Axis `Customers[City]`; Values Total Revenue; Legend `Products[Category]` | Where do we sell? |
| Donut | Legend `Customers[Gender]`; Values Total Revenue | Who buys? (indicative, see data note) |
| Slicers | `Calendar[Year]`, `Customers[City]`, `Customers[Gender]`, `Products[Category]` | Filter everything |

**Page 2: Customers**
| Visual | Fields | Question |
|---|---|---|
| Table | `Customers[Name]`, `City`, Total Sales (Transactions), Total Revenue, Customer CLV (Profit), Customer Segment. Visual-level filter: **Top N = 10 by Total Revenue** | Who are the top 10 customers? |
| Bar | Axis `Customers[City]`; Values Total Revenue / Total Customers | Which city is most valuable per customer? |
| Card | High Value Cut-off (Top 20%) | What spend makes a "high-value" customer? |

**Page 3: Products & Profit**
| Visual | Fields | Question |
|---|---|---|
| Bar (Top N = 10) | `Products[ProductName]`, Total Revenue | What earns the most? |
| Scatter | X Total Revenue, Y Profit Margin %, Details `ProductName`, Legend `Category` | Which big sellers have thin margins? |
| Table | ProductName, Total Units, Total Revenue, Total Profit, Profit Margin % (conditional format: margin < 20% orange) | Where should pricing be reviewed? |
| Column | Axis Category; Values Profit Margin % | Which category is more profitable per riyal? |

**Page 4: Forecast**
| Visual | Fields |
|---|---|
| Line chart | Axis `Forecast_Trend[MonthStart]` (Date, continuous); Values Actual Revenue (trend), Forecast Revenue (trend), Forecast Lower 80%, Forecast Upper 80% |
| Text box | "Indicative only: 12 months of history; a 3-month back-test showed the model did not beat a plain average. Treat as a range (about 30k-70k/month), not a point prediction." |

Keep the limitation text on the page; it is part of being honest with stakeholders.

## Validation: Power BI must show these numbers
Clear all slicers first. If a number differs, check relationships and column data types (most common cause: `SaleDate` left as text).

| Measure / slice | Expected |
|---|---|
| Total Revenue | **619,615** |
| Total Profit | **125,915** |
| Profit Margin % | **20.3%** |
| Total Units | **920** |
| Total Sales (Transactions) | **300** |
| Total Customers | **50** |
| Electronics / Accessories revenue | **571,070 / 48,545** |
| Electronics / Accessories profit | 109,230 / 16,685 |
| Revenue by city | Dubai 154,160 · Doha 138,475 · Muscat 127,775 · Riyadh 109,175 · Abu Dhabi 90,030 |
| Revenue by gender | Female 359,320 · Male 260,295 |
| Sep-2023 revenue (peak) | 84,830 |
| Top customer | Rohit Sharma (C046): revenue 29,060, profit 5,845 |
| High Value Cut-off (Top 20%) | 19,288 |
| Segment counts (High / Mid / Low) | 10 / 30 / 10 |

These match MySQL, Python and Excel (see `Reconciliation` sheet in the workbook).

## Notes
- The Year slicer has one value (2023) because the data covers one year; it is kept because the document asks for it and it works as soon as more years are loaded.
- 24 of 300 sales are dated before the customer's JoinDate (data anomaly, see documentation). Nothing was deleted.
- Power BI also has a built-in Analytics-pane forecast; we feed the Python forecast instead so all tools tell the same story.
