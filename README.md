# Sales & Customer Behavior Analysis for a Retail Store

End-to-end analytics project for an electronics retailer using **Excel, MySQL, Python and Power BI**, built from the supplied `Customers.csv`, `Products.csv` and `Sales.csv` (50 customers, 30 products, 300 sales, Jan-Dec 2023).

**Headline results:** revenue **619,615 QAR**, profit **125,915** (20.3% margin), 920 units, 50 active customers. Electronics = 92% of revenue at a 19% margin; Accessories = 8% of revenue at 34%. Top 10 customers = 37.5% of revenue. Revenue peaks Aug-Sep and collapses in Oct-Nov.

> The same numbers come out of Excel, MySQL and Python (12 control totals reconciled: see `excel/Retail_Sales_Analysis.xlsx` > `Reconciliation`). Power BI could not be built in this environment; a full build guide with expected values is provided.

## Folder structure
```
Retail_Sales_Customer_Analytics/
├── data/raw/                 original CSVs (never modified)
├── data/cleaned/             cleaned CSVs + Sales_Master.csv (merged, with Revenue/Profit)
├── python/                   01_clean_data, 02_analysis, 03_build_excel, 04_export_powerbi
├── sql/                      01_schema_and_load, 02_analysis_queries, 03_stored_procedures, 04_validation_metrics
├── excel/                    Retail_Sales_Analysis.xlsx
├── powerbi/                  POWER_BI_GUIDE.md, measures.dax, data/ (files to import)
├── outputs/                  result tables (CSV), charts/ (PNG), summary_kpis.json
├── documentation/            PROJECT_DOCUMENTATION.md, INTERVIEW_GUIDE.md, data_quality_report.txt
├── requirements.txt
└── run_all.sh                rebuilds everything from data/raw
```

## Quick start
```bash
pip install -r requirements.txt
bash run_all.sh            # cleaning > analysis > (MySQL if installed) > Excel > Power BI files
```
Or step by step:
| Step | Command (from project root) | Produces |
|---|---|---|
| 1 | `python python/01_clean_data.py` | `data/cleaned/*`, `documentation/data_quality_report.txt` |
| 2 | `python python/02_analysis.py` | `outputs/*.csv`, `outputs/charts/*.png`, `summary_kpis.json` |
| 3 | `mysql -u root -p --local-infile=1 < sql/01_schema_and_load.sql` then `02_analysis_queries.sql`, `03_stored_procedures.sql`, `04_validation_metrics.sql` (see below) | database `retail_analytics` |
| 4 | `python python/03_build_excel.py` | `excel/Retail_Sales_Analysis.xlsx` |
| 5 | `python python/04_export_powerbi.py` | `powerbi/data/*` |
| 6 | Follow `powerbi/POWER_BI_GUIDE.md` | the Power BI report |

**MySQL:** run the `.sql` files from the project root (the `LOAD DATA LOCAL INFILE` paths are relative). In MySQL Workbench, enable `OPT_LOCAL_INFILE=1` in the connection's Advanced tab, or use the Table Data Import Wizard on the three cleaned CSVs. Stored procedures: `CALL sp_monthly_revenue_report(2023); CALL sp_top_customers(10); CALL sp_category_summary(); CALL sp_executive_summary();`
`04_validation_metrics.sql` feeds the Excel reconciliation: `mysql -u root -B retail_analytics < sql/04_validation_metrics.sql > outputs/mysql_validation_metrics.tsv`.

**Excel note:** the shipped workbook has cached values. If you regenerate it with `03_build_excel.py`, Excel recalculates all formulas when you open it (values appear immediately in Excel/Sheets; quick-preview tools may show zeros until recalculated).

## Read next
- `documentation/PROJECT_DOCUMENTATION.md`: method, results, insights, recommendations, limitations, test evidence, file-by-file guide
- `documentation/INTERVIEW_GUIDE.md`: resume text, 2-minute pitch, concepts, Q&A
