#!/usr/bin/env bash
# Rebuilds everything from data/raw. Run from the project root:  bash run_all.sh
# MySQL steps run only if the `mysql` client is available and a local server accepts root (edit -u/-p as needed).
set -e
python3 python/01_clean_data.py
python3 python/02_analysis.py
if command -v mysql >/dev/null 2>&1; then
  mysql -u root --local-infile=1 < sql/01_schema_and_load.sql
  mysql -u root < sql/03_stored_procedures.sql
  mysql -u root -B retail_analytics < sql/04_validation_metrics.sql > outputs/mysql_validation_metrics.tsv
else
  echo "mysql client not found: skipping SQL steps (outputs/mysql_validation_metrics.tsv must exist for step 3)"
fi
python3 python/03_build_excel.py
python3 python/04_export_powerbi.py
echo "Done. Open excel/Retail_Sales_Analysis.xlsx, then build the Power BI report with powerbi/POWER_BI_GUIDE.md"
