"""
04_export_powerbi.py
Purpose : Prepare the files Power BI imports: the 3 cleaned tables + a Forecast_Trend table (actual + forecast by month).
Run     : python python/04_export_powerbi.py   (after 01 and 02, from the project root)
Outputs : powerbi/data/{Customers,Products,Sales,Forecast_Trend}.csv
"""
import shutil
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
DST = ROOT / "powerbi/data"
DST.mkdir(parents=True, exist_ok=True)
for n in ["Customers", "Products", "Sales"]:
    shutil.copy(ROOT / f"data/cleaned/{n}.csv", DST / f"{n}.csv")

monthly = pd.read_csv(ROOT / "outputs/monthly_revenue.csv")
fc = pd.read_csv(ROOT / "outputs/forecast_6_months.csv")
actual = pd.DataFrame({"MonthStart": pd.to_datetime(monthly.Month + "-01"), "ActualRevenue": monthly.Revenue})
fore = pd.DataFrame({"MonthStart": pd.to_datetime(fc.Month + "-01"), "ForecastRevenue": fc.ForecastRevenue,
                     "Lower80": fc.Lower80, "Upper80": fc.Upper80})
trend = pd.concat([actual, fore], ignore_index=True)
trend["MonthStart"] = trend["MonthStart"].dt.strftime("%Y-%m-%d")
trend["Type"] = trend["ActualRevenue"].notna().map({True: "Actual", False: "Forecast (indicative)"})
trend.to_csv(DST / "Forecast_Trend.csv", index=False)
print(trend)
