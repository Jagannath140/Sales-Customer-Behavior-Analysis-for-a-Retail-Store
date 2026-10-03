"""
02_analysis.py
Purpose : EDA, revenue/product/category/customer/profit analysis, CLV, monthly trend and forecast.
Run     : python python/02_analysis.py      (after 01_clean_data.py, from the project root)
Outputs : outputs/*.csv (result tables), outputs/charts/*.png, outputs/summary_kpis.json
Currency: the project document uses QAR; the data has no currency column, so we label amounts QAR.
"""
import json
import warnings
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from statsmodels.tsa.arima.model import ARIMA

warnings.filterwarnings("ignore")
ROOT = Path(__file__).resolve().parents[1]
OUT, CH = ROOT / "outputs", ROOT / "outputs/charts"
CH.mkdir(parents=True, exist_ok=True)
sns.set_theme(style="whitegrid", context="notebook")
BLUE, ORANGE = "#1f4e79", "#e07b39"

# ---------------------------------------------------------------- 1. LOAD
customers = pd.read_csv(ROOT / "data/cleaned/Customers.csv", parse_dates=["JoinDate"])
products = pd.read_csv(ROOT / "data/cleaned/Products.csv")
sales = pd.read_csv(ROOT / "data/cleaned/Sales.csv", parse_dates=["SaleDate"])

# ---------------------------------------------------------------- 2. MERGE
df = sales.merge(customers, on="CustomerID").merge(products, on="ProductID")
df["Revenue"] = df["Quantity"] * df["UnitPrice"]
df["Profit"] = df["Quantity"] * (df["UnitPrice"] - df["CostPrice"])
df["Month"] = df["SaleDate"].dt.to_period("M")

# ---------------------------------------------------------------- 3. EDA
print("Shape:", df.shape)
print("Missing values:\n", df.isna().sum()[df.isna().sum() > 0] if df.isna().any().any() else "none")
print(df[["Quantity", "UnitPrice", "Revenue", "Profit"]].describe().round(1))

kpi = {
    "total_revenue": int(df.Revenue.sum()),
    "total_profit": int(df.Profit.sum()),
    "profit_margin_pct": round(df.Profit.sum() / df.Revenue.sum() * 100, 2),
    "total_units": int(df.Quantity.sum()),
    "transactions": int(len(df)),
    "customers_who_bought": int(df.CustomerID.nunique()),
    "products_sold": int(df.ProductID.nunique()),
    "avg_transaction_value": round(df.Revenue.mean(), 2),
}
(OUT / "summary_kpis.json").write_text(json.dumps(kpi, indent=2))
print(kpi)

# ---------------------------------------------------------------- 4. REVENUE: monthly trend
monthly = df.groupby("Month").agg(Revenue=("Revenue", "sum"), Profit=("Profit", "sum"),
                                  Units=("Quantity", "sum"), Transactions=("SalesID", "count")).reset_index()
monthly["MarginPct"] = (monthly.Profit / monthly.Revenue * 100).round(2)
monthly["MoM_Growth_Pct"] = (monthly.Revenue.pct_change() * 100).round(1)
monthly["Month"] = monthly["Month"].astype(str)
monthly.to_csv(OUT / "monthly_revenue.csv", index=False)

fig, ax = plt.subplots(figsize=(10, 5))
ax.plot(monthly.Month, monthly.Revenue, marker="o", color=BLUE, label="Revenue")
ax.plot(monthly.Month, monthly.Profit, marker="o", color=ORANGE, label="Profit")
ax.set_title("Monthly Revenue and Profit, 2023 (Peak: Aug-Sep)")
ax.set_ylabel("QAR"); ax.tick_params(axis="x", rotation=45); ax.legend()
ax.yaxis.set_major_formatter(lambda v, _: f"{v/1000:.0f}k")
plt.tight_layout(); plt.savefig(CH / "01_monthly_revenue_profit.png", dpi=150); plt.close()

# ---------------------------------------------------------------- 5. CATEGORY
cat = df.groupby("Category").agg(Revenue=("Revenue", "sum"), Profit=("Profit", "sum"), Units=("Quantity", "sum")).reset_index()
cat["MarginPct"] = (cat.Profit / cat.Revenue * 100).round(2)
cat["RevenueShare_Pct"] = (cat.Revenue / cat.Revenue.sum() * 100).round(1)
cat.to_csv(OUT / "category_performance.csv", index=False)

fig, axes = plt.subplots(1, 2, figsize=(11, 4.5))
sns.barplot(data=cat, x="Category", y="Revenue", color=BLUE, ax=axes[0])
axes[0].set_title("Revenue by Category"); axes[0].set_ylabel("QAR")
sns.barplot(data=cat, x="Category", y="MarginPct", color=ORANGE, ax=axes[1])
axes[1].set_title("Profit Margin % by Category"); axes[1].set_ylabel("%")
for a in axes:
    for c in a.containers: a.bar_label(c, fmt="%.0f")
plt.tight_layout(); plt.savefig(CH / "02_category_revenue_margin.png", dpi=150); plt.close()

# ---------------------------------------------------------------- 6. PRODUCT
prod = df.groupby(["ProductID", "ProductName", "Category"]).agg(
    Units=("Quantity", "sum"), Revenue=("Revenue", "sum"), Profit=("Profit", "sum")).reset_index()
prod["MarginPct"] = (prod.Profit / prod.Revenue * 100).round(2)
prod = prod.sort_values("Revenue", ascending=False)
prod.to_csv(OUT / "product_performance.csv", index=False)

fig, axes = plt.subplots(1, 2, figsize=(13, 5))
top_rev = prod.head(10).iloc[::-1]
axes[0].barh(top_rev.ProductName, top_rev.Revenue, color=BLUE); axes[0].set_title("Top 10 Products by Revenue"); axes[0].set_xlabel("QAR")
top_units = prod.sort_values("Units", ascending=False).head(10).iloc[::-1]
axes[1].barh(top_units.ProductName, top_units.Units, color=ORANGE); axes[1].set_title("Top 10 Products by Units Sold"); axes[1].set_xlabel("Units")
plt.tight_layout(); plt.savefig(CH / "03_top_products.png", dpi=150); plt.close()

fig, ax = plt.subplots(figsize=(8, 5.5))
sns.scatterplot(data=prod, x="Revenue", y="MarginPct", hue="Category", size="Units", sizes=(40, 300), ax=ax)
for _, r in prod.iterrows():
    if r.Revenue > prod.Revenue.quantile(.8) or r.MarginPct < 20: ax.annotate(r.ProductName, (r.Revenue, r.MarginPct), fontsize=8, xytext=(4, 4), textcoords="offset points")
ax.set_title("Product Revenue vs Margin % (low-margin / high-revenue = review pricing)")
plt.tight_layout(); plt.savefig(CH / "04_product_revenue_vs_margin.png", dpi=150); plt.close()

# ---------------------------------------------------------------- 7. CUSTOMER + CLV
# Only 12 months of data and no churn information, so a forward-looking CLV (needs lifespan/retention)
# would rest on invented assumptions. We therefore report HISTORICAL CLV = profit (and revenue) each
# customer generated in the observed window. It is the honest, explainable measure this data supports.
cust = df.groupby(["CustomerID", "Name", "City", "Gender"]).agg(
    Orders=("SalesID", "count"), Units=("Quantity", "sum"), Revenue=("Revenue", "sum"), Profit=("Profit", "sum"),
    FirstPurchase=("SaleDate", "min"), LastPurchase=("SaleDate", "max")).reset_index()
cust["AvgOrderValue"] = (cust.Revenue / cust.Orders).round(2)
cust["CLV_Profit"] = cust.Profit
cust["CLV_Revenue"] = cust.Revenue
q80, q20 = cust.Revenue.quantile(.8), cust.Revenue.quantile(.2)
cust["Segment"] = np.where(cust.Revenue >= q80, "High value", np.where(cust.Revenue <= q20, "Low value", "Mid value"))
cust = cust.sort_values("CLV_Revenue", ascending=False)
cust.to_csv(OUT / "customer_clv.csv", index=False)
top10 = cust.head(10)
top10_share = round(top10.Revenue.sum() / cust.Revenue.sum() * 100, 1)
kpi.update({"top10_customers_revenue_share_pct": top10_share, "high_value_threshold_revenue": float(q80),
            "avg_clv_profit": round(cust.CLV_Profit.mean(), 2), "customers_over_5000_revenue": int((cust.Revenue > 5000).sum())})

fig, ax = plt.subplots(figsize=(9, 5))
t = top10.iloc[::-1]
ax.barh(t.Name, t.Revenue, color=BLUE); ax.set_title(f"Top 10 Customers by Revenue ({top10_share}% of total)"); ax.set_xlabel("QAR")
plt.tight_layout(); plt.savefig(CH / "05_top10_customers.png", dpi=150); plt.close()

city = df.groupby("City").agg(Revenue=("Revenue", "sum"), Profit=("Profit", "sum"), Customers=("CustomerID", "nunique")).reset_index()
city["RevenuePerCustomer"] = (city.Revenue / city.Customers).round(0)
city = city.sort_values("Revenue", ascending=False); city.to_csv(OUT / "city_performance.csv", index=False)
gender = df.groupby("Gender").agg(Revenue=("Revenue", "sum"), Customers=("CustomerID", "nunique")).reset_index()
gender.to_csv(OUT / "gender_performance.csv", index=False)

fig, axes = plt.subplots(1, 2, figsize=(12, 4.5))
sns.barplot(data=city, x="City", y="Revenue", color=BLUE, ax=axes[0]); axes[0].set_title("Revenue by City"); axes[0].set_ylabel("QAR")
axes[1].pie(gender.Revenue, labels=gender.Gender, autopct="%1.0f%%", colors=[BLUE, ORANGE], startangle=90)
axes[1].set_title("Revenue Share by Gender (indicative: see data note)")
plt.tight_layout(); plt.savefig(CH / "06_city_gender.png", dpi=150); plt.close()

# Purchase behaviour
basket = df.groupby("Quantity").size().reset_index(name="Transactions")
fig, ax = plt.subplots(figsize=(6, 4))
sns.barplot(data=basket, x="Quantity", y="Transactions", color=BLUE, ax=ax); ax.set_title("Units per Transaction")
plt.tight_layout(); plt.savefig(CH / "07_quantity_distribution.png", dpi=150); plt.close()

# ---------------------------------------------------------------- 8. FORECAST
# Honest approach: 12 monthly points is very little for ARIMA (needs ~30-50+). We (a) fit the ARIMA(2,1,2)
# from the project document, (b) fit smaller candidates, (c) back-test on the last 3 months against a simple
# average baseline, and (d) publish a forecast clearly labelled INDICATIVE with wide intervals.
s = monthly.set_index(pd.PeriodIndex(monthly.Month, freq="M").to_timestamp())["Revenue"].astype(float)
s.index.freq = "MS"
train, test = s.iloc[:-3], s.iloc[-3:]
cands = {"Baseline: mean of training": None, "ARIMA(1,0,0)": (1, 0, 0), "ARIMA(0,1,1)": (0, 1, 1),
         "ARIMA(1,1,1)": (1, 1, 1), "ARIMA(2,1,2) [project doc]": (2, 1, 2)}
rows = []
for name, order in cands.items():
    try:
        if order is None:
            pred = pd.Series(train.mean(), index=test.index); aic = np.nan
            full_aic = np.nan
        else:
            pred = ARIMA(train, order=order).fit().forecast(3)
            full_aic = ARIMA(s, order=order).fit().aic
        rows.append({"Model": name, "BacktestMAE_last3m": round(float(np.mean(np.abs(test.values - pred.values))), 0),
                     "AIC_full_series": None if order is None else round(full_aic, 1)})
    except Exception as e:
        rows.append({"Model": name, "BacktestMAE_last3m": None, "AIC_full_series": None})
bt = pd.DataFrame(rows); bt.to_csv(OUT / "forecast_model_comparison.csv", index=False)
print(bt)

best = bt.dropna(subset=["AIC_full_series"]).sort_values("BacktestMAE_last3m").iloc[0]["Model"]
order = cands[best]
fit = ARIMA(s, order=order).fit()
fc = fit.get_forecast(6)
ci = fc.conf_int(alpha=0.2)  # 80% interval
fdf = pd.DataFrame({"Month": fc.predicted_mean.index.strftime("%Y-%m"), "ForecastRevenue": fc.predicted_mean.round(0).values,
                    "Lower80": ci.iloc[:, 0].clip(lower=0).round(0).values, "Upper80": ci.iloc[:, 1].round(0).values,
                    "Model": best, "Status": "Indicative only (12 months of history)"})
fdf.to_csv(OUT / "forecast_6_months.csv", index=False)
baseline_mae = float(bt.loc[bt.Model.str.startswith("Baseline"), "BacktestMAE_last3m"].iloc[0])
best_mae = float(bt.loc[bt.Model == best, "BacktestMAE_last3m"].iloc[0])
kpi.update({"forecast_model": best, "forecast_backtest_mae": best_mae, "baseline_backtest_mae": baseline_mae,
            "forecast_beats_baseline": bool(best_mae < baseline_mae)})
print(fdf)

fig, ax = plt.subplots(figsize=(10, 5))
ax.plot(s.index, s.values, marker="o", color=BLUE, label="Actual 2023")
fi = pd.to_datetime(fdf.Month + "-01")
ax.plot(fi, fdf.ForecastRevenue, marker="o", ls="--", color=ORANGE, label=f"Forecast ({best})")
ax.fill_between(fi, fdf.Lower80, fdf.Upper80, color=ORANGE, alpha=.2, label="80% interval")
ax.set_title("Revenue Forecast, Next 6 Months (INDICATIVE: only 12 months of history)")
ax.set_ylabel("QAR"); ax.legend(); ax.yaxis.set_major_formatter(lambda v, _: f"{v/1000:.0f}k")
plt.tight_layout(); plt.savefig(CH / "08_forecast.png", dpi=150); plt.close()

(OUT / "summary_kpis.json").write_text(json.dumps(kpi, indent=2))
print(json.dumps(kpi, indent=2))
