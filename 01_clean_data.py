"""
01_clean_data.py
Purpose : Load raw CSVs, validate + clean them, save cleaned copies and a master (merged) table.
Run     : python python/01_clean_data.py     (from the project root)
Outputs : data/cleaned/{Customers,Products,Sales,Sales_Master}.csv
          documentation/data_quality_report.txt
Raw files in data/raw are NEVER modified.
"""
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RAW, CLEAN, DOCS = ROOT / "data/raw", ROOT / "data/cleaned", ROOT / "documentation"
CLEAN.mkdir(parents=True, exist_ok=True)
DOCS.mkdir(parents=True, exist_ok=True)
log = []


def note(msg):
    print(msg)
    log.append(msg)


customers = pd.read_csv(RAW / "Customers.csv")
products = pd.read_csv(RAW / "Products.csv")
sales = pd.read_csv(RAW / "Sales.csv")
note(f"Loaded rows -> customers={len(customers)}, products={len(products)}, sales={len(sales)}")

# ---- 1. Text standardisation (trim spaces, consistent case). Excel equivalent: TRIM(), PROPER()
for df, cols in [(customers, ["Name", "Gender", "City"]), (products, ["ProductName", "Category"])]:
    for c in cols:
        before = df[c].copy()
        df[c] = df[c].str.strip().str.replace(r"\s+", " ", regex=True)
        if c in ("Name", "City"):
            df[c] = df[c].str.title()
        note(f"Text clean {c}: {int((before != df[c]).sum())} values changed")
for df, cols in [(customers, ["CustomerID"]), (products, ["ProductID"]), (sales, ["SalesID", "CustomerID", "ProductID"])]:
    for c in cols:
        df[c] = df[c].str.strip().str.upper()

# ---- 2. Dates: raw format is M/D/YYYY -> ISO YYYY-MM-DD (MySQL DATE friendly). Excel: DATEVALUE()
customers["JoinDate"] = pd.to_datetime(customers["JoinDate"], format="%m/%d/%Y", errors="raise")
sales["SaleDate"] = pd.to_datetime(sales["SaleDate"], format="%m/%d/%Y", errors="raise")
note("Dates parsed with explicit %m/%d/%Y format, no unparseable values")

# ---- 3. Duplicates & missing values
for name, df, key in [("Customers", customers, "CustomerID"), ("Products", products, "ProductID"), ("Sales", sales, "SalesID")]:
    note(f"{name}: full-row duplicates={int(df.duplicated().sum())}, duplicate {key}={int(df[key].duplicated().sum())}, "
         f"missing cells={int(df.isna().sum().sum())}")
    df.drop_duplicates(inplace=True)  # no-op on this data, kept as a safeguard

# ---- 4. Validation of numerics and keys
assert customers["Age"].between(18, 100).all(), "Age out of range"
assert (products["UnitPrice"] > 0).all() and (products["CostPrice"] > 0).all()
assert (products["UnitPrice"] > products["CostPrice"]).all(), "Product sold below cost"
assert (sales["Quantity"] > 0).all()
assert sales["CustomerID"].isin(customers["CustomerID"]).all(), "Orphan CustomerID in Sales"
assert sales["ProductID"].isin(products["ProductID"]).all(), "Orphan ProductID in Sales"
note(f"Validation passed: ages {customers.Age.min()}-{customers.Age.max()}, quantity {sales.Quantity.min()}-{sales.Quantity.max()}, "
     f"no orphan keys, every product has price > cost")

# ---- 5. Anomaly detection (flag only, do NOT delete legitimate data)
chk = sales.merge(customers[["CustomerID", "JoinDate"]], on="CustomerID")
n_before_join = int((chk["SaleDate"] < chk["JoinDate"]).sum())
note(f"ANOMALY (flagged, kept): {n_before_join} of {len(sales)} sales are dated BEFORE the customer's JoinDate. "
     f"Likely synthetic/independent generation of the two tables. Flag column 'SaleBeforeJoin' added in master table.")
note(f"Sales date range: {sales.SaleDate.min().date()} to {sales.SaleDate.max().date()} (single year, 12 months)")
note(f"Customers who never bought: {len(set(customers.CustomerID) - set(sales.CustomerID))}; "
     f"products never sold: {len(set(products.ProductID) - set(sales.ProductID))}")
mism = customers[customers["Name"].str.split().str[0].isin(["Ahmed", "Aisha", "Michael", "Rajesh", "Omar", "Hassan"])]
note("NOTE: several Name/Gender pairs look inconsistent (e.g. C001 'Ahmed Khan' = Female). Values kept as supplied; "
     "gender-based insights should be treated as indicative only.")

# ---- 6. Master table (Sales + Customers + Products) with derived fields
m = sales.merge(customers, on="CustomerID", how="left", validate="m:1").merge(products, on="ProductID", how="left", validate="m:1")
m["Revenue"] = m["Quantity"] * m["UnitPrice"]
m["Cost"] = m["Quantity"] * m["CostPrice"]
m["Profit"] = m["Revenue"] - m["Cost"]
m["Year"] = m["SaleDate"].dt.year
m["Month"] = m["SaleDate"].dt.month
m["YearMonth"] = m["SaleDate"].dt.strftime("%Y-%m")
m["SaleBeforeJoin"] = (m["SaleDate"] < m["JoinDate"]).map({True: "Yes", False: "No"})
m = m[["SalesID", "SaleDate", "Year", "Month", "YearMonth", "CustomerID", "Name", "Age", "Gender", "City", "JoinDate",
       "ProductID", "ProductName", "Category", "Quantity", "UnitPrice", "CostPrice", "Revenue", "Cost", "Profit", "SaleBeforeJoin"]]
assert len(m) == len(sales), "Merge changed row count"
note(f"Master table: {len(m)} rows. Revenue={m.Revenue.sum():,}, Profit={m.Profit.sum():,}, Units={m.Quantity.sum():,}")

# ---- 7. Save (ISO dates)
for name, df in [("Customers", customers), ("Products", products), ("Sales", sales), ("Sales_Master", m)]:
    out = df.copy()
    for c in out.select_dtypes("datetime").columns:
        out[c] = out[c].dt.strftime("%Y-%m-%d")
    out.to_csv(CLEAN / f"{name}.csv", index=False)
(DOCS / "data_quality_report.txt").write_text("\n".join(log), encoding="utf-8")
print("Saved cleaned files to", CLEAN)
