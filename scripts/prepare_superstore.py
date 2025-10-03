"""
Sales Analysis Starter Script
- Cleans "Sample - Superstore.csv"
- Creates ready CSVs for Tableau dashboards
"""

import pandas as pd
from pathlib import Path

base_path = Path(".")  # current folder
src = base_path / "Sample - Superstore.csv"

def load_csv(path):
    for enc in ["utf-8", "latin-1", "cp1252"]:
        try:
            return pd.read_csv(path, encoding=enc)
        except Exception:
            continue
    return pd.read_csv(path)

def main():
    df = load_csv(src).rename(columns=lambda c: c.strip())

    # Parse dates
    for col in ["Order Date", "Ship Date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # Numeric columns
    for col in ["Sales", "Profit", "Discount"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    if "Quantity" in df.columns:
        df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")

    # Time features
    if "Order Date" in df.columns:
        df["Order Year"] = df["Order Date"].dt.year
        df["Order Month"] = df["Order Date"].dt.month
        df["Order Month Name"] = df["Order Date"].dt.strftime("%b")
        df["Order Year-Month"] = df["Order Date"].dt.to_period("M").astype(str)

    out_dir = base_path / "superstore_outputs"
    out_dir.mkdir(exist_ok=True)

    # Export cleaned dataset
    df.to_csv(out_dir / "superstore_clean.csv", index=False)

    # Monthly
    if "Order Year-Month" in df.columns:
        monthly = df.groupby("Order Year-Month").agg(
            Sales=("Sales", "sum"),
            Profit=("Profit", "sum"),
            Orders=("Order ID", "nunique")
        ).reset_index()
        monthly.to_csv(out_dir / "monthly_sales.csv", index=False)

    # Category / Sub-Category
    cat_cols = [c for c in ["Category", "Sub-Category"] if c in df.columns]
    if cat_cols:
        cat = df.groupby(cat_cols).agg(Sales=("Sales", "sum"), Profit=("Profit", "sum")).reset_index()
        cat.to_csv(out_dir / "category_sales.csv", index=False)

    # Geo
    geo_cols = [c for c in ["Region", "State", "City"] if c in df.columns]
    if geo_cols:
        geo = df.groupby(geo_cols).agg(Sales=("Sales", "sum"), Profit=("Profit", "sum")).reset_index()
        geo.to_csv(out_dir / "geo_sales.csv", index=False)

    # Top products
    if "Product Name" in df.columns:
        top_products = (
            df.groupby("Product Name")
              .agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"), Quantity=("Quantity", "sum"))
              .reset_index()
              .sort_values("Sales", ascending=False)
              .head(20)
        )
        top_products.to_csv(out_dir / "top_products.csv", index=False)

    # Customers
    if "Customer ID" in df.columns:
        group_cols = ["Customer ID"]
        if "Customer Name" in df.columns:
            group_cols.append("Customer Name")
        customers = (
            df.groupby(group_cols)
              .agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"), Orders=("Order ID", "nunique"))
              .reset_index()
              .sort_values("Sales", ascending=False)
        )
        customers.to_csv(out_dir / "customers_summary.csv", index=False)

if __name__ == "__main__":
    main()
