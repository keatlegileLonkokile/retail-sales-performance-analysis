"""
analyze.py

Quick exploratory analysis of data/sales_data.csv used to sanity-check the
dataset and confirm the findings before they're written up in the Excel
report and README. Run this first if you want to see the numbers behind the
headline findings printed to the console.

Usage:
    python3 scripts/analyze.py
"""

import pandas as pd

df = pd.read_csv("data/sales_data.csv", parse_dates=["Month"])

print("=" * 60)
print("NORTHWIND RETAIL - SALES DATA QUICK ANALYSIS")
print("=" * 60)

print(f"\nRows: {len(df):,}  |  Date range: {df['Month'].min().date()} to {df['Month'].max().date()}")
print(f"Total Revenue: ${df['Revenue'].sum():,.2f}")
print(f"Total Profit:  ${df['Profit'].sum():,.2f}")
print(f"Overall Margin: {df['Profit'].sum() / df['Revenue'].sum():.1%}")

print("\n--- Revenue by Region ---")
print(df.groupby("Region")["Revenue"].sum().sort_values(ascending=False).round(0).to_string())

print("\n--- Revenue by Store ---")
print(df.groupby("Store")["Revenue"].sum().sort_values(ascending=False).round(0).to_string())

print("\n--- Miami: H1 (Aug-Jan) vs H2 (Feb-Jul) Revenue ---")
df["half"] = df["Month"].apply(lambda m: "H1" if m < pd.Timestamp("2026-02-01") else "H2")
miami = df[df.Store == "Miami"].groupby("half")["Revenue"].sum()
pct = (miami["H2"] - miami["H1"]) / miami["H1"]
print(miami.to_string())
print(f"Change: {pct:.1%}")

print("\n--- Dallas Sports & Outdoors vs. other stores' average ---")
dallas_so = df[(df.Store == "Dallas") & (df.Category == "Sports & Outdoors")]["Revenue"].sum()
other_so = (
    df[(df.Store != "Dallas") & (df.Category == "Sports & Outdoors")]
    .groupby("Store")["Revenue"].sum().mean()
)
print(f"Dallas: ${dallas_so:,.0f}  |  Other stores' average: ${other_so:,.0f}  "
      f"|  Gap: {(dallas_so - other_so) / other_so:.1%}")

print("\n--- Margin by Category (lowest first) ---")
cat = df.groupby("Category").agg(Revenue=("Revenue", "sum"), Profit=("Profit", "sum"))
cat["Margin"] = cat["Profit"] / cat["Revenue"]
print(cat.sort_values("Margin").round(3).to_string())

print("\n--- Electronics Monthly Revenue (holiday spike check) ---")
print(
    df[df.Category == "Electronics"]
    .groupby(df["Month"].dt.strftime("%Y-%m"))["Revenue"].sum().round(0).to_string()
)

print("\n--- Top 5 Products by Revenue ---")
print(df.groupby("Product")["Revenue"].sum().sort_values(ascending=False).head(5).round(0).to_string())

print("\n--- Bottom 5 Products by Revenue ---")
print(df.groupby("Product")["Revenue"].sum().sort_values().head(5).round(0).to_string())
