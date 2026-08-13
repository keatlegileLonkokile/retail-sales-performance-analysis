"""
generate_sales_data.py

Creates a synthetic one-year sales dataset for "Northwind Retail", a fictional
8-store retail chain, used in the Retail Sales Performance Analysis portfolio
project.

The data is generated with deliberate, realistic business patterns baked in
so the downstream analysis has real findings to surface:

  - Miami steadily loses revenue across the year (simulated new local
    competitor opening in the fall)
  - Dallas underperforms specifically in Sports & Outdoors (weak regional
    demand for that category)
  - Electronics spikes company-wide in Nov/Dec (holiday shopping season)
  - Home & Kitchen runs on a thin margin relative to other categories
    (a pricing/cost problem worth flagging)

All data is synthetic and generated for demonstration/portfolio purposes only.
"""

import numpy as np
import pandas as pd
from datetime import date

np.random.seed(7)

STORES = {
    "New York":   "Northeast",
    "Boston":     "Northeast",
    "Atlanta":    "Southeast",
    "Miami":      "Southeast",
    "Chicago":    "Midwest",
    "Los Angeles":"West",
    "Seattle":    "West",
    "Dallas":     "Southwest",
}

# category -> list of (product, unit_price, target_margin)
CATALOG = {
    "Electronics": [
        ("Wireless Earbuds", 59.99, 0.35),
        ("Bluetooth Speaker", 44.99, 0.38),
        ("Smartwatch", 119.99, 0.32),
        ("Tablet Stand", 24.99, 0.40),
        ("Portable Charger", 34.99, 0.36),
    ],
    "Apparel": [
        ("Men's T-Shirt", 19.99, 0.55),
        ("Women's Leggings", 34.99, 0.58),
        ("Denim Jacket", 69.99, 0.52),
        ("Running Shoes", 79.99, 0.50),
        ("Wool Sweater", 54.99, 0.56),
    ],
    "Home & Kitchen": [
        ("Coffee Maker", 49.99, 0.28),
        ("Non-Stick Pan Set", 64.99, 0.30),
        ("Throw Blanket", 29.99, 0.32),
        ("LED Desk Lamp", 27.99, 0.29),
        ("Storage Bins (Set of 3)", 22.99, 0.27),
    ],
    "Beauty & Personal Care": [
        ("Facial Cleanser", 14.99, 0.62),
        ("Moisturizer", 18.99, 0.60),
        ("Hair Dryer", 39.99, 0.58),
        ("Electric Toothbrush", 44.99, 0.55),
        ("Perfume Set", 42.99, 0.63),
    ],
    "Sports & Outdoors": [
        ("Yoga Mat", 24.99, 0.48),
        ("Camping Tent", 129.99, 0.42),
        ("Water Bottle", 16.99, 0.50),
        ("Resistance Bands", 19.99, 0.52),
        ("Hiking Backpack", 74.99, 0.45),
    ],
}

MONTHS = pd.date_range("2025-08-01", periods=12, freq="MS")  # Aug 2025 - Jul 2026

rows = []

for store, region in STORES.items():
    # baseline store "size" multiplier so stores aren't all identical
    store_size = np.random.uniform(0.8, 1.3)

    for month_idx, month in enumerate(MONTHS):
        # Miami: steady decline across the year (new local competitor)
        if store == "Miami":
            miami_decline = 1.15 - (month_idx * 0.06)  # ~1.15 -> ~0.5
            store_month_mult = store_size * max(miami_decline, 0.35)
        else:
            store_month_mult = store_size * np.random.uniform(0.92, 1.08)

        for category, products in CATALOG.items():
            for product, price, margin in products:
                cost = round(price * (1 - margin) * np.random.uniform(0.97, 1.03), 2)

                base_units = np.random.poisson(lam=14)

                # Holiday spike for Electronics in Nov/Dec
                if category == "Electronics" and month.month in (11, 12):
                    base_units = int(base_units * np.random.uniform(2.2, 2.8))

                # Dallas underperforms specifically in Sports & Outdoors
                if store == "Dallas" and category == "Sports & Outdoors":
                    base_units = int(base_units * np.random.uniform(0.2, 0.4))

                units = max(int(round(base_units * store_month_mult)), 0)
                if units == 0:
                    continue

                revenue = round(units * price, 2)
                cost_total = round(units * cost, 2)
                profit = round(revenue - cost_total, 2)

                rows.append({
                    "Region": region,
                    "Store": store,
                    "Category": category,
                    "Product": product,
                    "Month": month.date().isoformat(),
                    "Units Sold": units,
                    "Unit Price": price,
                    "Unit Cost": cost,
                    "Revenue": revenue,
                    "Cost": cost_total,
                    "Profit": profit,
                })

df = pd.DataFrame(rows)
df.sort_values(["Month", "Region", "Store", "Category", "Product"], inplace=True)
df.to_csv("data/sales_data.csv", index=False)

print(f"Generated {len(df)} rows -> data/sales_data.csv")
print(f"Total revenue: ${df['Revenue'].sum():,.2f}")
print(f"Total profit:  ${df['Profit'].sum():,.2f}")
print(f"Date range: {df['Month'].min()} to {df['Month'].max()}")
