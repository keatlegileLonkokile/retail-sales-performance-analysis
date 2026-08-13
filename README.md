# Retail Sales Performance Analysis

A business analyst project. Take a year of raw retail sales data, work out what's really going on across regions, stores, categories and products and turn it into an Excel report a manager could open, trust and act on.

## Scenario

You're a business analyst at **Acewind Retail**, a made up chain of 8 stores across 5 SA regions selling five product categories. Leadership has one ask. Tell us what's working, what isn't and what we should do about it, using a year of sales data from August 2025 to July 2026.

The goal isn't just to summarise the numbers. It's to find specific findings you can back up, like a declining store, a category that's underperforming in one region, a category with thin margins and a seasonal pattern and turn each one into a real recommendation. Every number in the workbook comes from a formula, not a number typed in by hand.

## Approach

1. **Data**: Build a realistic transaction dataset (2,400 rows, one row per store, product and month) using `pandas` and `numpy`.
2. **Explore**: Run a quick pandas script (`scripts/analyse.py`) to check the data makes sense before trusting any findings.
3. **Build**: Build a multi sheet Excel workbook (`scripts/build_workbook.py`, `openpyxl`) where every summary number is a live formula (`SUMIFS`, `INDEX`/`MATCH`, `LARGE`/`SMALL`, `AVERAGE`) pulling from the raw data. Nothing is hardcoded from Python except the raw transactions themselves.
4. **Recalculate**: Run the workbook through LibreOffice to compute the formula values (9,864 formulas, zero errors) so it opens correctly straight away in Excel or Sheets.
5. **Report**: An Insights & Recommendations sheet with a live key metrics table plus five specific, numbered findings and recommendations.

## Project structure

```
retail-sales-performance-analysis/
├── data/
│   └── sales_data.csv                    # synthetic transaction level data
├── scripts/
│   ├── generate_sales_data.py            # builds the synthetic dataset
│   ├── analyze.py                        # quick pandas exploration and sanity check
│   └── build_workbook.py                 # builds the formula driven Excel report
├── output/
│   └── Sales_Performance_Report.xlsx     # the deliverable
├── requirements.txt
└── README.md
```

## Setup and usage

```bash
git clone <your-repo-url>
cd retail-sales-performance-analysis
pip install -r requirements.txt

# optional: regenerate the synthetic dataset
python3 scripts/generate_sales_data.py

# optional: quick console analysis
python3 scripts/analyze.py

# build the Excel report
python3 scripts/build_workbook.py
```

Open `output/Sales_Performance_Report.xlsx` in Excel, Google Sheets or LibreOffice. Every sheet recalculates from the Raw Data tab if you change any number.

## About the sample data

`data/sales_data.csv` is made up (`scripts/generate_sales_data.py` builds it). No real company, stores or products. It's built with a few patterns on purpose, so there's something real to find:

- **Johannesburg** loses revenue steadily across the year, simulating a new local competitor
- **Capetown** underperforms specifically in Sports & Outdoors. Weak demand for that one category, not a store wide problem
- **Electronics** spikes company wide in November and December, holiday shopping
- **Home & Kitchen** runs on a thinner margin than every other category, a pricing or cost issue

## Workbook contents

| Sheet | What it shows |
|---|---|
| Raw Data | All 2,400 transactions, with Revenue, Cost, Profit and Margin as formulas |
| Regional Summary | Revenue, cost, profit and margin by region, with a chart |
| Category Summary | The same, by category, with a margin chart |
| Monthly Trend | Total revenue vs Electronics revenue by month, with a trend line chart |
| Store Performance | First half vs second half revenue per store, using SUMIFS with date range criteria, with a clustered bar chart. This is what surfaces Miami's decline |
| Product Summary | Revenue, profit and margin per product, plus live top 5 and bottom 5 tables using LARGE, SMALL, INDEX and MATCH |
| Insights & Recommendations | A live key metrics table (top and bottom region, most declining store, thinnest margin category, holiday lift, all formula driven) plus five written findings and recommendations |

## Key findings

**Total revenue for the year:** around R1.73 million. **Total profit:** around R743,000. **Overall margin:** around 43%.

1. **Johannesburg is in a steady decline that's getting worse.** Revenue fell about 43% from the first half of the year to the second half, the steepest drop of any store, while every other store in the region held steady. Recommend an on the ground review of competitor activity, pricing and foot traffic, plus a local promotion, before considering anything bigger like closing the store.
2. **Capetown is badly underperforming in Sports & Outdoors specifically.** About 75% below the average of the other stores in that one category, while everything else at that store is normal. Recommend an assortment and shelf placement audit for that category rather than a store wide response.
3. **Home & Kitchen has the thinnest margin of any category**, around 29% against 53 to 59% for Apparel and Beauty & Personal Care. Recommend a supplier and cost review, and testing a modest price increase on the highest volume products.
4. **Electronics spikes sharply in November and December.** Recommend building this into inventory and staffing plans ahead of the fourth quarter instead of reacting to it, plus a cross sell push for accessories and warranties during that window.
5. **Apparel and Beauty & Personal Care are the strongest performers overall**, solid revenue and the highest margins. Recommend protecting their shelf space and marketing budget, and using their pricing approach as a template when fixing the weaker categories.

## Possible extensions

- Add a Power Query or Power BI version so the dashboard refreshes on its own instead of being a static file
- Bring in customer level data (like a loyalty program) to look at performance by customer type, not just store or category
- Run a pricing test on Home & Kitchen products to check finding 3 before rolling out a price change
- Automate the monthly refresh so this becomes a recurring report instead of a one off

## Disclaimer

All data in this repository is made up, for learning and portfolio purposes only. No real company, stores, employees or customers are represented.
