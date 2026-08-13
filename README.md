# Retail Sales Performance Analysis

A business analyst project: take a year of raw retail sales data, find out what's actually going on across regions, stores, categories, and products, and turn it into a formula-driven Excel report a manager could open, trust, and act on.

## Scenario

You're a business analyst at **Northwind Retail**, a fictional chain of 8 stores across 5 US regions selling five product categories. Leadership has one ask: "Tell us what's working, what isn't, and what we should do about it" — using a year of transaction-level sales data (Aug 2025 - Jul 2026).

The goal isn't just to summarize the numbers — it's to find the specific, defensible findings (a declining store, an underperforming category-region combination, a thin-margin category, a seasonal pattern) and turn each one into a concrete recommendation, backed by a workbook where every number is a live formula rather than a pasted-in result.

## Approach

1. **Data**: Generate a realistic transaction dataset (2,400 rows: store x product x month) with `pandas`/`numpy`.
2. **Explore**: Run a quick pandas pass (`scripts/analyze.py`) to sanity-check the data and confirm findings before committing to them.
3. **Build**: Construct a multi-sheet Excel workbook (`scripts/build_workbook.py`, `openpyxl`) where every summary number is a live formula (`SUMIFS`, `INDEX`/`MATCH`, `LARGE`/`SMALL`, `AVERAGE`) referencing the raw data — nothing is hardcoded from Python except the transactions themselves.
4. **Recalculate**: Run the workbook through LibreOffice headless to compute cached formula values (9,864 formulas, 0 errors) so it opens correctly in Excel/Sheets without a manual "recalculate all."
5. **Report**: An Insights & Recommendations sheet with a live "Key Metrics" table plus five specific, numbered findings and recommendations.

## Project structure

```
retail-sales-performance-analysis/
├── data/
│   └── sales_data.csv                    # synthetic transaction-level data
├── scripts/
│   ├── generate_sales_data.py            # builds the synthetic dataset
│   ├── analyze.py                        # quick pandas exploration / sanity check
│   └── build_workbook.py                 # builds the formula-driven Excel report
├── output/
│   └── Sales_Performance_Report.xlsx     # the deliverable
├── requirements.txt
└── README.md
```

## Setup & usage

```bash
git clone <your-repo-url>
cd retail-sales-performance-analysis
pip install -r requirements.txt

# (optional) regenerate the synthetic dataset
python3 scripts/generate_sales_data.py

# (optional) quick console analysis
python3 scripts/analyze.py

# build the Excel report
python3 scripts/build_workbook.py
```

Open `output/Sales_Performance_Report.xlsx` in Excel, Google Sheets, or LibreOffice — every sheet recalculates from the `Raw Data` tab if you change any number.

## About the sample data

`data/sales_data.csv` is **synthetically generated** (`scripts/generate_sales_data.py`) — no real company, stores, or products. It's built with a few deliberate patterns so the analysis has real signal to find, not noise:

- **Miami** loses revenue steadily across the year (simulating a new local competitor)
- **Dallas** specifically underperforms in the Sports & Outdoors category (weak regional demand for that category only — not a store-wide problem)
- **Electronics** spikes company-wide in November/December (holiday shopping)
- **Home & Kitchen** runs on a thinner margin than every other category (a pricing/cost issue)

## Workbook contents

| Sheet | What it shows |
|---|---|
| Raw Data | All 2,400 transactions, with Revenue/Cost/Profit/Margin as formulas |
| Regional Summary | Revenue/cost/profit/margin by region, with a chart |
| Category Summary | Same, by category, with a margin chart |
| Monthly Trend | Total vs. Electronics revenue by month, with a trend line chart |
| Store Performance | H1 vs. H2 revenue per store (`SUMIFS` with date-range criteria) with a clustered bar chart — this is what surfaces Miami's decline |
| Product Summary | Revenue/profit/margin per product, plus live Top-5 / Bottom-5 tables (`LARGE`/`SMALL` + `INDEX`/`MATCH`) |
| Insights & Recommendations | A live "Key Metrics" table (top/bottom region, most-declining store, thinnest-margin category, holiday lift — all formula-driven) plus five written findings and recommendations |

## Key findings

**Total FY revenue:** ~$1.73M | **Total FY profit:** ~$743K | **Overall margin:** ~43%

1. **Miami is in a steady, accelerating decline.** Revenue fell ~43% from H1 to H2 — the steepest drop of any store, while its regional peers held steady. Recommend an on-the-ground review (competitor activity, pricing, foot traffic) and a targeted local promotion before considering a footprint change.
2. **Dallas is dramatically underperforming in Sports & Outdoors specifically** — about 75% below the average of other stores in that category alone, while normal everywhere else. Recommend an assortment/placement audit for that category rather than a store-wide response.
3. **Home & Kitchen carries the thinnest margin** of any category (~29% vs. 53-59% for Apparel and Beauty & Personal Care). Recommend a supplier/cost review and testing a modest price increase on the highest-volume SKUs.
4. **Electronics spikes sharply in November/December** — recommend building this into inventory and staffing plans ahead of Q4, plus a cross-sell push for accessories/warranties during that window.
5. **Apparel and Beauty & Personal Care are the strongest all-around performers** (solid revenue, highest margins) — recommend protecting their shelf space/marketing spend and using their pricing approach as a template for the weaker categories.

## Possible extensions

- Add a Power Query / Power BI version for a live, refreshable dashboard instead of a static workbook.
- Customer-level data (loyalty program) to segment performance by customer type, not just store/category.
- A pricing elasticity test on Home & Kitchen SKUs to validate finding #3 before rolling out a price change.
- Automate the monthly refresh so this becomes a recurring report rather than a one-off.

## Disclaimer

All data in this repository is synthetically generated for educational/portfolio purposes. No real company, stores, employees, or customers are represented.
