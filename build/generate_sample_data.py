"""
Generates a realistic SAMPLE sales transactions workbook so the dashboard
pipeline (extract_data.py -> build_dashboard.py) has something real to run
against out of the box.

This script is NOT part of the normal build pipeline — it's a one-off (or
occasional re-run) data generator. Once you have real sales data, replace
`data/Sales Data.xlsx` with your own file (same sheet/table/column layout,
see documentation/Documentation.md) and stop using this script.

Run from inside build/:
    python3 generate_sample_data.py
"""

import random
from datetime import date, timedelta

import numpy as np
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo

random.seed(42)
np.random.seed(42)

# ---------------------------------------------------------------------------
# Reference data: category -> products, plus a base unit price per product
# ---------------------------------------------------------------------------
CATALOG = {
    "Electronics": {
        "Wireless Earbuds": 59.99,
        "4K Monitor": 279.99,
        "Bluetooth Speaker": 44.99,
        "Smartwatch": 189.99,
        "Laptop Stand": 34.99,
        "Smartphone Case": 19.99,
    },
    "Apparel": {
        "Men's T-Shirt": 18.50,
        "Women's Jeans": 54.00,
        "Running Shoes": 89.99,
        "Winter Jacket": 129.00,
        "Baseball Cap": 16.99,
    },
    "Home & Garden": {
        "Ceramic Mug Set": 24.99,
        "LED Desk Lamp": 32.50,
        "Throw Blanket": 39.99,
        "Garden Hose": 27.99,
        "Planter Pot Set": 21.99,
    },
    "Sports & Outdoors": {
        "Yoga Mat": 29.99,
        "Camping Tent (2-person)": 119.00,
        "Insulated Water Bottle": 22.50,
        "Resistance Bands Set": 19.99,
        "Hiking Backpack": 74.99,
    },
    "Beauty": {
        "Face Moisturizer": 26.00,
        "Shampoo Bar": 12.99,
        "Lip Balm Set": 9.99,
        "Vitamin C Serum": 32.00,
    },
    "Grocery": {
        "Organic Coffee Beans (1lb)": 14.99,
        "Granola Bars (Box of 12)": 8.49,
        "Cold-Pressed Olive Oil": 17.99,
        "Herbal Tea Sampler": 11.49,
    },
}

# Rough cost-of-goods ratio per category (drives Profit), with a bit of noise
COGS_RATIO = {
    "Electronics": 0.62,
    "Apparel": 0.48,
    "Home & Garden": 0.55,
    "Sports & Outdoors": 0.52,
    "Beauty": 0.40,
    "Grocery": 0.65,
}

# Category-level growth trend across the date range (index at start vs end)
# >1 means the category is growing year over year, <1 means it's declining
CATEGORY_TREND = {
    "Electronics": 1.35,
    "Apparel": 1.05,
    "Home & Garden": 1.15,
    "Sports & Outdoors": 1.20,
    "Beauty": 1.45,
    "Grocery": 0.95,
}

SALE_LOCATIONS = ["Online", "Retail", "Wholesale"]
SALE_LOCATION_WEIGHTS = [0.55, 0.30, 0.15]

SALESPEOPLE = [
    "Priya Shah", "Daniel Osei", "Maria Gonzalez", "Tom Whitfield",
    "Aisha Khan", "Liam O'Connor", "Yuki Tanaka", "Sofia Rossi",
    "Ben Carter", "Grace Nwosu",
]

CUSTOMER_SEGMENTS = ["Consumer", "Small Business", "Enterprise"]
SEGMENT_WEIGHTS = [0.65, 0.25, 0.10]

START = date(2026, 1, 1)
END = date(2026, 8, 31)
TOTAL_DAYS = (END - START).days


def seasonality(d: date) -> float:
    """Multiplier: holiday bump (Nov/Dec) + back-to-school bump (Aug/Sep)."""
    m = d.month
    factor = 1.0
    if m in (11, 12):
        factor *= 1.55
    if m in (8, 9):
        factor *= 1.20
    if m in (1, 2):
        factor *= 0.85
    # mild day-of-week effect: weekends higher for retail/online
    if d.weekday() >= 5:
        factor *= 1.10
    return factor


def trend_multiplier(category: str, d: date) -> float:
    progress = (d - START).days / TOTAL_DAYS  # 0 -> 1 across the whole range
    start_mult, end_mult = 1.0, CATEGORY_TREND[category]
    return start_mult + (end_mult - start_mult) * progress


rows = []
order_id = 100000

d = START
while d <= END:
    day_seasonality = seasonality(d)
    # number of orders on this day scales with seasonality + mild randomness
    n_orders = max(0, int(np.random.poisson(lam=7 * day_seasonality)))
    for _ in range(n_orders):
        category = random.choices(
            list(CATALOG.keys()),
            weights=[trend_multiplier(c, d) for c in CATALOG.keys()],
        )[0]
        product = random.choice(list(CATALOG[category].keys()))
        base_price = CATALOG[category][product]

        # small price noise (promotions etc.)
        unit_price = round(base_price * np.random.uniform(0.90, 1.05), 2)
        units = int(np.random.choice([1, 1, 1, 2, 2, 3, 4, 6], p=[
            0.32, 0.18, 0.15, 0.14, 0.09, 0.06, 0.04, 0.02
        ]))

        sale_location = random.choices(SALE_LOCATIONS, weights=SALE_LOCATION_WEIGHTS)[0]
        if sale_location == "Wholesale":
            units *= random.choice([4, 6, 8, 10])
            unit_price = round(unit_price * 0.80, 2)  # wholesale discount

        revenue = round(unit_price * units, 2)
        cogs_ratio = np.clip(
            np.random.normal(COGS_RATIO[category], 0.04), 0.30, 0.85
        )
        cost = round(revenue * cogs_ratio, 2)
        profit = round(revenue - cost, 2)

        salesperson = random.choice(SALESPEOPLE)
        segment = random.choices(CUSTOMER_SEGMENTS, weights=SEGMENT_WEIGHTS)[0]

        order_id += 1
        rows.append([
            f"SO-{order_id}",
            d,  # PurchaseDate
            d,  # SaleDate (same as PurchaseDate for this sample dataset)
            category,
            product,
            sale_location,
            segment,
            salesperson,
            units,
            unit_price,
            revenue,
            cost,
            profit,
        ])
    d += timedelta(days=1)

print(f"Generated {len(rows):,} sample transactions "
      f"({START.isoformat()} to {END.isoformat()})")

# ---------------------------------------------------------------------------
# Write to an .xlsx workbook as a proper Excel Table, matching the
# convention used elsewhere in this project (named tables extract_data.py
# can reference directly instead of guessing at ranges).
# ---------------------------------------------------------------------------
HEADERS = [
    "OrderID", "PurchaseDate", "SaleDate", "Category", "Product", "SaleLocation",
    "CustomerSegment", "Salesperson", "Units", "UnitPrice", "Revenue",
    "Cost", "Profit",
]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Sales"
ws.append(HEADERS)
for row in rows:
    ws.append(row)

# Format the date columns as dates, currency columns with 2dp
for col_name in ("PurchaseDate", "SaleDate"):
    c = HEADERS.index(col_name) + 1
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=c).number_format = "yyyy-mm-dd"
for col_name in ("UnitPrice", "Revenue", "Cost", "Profit"):
    c = HEADERS.index(col_name) + 1
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=c).number_format = '"£"#,##0.00'

last_row = ws.max_row
last_col_letter = ws.cell(row=1, column=len(HEADERS)).column_letter
table_ref = f"A1:{last_col_letter}{last_row}"

table = Table(displayName="tblSalesTransactions", ref=table_ref)
table.tableStyleInfo = TableStyleInfo(
    name="TableStyleMedium2", showRowStripes=True
)
ws.add_table(table)

# Reasonable column widths
widths = [12, 14, 14, 16, 26, 12, 16, 16, 8, 11, 12, 11, 11]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w

out_path = "../data/Sales Data.xlsx"
wb.save(out_path)
print(f"Wrote {out_path}")
