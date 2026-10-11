#!/usr/bin/env python3
"""
Generate a synthetic sales export to feed the weekly report.

All data is made up (fictional products, reps and orders). The data covers
12 weeks ending on the most recent Sunday, with a few realistic quirks so the
report has something to say: a busy weekend pattern, steady growth, one product
having a big week, one product having a slow week, a bump in refunds, and a
handful of invalid rows to show that the report's validation works.

Usage:
    python generate_sample_sales.py                    # -> data/sales.csv
    python generate_sample_sales.py --weeks 16 --seed 7
    python generate_sample_sales.py --end-date 2026-09-27
"""

import argparse
import csv
import random
from datetime import date, timedelta
from pathlib import Path

FIELDS = ["order_id", "order_date", "product", "category", "region",
          "sales_rep", "quantity", "unit_price", "discount_pct", "status"]

# name, category, price, popularity weight
PRODUCTS = [
    ("Wireless Mouse", "Electronics", 29.99, 10),
    ("Mechanical Keyboard", "Electronics", 89.99, 6),
    ("USB-C Hub", "Electronics", 49.99, 8),
    ("Noise-Cancelling Headphones", "Electronics", 199.99, 4),
    ("Standing Desk", "Home Office", 399.00, 2),
    ("Ergonomic Chair", "Home Office", 249.00, 3),
    ("Desk Lamp", "Home Office", 39.99, 7),
    ("Monitor Arm", "Home Office", 79.99, 5),
    ("Laptop Sleeve", "Accessories", 34.99, 9),
    ("Webcam", "Accessories", 69.99, 6),
    ("Phone Stand", "Accessories", 19.99, 10),
    ("Cable Organizer", "Accessories", 12.99, 8),
]
REGIONS = ["Toronto", "Ottawa", "Montreal", "Calgary", "Vancouver"]
REGION_WEIGHTS = [10, 4, 5, 3, 4]
REPS = ["Alex Chen", "Priya Nair", "Marcus Bell", "Sofia Reyes", "Jordan Lee"]
# Mon .. Sun: weekends are the busiest
DAY_MULTIPLIER = [0.85, 0.9, 0.95, 1.0, 1.15, 1.45, 1.2]


def last_sunday(on_or_before: date) -> date:
    return on_or_before - timedelta(days=(on_or_before.weekday() + 1) % 7)


def generate(weeks: int = 12, end: date = None, seed: int = 42,
             avg_orders_per_day: int = 28) -> list:
    rng = random.Random(seed)
    end = end or last_sunday(date.today())
    start = end - timedelta(days=weeks * 7 - 1)  # always a Monday

    rows, order_num = [], 1
    for offset in range(weeks * 7):
        day = start + timedelta(days=offset)
        weeks_ago = (end - day).days // 7          # 0 = the final week
        growth = 1 + 0.012 * (weeks - 1 - weeks_ago)
        base = avg_orders_per_day * growth * DAY_MULTIPLIER[day.weekday()]
        n_orders = max(0, int(rng.gauss(base, base * 0.12)))

        # the story: headphones have a big final week, desk lamps a slow one
        weights = []
        for name, _, _, weight in PRODUCTS:
            if weeks_ago == 0 and name == "Noise-Cancelling Headphones":
                weight *= 2.0
            if weeks_ago == 0 and name == "Desk Lamp":
                weight *= 0.5
            weights.append(weight)
        refund_rate = 0.065 if weeks_ago == 0 else 0.03

        for _ in range(n_orders):
            order_id = f"SO-{order_num:06d}"
            order_num += 1
            region = rng.choices(REGIONS, weights=REGION_WEIGHTS)[0]
            rep = rng.choice(REPS)
            status = "Refunded" if rng.random() < refund_rate else "Completed"
            n_lines = rng.choices([1, 2, 3], weights=[75, 20, 5])[0]

            chosen = []
            while len(chosen) < n_lines:
                product = rng.choices(PRODUCTS, weights=weights)[0]
                if product not in chosen:
                    chosen.append(product)

            for name, category, price, _ in chosen:
                rows.append({
                    "order_id": order_id,
                    "order_date": day.isoformat(),
                    "product": name,
                    "category": category,
                    "region": region,
                    "sales_rep": rep,
                    "quantity": rng.choices([1, 2, 3, 4], weights=[70, 20, 7, 3])[0],
                    "unit_price": f"{price:.2f}",
                    "discount_pct": rng.choice([0] * 17 + [5, 10, 15]),
                    "status": status,
                })

    # a handful of invalid rows, to show the report skips and reports them
    for row in rng.sample(rows, k=8):
        problem = rng.choice(["qty", "date", "price", "discount", "product"])
        if problem == "qty":
            row["quantity"] = ""
        elif problem == "date":
            row["order_date"] = "not a date"
        elif problem == "price":
            row["unit_price"] = "-5.00"
        elif problem == "discount":
            row["discount_pct"] = 150
        else:
            row["product"] = ""
    return rows


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate a sample sales CSV")
    parser.add_argument("--weeks", type=int, default=12, help="weeks of data (default 12)")
    parser.add_argument("--seed", type=int, default=42, help="random seed (default 42)")
    parser.add_argument("--end-date", type=date.fromisoformat, default=None,
                        help="last day of data, a Sunday (default: most recent Sunday)")
    parser.add_argument("--output", type=Path, default=Path("data/sales.csv"))
    args = parser.parse_args()

    rows = generate(args.weeks, args.end_date, args.seed)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows(rows)
    print(f"Wrote {len(rows):,} sales rows to {args.output}")


if __name__ == "__main__":
    main()
