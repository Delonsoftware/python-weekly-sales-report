# Automated Weekly Sales Report Generator

A Python tool that turns a raw sales CSV into a **ready-to-send weekly report in one
command**: a formatted Excel workbook with charts, plus a plain-English summary you can
paste straight into an email.

A task that takes 30+ minutes by hand (filtering, pivoting, comparing to last week,
writing it up) runs in a fraction of a second, and can be scheduled to run every Monday.

Includes a generator for realistic (fully synthetic) sales data so you can try it in
under a minute.

## What the report contains

**Excel workbook** (`weekly_sales_report_<week>.xlsx`):

| Sheet | Contents |
|---|---|
| Summary | KPIs vs last week (revenue, orders, units, average order value, refunds), written highlights, daily and trend charts |
| Daily | Revenue, orders and units for each day, with last week's revenue alongside |
| Top Products | Revenue, share, and change vs last week for every product |
| Categories / Regions / Sales Reps | The same breakdown for each |
| Weekly Trend | Last 8 weeks of revenue and orders |
| Skipped Rows | Any invalid rows in the input, with the reason (nothing is silently ignored) |

**Plain-English summary** (`weekly_summary_<week>.txt`), for example:

> Revenue for the week was $30,394 from 233 orders, up 8.5% from last week ($28,013).
> Noise-Cancelling Headphones was the top seller with $7,760 (26% of revenue).
> Biggest drop: Monitor Arm, down 56% to $1,080 (from $2,440).
> Refunds: 17 orders ($1,772), 6.8% of orders, up from 2.6% last week. Worth a look.

## Setup

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Usage

```bash
python generate_sample_sales.py                    # creates data/sales.csv (12 weeks)
python sales_report.py                             # report on the latest complete week
python sales_report.py --week-start 2026-09-14     # a specific week (any date inside it)
python sales_report.py --input my_sales.csv --output-dir reports
```

Use your own data: the CSV needs `order_id, order_date, product, quantity, unit_price`.
Optional columns are `category, region, sales_rep, discount_pct, status`
(status: Completed, Refunded or Cancelled; blank counts as Completed).

## How the numbers are calculated

- Weeks run Monday to Sunday. By default the report covers the latest *complete* week in the data.
- Line revenue = quantity x unit price x (1 - discount), rounded to the cent.
- Refunded orders are excluded from revenue and reported separately. Cancelled orders are ignored.
- "Biggest gain / drop" only considers products with meaningful sales last week
  (default: at least $250, change with `--min-prior`), so a product going from $1 to $20 doesn't dominate.
- Rows with a missing product/ID, invalid date, quantity, price or discount are skipped
  and listed on the **Skipped Rows** sheet.

## Run it automatically every Monday morning (Mac)

```bash
crontab -e
```

Add this line (adjust the folder path), which runs at 8:00 every Monday:

```
0 8 * * 1 cd /path/to/project && ./venv/bin/python sales_report.py >> cron.log 2>&1
```

On macOS you may need to grant Full Disk Access to `cron` in System Settings if your
project lives in Documents or Desktop.

## Tests

```bash
python test_sales_report.py
```

Covers week selection, revenue math (discounts, refunds, cancellations), percent-change
edge cases, invalid-row handling, the "biggest mover" noise filter, and a full end-to-end
run that checks the headline revenue against an independent calculation.

## Tech

Python, pandas, openpyxl (native Excel charts).

## Note

Built as a personal sample project. All data is synthetic: fictional products, reps and orders.
