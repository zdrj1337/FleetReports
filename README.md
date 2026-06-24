# Fleet Payment Report Generator

A command-line tool that turns the weekly driver data exported by **Uber** and **Bolt**
into clean, per-driver payment reports in Excel — plus a fleet-wide summary with a
per-platform breakdown and a chart. Built as a portfolio project, modelled on a real
ride-hailing fleet I drove for, where the fleet owner receives raw exports from each
platform and has to produce a payment report for every driver.

The interesting part: the two platforms export data in **different shapes** (different
column names and structure). The tool normalises both into one common format, applies the
fleet's payment rules, and generates the reports.

## What it does

1. **Reads** the weekly exports from both platforms (`uber.csv`, `bolt.csv`).
2. **Normalises** the two different layouts into one common schema (a small ETL step).
3. **Calculates**, per driver: platform commission, the fleet's own 7% commission, net
   earnings, and the final bank transfer (net + tips + bonus − cash the driver already kept).
4. **Generates** an Excel report per driver, plus a `_fleet_summary.xlsx` that contains a
   fleet overview, a **By Platform** breakdown sheet, and a bar chart of the bank transfer
   per driver. Every derived value is written as a live Excel **formula**, so the math is
   transparent and the files recalculate.

## How it works

The platforms don't agree on column names — Bolt says `Bolt commission`, Uber says
`Uber service fee`; Bolt has `Bonus`, Uber has `Promotions`; and so on. `fleet_logic.py`
maps each export into a single internal table:

```
driver | platform | gross | platform_commission | cash | tips | bonus
```

Drivers can be on one platform or both — the tool combines a driver's rows across platforms
(and handles drivers who only drove for one of them). `generate_reports.py` then writes the
Excel files.

## Sample driver report

![Sample driver report](docs/sample_driver_report.png)

## Fleet summary

![Fleet summary](docs/fleet_summary.png)

## Running it

Requires Python 3.10+.

```bash
# 1. (recommended) virtual environment
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate

# 2. install dependencies
pip install -r requirements.txt

# 3a. generate the reports for one period
python generate_reports.py 2026-W10

# 3b. or generate every period found under data/
python generate_reports.py --all
```

The reports appear under `reports/<year>/<week>/` — one `.xlsx` per driver plus
`_fleet_summary.xlsx`. Sample data is included for many weeks (2025-W40 → 2026-W20), so it
runs out of the box; `2026-W10` is a small, hand-checked showcase week.

To create fresh synthetic input data:

```bash
python generate_sample_data.py            # regenerate all sample weeks
python generate_sample_data.py 2026-W05   # a single week
```

## Web app

The project also ships a small Flask web interface (`app.py`) on top of the same logic.
Pick a week and generate the reports; the page shows the fleet totals and a per-driver
table, where each report can be viewed as a PDF in the browser or downloaded as Excel (or
all of them together as a zip).

```bash
pip install -r requirements.txt
python app.py
# open http://127.0.0.1:5000
```

## Project structure

```
.
├── app.py                    # Flask web interface (pick a week, view PDF, download Excel)
├── generate_reports.py       # builds the Excel reports (formatting, formulas, chart)
├── pdf_report.py             # builds the PDF reports (reportlab) for in-browser viewing
├── fleet_logic.py            # reads & normalises the two platform exports (pandas)
├── generate_sample_data.py   # makes synthetic Bolt/Uber exports for testing
├── templates/index.html      # the web page
├── requirements.txt
├── data/                     # raw weekly exports, archived per period
│   ├── 2025/
│   │   └── W40/ ... W52/
│   │       ├── bolt.csv
│   │       └── uber.csv
│   └── 2026/
│       └── W01/ ... W20/
└── reports/                  # generated reports (reports/<year>/<week>/)
```

Inputs are organised by period and kept rather than overwritten, so there's a traceable
history of every week's data — the way a real fleet would need it for accounting.

## Notes

All drivers and figures are fictional. The payment rules (platform commission, the fleet's
7% commission, cash reconciliation) are modelled on a real weekly fleet report.

## Roadmap

- Optional email step: prepare each driver's email with a payment summary.
- Support for additional platforms.
- Live hosted demo.

## Author

**Florin-Traian Zadorojneac** — Automation & Applied Informatics student, Galați, Romania.
