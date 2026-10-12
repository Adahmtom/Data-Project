# Automated Monthly KPI Report

A script that builds a monthly KPI workbook in Excel for FormWise, a made-up online legal-document subscription
service. The data is simulated by generate_data.py.

The goal was to replace a manual monthly report with something anyone can re-run, and to let people look at any
month themselves without asking for a new report.

## How it works
`build_report.py` pulls the numbers with SQL and writes a workbook where the KPIs are Excel formulas, not pasted
values. Everything is driven by one cell, the report month, which has a dropdown.

| Sheet | Contents |
|---|---|
| KPI Summary | This month vs last month vs same month last year, green/red changes, and a channel by device conversion table built with SUMIFS. Includes a check that the table totals match the KPIs. |
| Monthly | 18-month trend with conversion and revenue charts |
| Data Quality | Flags months where purchase tracking and billing disagree by more than a set tolerance (3% by default, editable) |
| Data | Raw extract the formulas read from |

It also writes a short text summary to summary_YYYY-MM.md.

## Run it
```
pip install -r requirements.txt
python build_report.py --month 2026-06
```
The database is built automatically on the first run. A sample report for June 2026 is in `output/`.

Tools: Python, SQL (SQLite), pandas, openpyxl, Excel
