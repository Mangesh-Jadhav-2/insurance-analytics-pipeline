# Insurance Analytics Pipeline

A set of Python and SQL scripts I built for an end-to-end BFSI/insurance
analytics pipeline: cleaning raw policy and branch-expense data, loading it
for SQL/BI consumption, and generating relationship-manager (RM) and branch
performance reports feeding into a Power BI dashboard.

All scripts here are generalized versions of production scripts, company
names, internal file paths, and any sensitive data have been removed or
replaced with generic placeholders and CLI arguments.

## Pipeline overview

```
Raw Excel exports (policy, salary, branch expense)
        |
        v
  policy_data_sql_prep.py   ---->  cleaned CSV, SQL-load-ready
  branch_cost_to_csv.py     ---->  cleaned CSV, SQL-load-ready
        |
        v
   SQL Server (rollups, joins, pivots)
        |
        v
  rm_efficiency_report.py         --->  formatted Excel report (color-coded)
  weighted_average_calc.py        --->  segment-weighted branch ranking
  branch_excel_splitter.py        --->  per-branch Excel files
        |
        v
      Power BI dashboard
```

## Python scripts (/python)

| Script | What it does |
|---|---|
| policy_data_sql_prep.py | Cleans raw policy-level Excel exports for SQL loading: normalizes float-mangled ID columns, builds a composite primary key, parses dates, computes retention. |
| branch_cost_to_csv.py | Maps a branch operating-expense Excel report's headers to SQL column names and coerces types into a load-ready CSV. |
| rm_efficiency_report.py | Consolidates raw transaction data per RM/month and computes an Efficiency metric (Retention / Salary), with color-coded Excel output flagging declining and underperforming RMs. |
| weighted_average_calc.py | Calculates a segment-weighted premium average per branch (Health weighted higher than Motor, etc.) and ranks branches accordingly. |
| branch_excel_splitter.py | Splits one workbook into per-branch Excel files, preserving header formatting and column widths. |

Each script takes --input / --output arguments (or prompts interactively), no hardcoded paths.

## SQL (/sql)

| Script | What it does |
|---|---|
| rm_efficiency_monthly_pivot.sql | CTE + conditional aggregation pivoting Active POS, NOP, Premium, Payin/Payout and Retention into one column per metric-per-month, per RM. |
| rm_efficiency_with_payscale.sql | Full joins RM transaction performance against monthly salary data as the input to efficiency calculations. |
| pincode_premium_nop.sql | Validates pincode formatting and rolls up policy count/premium by segment and pincode, with business-rule variants for Life segments. |

## Stack

Python (pandas, openpyxl), SQL Server, Power BI.

---
Table and column names above are illustrative, connect these scripts to your own schema.
