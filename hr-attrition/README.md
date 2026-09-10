# HR Attrition Analytics Dashboard

End-to-end pipeline: messy HR data -> SQL cleaning -> Python transform -> Excel dashboard.

## What it does
- Cleans 260+ employee records (fixes gender/department casing, invalid ages, negative salaries, invalid satisfaction scores, duplicates)
- Calculates: Attrition Rate, Avg Tenure, Avg Satisfaction, Avg Salary, Attrition Rate by Department
- Excel dashboard with live formulas (COUNTIFS/AVERAGEIF) + bar chart + pie chart

## Run it
```bash
pip install pandas openpyxl
python python/01_generate_data.py
python python/02_pipeline_and_dashboard.py
```
Output: `exports/HR_Attrition_Dashboard.xlsx`

## Tech stack
SQL (SQLite) for cleaning · Python (Pandas) for transformation · Excel (openpyxl) for reporting

## Use YOUR real data
Replace `raw_data/hr_employees.csv` with your actual HR export (same
column structure) and re-run — the pipeline and dashboard update automatically.
