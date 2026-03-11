# ⚙️ Automated Revenue Reporting Pipeline — Python + BigQuery + openpyxl

![Python](https://img.shields.io/badge/Python-3.10-blue?logo=python)
![BigQuery](https://img.shields.io/badge/Google%20BigQuery-SQL-blue?logo=googlebigquery)
![openpyxl](https://img.shields.io/badge/openpyxl-Excel%20Automation-green)
![Pandas](https://img.shields.io/badge/Pandas-2.0-lightblue?logo=pandas)

An end-to-end automated revenue reporting pipeline that replaces manual Excel workflows — pulling live data from **Google BigQuery**, processing with **pandas**, and generating a fully formatted multi-sheet Excel report with **openpyxl**. Saves 12+ analyst hours per week.

---

## 📌 Business Problem

The analytics team was spending 12+ hours every week manually:
- Pulling revenue data from BigQuery
- Pasting into Excel
- Formatting tables, headers, and KPI summaries
- Saving and distributing to stakeholders

**Goal:** Automate the entire pipeline — from BigQuery query to formatted Excel delivery — so analysts run one script and the report generates itself.

---

## 🗂️ Project Structure

```
automated-revenue-pipeline/
│
├── pipeline.py                  # Full automation script
├── requirements.txt             # Dependencies
└── README.md
```

---

## 🛠️ Tech Stack

| Tool | Usage |
|------|-------|
| Google BigQuery | Live revenue data source |
| Python (pandas) | Data aggregation and transformation |
| openpyxl | Excel report generation and formatting |
| Google Colab | Execution environment |

---

## 🔄 Pipeline Flow

```
BigQuery (ecommerce_rfm.orders)
        ↓
3 SQL queries (weekly, monthly, summary)
        ↓
pandas DataFrames
        ↓
openpyxl — 4 formatted sheets
        ↓
revenue_report_YYYYMMDD_HHMM.xlsx
```

---

## 📊 Report Structure

The generated Excel file contains 4 sheets:

### Sheet 1 — Cover Page
- Report title and generation timestamp (auto-updated every run)
- 4 KPI boxes: Total Revenue, Total Orders, Brands, Avg Order Value
- Clean corporate design with dark blue header theme

### Sheet 2 — Brand Summary
- Total revenue, orders, avg order value per brand
- Colour-coded alternating rows
- Formatted currency columns

### Sheet 3 — Monthly Revenue
- Month-by-month revenue breakdown per brand
- Sortable and filterable
- Full 2-year history

### Sheet 4 — Weekly Revenue
- Week-by-week granular revenue data
- Useful for spotting short-term anomalies

---

## 🔍 BigQuery Queries

### Weekly Revenue
```sql
SELECT
    brand,
    DATE_TRUNC(order_date, WEEK)     AS week_start,
    COUNT(DISTINCT order_id)         AS total_orders,
    ROUND(SUM(grand_total), 2)       AS weekly_revenue,
    ROUND(AVG(grand_total), 2)       AS avg_order_value
FROM `ecommerce_rfm.orders`
WHERE status = 'complete'
GROUP BY brand, week_start
ORDER BY week_start, brand
```

### Monthly Revenue
```sql
SELECT
    brand,
    DATE_TRUNC(order_date, MONTH)    AS month_start,
    COUNT(DISTINCT order_id)         AS total_orders,
    ROUND(SUM(grand_total), 2)       AS monthly_revenue,
    ROUND(AVG(grand_total), 2)       AS avg_order_value
FROM `ecommerce_rfm.orders`
WHERE status = 'complete'
GROUP BY brand, month_start
ORDER BY month_start, brand
```

---

## 📊 Key Results

| Metric | Value |
|--------|-------|
| Total Revenue (3 brands) | $91,867,502 |
| Total Orders | 225,313 |
| Avg Order Value | $407.73 |
| Analyst hours saved per week | 12+ hours |
| Report generation time | < 60 seconds |
| Manual steps required | 0 |

---

## 🚀 How to Run

### Setup
```bash
pip install openpyxl google-cloud-bigquery pandas db-dtypes
```

### In Google Colab
```python
# 1. Authenticate
from google.colab import auth
auth.authenticate_user()

# 2. Set your project
PROJECT_ID = "your-gcp-project-id"

# 3. Run pipeline
filename = build_revenue_report(df_summary, df_weekly, df_monthly)

# 4. Download
from google.colab import files
files.download(filename)
```

### Output
A file named `revenue_report_YYYYMMDD_HHMM.xlsx` is generated automatically with the current timestamp.

---

## 💡 Key Learnings & Interview Talking Points

- **openpyxl vs pandas to_excel** — pandas `to_excel()` dumps raw data; openpyxl gives full control over fonts, colors, borders, merged cells, and KPI boxes
- **Timestamped filenames** — `f"revenue_report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"` ensures no file is ever overwritten
- **DATE_TRUNC in BigQuery** — cleaner than EXTRACT for weekly/monthly aggregations; works directly with pandas datetime
- **The real value** is not the report itself — it's removing 12 hours of repetitive human work per week, which compounds to 600+ hours per year saved
- **Automation mindset:** always ask "what is the human doing repeatedly that code could do instead?"

---

## 👤 Author

Built as part of a portfolio replicating real-world e-commerce analytics automation workflows.
