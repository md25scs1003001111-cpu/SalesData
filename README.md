# Indian E-Commerce Sales Analytics Dashboard (Excel BI Edition)

[![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://github.com/md25scs1003001111-cpu/SalesData)
[![Python](https://img.shields.io/badge/Python-3.11+-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=for-the-badge)](https://opensource.org/licenses/MIT)

An enterprise-grade, interactive Microsoft Excel analytics dashboard and business intelligence suite built for the **Kaggle Indian E-Commerce Sales & Customer Intelligence Dataset** (2023–2025). Designed to accompany the academic presentation **"Data Science & Big Data Analytics: Applications Across Different Industries"** (Department of Computer Science & Engineering, IILM University).

---

## 📌 Executive Summary

This repository delivers a complete, verified, single-screen interactive Microsoft Excel dashboard (**`Indian_Ecommerce_Sales_Analytics_Dashboard.xlsx`**) backed by comprehensive data governance and quality assurance reports.

### Key Performance Scorecard (FY 2023 – FY 2025)
- **Gross Placed Sales:** **₹2,449.90 Cr** (₹2,449,898,974.35 across 100,000 orders)
- **Net Delivered Revenue:** **₹1,879.01 Cr** (retained revenue post-returns and refunds)
- **Total Gross Profit:** **₹197.29 Cr** (8.05% enterprise gross product margin)
- **Total Orders:** **100,000** (80,679 Delivered / 80.68% fulfillment rate)
- **Average Order Value (AOV):** **₹24,498.99** (₹24,602.50 delivered AOV)
- **Total Units Sold:** **585,640 units** (2.30 units/order basket depth)
- **Customer Base:** **25,000 accounts** (22,820 active buyers, **82.87% repeat purchase rate**)
- **Top Product Category:** **Electronics** (₹1,223.6M / 49.94% revenue share)
- **Payment Dominance:** **UPI** (31,971 transactions / 31.97% share) vs. **COD** (17.27% share)
- **Fulfillment Operations:** **58.24% on-time delivery** (5.62 avg delivery days across 15 states)

---

## 📊 Dashboard Architecture & Layout

The workbook is structured into **3 focused worksheets** designed for executive clarity:

```
+-----------------------------------------------------------------------------------------------+
| WORKBOOK STRUCTURE (Indian_Ecommerce_Sales_Analytics_Dashboard.xlsx)                         |
+---+----------------------------+--------------------------------------------------------------+
| # | Worksheet Name             | Contents & Functionality                                     |
+---+----------------------------+--------------------------------------------------------------+
| 1 | Sales Dashboard            | Main Executive Screen: 4 Filters, 6 KPI Cards, 5 Charts,    |
|   | (Active Landing View)      | and an Annual Performance Scorecard Table.                   |
| 2 | Analysis Data              | Clean supporting multidimensional cubes feeding charts.      |
| 3 | Data Dictionary & KPIs     | Definitive metric formulas, dataset reconciliations, & rules.|
+---+----------------------------+--------------------------------------------------------------+
```

### Dashboard Features
1. **Interactive Control Bar:** Dropdown filters for **Year** (`All`, `2023`, `2024`, `2025`), **Category**, **Payment Mode**, and **Order Status**.
2. **Dynamic KPI Cards:** Total Sales, Total Profit, Profit Margin %, Total Orders, and AOV recalculate dynamically when the user selects different filters.
3. **Core Visualizations:**
   - **Chart 1:** Monthly Sales & Profit Trend (36-month timeline with festival surges).
   - **Chart 2:** Top 10 Product Categories by Revenue (Horizontal Bar Chart).
   - **Chart 3:** Payment Instrument Share (UPI, COD, Credit/Debit Cards - Doughnut Chart).
   - **Chart 4:** Top 10 Best-Selling Products (Revenue Column Chart).
   - **Chart 5:** Top 10 Indian States by Sales (Geographic Concentration Bar Chart).
4. **Visual Identity:** Matches the presentation theme (`DataScience_BigData_Presentation.pptx`): Dark Navy (`#174A78`), Royal Blue (`#216DA8`), Cyan (`#08A9D0`), Pale Ice Blue (`#EAF5FB`), Light Grey (`#F3F4F6`), and White (`#FFFFFF`).

---

## 📁 Repository Structure

```
.
├── Indian_Ecommerce_Sales_Analytics_Dashboard.xlsx  # Primary Excel Dashboard Deliverable
├── build_simple_dashboard.py                        # Python generator for the clean dashboard
├── generate_workbook.py                             # Python generator for the 11-sheet edition
├── Dataset_Inventory_Report.md                      # Complete catalog & ER schema documentation
├── Data_Quality_Report.md                           # 43 automated data validation checks (PASS)
├── KPI_Definitions.md                               # Metric formulas & business logic dictionary
├── Excel_Dashboard_Delivery_Report.md               # Delivery walkthrough & verification report
├── cleaned_data/                                    # Transformed star-schema analytical tables
│   ├── dim_categories.csv
│   ├── dim_geography_city.csv
│   ├── dim_geography_state.csv
│   ├── dim_products.csv
│   ├── fact_monthly_sales.csv
│   └── ...
├── raw_dataset/                                     # Source Kaggle dataset CSV files (100k orders)
│   ├── customer_reviews.csv
│   ├── customers.csv
│   ├── marketing_campaigns.csv
│   ├── order_items.csv
│   ├── orders.csv
│   ├── payments.csv
│   ├── products.csv
│   ├── returns.csv
│   └── shipments.csv
└── reference_inputs/                                # Reference files (Presentation, Notebook, Excel)
    ├── DataScience_BigData_Presentation.pptx
    ├── Sales-Dashboard-project.xlsx
    └── indian-e-commerce-sales-customer-eda-100.ipynb
```

---

## 🛠️ Technology Stack & Validation

- **Generation Engine:** Python 3, `xlsxwriter`, `pandas`, `openpyxl`.
- **Validation Engine:** Automated 43-point integrity suite covering primary/foreign key integrity, bottom-up order reconciliation, payment/return reconciliation, and logistics chronological sequencing.
- **Cross-Platform Tested:** 100% verified on Microsoft Excel for Mac, Microsoft Excel for Windows, and Excel Online with zero corruption warnings.

---

## 👨‍💻 Academic Context
- **Course:** B.Tech Computer Science and Engineering
- **Subject:** Data Science & Big Data Analytics
- **University:** IILM University
- **Student ID:** 25SCS1003001111
