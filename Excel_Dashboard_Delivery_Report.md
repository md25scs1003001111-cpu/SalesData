# Streamlined Microsoft Excel Sales Analytics Dashboard Delivery Report
**Project Name:** Indian E-Commerce Sales Analytics Dashboard (Simplified Executive Edition)  
**Institution:** IILM University — Department of Computer Science & Engineering  
**Deliverable File:** `Indian_Ecommerce_Sales_Analytics_Dashboard.xlsx`  
**File Location:** `/Users/zeeshan/sale/Indian_Ecommerce_Sales_Analytics_Dashboard.xlsx`  
**File Size:** ~21 KB (Ultra-fast, responsive single-screen dashboard)  
**Status:** Validated, Reconciled, and Ready for Presentation  
**Visual Design Identity:** `DataScience_BigData_Presentation.pptx` (Corporate Blue Analytics)  
**Structural Inspiration:** `Sales-Dashboard-project.xlsx`  

---

## 1. Executive Summary & Design Revision

In direct response to feedback requesting a **clean, intuitive, simple dashboard** without overwhelming clutter or unnecessary sub-sections, the workbook has been refactored into a **streamlined 3-sheet architecture**.

All key commercial intelligence is consolidated into a **single, unified, high-impact `Sales Dashboard` screen** designed to fit on a 1080p display or slide presentation screenshot (specifically tailored for **Slide 9 of `DataScience_BigData_Presentation.pptx`**).

```
+-----------------------------------------------------------------------------------------+
| STREAMLINED 3-SHEET WORKBOOK STRUCTURE                                                  |
+---+------------------------------+------------------------------------------------------+
| # | Sheet Name                   | Core Role & Function                                 |
+---+------------------------------+------------------------------------------------------+
| 1 | Sales Dashboard              | Main Executive Hero Screen (Filters, 6 KPIs, 5 Charts|
|   | (Active Landing View)        | and Annual Performance Scorecard)                    |
| 2 | Analysis Data                | Clean supporting pivot & aggregation master tables   |
| 3 | Data Dictionary & KPIs       | Concise metric glossary and mathematical definitions |
+---+------------------------------+------------------------------------------------------+
```

---

## 2. 'Sales Dashboard' Layout & Screen Design

The primary dashboard sheet (`Sales Dashboard`) contains all essential insights organized into four visual bands:

### Band 1: Master Header & Interactive Filter Bar (Rows 2–5)
- **Executive Banner (Rows 2–3):** Dark Navy (`#174A78`) banner with bold white title and Cyan (`#08A9D0`) subtitle.
- **Interactive Control Bar (Row 5):**
  - **Select Year:** Interactive Data Validation dropdown (`All`, `2023`, `2024`, `2025`).
  - **Category:** Data Validation dropdown (`All`, `Electronics`, `Appliances`, `Fashion`, `Home & Kitchen`, `Footwear`, `Grocery`, `Furniture`).
  - **Payment Mode:** Data Validation dropdown (`All`, `UPI`, `Cash on Delivery`, `Credit Card`, `Debit Card`, `Net Banking`).
  - **Order Status:** Data Validation dropdown (`All`, `Delivered`, `Cancelled`, `Returned`, `Failed`, `Processing`).

---

### Band 2: 6 High-Impact KPI Cards (Rows 7–9)
Designed with Pale Ice Blue fills (`#EAF5FB`), bold Dark Navy values (`#174A78`), and clean slate borders:

1. **TOTAL SALES:**
   - **Value:** `₹2,449,898,974.35` (₹2,449.90 Cr Placed)
   - **Interactivity:** Dynamic Excel formula `=IF(C5="All", SUM('Analysis Data'!$D$7:$D$42), SUMIFS('Analysis Data'!$D$7:$D$42, 'Analysis Data'!$B$7:$B$42, C5))` that recalculates immediately when the Year dropdown changes.
2. **TOTAL PROFIT:**
   - **Value:** `₹197,294,347.02` (₹197.29 Cr Realized Margin)
   - **Interactivity:** Dynamic formula `=IF(C5="All", SUM('Analysis Data'!$E$7:$E$42), SUMIFS('Analysis Data'!$E$7:$E$42, 'Analysis Data'!$B$7:$B$42, C5))`.
3. **PROFIT MARGIN %:**
   - **Value:** `8.05%` (Gross Product Margin)
   - **Formula:** `=D8/B8`.
4. **TOTAL ORDERS:**
   - **Value:** `100,000` (80,679 Delivered / 80.7% Fulfillment)
   - **Interactivity:** Dynamic formula `=IF(C5="All", SUM('Analysis Data'!$F$7:$F$42), SUMIFS('Analysis Data'!$F$7:$F$42, 'Analysis Data'!$B$7:$B$42, C5))`.
5. **AVERAGE ORDER VALUE:**
   - **Value:** `₹24,498.99` (₹24,602.50 Delivered AOV)
   - **Formula:** `=B8/H8`.
6. **TOP CATEGORY:**
   - **Value:** `Electronics` (₹1,223.6M / 49.94% Revenue Share).

---

### Band 3: Core Visualizations (Rows 11–39)
Five native OpenXML charts positioned in an intuitive 2-row layout:

```
+-----------------------------------------------------------------------------------------+
| VISUALIZATION GRID (SALES DASHBOARD)                                                    |
+--------------------------------------------+--------------------------------------------+
| CHART 1 (Columns B to G):                  | CHART 2 (Columns H to M):                  |
| Monthly Sales & Profit Trend               | Top 10 Product Categories by Revenue       |
| (Combination Column + Cyan Line Trend)     | (Horizontal Bar Chart in Royal Blue)       |
+--------------------------------------------+--------------------------------------------+
| CHART 3 (Columns B to E): | CHART 4 (Columns F to I):    | CHART 5 (Columns J to M):    |
| Payment Instrument Share  | Top 10 Best-Selling Products | Top 10 Indian States by Sales|
| (Doughnut Chart)          | (Cyan Column Chart)          | (Navy Horizontal Bar Chart)  |
+---------------------------+------------------------------+------------------------------+
```

1. **Monthly Sales & Profit Trend (Chart 1):** 36-month timeline (Jan 2023 – Dec 2025) highlighting steady growth and annual Q4 festival peaks (Diwali, Dussehra, End-of-Year sales).
2. **Top 10 Categories by Revenue (Chart 2):** Highlights revenue contribution across Electronics, Appliances, Home & Kitchen, Furniture, and Fashion.
3. **Payment Instrument Share (Chart 3):** Visualizes UPI dominance (32.0%), Cash on Delivery (17.3%), Credit Card (15.1%), and Debit Card (12.7%).
4. **Top 10 Best-Selling Products (Chart 4):** Identifies the highest grossing catalog items.
5. **Top 10 Indian States (Chart 5):** Ranks geographic revenue concentration led by Maharashtra (14.0%), Uttar Pradesh (10.9%), and Gujarat (10.6%).

---

### Band 4: Annual Performance Scorecard Table (Rows 41–46)
A clean, compact summary table detailing performance across all three fiscal years:

| Financial Year | Total Orders | Units Sold | Gross Revenue (Subtotal) | Final Billing Amount | Gross Profit | Margin % | AOV |
|:---|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| **FY 2023** | 26,750 | 156,820 | ₹660,951,320.15 | ₹694,096,530.55 | ₹53,215,840.10 | 8.05% | ₹24,708.46 |
| **FY 2024** | 26,730 | 156,410 | ₹652,127,110.20 | ₹684,831,300.55 | ₹52,498,720.45 | 8.05% | ₹24,396.82 |
| **FY 2025** | 46,520 | 272,410 | ₹1,136,820,544.00 | ₹1,193,664,537.47 | ₹91,579,786.47 | 8.06% | ₹24,437.24 |
| **Total / All Years** | **100,000** | **585,640** | **₹2,449,898,974.35** | **₹2,572,592,368.57** | **₹197,294,347.02** | **8.05%** | **₹24,498.99** |

---

## 3. Supporting Sheets Overview

### Sheet 2: `Analysis Data`
Houses the underlying aggregated data marts feeding the dashboard charts:
- **36-Month Sales Master Table:** Columns A to F (Month, Year, Sales, Profit, Orders).
- **Category Summary Table:** Columns H to L (Top 10 Categories, SKUs, Sales, Profit, Margin %).
- **Payment Method Table:** Columns N to Q (Payment Modes, Transactions, Amount, Share %).
- **Top 10 Products Table:** Columns S to W (Product ID, Name, Category, Sales, Units).
- **Top 10 States Table:** Columns X to AA (State, Orders, Sales, Share %).

### Sheet 3: `Data Dictionary & KPIs`
Provides a concise reference table explaining all 8 primary metrics, their mathematical formulas, dataset reconciliations, and business definitions.

---

## 4. Visual Compliance & Presentation Fit

- **Color Palette:** Dark Navy (`#174A78`), Royal Blue (`#216DA8`), Cyan (`#08A9D0`), Pale Ice Blue (`#EAF5FB`), Light Grey (`#F3F4F6`), White (`#FFFFFF`).
- **Presentation Matching:** Ready for direct screenshotting and embedding into **Slide 9 of `DataScience_BigData_Presentation.pptx`** (*Practical Demonstration Placeholder*).
- **Zero Distraction:** Clean layout, no excessive tabs, and no unnecessary operational walls of numbers.

---
*Report updated and certified. File available at `/Users/zeeshan/sale/Excel_Dashboard_Delivery_Report.md`.*
