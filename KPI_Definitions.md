# Enterprise E-Commerce KPI Definitions & Business Logic Dictionary
**Project:** Indian E-Commerce Sales & Customer Intelligence Analytics  
**Institution:** IILM University — Department of Computer Science & Engineering  
**Deliverable:** Definitive Metric Dictionary, Mathematical Formulas & Accounting Policies  
**Dashboard Workbook:** `Indian_Ecommerce_Sales_Analytics_Dashboard.xlsx`  
**Standard:** Enterprise BI & E-Commerce Financial Governance  

---

## 1. Executive Summary

This document establishes the official definitions, computational formulas, aggregation grains, and reconciliation standards for all metrics displayed in the Microsoft Excel dashboard. 

In e-commerce business intelligence, ambiguity surrounding terminology (e.g., confusing gross placed orders with net delivered sales, or item-level return rate with order-level return rate) leads to distorted financial assessments. This dictionary guarantees unambiguous semantic consistency across all 11 worksheets of the workbook.

---

## 2. Core Revenue & Financial Metrics

### 2.1. Gross Placed Revenue (Subtotal)
- **Mathematical Formula:**
  $$\text{Gross Revenue} = \sum_{i \in \text{Orders}} \text{orders.subtotal}_i = \sum_{j \in \text{Order Items}} \text{order\_items.item\_revenue}_j$$
- **Dataset Value:** **₹2,449,898,974.35** (₹2,449.90 Cr)
- **Granularity:** Order Header / Line Item Basket
- **Unit of Measurement:** Indian Rupee (INR ₹)
- **Business Meaning:** Total gross catalog value of all goods ordered by customers before promotional coupon deductions at checkout, shipping freight fees, and Goods and Services Tax (GST). Reconciles bottom-up to line-item revenues with zero variance.

### 2.2. Total Final Billing Amount
- **Mathematical Formula:**
  $$\text{Final Amount} = \sum_{i \in \text{Orders}} (\text{subtotal}_i - \text{discount\_amount}_i + \text{shipping\_fee}_i + \text{tax\_amount}_i)$$
- **Dataset Value:** **₹2,572,592,368.57** (₹2,572.59 Cr)
- **Granularity:** Order Header
- **Unit of Measurement:** Indian Rupee (INR ₹)
- **Business Meaning:** The complete invoice payable amount charged to customer payment gateways, including shipping charges (waived for orders over ₹999 or Premium subscribers) and category GST tax slabs.

### 2.3. Delivered Gross Revenue
- **Mathematical Formula:**
  $$\text{Delivered Revenue} = \sum_{i \in \text{Orders}, \text{status} = \text{'Delivered'}} \text{orders.subtotal}_i$$
- **Dataset Value:** **₹1,984,905,302.15** (₹1,984.91 Cr)
- **Share of Gross Revenue:** **81.02%**
- **Business Meaning:** Value of orders successfully fulfilled, delivered to customer doorsteps, and retained without full order return.

### 2.4. Net Revenue (Accounting Top-Line)
- **Mathematical Formula:**
  $$\text{Net Revenue} = \text{Delivered Subtotal Revenue} - \sum_{k \in \text{Returns}} \text{returns.refund\_amount}_k$$
- **Dataset Value:** **₹1,879,006,920.35** (₹1,879.01 Cr)
- **Calculation Details:** ₹1,984,905,302.15 (Delivered) - ₹105,898,381.80 (Returns Refund) = ₹1,879,006,920.35
- **Unit of Measurement:** Indian Rupee (INR ₹)
- **Business Meaning:** True top-line retained revenue after factoring out reverse logistics refunds. This represents the primary revenue baseline for income statements.

### 2.5. Gross Profit (COGS Margin)
- **Mathematical Formula:**
  $$\text{Gross Profit} = \sum_{j \in \text{Order Items}} (\text{order\_items.item\_revenue}_j - \text{order\_items.item\_cost}_j)$$
- **Dataset Value (All Orders):** **₹197,294,347.02** (₹197.29 Cr)
- **Dataset Value (Delivered Orders):** **₹158,628,938.94** (₹158.63 Cr)
- **Granularity:** Line Item
- **Unit of Measurement:** Indian Rupee (INR ₹)
- **Accounting Policy:** Represents merchant product contribution margin (Item Revenue minus Product Procurement Cost). Operating expenditures (marketing ad spend ₹187.1M, payment gateway fees ₹11.2M, reverse logistics freight) are tracked in separate functional modules and NOT deducted from line-item gross profit.

### 2.6. Gross Profit Margin %
- **Mathematical Formula:**
  $$\text{Gross Margin \%} = \frac{\text{Gross Profit}}{\text{Gross Revenue}} \times 100$$
- **Dataset Value:** **8.05%** (Overall Placed) | **7.99%** (Delivered Items)
- **Category Dispersion:** Highly stratified across categories:
  - *High-Margin Drivers:* Fashion (33.14%), Footwear (29.42%), Bags (28.59%), Beauty (26.50%), Personal Care (24.66%).
  - *Low-Margin Volume Drivers:* Electronics (0.69%), Appliances (3.89%), Grocery (5.61%).

---

## 3. Order & Fulfillment Metrics

### 3.1. Total Orders Placed
- **Mathematical Formula:** $\text{Total Orders} = \text{COUNT}(\text{orders.order\_id})$
- **Dataset Value:** **100,000 Orders**
- **Temporal Distribution:** FY23: 26,750 | FY24: 26,730 | FY25: 46,520.

### 3.2. Order Fulfillment Status Pipeline
- **Delivered Orders:** 80,679 (80.68%) — Successfully fulfilled.
- **Cancelled Orders:** 8,073 (8.07%) — Cancelled pre-fulfillment (₹192.97M subtotal).
- **Returned Orders:** 6,681 (6.68%) — Delivered then fully returned (₹161.77M subtotal).
- **Failed Orders:** 2,880 (2.79%) — Payment or system transaction failure (₹68.39M subtotal).
- **Processing Orders:** 1,687 (1.71%) — In warehouse pick/pack at reporting cutoff (₹41.87M subtotal).

### 3.3. Average Order Value (AOV)
- **Mathematical Formula:**
  $$\text{AOV} = \frac{\text{Total Subtotal Revenue}}{\text{Total Orders Placed}} = \frac{₹2,449,898,974.35}{100,000} = \mathbf{₹24,498.99}$$
- **Delivered AOV:** $\frac{₹1,984,905,302.15}{80,679} = \mathbf{₹24,602.50}$

### 3.4. Total Units Sold & Basket Depth
- **Total Units Sold:** **585,640 units**
- **Average Basket Depth (Units / Order):** $\frac{585,640}{100,000} = \mathbf{2.30\text{ units/order}}$
- **Average Unit Selling Price:** $\frac{₹2,449,898,974.35}{585,640} = \mathbf{₹4,183.28\text{ per unit}}$

---

## 4. Customer Behavior & Retention Metrics

### 4.1. Total Registered vs. Active Purchasing Customers
- **Total Registered Accounts:** **25,000 customers**
- **Purchasing Customers (Orders $\ge 1$):** **22,820 customers** (91.28% activation rate)
- **Inactive / Never Purchased:** **2,180 customers** (8.72%) (2,300 in customer master status table reflecting late-period churn rules).

### 4.2. Repeat Customer Rate
- **Numerator:** Unique customers with $\ge 2$ lifetime orders (18,912 customers).
- **Denominator (Purchasing Cohort):** 22,820 purchasing customers.
- **Repeat Customer Rate (Standard):**
  $$\text{Repeat Rate} = \frac{18,912}{22,820} = \mathbf{82.87\%}$$
- **Repeat Rate (All Registered Accounts):** $\frac{18,912}{25,000} = \mathbf{75.65\%}$
- **Significance:** Demonstrates exceptional brand stickiness and customer lifetime engagement.

### 4.3. Customer Lifetime Spend (ARPU)
- **Mathematical Formula:**
  $$\text{Average Spend per Buyer} = \frac{\text{Total Spend}}{\text{Purchasing Customers}} = \frac{₹2,449,898,974.35}{22,820} = \mathbf{₹107,357.68}$$

### 4.4. Customer Churn Rate
- **Definition:** Customer has placed zero orders in $>180$ days preceding the dataset end date (`2025-12-31`).
- **Churned Customer Count:** **13,365 customers**
- **Enterprise Churn Rate:** $\frac{13,365}{25,000} = \mathbf{53.46\%}$
- **At Risk Cohort:** 2,101 customers (last order between 90 and 180 days; prime winback target).

---

## 5. Reverse Logistics & Fulfillment SLA Metrics

### 5.1. Order-Level Return Rate
- **Formula:** $\frac{\text{Returned Orders}}{\text{Total Placed Orders}} = \frac{6,681}{100,000} = \mathbf{6.68\%}$
- **Fulfilled Return Rate:** $\frac{\text{Returned Orders}}{\text{Delivered} + \text{Returned}} = \frac{6,681}{87,360} = \mathbf{7.65\%}$

### 5.2. Item-Level Return Rate
- **Formula:** $\frac{\text{Returned Line Items}}{\text{Total Order Items}} = \frac{12,075}{254,331} = \mathbf{4.75\%}$
- **Total Reverse Logistics Refunds:** **₹105,898,381.80**

### 5.3. On-Time Delivery Rate & Logistics Delay Rate
- **Total Dispatched Shipments:** **89,681 shipments**
- **Delayed Shipments (`delayed_flag == 1`):** **37,448 shipments** (41.76%)
- **On-Time Delivery Rate (SLA Compliance):**
  $$\text{On-Time Delivery \%} = 1 - 0.4176 = \mathbf{58.24\%}$$
- **Average Transit Time:** **5.62 Days** (Standard: 5.8d, Express: 3.8d, Next-Day: 1.9d, Economy: 7.2d).

---

## 6. Marketing Attribution & Efficiency Metrics

### 6.1. Return on Ad Spend / Campaign ROI
- **Formula:**
  $$\text{Campaign ROI} = \frac{\text{Attributed Revenue Generated}}{\text{Campaign Ad Spend Cost}} = \frac{₹436,296,290.91}{₹187,120,911.26} = \mathbf{2.33\times}$$
- **Channel Stratification:**
  - *Direct / Mobile App:* **6.20x** (₹7.6M spend $\to$ ₹47.2M revenue)
  - *Organic Search:* **5.31x** (₹14.4M spend $\to$ ₹76.7M revenue)
  - *Referral:* **5.09x** (₹9.4M spend $\to$ ₹47.8M revenue)
  - *Affiliate:* **2.22x** (₹11.0M spend $\to$ ₹24.4M revenue)
  - *Paid Search:* **1.89x** (₹36.3M spend $\to$ ₹68.6M revenue)
  - *Facebook Ads:* **1.08x** (₹37.9M spend $\to$ ₹40.9M revenue)

### 6.2. Cost Per Acquisition (CPA / CAC)
- **Formula:** $\frac{\text{Campaign Ad Spend}}{\text{Total Conversions}} = \frac{₹187,120,911.26}{17,179\text{ orders}} = \mathbf{₹10,892.42\text{ per order}}$

---

## 7. Customer Satisfaction & Feedback Metrics

### 7.1. Average Review Rating
- **Formula:** $\frac{\sum \text{rating}}{\text{Total Reviews}} = \frac{286,093}{77,530} = \mathbf{3.69\text{ / 5.0 Stars}}$
- **Star Rating Distribution:**
  - 5 Stars: 15,046 (19.41%)
  - 4 Stars: 31,265 (40.33%)
  - 3 Stars: 24,208 (31.22%)
  - 2 Stars: 6,311 (8.14%)
  - 1 Star: 700 (0.90%)

### 7.2. Review Sentiment Share
- **Positive Sentiment:** **46,168 reviews (59.71%)**
- **Neutral Sentiment:** **24,132 reviews (31.22%)**
- **Negative Sentiment:** **6,997 reviews (9.05%)**
- **Verified Purchase Proportion:** **92.00% (71,328 reviews)**

---
*Metric definitions and governance certified. Document saved in `/Users/zeeshan/sale/KPI_Definitions.md`.*
