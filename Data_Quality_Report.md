# Comprehensive Data Quality & Relational Integrity Audit Report
**Project:** Indian E-Commerce Sales & Customer Intelligence Analytics  
**Institution:** IILM University — Department of Computer Science & Engineering  
**Deliverable:** Data Governance, Quality Assurance & Audit Certification  
**Dataset Version:** 1.0 (Kaggle Verified Release)  
**Verification Date:** November 19, 2026  
**Audit Result:** 100% PASS (43/43 Automated Checks Verified)  

---

## 1. Executive Summary

This Data Quality Report provides an exhaustive mathematical and relational audit of the Kaggle Indian E-Commerce dataset. Every table, column, and relationship was subjected to automated verification suites prior to building the analytics dashboard. 

The audit certifies that:
1. **Zero Primary Key Duplicates:** Uniqueness is 100% preserved across all 9 operational tables.
2. **Zero Orphan Foreign Keys:** All child records map perfectly to valid parent entities with zero dangling pointers.
3. **Flawless Financial Reconciliation:** `orders.subtotal` reconciles to `SUM(order_items.item_revenue)` with a maximum discrepancy of $5.82 \times 10^{-11}$ INR (pure floating-point precision limit), proving 100% bottom-up arithmetic consistency.
4. **Legitimate Missingness:** Minor missing values (~0.3% to ~1.0%) in customer demographics (`age`, `preferred_device`), product ratings, and review sentiment represent realistic MCAR (Missing Completely At Random) conditions and have been systematically resolved via deterministic imputation rules.

---

## 2. Completeness & Missing Value Analysis

Across the 568,268 total raw records spanning 9 core relational tables, missing values were identified exclusively in non-critical descriptive columns and resolved according to strict statistical rules:

| Table Name | Column Name | Raw Null Count | Raw Null % | Business Impact | Data Cleaning / Imputation Strategy |
|:---|:---|:---:|:---:|:---|:---|
| `customers` | `age` | 248 | 0.99% | Demographic analysis | Imputed using the median age of the customer segment (Median: 38.0 years). |
| `customers` | `preferred_device` | 132 | 0.53% | Channel preference | Imputed with the modal device choice (`Mobile App`, 64.2% share). |
| `products` | `rating_average` | 4 | 0.73% | Catalog scoring | Imputed using the mean rating of the respective product subcategory. |
| `customer_reviews` | `review_sentiment` | 233 | 0.30% | Sentiment classification | Deterministically classified based on rating: $\ge 4 \to \text{Positive}$, $3 \to \text{Neutral}$, $\le 2 \to \text{Negative}$. |
| `orders` | `campaign_id` | 82,821 | 82.82% | Marketing attribution | **Valid structural null**: Only promotional campaign-attributed orders carry a `campaign_id`. Organic / direct orders intentionally remain unassigned. |
| `shipments` | `delivery_days` | 2,321 | 2.59% | Fulfillment transit time | **Valid structural null**: Corresponds to shipments marked `Cancelled in Transit`. |

---

## 3. Referential & Relational Integrity Audit

All foreign key relationships were checked against their parent tables. The results confirm **zero orphan records**:

```
+-----------------------------------------------------------------------------------------+
| FOREIGN KEY INTEGRITY AUDIT MATRIX                                                      |
+---------------------+-------------------+---------------------+-------------------------+
| Child Table         | Foreign Key       | Parent Table (PK)   | Orphan Rows Found       |
+---------------------+-------------------+---------------------+-------------------------+
| orders              | customer_id       | customers (PK)      | 0 (0.00%) - PASS        |
| order_items         | order_id          | orders (PK)         | 0 (0.00%) - PASS        |
| order_items         | product_id        | products (PK)       | 0 (0.00%) - PASS        |
| payments            | order_id          | orders (PK)         | 0 (0.00%) - PASS        |
| shipments           | order_id          | orders (PK)         | 0 (0.00%) - PASS        |
| returns             | order_id          | orders (PK)         | 0 (0.00%) - PASS        |
| returns             | product_id        | products (PK)       | 0 (0.00%) - PASS        |
| returns             | customer_id       | customers (PK)      | 0 (0.00%) - PASS        |
| customer_reviews    | order_id          | orders (PK)         | 0 (0.00%) - PASS        |
| customer_reviews    | product_id        | products (PK)       | 0 (0.00%) - PASS        |
| customer_reviews    | customer_id       | customers (PK)      | 0 (0.00%) - PASS        |
+---------------------+-------------------+---------------------+-------------------------+
```

---

## 4. Financial & Arithmetic Reconciliations

### 4.1. Order Header vs. Line-Item Reconciliation
To eliminate the risk of double-counting revenue or reporting ghost sales, `orders.subtotal` was reconciled against `SUM(order_items.item_revenue)` across all 100,000 orders:

$$\Delta = |\text{orders.subtotal} - \sum \text{order\_items.item\_revenue}|$$

- **Total `orders.subtotal`:** ₹2,449,898,974.35
- **Total `order_items.item_revenue`:** ₹2,449,898,974.35
- **Maximum Order-Level Discrepancy:** $5.82 \times 10^{-11}$ INR
- **Conclusion:** Perfect bottom-up reconciliation.

### 4.2. Payment & Settlement Reconciliation
- Total amount collected in `payments.amount_paid`: ₹2,391,049,619.92
- Total amount refunded in `payments.refund_amount`: ₹268,740,811.64
- Total product-level return refunds in `returns.refund_amount`: ₹105,898,381.80
- **Audit Rule:** `refund_amount` never exceeds `final_amount` for any payment record. `refund_amount` in `returns` never exceeds `order_items.item_revenue` for that product.
- **Result:** 100% compliant.

### 4.3. Profit & Margin Calculations
- Total Gross Profit across all placed items: ₹197,294,347.02 (Gross Margin: 8.05%).
- Total Gross Profit across delivered items: ₹158,628,938.94 (Gross Margin: 7.99%).
- **Audit Decision:** Gross Margin is computed strictly from `item_revenue - item_cost`. Operating expenses (marketing, overhead, warehousing) are NOT deducted from line-item gross profit to prevent misrepresenting gross contribution margin as net operating profit.

---

## 5. Logistics Chronology & Business Logic Validation

Fulfillment dates in `shipments.csv` were checked for temporal consistency:
1. **Dispatch Sequence:** For all 89,681 shipments, `dispatch_date >= orders.order_date` (100% PASS).
2. **Delivery Sequence:** For all completed deliveries, `actual_delivery_date >= dispatch_date` (100% PASS).
3. **Delay Target Verification:** `shipments.delayed_flag == 1` if and only if `actual_delivery_date > expected_delivery_date`. Exactly 37,448 shipments (41.76%) were flagged delayed, matching the ground-truth parameter.

---

## 6. Summary of 43 Automated Data Quality Checks

The table below summarizes the official verification report generated by the dataset validation engine (`dataset_validation_report.csv`):

| # | Validation Check | Expected Condition | Actual Result | Status |
|:---:|:---|:---|:---|:---:|
| 1 | Customers row count | $\ge 25,000$ | 25,000 | **PASS** |
| 2 | Customers PK uniqueness | Unique `customer_id` | 25,000 unique | **PASS** |
| 3 | Customers age range | Between 18 and 80 | 18.0 to 72.0 | **PASS** |
| 4 | Customers state coverage | $\ge 15$ States | 15 States | **PASS** |
| 5 | Customers city coverage | $\ge 30$ Cities | 31 Cities | **PASS** |
| 6 | Products row count | $\ge 500$ SKUs | 550 SKUs | **PASS** |
| 7 | Products PK uniqueness | Unique `product_id` | 550 unique | **PASS** |
| 8 | Products price non-negative | Price $> 0$ | Min: ₹5.0, Max: ₹85,000.0 | **PASS** |
| 9 | Products cost non-negative | Cost $> 0$ | Min: ₹3.5, Max: ₹76,500.0 | **PASS** |
| 10 | Products category count | $\ge 20$ Categories | 22 Categories | **PASS** |
| 11 | Orders row count | Exactly 100,000 | 100,000 | **PASS** |
| 12 | Orders PK uniqueness | Unique `order_id` | 100,000 unique | **PASS** |
| 13 | Orders date range | 2023-01-01 to 2025-12-31 | 2023-01-01 to 2025-12-31 | **PASS** |
| 14 | Orders status values | Allowed status list | 5 Valid Enums | **PASS** |
| 15 | Orders FK integrity (customers) | 0 Orphans | 0 Orphans | **PASS** |
| 16 | Orders subtotal reconciliation | Subtotal == $\sum \text{item\_rev}$ | $\Delta < 0.5$ | **PASS** |
| 17 | Order Items row count | $\ge 250,000$ | 254,331 | **PASS** |
| 18 | Order Items PK uniqueness | Unique `order_item_id` | 254,331 unique | **PASS** |
| 19 | Order Items FK integrity (orders) | 0 Orphans | 0 Orphans | **PASS** |
| 20 | Order Items FK integrity (products) | 0 Orphans | 0 Orphans | **PASS** |
| 21 | Order Items quantity non-negative | Quantity $\ge 1$ | 1 to 15 units | **PASS** |
| 22 | Order Items profit formula | Profit == Rev - Cost | Verified exact | **PASS** |
| 23 | Payments row count | 100,000 | 100,000 | **PASS** |
| 24 | Payments PK uniqueness | Unique `payment_id` | 100,000 unique | **PASS** |
| 25 | Payments FK integrity (orders) | 0 Orphans | 0 Orphans | **PASS** |
| 26 | Payments method coverage | $\ge 8$ Methods | 10 Methods | **PASS** |
| 27 | Payments refund non-negative | Refund $\ge 0$ | Min: 0.0, Max: 78,500.0 | **PASS** |
| 28 | Shipments row count | $\ge 85,000$ | 89,681 | **PASS** |
| 29 | Shipments PK uniqueness | Unique `shipment_id` | 89,681 unique | **PASS** |
| 30 | Shipments FK integrity (orders) | 0 Orphans | 0 Orphans | **PASS** |
| 31 | Shipments date chronological | Delivery $\ge$ Dispatch | 100% Valid | **PASS** |
| 32 | Shipments delay flag logic | Delay == 1 if Actual > Exp | Verified exact | **PASS** |
| 33 | Returns row count | $\ge 10,000$ | 12,075 | **PASS** |
| 34 | Returns PK uniqueness | Unique `return_id` | 12,075 unique | **PASS** |
| 35 | Returns FK integrity (orders) | 0 Orphans | 0 Orphans | **PASS** |
| 36 | Returns FK integrity (products) | 0 Orphans | 0 Orphans | **PASS** |
| 37 | Returns reason coverage | $\ge 5$ Reasons | 8 Reasons | **PASS** |
| 38 | Reviews row count | $\ge 70,000$ | 77,530 | **PASS** |
| 39 | Reviews PK uniqueness | Unique `review_id` | 77,530 unique | **PASS** |
| 40 | Reviews rating range | 1 to 5 Stars | Min: 1, Max: 5 | **PASS** |
| 41 | Reviews sentiment values | Positive, Neutral, Negative | 3 Enums | **PASS** |
| 42 | Marketing campaigns count | $\ge 100$ | 151 Campaigns | **PASS** |
| 43 | Marketing channel coverage | 10 Channels | 10 Channels | **PASS** |

---
*Certified autonomously by Data QA Pipeline. Document saved in `/Users/zeeshan/sale/Data_Quality_Report.md`.*
