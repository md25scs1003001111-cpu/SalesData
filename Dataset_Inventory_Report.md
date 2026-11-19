# Comprehensive Dataset Inventory & Schema Architecture Report
**Project:** Indian E-Commerce Sales & Customer Intelligence Analytics  
**Institution:** IILM University — Department of Computer Science & Engineering  
**Course:** B.Tech Computer Science & Engineering (Data Science & Big Data Analytics)  
**Deliverable:** Master Data Inventory & Architecture Specification  
**Status:** Validated & Reconciled  
**Dataset Source:** Kaggle Indian E-Commerce Marketplace Intelligence Dataset (`shiyalkishan01`)  
**Temporal Coverage:** January 1, 2023 – December 31, 2025 (3 Full Calendar Years / 1,096 Days)  

---

## 1. Executive Summary

This inventory report provides a comprehensive technical audit of the Indian E-Commerce Marketplace Dataset utilized to construct the enterprise analytics dashboard. The raw dataset comprises **11 relational CSV files** encapsulating **100,000 orders**, **254,331 order line items**, **25,000 customers**, **550 products across 22 categories**, **100,000 payments**, **89,681 shipments**, **12,075 returns**, **77,530 verified reviews**, and **151 marketing campaigns across 10 acquisition channels**, spanning **15 Indian states and 31 major commercial cities**.

The dataset was engineered with rigorous relational consistency, incorporating latent real-world behavioral drivers such as Indian festival demand spikes (Diwali, Dussehra, Navratri, Eid, Holi, End-of-Year sales), regional logistics lead times, category-specific margin structures, customer churn definitions, and return propensities.

---

## 2. Master Table Inventory & Cardinality

| Table Name | File Name | Total Rows | Columns | Primary Key | Foreign Keys | Functional Grain |
|:---|:---|:---:|:---:|:---|:---|:---|
| **Orders** | `orders.csv` | 100,000 | 16 | `order_id` | `customer_id`, `campaign_id` | 1 row per order transaction |
| **Order Items** | `order_items.csv` | 254,331 | 9 | `order_item_id` | `order_id`, `product_id` | 1 row per product line item within an order |
| **Customers** | `customers.csv` | 25,000 | 18 | `customer_id` | *None* | 1 row per registered customer |
| **Products** | `products.csv` | 550 | 14 | `product_id` | *None* | 1 row per catalog SKU |
| **Payments** | `payments.csv` | 100,000 | 8 | `payment_id` | `order_id` | 1 row per order payment attempt / settlement |
| **Shipments** | `shipments.csv` | 89,681 | 11 | `shipment_id` | `order_id` | 1 row per fulfilled / dispatched order |
| **Returns** | `returns.csv` | 12,075 | 8 | `return_id` | `order_id`, `customer_id`, `product_id` | 1 row per returned product item |
| **Customer Reviews** | `customer_reviews.csv` | 77,530 | 8 | `review_id` | `order_id`, `customer_id`, `product_id` | 1 row per product review submitted |
| **Marketing Campaigns** | `marketing_campaigns.csv` | 151 | 14 | `campaign_id` | *None* | 1 row per promotional marketing campaign |
| **Data Dictionary** | `data_dictionary.csv` | 106 | 8 | `(table_name, column_name)` | *None* | Schema documentation reference metadata |
| **Validation Report** | `dataset_validation_report.csv` | 43 | 5 | `validation_check` | *None* | Pre-computed dataset audit results |

---

## 3. Detailed Table Specifications & Column Dictionaries

### 3.1. `orders.csv` (100,000 Records)
Represents the transactional header for each order placed on the platform.
- `order_id` (String, PK): Unique order identifier (`ORD1000000` to `ORD1099999`).
- `customer_id` (String, FK): Foreign key referencing `customers.csv`.
- `order_date` (Date): Date of transaction (`2023-01-01` to `2025-12-31`).
- `order_time` (Time): Timestamp of order placement (`HH:MM:SS`).
- `order_status` (Enum): Transaction fulfillment status (`Delivered`, `Cancelled`, `Returned`, `Failed`, `Processing`).
- `shipping_method` (Enum): Delivery speed tier selected (`Standard`, `Express`, `Next-Day`, `Economy/Surface`).
- `delivery_city` (String): Destination city across 31 Indian urban centers.
- `delivery_state` (String): Destination state across 15 Indian states.
- `coupon_code` (String): Promotional discount voucher applied (e.g., `FEST20`, `WELCOME10`, or `None`).
- `discount_percentage` (Float): Percentage discount deducted from base prices.
- `marketing_channel` (Enum): First/last-touch acquisition channel attributing the order.
- `subtotal` (Float): Gross item revenue sum before tax and shipping fees (reconciled exactly against `order_items`).
- `shipping_fee` (Float): Freight fee charged to the customer (waived for orders over ₹999 or Premium members).
- `tax_amount` (Float): Goods and Services Tax (GST) calculated per item category tax slab.
- `final_amount` (Float): Net payable billing amount (`subtotal - discount + shipping_fee + tax_amount`).
- `campaign_id` (String, FK, Nullable): Foreign key linking attributed orders to specific marketing initiatives.

### 3.2. `order_items.csv` (254,331 Records)
Granular product-level line items comprising the basket of each order.
- `order_item_id` (String, PK): Unique line item identifier (`OI10000000` to `OI10254330`).
- `order_id` (String, FK): Parent order reference linking to `orders.csv`.
- `product_id` (String, FK): Catalog SKU identifier linking to `products.csv`.
- `quantity` (Integer): Units ordered per SKU (ranging from 1 to 15).
- `unit_price` (Float): Retail selling price per unit at transaction time.
- `discount_percentage` (Float): Line-item specific promotional discount.
- `item_revenue` (Float): Net revenue realized from the line item (`quantity * unit_price * (1 - discount)`).
- `item_cost` (Float): Cost of Goods Sold (COGS) incurred to supply the line item.
- `profit` (Float): Gross margin contribution (`item_revenue - item_cost`).

### 3.3. `customers.csv` (25,000 Records)
Customer master repository with demographic attributes, preferences, and RFM-style aggregates.
- `customer_id` (String, PK): Unique customer identifier (`CUST100000` to `CUST124999`).
- `customer_signup_date` (Date): Platform onboarding registration date.
- `gender` (Enum): Demographic gender (`Male`, `Female`, `Other`).
- `age` (Float): Customer age in years (18 to 72; median: 38.0).
- `age_group` (Enum): Standard age bracket (`18-24`, `25-34`, `35-44`, `45-54`, `55+`).
- `state` / `city` / `pincode_prefix`: Registered geographic residency in India.
- `customer_segment` (Enum): Value cohort (`Regular`, `Budget`, `New`, `Premium`).
- `preferred_device` (Enum): Platform touchpoint (`Mobile App`, `Desktop`, `Mobile Web`).
- `preferred_payment_method` (Enum): Primary payment choice (`UPI`, `Cash on Delivery`, `Credit Card`, `Debit Card`, etc.).
- `acquisition_channel` (Enum): Origin marketing channel responsible for initial conversion.
- `total_orders` (Integer): Recomputed lifetime count of placed orders.
- `total_spend` (Float): Recomputed cumulative monetary spend in INR.
- `last_order_date` (Date): Date of most recent purchase.
- `average_order_value` (Float): Customer-specific AOV (`total_spend / total_orders`).
- `customer_status` (Enum): Lifecycle stage (`Active`, `At Risk`, `Churned`, `Never Purchased`).
- `loyalty_tier` (Enum): Gamified rewards classification (`Bronze`, `Silver`, `Gold`, `Platinum`).

### 3.4. `products.csv` (550 Records)
Product master catalog spanning 22 consumer retail categories.
- `product_id` (String, PK): SKU identifier (`PROD10000` to `PROD10549`).
- `product_name` (String): Formatted catalog title.
- `category` (Enum): Primary category classification (22 distinct categories).
- `subcategory` (String): Niche classification within category.
- `brand` (String): Merchant brand name.
- `price` (Float): Standard catalog retail price (MSRP).
- `cost_price` (Float): Procurement wholesale cost price.
- `discount_range` (String): Standard promotional discount corridor (e.g., `0-15%`, `10-30%`).
- `rating_average` (Float): Rolling average customer review rating (1.0 to 5.0).
- `rating_count` (Integer): Cumulative volume of submitted product reviews.
- `stock_quantity` (Integer): Real-time inventory available across regional fulfillment centers.
- `product_launch_date` (Date): First commercial listing date.
- `product_type` (Enum): Market tiering (`Mass`, `Mid-Market`, `Premium`).
- `return_rate_baseline` (Float): Latent category propensity factor governing return probability.

### 3.5. `payments.csv` (100,000 Records)
Transaction settlement logs corresponding 1:1 with placed orders.
- `payment_id` (String, PK): Unique financial ledger transaction ID (`PAY1000000` to `PAY1099999`).
- `order_id` (String, FK): Associated order reference.
- `payment_date` (Date): Transaction processing date.
- `payment_method` (Enum): Payment gateway channel (`UPI`, `Cash on Delivery`, `Credit Card`, `Debit Card`, `Net Banking`, `Wallet`, `EMI`, `Pay Later`, `Gift Card`, `Cardless EMI`).
- `payment_status` (Enum): Gateway outcome (`Success`, `Refunded`, `Failed`, `Not Charged`, `Pending`).
- `amount_paid` (Float): Net settled amount collected from customer.
- `transaction_fee` (Float): Gateway processing fee absorbed by the business.
- `refund_amount` (Float): Total monetary value remitted back to customer upon return/cancellation.

### 3.6. `shipments.csv` (89,681 Records)
Outbound fulfillment logs for orders transitioning into the physical logistics pipeline.
- `shipment_id` (String, PK): Unique consignment tracking number (`SHIP1000000` to `SHIP1089680`).
- `order_id` (String, FK): Order fulfilled.
- `warehouse_city` (String): Origin hub (e.g., Chennai, Bengaluru, Mumbai, Delhi, Kolkata, Hyderabad).
- `delivery_city` (String): Destination delivery terminal.
- `shipping_method` (Enum): Delivery speed tier.
- `dispatch_date` / `expected_delivery_date` / `actual_delivery_date` (Dates): Logistics milestones.
- `delivery_days` (Float): Actual elapsed transit time in days.
- `delivery_status` (Enum): Final delivery outcome (`Delivered`, `Delivered then Returned`, `Cancelled in Transit`).
- `delayed_flag` (Integer): Binary indicator (`1` if `actual_delivery_date > expected_delivery_date`, else `0`).

### 3.7. `returns.csv` (12,075 Records)
Reverse logistics documentation for item-level return events.
- `return_id` (String, PK): Reverse RMA identifier (`RET100000` to `RET112074`).
- `order_id` (String, FK): Parent order reference.
- `customer_id` (String, FK): Returning customer reference.
- `product_id` (String, FK): Specific SKU returned.
- `return_date` (Date): Date return request was authorized.
- `return_reason` (Enum): Customer-cited reason (`Size issue`, `Changed mind`, `Product damaged`, `Quality issue`, `Product not as expected`, `Wrong product`, `Late delivery`, `Other`).
- `refund_amount` (Float): Value refunded for the returned item.
- `return_status` (Enum): Return lifecycle status (`Refund Completed`, `Rejected`, `In Inspection`).

### 3.8. `customer_reviews.csv` (77,530 Records)
Post-purchase customer satisfaction ratings and sentiment classifications.
- `review_id` (String, PK): Customer review record ID (`REV100000` to `REV177529`).
- `order_id` / `customer_id` / `product_id` (FKs): Composite entity references.
- `review_date` (Date): Review submission date.
- `rating` (Integer): Discrete star rating (1 to 5).
- `review_sentiment` (Enum): Categorical sentiment classification (`Positive`, `Neutral`, `Negative`).
- `verified_purchase` (Boolean): Flag verifying authenticated transaction history (`True` or `False`).

### 3.9. `marketing_campaigns.csv` (151 Records)
Marketing campaign attribution, ad spend, conversion, and return on investment ledger.
- `campaign_id` (String, PK): Promotional campaign ID (`CAMP1001` to `CAMP1151`).
- `campaign_name` (String): Descriptive marketing campaign title.
- `channel` (Enum): Acquisition channel (`Organic Search`, `Paid Search`, `Referral`, `Direct/App`, `Email Marketing`, `Social Media - Facebook`, `Social Media - Instagram`, `Influencer Marketing`, `Affiliate`, `SMS Marketing`).
- `campaign_start_date` / `campaign_end_date` (Dates): Flight duration dates.
- `campaign_type` (Enum): Objective tier (`Festive Special`, `New Product Launch`, `Clearance`, `Flash Sale`, etc.).
- `target_segment` (Enum): Targeted customer demographic segment.
- `discount_percentage` (Float): Campaign promotional discount incentive.
- `impressions` / `clicks` / `conversions` (Integers): Funnel conversion counts.
- `revenue_generated` (Float): Attributed gross revenue tied directly to participating orders.
- `campaign_cost` (Float): Total advertising and media budget spent.
- `roi` (Float): Return on Investment ratio (`revenue_generated / campaign_cost`).

---

## 4. Entity-Relationship Schema Architecture

The relational schema implements a **Snowflake / Multi-Fact Star Schema** optimized for high-performance analytical aggregation:

```
                      +-------------------+
                      |   CUSTOMERS (D)   |
                      |   (25,000 Rows)   |
                      +---------+---------+
                                |
             +------------------+------------------+
             |                                     |
             v                                     v
   +-------------------+                 +-------------------+
   |    ORDERS (F)     |<----------------|   CAMPAIGNS (D)   |
   |  (100,000 Rows)   |  (campaign_id)  |    (151 Rows)     |
   +---------+---------+                 +-------------------+
             |
             +--------------------+--------------------+
             |                    |                    |
             v                    v                    v
   +-------------------+  +-------------------+  +-------------------+
   | ORDER_ITEMS (F)   |  |   PAYMENTS (F)    |  |   SHIPMENTS (F)   |
   |  (254,331 Rows)   |  |  (100,000 Rows)   |  |  (89,681 Rows)    |
   +---------+---------+  +-------------------+  +-------------------+
             |
             +--------------------+
             |                    |
             v                    v
   +-------------------+  +-------------------+
   |   PRODUCTS (D)    |  |    RETURNS (F)    |
   |    (550 Rows)     |  |   (12,075 Rows)   |
   +---------+---------+  +-------------------+
             |
             v
   +-------------------+
   |  REVIEWS (F/D)    |
   |   (77,530 Rows)   |
   +-------------------+
```

---

## 5. Analytical Grain & Avoidance of Duplicate Counting

To guarantee zero revenue distortion or duplicated metrics:
1. **Gross Revenue & AOV** are aggregated strictly at the `orders` header grain (`subtotal`) or bottom-up from `order_items` (`item_revenue`). Joining `orders` to `order_items` without aggregating line items first would multiply order metrics by 2.54x!
2. **Gross Profit & COGS** exist exclusively at the `order_items` level and are summed directly without cross-joining to `shipments` or `payments`.
3. **Net Revenue** accounts for product returns by subtracting `returns.refund_amount` exclusively from delivered order totals.
4. **Logistics Performance** is computed on `shipments` (89,681 orders dispatched), ensuring that cancelled and failed orders (which never ship) do not artificially skew delivery lead times or on-time fulfillment rates.

---
*Report compiled autonomously following data verification. Master copy saved in `/Users/zeeshan/sale/Dataset_Inventory_Report.md`.*
