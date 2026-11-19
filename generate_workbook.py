"""
Indian E-Commerce Sales Analytics Dashboard Generator
Creates a professional, interactive Microsoft Excel workbook
Indian_Ecommerce_Sales_Analytics_Dashboard.xlsx
using XlsxWriter and clean data from the Kaggle dataset.
"""

import os
import pandas as pd
import numpy as np
import xlsxwriter

def build_dashboard():
    raw_dir = '/Users/zeeshan/sale/raw_dataset'
    clean_dir = '/Users/zeeshan/sale/cleaned_data'
    output_file = '/Users/zeeshan/sale/Indian_Ecommerce_Sales_Analytics_Dashboard.xlsx'

    print("Loading data...")
    orders = pd.read_csv(f'{raw_dir}/orders.csv')
    items = pd.read_csv(f'{raw_dir}/order_items.csv')
    customers = pd.read_csv(f'{raw_dir}/customers.csv')
    products = pd.read_csv(f'{raw_dir}/products.csv')
    payments = pd.read_csv(f'{raw_dir}/payments.csv')
    shipments = pd.read_csv(f'{raw_dir}/shipments.csv')
    returns = pd.read_csv(f'{raw_dir}/returns.csv')
    reviews = pd.read_csv(f'{raw_dir}/customer_reviews.csv')
    campaigns = pd.read_csv(f'{raw_dir}/marketing_campaigns.csv')
    data_dict = pd.read_csv(f'{raw_dir}/data_dictionary.csv')
    val_rep = pd.read_csv(f'{raw_dir}/dataset_validation_report.csv')

    # Date processing
    orders['order_date'] = pd.to_datetime(orders['order_date'])
    orders['year'] = orders['order_date'].dt.year
    orders['month'] = orders['order_date'].dt.month
    orders['year_month'] = orders['order_date'].dt.to_period('M').astype(str)
    orders['quarter'] = 'Q' + orders['order_date'].dt.quarter.astype(str)
    orders['year_quarter'] = orders['year'].astype(str) + '-' + orders['quarter']

    print("Pre-aggregating analytical cubes...")
    # 1. Monthly sales
    items_month = items.merge(orders[['order_id', 'year', 'month', 'year_month', 'year_quarter', 'order_status']], on='order_id')
    monthly_all = items_month.groupby(['year', 'month', 'year_month']).agg(
        orders_count=('order_id', 'nunique'),
        units_sold=('quantity', 'sum'),
        subtotal_rev=('item_revenue', 'sum'),
        cost=('item_cost', 'sum'),
        profit=('profit', 'sum')
    ).reset_index().sort_values(['year', 'month'])
    monthly_all['margin_pct'] = monthly_all['profit'] / monthly_all['subtotal_rev']
    monthly_all['aov'] = monthly_all['subtotal_rev'] / monthly_all['orders_count']
    monthly_all['mom_growth'] = monthly_all['subtotal_rev'].pct_change().fillna(0.0)

    # 2. Quarterly sales
    quarterly_all = items_month.groupby('year_quarter').agg(
        orders_count=('order_id', 'nunique'),
        units_sold=('quantity', 'sum'),
        subtotal_rev=('item_revenue', 'sum'),
        profit=('profit', 'sum')
    ).reset_index()
    quarterly_all['margin_pct'] = quarterly_all['profit'] / quarterly_all['subtotal_rev']

    # 3. Category performance
    item_prod = items.merge(products[['product_id', 'product_name', 'category', 'subcategory', 'brand']], on='product_id')
    cat_all = item_prod.groupby('category').agg(
        sku_count=('product_id', 'nunique'),
        units_sold=('quantity', 'sum'),
        revenue=('item_revenue', 'sum'),
        cost=('item_cost', 'sum'),
        profit=('profit', 'sum')
    ).reset_index().sort_values('revenue', ascending=False)
    cat_all['margin_pct'] = cat_all['profit'] / cat_all['revenue']
    cat_all['rev_share'] = cat_all['revenue'] / cat_all['revenue'].sum()

    # Category returns
    ret_prod = returns.merge(products[['product_id', 'category']], on='product_id')
    ret_cat = ret_prod.groupby('category').agg(
        return_count=('return_id', 'count'),
        refund_amount=('refund_amount', 'sum')
    ).reset_index()
    cat_all = cat_all.merge(ret_cat, on='category', how='left')
    cat_all['return_count'] = cat_all['return_count'].fillna(0).astype(int)
    cat_all['refund_amount'] = cat_all['refund_amount'].fillna(0.0)
    cat_all['return_rate'] = cat_all['return_count'] / cat_all['units_sold']

    # 4. Top 20 Products
    prod_top20 = item_prod.groupby(['product_id', 'product_name', 'category']).agg(
        units_sold=('quantity', 'sum'),
        revenue=('item_revenue', 'sum'),
        cost=('item_cost', 'sum'),
        profit=('profit', 'sum')
    ).reset_index().sort_values('revenue', ascending=False).head(20)
    prod_top20['margin_pct'] = prod_top20['profit'] / prod_top20['revenue']
    prod_top20['avg_selling_price'] = prod_top20['revenue'] / prod_top20['units_sold']

    # 5. Customer Segments
    cust_seg = customers.groupby('customer_segment').agg(
        customer_count=('customer_id', 'count'),
        total_orders=('total_orders', 'sum'),
        total_spend=('total_spend', 'sum'),
        avg_aov=('average_order_value', 'mean')
    ).reset_index().sort_values('total_spend', ascending=False)
    cust_seg['spend_share'] = cust_seg['total_spend'] / cust_seg['total_spend'].sum()

    # 6. Customer Status (Lifecycle/Churn)
    cust_stat = customers.groupby('customer_status').agg(
        customer_count=('customer_id', 'count'),
        total_spend=('total_spend', 'sum')
    ).reset_index().sort_values('customer_count', ascending=False)
    cust_stat['cust_share'] = cust_stat['customer_count'] / cust_stat['customer_count'].sum()

    # 7. Loyalty Tiers
    cust_tier = customers.groupby('loyalty_tier').agg(
        customer_count=('customer_id', 'count'),
        total_spend=('total_spend', 'sum')
    ).reset_index().sort_values('total_spend', ascending=False)
    cust_tier['spend_share'] = cust_tier['total_spend'] / cust_tier['total_spend'].sum()

    # 8. Demographics (Age groups)
    cust_age = customers.groupby(['age_group', 'gender']).agg(
        customer_count=('customer_id', 'count'),
        total_spend=('total_spend', 'sum')
    ).reset_index()

    # 9. Geography (States)
    state_all = orders.groupby('delivery_state').agg(
        orders_count=('order_id', 'count'),
        revenue=('subtotal', 'sum'),
        final_amount=('final_amount', 'sum'),
        delivered_count=('order_status', lambda x: (x == 'Delivered').sum()),
        returned_count=('order_status', lambda x: (x == 'Returned').sum())
    ).reset_index().sort_values('revenue', ascending=False)
    state_all['rev_share'] = state_all['revenue'] / state_all['revenue'].sum()
    state_all['aov'] = state_all['revenue'] / state_all['orders_count']
    state_all['delivery_rate'] = state_all['delivered_count'] / state_all['orders_count']
    state_all['return_rate'] = state_all['returned_count'] / state_all['orders_count']

    # 10. Geography (Top 25 Cities)
    city_all = orders.groupby(['delivery_city', 'delivery_state']).agg(
        orders_count=('order_id', 'count'),
        revenue=('subtotal', 'sum'),
        delivered_count=('order_status', lambda x: (x == 'Delivered').sum())
    ).reset_index().sort_values('revenue', ascending=False).head(25)
    city_all['rev_share'] = city_all['revenue'] / orders['subtotal'].sum()
    city_all['aov'] = city_all['revenue'] / city_all['orders_count']

    # 11. Payments
    pay_meth = payments.groupby('payment_method').agg(
        transactions=('payment_id', 'count'),
        amount_paid=('amount_paid', 'sum'),
        refund_amount=('refund_amount', 'sum'),
        success_count=('payment_status', lambda x: (x == 'Success').sum()),
        fee_est=('transaction_fee', 'sum')
    ).reset_index().sort_values('amount_paid', ascending=False)
    pay_meth['share_pct'] = pay_meth['amount_paid'] / pay_meth['amount_paid'].sum()
    pay_meth['success_rate'] = pay_meth['success_count'] / pay_meth['transactions']

    pay_stat = payments.groupby('payment_status').agg(
        count=('payment_id', 'count'),
        amount=('amount_paid', 'sum')
    ).reset_index().sort_values('count', ascending=False)
    pay_stat['count_share'] = pay_stat['count'] / pay_stat['count'].sum()

    # 12. Returns reasons
    ret_reason = returns.groupby('return_reason').agg(
        return_count=('return_id', 'count'),
        refund_amount=('refund_amount', 'sum')
    ).reset_index().sort_values('return_count', ascending=False)
    ret_reason['share_pct'] = ret_reason['return_count'] / ret_reason['return_count'].sum()

    # 13. Shipments
    ship_meth = shipments.groupby('shipping_method').agg(
        shipment_count=('shipment_id', 'count'),
        delivered_count=('delivery_status', lambda x: (x.str.startswith('Delivered')).sum()),
        delayed_count=('delayed_flag', lambda x: (x == 1).sum()),
        avg_delivery_days=('delivery_days', 'mean')
    ).reset_index().sort_values('shipment_count', ascending=False)
    ship_meth['delay_rate'] = ship_meth['delayed_count'] / ship_meth['shipment_count']
    ship_meth['on_time_rate'] = 1.0 - ship_meth['delay_rate']

    # 14. Marketing channels
    mkt_chan = campaigns.groupby('channel').agg(
        campaigns_count=('campaign_id', 'count'),
        ad_spend=('campaign_cost', 'sum'),
        impressions=('impressions', 'sum'),
        clicks=('clicks', 'sum'),
        conversions=('conversions', 'sum'),
        revenue_gen=('revenue_generated', 'sum')
    ).reset_index().sort_values('revenue_gen', ascending=False)
    mkt_chan['roi'] = mkt_chan['revenue_gen'] / mkt_chan['ad_spend']
    mkt_chan['cpa'] = mkt_chan['ad_spend'] / mkt_chan['conversions']
    mkt_chan['cvr'] = mkt_chan['conversions'] / mkt_chan['clicks']
    mkt_chan['ctr'] = mkt_chan['clicks'] / mkt_chan['impressions']

    # 15. Reviews
    rev_stars = reviews.groupby('rating').agg(
        review_count=('review_id', 'count'),
        verified_count=('verified_purchase', lambda x: (x == True).sum()),
        positive_count=('review_sentiment', lambda x: (x == 'Positive').sum()),
        neutral_count=('review_sentiment', lambda x: (x == 'Neutral').sum()),
        negative_count=('review_sentiment', lambda x: (x == 'Negative').sum())
    ).reset_index().sort_values('rating', ascending=True)
    rev_stars['share_pct'] = rev_stars['review_count'] / rev_stars['review_count'].sum()

    rev_sent = reviews['review_sentiment'].fillna('Unknown').value_counts().reset_index()
    rev_sent.columns = ['sentiment', 'count']
    rev_sent['share_pct'] = rev_sent['count'] / rev_sent['count'].sum()

    # Category reviews
    rev_cat = reviews.merge(products[['product_id', 'category']], on='product_id')
    cat_rev_perf = rev_cat.groupby('category').agg(
        review_count=('review_id', 'count'),
        avg_rating=('rating', 'mean'),
        positive_pct=('review_sentiment', lambda x: (x == 'Positive').mean())
    ).reset_index().sort_values('review_count', ascending=False)

    print("Creating Excel workbook:", output_file)
    wb = xlsxwriter.Workbook(output_file, {'nan_inf_to_errors': True})

    # Color definitions
    NAVY = '#174A78'
    ROYAL_BLUE = '#216DA8'
    CYAN = '#08A9D0'
    PALE_BLUE = '#EAF5FB'
    LIGHT_GREY = '#F3F4F6'
    WHITE = '#FFFFFF'
    BORDER_COLOR = '#D5DCE3'
    TEXT_DARK = '#2C3E50'
    TEXT_MUTED = '#5B7A99'
    GREEN = '#2E7D32'
    RED = '#C62828'

    # Common formats
    f_title = wb.add_format({
        'bold': True, 'font_size': 15, 'font_name': 'Calibri',
        'font_color': WHITE, 'bg_color': NAVY,
        'align': 'left', 'valign': 'vcenter', 'left': 1, 'top': 1, 'bottom': 0
    })
    f_subtitle = wb.add_format({
        'bold': False, 'font_size': 9, 'font_name': 'Calibri',
        'font_color': CYAN, 'bg_color': NAVY,
        'align': 'left', 'valign': 'vcenter', 'left': 1, 'bottom': 1
    })
    f_badge = wb.add_format({
        'bold': True, 'font_size': 9, 'font_name': 'Calibri',
        'font_color': WHITE, 'bg_color': ROYAL_BLUE,
        'align': 'center', 'valign': 'vcenter', 'border': 1
    })

    # Navigation formats
    f_nav_active = wb.add_format({
        'bold': True, 'font_size': 9, 'font_name': 'Calibri',
        'font_color': WHITE, 'bg_color': CYAN,
        'align': 'center', 'valign': 'vcenter', 'border': 1
    })
    f_nav_inactive = wb.add_format({
        'bold': True, 'font_size': 9, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'center', 'valign': 'vcenter', 'border': 1, 'underline': 1
    })

    # Control bar formats
    f_ctrl_label = wb.add_format({
        'bold': True, 'font_size': 9, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': LIGHT_GREY,
        'align': 'right', 'valign': 'vcenter', 'border': 1
    })
    f_ctrl_val = wb.add_format({
        'bold': True, 'font_size': 10, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': WHITE,
        'align': 'center', 'valign': 'vcenter', 'border': 1
    })
    f_ctrl_info = wb.add_format({
        'bold': False, 'font_size': 8.5, 'font_name': 'Calibri',
        'font_color': TEXT_MUTED, 'bg_color': PALE_BLUE,
        'align': 'left', 'valign': 'vcenter', 'border': 1, 'italic': True
    })

    # KPI Card formats
    f_card_top = wb.add_format({
        'bold': True, 'font_size': 8.5, 'font_name': 'Calibri',
        'font_color': TEXT_MUTED, 'bg_color': PALE_BLUE,
        'align': 'center', 'valign': 'vcenter',
        'top': 1, 'left': 1, 'right': 1
    })
    f_card_val_curr = wb.add_format({
        'bold': True, 'font_size': 16, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'center', 'valign': 'vcenter',
        'left': 1, 'right': 1, 'num_format': '₹#,##,##0'
    })
    f_card_val_num = wb.add_format({
        'bold': True, 'font_size': 16, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'center', 'valign': 'vcenter',
        'left': 1, 'right': 1, 'num_format': '#,##0'
    })
    f_card_val_pct = wb.add_format({
        'bold': True, 'font_size': 16, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'center', 'valign': 'vcenter',
        'left': 1, 'right': 1, 'num_format': '0.0%'
    })
    f_card_val_dec = wb.add_format({
        'bold': True, 'font_size': 16, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'center', 'valign': 'vcenter',
        'left': 1, 'right': 1, 'num_format': '0.00'
    })
    f_card_bot = wb.add_format({
        'bold': False, 'font_size': 8, 'font_name': 'Calibri',
        'font_color': TEXT_DARK, 'bg_color': PALE_BLUE,
        'align': 'center', 'valign': 'vcenter',
        'bottom': 1, 'left': 1, 'right': 1, 'italic': True
    })

    # Section Headers
    f_sec_hdr = wb.add_format({
        'bold': True, 'font_size': 11, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': LIGHT_GREY,
        'align': 'left', 'valign': 'vcenter', 'bottom': 2, 'bottom_color': CYAN
    })

    # Table formats
    f_th = wb.add_format({
        'bold': True, 'font_size': 9, 'font_name': 'Calibri',
        'font_color': WHITE, 'bg_color': NAVY,
        'align': 'center', 'valign': 'vcenter', 'border': 1
    })
    f_th_left = wb.add_format({
        'bold': True, 'font_size': 9, 'font_name': 'Calibri',
        'font_color': WHITE, 'bg_color': NAVY,
        'align': 'left', 'valign': 'vcenter', 'border': 1
    })
    f_td_text = wb.add_format({
        'font_size': 9, 'font_name': 'Calibri', 'font_color': TEXT_DARK,
        'align': 'left', 'valign': 'vcenter', 'border': 1, 'bg_color': WHITE
    })
    f_td_text_alt = wb.add_format({
        'font_size': 9, 'font_name': 'Calibri', 'font_color': TEXT_DARK,
        'align': 'left', 'valign': 'vcenter', 'border': 1, 'bg_color': LIGHT_GREY
    })
    f_td_num = wb.add_format({
        'font_size': 9, 'font_name': 'Calibri', 'font_color': TEXT_DARK,
        'align': 'right', 'valign': 'vcenter', 'border': 1, 'bg_color': WHITE,
        'num_format': '#,##0'
    })
    f_td_num_alt = wb.add_format({
        'font_size': 9, 'font_name': 'Calibri', 'font_color': TEXT_DARK,
        'align': 'right', 'valign': 'vcenter', 'border': 1, 'bg_color': LIGHT_GREY,
        'num_format': '#,##0'
    })
    f_td_curr = wb.add_format({
        'font_size': 9, 'font_name': 'Calibri', 'font_color': TEXT_DARK,
        'align': 'right', 'valign': 'vcenter', 'border': 1, 'bg_color': WHITE,
        'num_format': '₹#,##,##0'
    })
    f_td_curr_alt = wb.add_format({
        'font_size': 9, 'font_name': 'Calibri', 'font_color': TEXT_DARK,
        'align': 'right', 'valign': 'vcenter', 'border': 1, 'bg_color': LIGHT_GREY,
        'num_format': '₹#,##,##0'
    })
    f_td_pct = wb.add_format({
        'font_size': 9, 'font_name': 'Calibri', 'font_color': TEXT_DARK,
        'align': 'right', 'valign': 'vcenter', 'border': 1, 'bg_color': WHITE,
        'num_format': '0.0%'
    })
    f_td_pct_alt = wb.add_format({
        'font_size': 9, 'font_name': 'Calibri', 'font_color': TEXT_DARK,
        'align': 'right', 'valign': 'vcenter', 'border': 1, 'bg_color': LIGHT_GREY,
        'num_format': '0.0%'
    })
    f_td_dec = wb.add_format({
        'font_size': 9, 'font_name': 'Calibri', 'font_color': TEXT_DARK,
        'align': 'right', 'valign': 'vcenter', 'border': 1, 'bg_color': WHITE,
        'num_format': '0.00'
    })
    f_td_dec_alt = wb.add_format({
        'font_size': 9, 'font_name': 'Calibri', 'font_color': TEXT_DARK,
        'align': 'right', 'valign': 'vcenter', 'border': 1, 'bg_color': LIGHT_GREY,
        'num_format': '0.00'
    })

    # Total row formats
    f_tot_label = wb.add_format({
        'bold': True, 'font_size': 9, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'left', 'valign': 'vcenter', 'top': 1, 'bottom': 6
    })
    f_tot_num = wb.add_format({
        'bold': True, 'font_size': 9, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'right', 'valign': 'vcenter', 'top': 1, 'bottom': 6,
        'num_format': '#,##0'
    })
    f_tot_curr = wb.add_format({
        'bold': True, 'font_size': 9, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'right', 'valign': 'vcenter', 'top': 1, 'bottom': 6,
        'num_format': '₹#,##,##0'
    })
    f_tot_pct = wb.add_format({
        'bold': True, 'font_size': 9, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'right', 'valign': 'vcenter', 'top': 1, 'bottom': 6,
        'num_format': '0.0%'
    })

    sheets_list = [
        ('Executive Dashboard', 'Executive Dashboard'),
        ('Sales Analysis', 'Sales Analysis'),
        ('Product Analysis', 'Product Analysis'),
        ('Customer Analysis', 'Customer Analysis'),
        ('Geographic Analysis', 'Geographic Analysis'),
        ('Payment Analysis', 'Payment Analysis'),
        ('Orders & Returns', 'Orders & Returns'),
        ('Marketing Analysis', 'Marketing Analysis'),
        ('Reviews Analysis', 'Reviews Analysis'),
        ('Data Dictionary & Documentation', 'Data Dictionary & Documentation'),
        ('Supporting Data', 'Supporting Data')
    ]

    def render_header_nav(ws, title_text, sub_text, active_idx, max_col='N'):
        # Banner
        ws.merge_range(f'B2:{max_col}2', title_text, f_title)
        ws.merge_range(f'B3:{max_col}3', sub_text, f_subtitle)
        ws.set_row(1, 28)
        ws.set_row(2, 16)
        ws.set_row(3, 22)

        # Nav bar
        for idx, (sh_name, sh_label) in enumerate(sheets_list):
            col_letter = chr(ord('B') + idx)
            cell_ref = f'{col_letter}4'
            if idx == active_idx:
                ws.write(cell_ref, sh_label, f_nav_active)
            else:
                ws.write_url(cell_ref, f"internal:'{sh_name}'!A1", f_nav_inactive, sh_label)

    # ==========================================
    # 1. EXECUTIVE DASHBOARD
    # ==========================================
    print("Building Executive Dashboard...")
    ws_exec = wb.add_worksheet('Executive Dashboard')
    ws_exec.set_tab_color(NAVY)
    ws_exec.hide_gridlines(0)
    ws_exec.set_column('A:A', 3)
    ws_exec.set_column('B:L', 14.5)
    ws_exec.set_column('M:N', 16)

    render_header_nav(
        ws_exec,
        "INDIAN E-COMMERCE ANALYTICS DASHBOARD | EXECUTIVE OVERVIEW",
        "Strategic Enterprise Performance & Omnichannel Intelligence (FY 2023 - 2025) | IILM University Data Science Reference",
        0, 'N'
    )

    # Interactive Control Bar (Row 6)
    ws_exec.set_row(5, 22)
    ws_exec.write('B6', "Analysis Year:", f_ctrl_label)
    ws_exec.write('C6', "All", f_ctrl_val)
    ws_exec.data_validation('C6', {'validate': 'list', 'source': ['All', '2023', '2024', '2025']})

    ws_exec.write('D6', "Category Focus:", f_ctrl_label)
    ws_exec.write('E6', "All", f_ctrl_val)
    ws_exec.data_validation('E6', {'validate': 'list', 'source': ['All', 'Electronics', 'Appliances', 'Fashion', 'Home & Kitchen', 'Footwear', 'Grocery', 'Furniture']})

    ws_exec.write('F6', "Order Status:", f_ctrl_label)
    ws_exec.write('G6', "All", f_ctrl_val)
    ws_exec.data_validation('G6', {'validate': 'list', 'source': ['All', 'Delivered', 'Cancelled', 'Returned', 'Failed', 'Processing']})

    ws_exec.merge_range('H6:N6', "Scope: 100,000 Transactions | 25,000 Customers | 550 Products | 15 States | 31 Cities", f_ctrl_info)

    # KPI Cards (Rows 8-10)
    # Card 1: Gross Revenue
    ws_exec.write('B8', "GROSS REVENUE", f_card_top)
    ws_exec.write_formula('B9', '=IF(C6="All",SUM(\'Supporting Data\'!$D$7:$D$42),SUMIFS(\'Supporting Data\'!$D$7:$D$42,\'Supporting Data\'!$B$7:$B$42,C6))', f_card_val_curr, 2449898974.35)
    ws_exec.write('B10', "₹2,449.90 Cr Placed", f_card_bot)

    # Card 2: Net Revenue
    ws_exec.write('C8', "NET REVENUE", f_card_top)
    ws_exec.write_formula('C9', '=B9*0.767', f_card_val_curr, 1879006920.35)
    ws_exec.write('C10', "₹1,879.01 Cr Delivered Net", f_card_bot)

    # Card 3: Total Orders
    ws_exec.write('D8', "TOTAL ORDERS", f_card_top)
    ws_exec.write_formula('D9', '=IF(C6="All",SUM(\'Supporting Data\'!$F$7:$F$42),SUMIFS(\'Supporting Data\'!$F$7:$F$42,\'Supporting Data\'!$B$7:$B$42,C6))', f_card_val_num, 100000)
    ws_exec.write('D10', "80,679 Delivered (80.7%)", f_card_bot)

    # Card 4: Customers
    ws_exec.write('E8', "TOTAL CUSTOMERS", f_card_top)
    ws_exec.write('E9', 25000, f_card_val_num)
    ws_exec.write('E10', "22,820 Active Purchasers", f_card_bot)

    # Card 5: Units Sold
    ws_exec.write('F8', "TOTAL UNITS SOLD", f_card_top)
    ws_exec.write_formula('F9', '=IF(C6="All",SUM(\'Supporting Data\'!$E$7:$E$42),SUMIFS(\'Supporting Data\'!$E$7:$E$42,\'Supporting Data\'!$B$7:$B$42,C6))', f_card_val_num, 585640)
    ws_exec.write('F10', "2.30 Units / Order", f_card_bot)

    # Card 6: AOV
    ws_exec.write('G8', "AVERAGE ORDER VALUE", f_card_top)
    ws_exec.write_formula('G9', '=B9/D9', f_card_val_curr, 24498.99)
    ws_exec.write('G10', "₹24,602.50 Delivered", f_card_bot)

    # Card 7: Gross Profit
    ws_exec.write('H8', "GROSS PROFIT", f_card_top)
    ws_exec.write_formula('H9', '=IF(C6="All",SUM(\'Supporting Data\'!$G$7:$G$42),SUMIFS(\'Supporting Data\'!$G$7:$G$42,\'Supporting Data\'!$B$7:$B$42,C6))', f_card_val_curr, 197294347.02)
    ws_exec.write('H10', "₹197.29 Cr Realized", f_card_bot)

    # Card 8: Gross Margin
    ws_exec.write('I8', "GROSS MARGIN %", f_card_top)
    ws_exec.write_formula('I9', '=H9/B9', f_card_val_pct, 0.08053)
    ws_exec.write('I10', "8.05% Gross Margin", f_card_bot)

    # Card 9: Return Rate
    ws_exec.write('J8', "RETURN RATE", f_card_top)
    ws_exec.write('J9', 0.06681, f_card_val_pct)
    ws_exec.write('J10', "6,681 Orders Returned", f_card_bot)

    # Card 10: Repeat Cust Rate
    ws_exec.write('K8', "REPEAT CUST RATE", f_card_top)
    ws_exec.write('K9', 0.8287, f_card_val_pct)
    ws_exec.write('K10', "18,912 Repeat Buyers", f_card_bot)

    # Card 11: Campaign ROI
    ws_exec.write('L8', "MARKETING BLENDED ROI", f_card_top)
    ws_exec.write('L9', 2.33, f_card_val_dec)
    ws_exec.write('L10', "₹436.3M Rev / ₹187.1M Spend", f_card_bot)

    # Card 12: Customer Rating
    ws_exec.write('M8', "AVG REVIEW RATING", f_card_top)
    ws_exec.write('M9', 3.69, f_card_val_dec)
    ws_exec.write('M10', "59.7% Positive Sentiment", f_card_bot)

    # Card 13: Delivery Days
    ws_exec.write('N8', "AVG DELIVERY DAYS", f_card_top)
    ws_exec.write('N9', 5.62, f_card_val_dec)
    ws_exec.write('N10', "58.2% On-Time Delivery", f_card_bot)

    ws_exec.set_row(7, 14)
    ws_exec.set_row(8, 26)
    ws_exec.set_row(9, 14)

    # Section 1: Executive Charts Header
    ws_exec.merge_range('B12:N12', "STRATEGIC EXECUTIVE VISUALIZATIONS", f_sec_hdr)
    ws_exec.set_row(11, 20)

    # CHARTS on Executive Dashboard:
    # Chart 1: 3-Year Monthly Revenue Trend (Line Chart)
    ch_exec_trend = wb.add_chart({'type': 'line'})
    ch_exec_trend.add_series({
        'name': 'Monthly Gross Revenue',
        'categories': "='Supporting Data'!$A$6:$A$41",
        'values': "='Supporting Data'!$D$6:$D$41",
        'line': {'color': NAVY, 'width': 2.25}
    })
    ch_exec_trend.set_title({'name': '3-Year Monthly Revenue Trajectory (2023 - 2025)', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch_exec_trend.set_x_axis({'name': 'Month', 'name_font': {'size': 8}, 'num_font': {'size': 7, 'rotation': -45}})
    ch_exec_trend.set_y_axis({'name': 'Revenue (INR)', 'name_font': {'size': 8}, 'major_gridlines': {'visible': True, 'line': {'color': BORDER_COLOR, 'dash_type': 'dash'}}})
    ch_exec_trend.set_legend({'position': 'top'})
    ch_exec_trend.set_size({'width': 580, 'height': 270})
    ws_exec.insert_chart('B13', ch_exec_trend)

    # Chart 2: Category Revenue Breakdown (Column Chart)
    ch_exec_cat = wb.add_chart({'type': 'column'})
    ch_exec_cat.add_series({
        'name': 'Revenue by Category',
        'categories': "='Supporting Data'!$I$6:$I$15",
        'values': "='Supporting Data'!$K$6:$K$15",
        'fill': {'color': ROYAL_BLUE}
    })
    ch_exec_cat.set_title({'name': 'Top 10 Product Categories by Revenue', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch_exec_cat.set_x_axis({'num_font': {'size': 7, 'rotation': -45}})
    ch_exec_cat.set_y_axis({'major_gridlines': {'visible': True, 'line': {'color': BORDER_COLOR, 'dash_type': 'dash'}}})
    ch_exec_cat.set_legend({'none': True})
    ch_exec_cat.set_size({'width': 440, 'height': 270})
    ws_exec.insert_chart('H13', ch_exec_cat)

    # Chart 3: Order Status Distribution (Doughnut)
    ch_exec_status = wb.add_chart({'type': 'doughnut'})
    ch_exec_status.add_series({
        'name': 'Order Status',
        'categories': "='Supporting Data'!$Q$6:$Q$10",
        'values': "='Supporting Data'!$R$6:$R$10",
        'points': [
            {'fill': {'color': NAVY}},
            {'fill': {'color': RED}},
            {'fill': {'color': '#E67E22'}},
            {'fill': {'color': '#95A5A6'}},
            {'fill': {'color': CYAN}}
        ]
    })
    ch_exec_status.set_title({'name': 'Order Status Distribution', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch_exec_status.set_legend({'position': 'right', 'font': {'size': 8}})
    ch_exec_status.set_size({'width': 330, 'height': 240})
    ws_exec.insert_chart('B28', ch_exec_status)

    # Chart 4: Payment Method Mix (Column)
    ch_exec_pay = wb.add_chart({'type': 'column'})
    ch_exec_pay.add_series({
        'name': 'Amount Paid',
        'categories': "='Supporting Data'!$U$6:$U$12",
        'values': "='Supporting Data'!$W$6:$W$12",
        'fill': {'color': CYAN}
    })
    ch_exec_pay.set_title({'name': 'Top Payment Methods by Value', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch_exec_pay.set_x_axis({'num_font': {'size': 7, 'rotation': -30}})
    ch_exec_pay.set_legend({'none': True})
    ch_exec_pay.set_size({'width': 360, 'height': 240})
    ws_exec.insert_chart('F28', ch_exec_pay)

    # Chart 5: Top 10 States by Sales (Bar)
    ch_exec_state = wb.add_chart({'type': 'bar'})
    ch_exec_state.add_series({
        'name': 'Sales by State',
        'categories': "='Supporting Data'!$AA$6:$AA$15",
        'values': "='Supporting Data'!$AC$6:$AC$15",
        'fill': {'color': NAVY}
    })
    ch_exec_state.set_title({'name': 'Top 10 States by Revenue', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch_exec_state.set_y_axis({'num_font': {'size': 7}})
    ch_exec_state.set_legend({'none': True})
    ch_exec_state.set_size({'width': 330, 'height': 240})
    ws_exec.insert_chart('K28', ch_exec_state)

    # Section 2: Executive Summary Table (Row 42+)
    ws_exec.merge_range('B42:N42', "EXECUTIVE ANNUAL & METRIC SCORECARD", f_sec_hdr)
    ws_exec.set_row(41, 20)

    exec_tbl_headers = [
        "Financial Year", "Orders Placed", "Delivered Orders", "Delivery Rate %",
        "Units Sold", "Gross Revenue (Subtotal)", "Total Tax & Shipping", "Total Final Amount",
        "Gross Profit", "Gross Margin %", "Order Return Rate %", "AOV"
    ]
    for c_idx, h in enumerate(exec_tbl_headers):
        ws_exec.write(42, 1 + c_idx, h, f_th)

    yearly_summary = [
        ("FY 2023", 26750, 21544, 21544/26750, 156820, 660951320.15, 33145210.40, 694096530.55, 53215840.10, 53215840.10/660951320.15, 0.0664, 660951320.15/26750),
        ("FY 2024", 26730, 21602, 21602/26730, 156410, 652127110.20, 32704190.35, 684831300.55, 52498720.45, 52498720.45/652127110.20, 0.0671, 652127110.20/26730),
        ("FY 2025", 46520, 37533, 37533/46520, 272410, 1136820544.00, 56843993.47, 1193664537.47, 91579786.47, 91579786.47/1136820544.00, 0.0669, 1136820544.00/46520)
    ]

    for r_idx, row in enumerate(yearly_summary):
        r_num = 43 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        fmt_n = f_td_num_alt if r_idx % 2 == 1 else f_td_num
        fmt_c = f_td_curr_alt if r_idx % 2 == 1 else f_td_curr
        fmt_p = f_td_pct_alt if r_idx % 2 == 1 else f_td_pct

        ws_exec.write(r_num, 1, row[0], fmt_t)
        ws_exec.write(r_num, 2, row[1], fmt_n)
        ws_exec.write(r_num, 3, row[2], fmt_n)
        ws_exec.write(r_num, 4, row[3], fmt_p)
        ws_exec.write(r_num, 5, row[4], fmt_n)
        ws_exec.write(r_num, 6, row[5], fmt_c)
        ws_exec.write(r_num, 7, row[6], fmt_c)
        ws_exec.write(r_num, 8, row[7], fmt_c)
        ws_exec.write(r_num, 9, row[8], fmt_c)
        ws_exec.write(r_num, 10, row[9], fmt_p)
        ws_exec.write(r_num, 11, row[10], fmt_p)
        ws_exec.write(r_num, 12, row[11], fmt_c)

    # Total Row
    tot_r = 46
    ws_exec.write(tot_r, 1, "Total / Enterprise", f_tot_label)
    ws_exec.write(tot_r, 2, "=SUM(C44:C46)", f_tot_num)
    ws_exec.write(tot_r, 3, "=SUM(D44:D46)", f_tot_num)
    ws_exec.write(tot_r, 4, "=D47/C47", f_tot_pct)
    ws_exec.write(tot_r, 5, "=SUM(F44:F46)", f_tot_num)
    ws_exec.write(tot_r, 6, "=SUM(G44:G46)", f_tot_curr)
    ws_exec.write(tot_r, 7, "=SUM(H44:H46)", f_tot_curr)
    ws_exec.write(tot_r, 8, "=SUM(I44:I46)", f_tot_curr)
    ws_exec.write(tot_r, 9, "=SUM(J44:J46)", f_tot_curr)
    ws_exec.write(tot_r, 10, "=J47/G47", f_tot_pct)
    ws_exec.write(tot_r, 11, 0.0668, f_tot_pct)
    ws_exec.write(tot_r, 12, "=G47/C47", f_tot_curr)

    # Strategic Notes
    ws_exec.merge_range('B49:N52',
        "KEY EXECUTIVE FINDINGS & STRATEGIC TAKEAWAYS:\n"
        "1. Rapid Growth in FY 2025: Revenue expanded +74.3% YoY in 2025, driven by Q4 festive sales (Diwali, Dussehra, Year-End) surpassing ₹526M in Q4 alone.\n"
        "2. Margin Structure: Enterprise gross margin is 8.05% (₹197.3M). Fashion yields the highest margin (33.1%), whereas Electronics generates 49.9% of total volume at thin margins (0.69%).\n"
        "3. High Customer Retention: 82.87% of purchasing customers are repeat buyers. Premium & Regular segments account for 78.4% of lifetime spend.\n"
        "4. Logistics Opportunity: 41.76% of shipments experienced delivery delays due to regional fulfillment bottlenecks, directly driving 8.0% of customer return reasons.",
        wb.add_format({'font_size': 8.5, 'font_name': 'Calibri', 'font_color': NAVY, 'bg_color': PALE_BLUE, 'border': 1, 'valign': 'top'})
    )

    # Freeze panes below nav
    ws_exec.freeze_panes(4, 0)

    # ==========================================
    # 2. SALES ANALYSIS
    # ==========================================
    print("Building Sales Analysis sheet...")
    ws_sales = wb.add_worksheet('Sales Analysis')
    ws_sales.set_tab_color(ROYAL_BLUE)
    ws_sales.hide_gridlines(0)
    ws_sales.set_column('A:A', 3)
    ws_sales.set_column('B:L', 15)

    render_header_nav(
        ws_sales,
        "INDIAN E-COMMERCE SALES REVENUE & TREND ANALYSIS",
        "36-Month Historical Revenue, Order Velocity, Seasonality, and Growth Decomposition",
        1, 'L'
    )

    # KPI Cards
    ws_sales.write('B7', "FY23 REVENUE", f_card_top)
    ws_sales.write('B8', 660951320.15, f_card_val_curr)
    ws_sales.write('B9', "26,750 Orders", f_card_bot)

    ws_sales.write('C7', "FY24 REVENUE", f_card_top)
    ws_sales.write('C8', 652127110.20, f_card_val_curr)
    ws_sales.write('C9', "26,730 Orders (-1.3%)", f_card_bot)

    ws_sales.write('D7', "FY25 REVENUE", f_card_top)
    ws_sales.write('D8', 1136820544.00, f_card_val_curr)
    ws_sales.write('D9', "46,520 Orders (+74.3%)", f_card_bot)

    ws_sales.write('E7', "3-YEAR CAGR", f_card_top)
    ws_sales.write('E8', 0.3115, f_card_val_pct)
    ws_sales.write('E9', "Compounded Annual Growth", f_card_bot)

    ws_sales.write('F7', "PEAK MONTH REV", f_card_top)
    ws_sales.write('F8', 235942310.50, f_card_val_curr)
    ws_sales.write('F9', "December 2025", f_card_bot)

    ws_sales.write('G7', "AVG DAILY SALES", f_card_top)
    ws_sales.write('G8', 2235309.28, f_card_val_curr)
    ws_sales.write('G9', "1,096 Total Days", f_card_bot)

    ws_sales.set_row(6, 14)
    ws_sales.set_row(7, 26)
    ws_sales.set_row(8, 14)

    # Sales Charts Header
    ws_sales.merge_range('B11:L11', "SALES TRENDS & SEASONALITY VISUALIZATIONS", f_sec_hdr)

    # Chart 1: 3-Year Monthly Revenue Progression
    ch_sales_trend = wb.add_chart({'type': 'line'})
    ch_sales_trend.add_series({
        'name': 'Monthly Revenue',
        'categories': "='Sales Analysis'!$B$26:$B$61",
        'values': "='Sales Analysis'!$E$26:$E$61",
        'line': {'color': NAVY, 'width': 2}
    })
    ch_sales_trend.add_series({
        'name': 'Monthly Profit',
        'categories': "='Sales Analysis'!$B$26:$B$61",
        'values': "='Sales Analysis'!$G$26:$G$61",
        'line': {'color': CYAN, 'width': 1.75}
    })
    ch_sales_trend.set_title({'name': '36-Month Revenue & Profit Trajectory', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch_sales_trend.set_x_axis({'num_font': {'size': 7, 'rotation': -45}})
    ch_sales_trend.set_legend({'position': 'top'})
    ch_sales_trend.set_size({'width': 560, 'height': 240})
    ws_sales.insert_chart('B12', ch_sales_trend)

    # Chart 2: Quarterly Revenue
    ch_quarter = wb.add_chart({'type': 'column'})
    ch_quarter.add_series({
        'name': 'Quarterly Revenue',
        'categories': "='Supporting Data'!$AE$6:$AE$17",
        'values': "='Supporting Data'!$AG$6:$AG$17",
        'fill': {'color': ROYAL_BLUE}
    })
    ch_quarter.set_title({'name': 'Quarterly Revenue Progression (Q1 2023 - Q4 2025)', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch_quarter.set_x_axis({'num_font': {'size': 7, 'rotation': -30}})
    ch_quarter.set_legend({'none': True})
    ch_quarter.set_size({'width': 440, 'height': 240})
    ws_sales.insert_chart('H12', ch_quarter)

    # 36-Month Sales Table Header
    ws_sales.merge_range('B24:L24', "MONTH-BY-MONTH SALES PERFORMANCE REGISTER (2023 - 2025)", f_sec_hdr)
    sales_tbl_headers = ["Year-Month", "Year", "Month", "Orders Count", "Units Sold", "Subtotal Revenue", "COGS Cost", "Gross Profit", "Margin %", "AOV", "MoM Growth %"]
    for c_idx, h in enumerate(sales_tbl_headers):
        ws_sales.write(24, 1 + c_idx, h, f_th)

    for r_idx, row in monthly_all.iterrows():
        r_num = 25 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        fmt_n = f_td_num_alt if r_idx % 2 == 1 else f_td_num
        fmt_c = f_td_curr_alt if r_idx % 2 == 1 else f_td_curr
        fmt_p = f_td_pct_alt if r_idx % 2 == 1 else f_td_pct

        ws_sales.write(r_num, 1, row['year_month'], fmt_t)
        ws_sales.write(r_num, 2, int(row['year']), fmt_t)
        ws_sales.write(r_num, 3, int(row['month']), fmt_t)
        ws_sales.write(r_num, 4, int(row['orders_count']), fmt_n)
        ws_sales.write(r_num, 5, int(row['units_sold']), fmt_n)
        ws_sales.write(r_num, 6, float(row['subtotal_rev']), fmt_c)
        ws_sales.write(r_num, 7, float(row['cost']), fmt_c)
        ws_sales.write(r_num, 8, float(row['profit']), fmt_c)
        ws_sales.write(r_num, 9, float(row['margin_pct']), fmt_p)
        ws_sales.write(r_num, 10, float(row['aov']), fmt_c)
        ws_sales.write(r_num, 11, float(row['mom_growth']), fmt_p)

    tot_s_r = 25 + len(monthly_all)
    ws_sales.write(tot_s_r, 1, "Total / Enterprise", f_tot_label)
    ws_sales.write(tot_s_r, 2, "-", f_tot_label)
    ws_sales.write(tot_s_r, 3, "-", f_tot_label)
    ws_sales.write(tot_s_r, 4, "=SUM(E26:E61)", f_tot_num)
    ws_sales.write(tot_s_r, 5, "=SUM(F26:F61)", f_tot_num)
    ws_sales.write(tot_s_r, 6, "=SUM(G26:G61)", f_tot_curr)
    ws_sales.write(tot_s_r, 7, "=SUM(H26:H61)", f_tot_curr)
    ws_sales.write(tot_s_r, 8, "=SUM(I26:I61)", f_tot_curr)
    ws_sales.write(tot_s_r, 9, "=I62/G62", f_tot_pct)
    ws_sales.write(tot_s_r, 10, "=G62/E62", f_tot_curr)
    ws_sales.write(tot_s_r, 11, "-", f_tot_label)

    ws_sales.freeze_panes(4, 0)

    # ==========================================
    # 3. PRODUCT ANALYSIS
    # ==========================================
    print("Building Product Analysis sheet...")
    ws_prod = wb.add_worksheet('Product Analysis')
    ws_prod.set_tab_color(NAVY)
    ws_prod.hide_gridlines(0)
    ws_prod.set_column('A:A', 3)
    ws_prod.set_column('B:L', 15)

    render_header_nav(
        ws_prod,
        "INDIAN E-COMMERCE PRODUCT & CATEGORY ANALYTICS",
        "Catalog Assortment, SKU Contribution, Gross Margins, and Product Return Rates",
        2, 'L'
    )

    # KPI Cards
    ws_prod.write('B7', "CATALOG SKUS", f_card_top)
    ws_prod.write('B8', 550, f_card_val_num)
    ws_prod.write('B9', "Across 22 Categories", f_card_bot)

    ws_prod.write('C7', "TOP CATEGORY", f_card_top)
    ws_prod.write('C8', 1223619210.45, f_card_val_curr)
    ws_prod.write('C9', "Electronics (49.9% Rev)", f_card_bot)

    ws_prod.write('D7', "HIGHEST MARGIN", f_card_top)
    ws_prod.write('D8', 0.3314, f_card_val_pct)
    ws_prod.write('D9', "Fashion (₹46.2M Profit)", f_card_bot)

    ws_prod.write('E7', "TOTAL UNITS SOLD", f_card_top)
    ws_prod.write('E8', 585640, f_card_val_num)
    ws_prod.write('E9', "1,065 Units / SKU", f_card_bot)

    ws_prod.write('F7', "AVG ITEM PRICE", f_card_top)
    ws_prod.write('F8', 4183.28, f_card_val_curr)
    ws_prod.write('F9', "Blended Selling Price", f_card_bot)

    ws_prod.write('G7', "TOTAL PROFIT", f_card_top)
    ws_prod.write('G8', 197294347.02, f_card_val_curr)
    ws_prod.write('G9', "₹358.7k Profit / SKU", f_card_bot)

    ws_prod.set_row(6, 14)
    ws_prod.set_row(7, 26)
    ws_prod.set_row(8, 14)

    # Product Charts Header
    ws_prod.merge_range('B11:L11', "PRODUCT & CATEGORY VISUALIZATIONS", f_sec_hdr)

    # Chart 1: Category Revenue vs Profit
    ch_prod_cat = wb.add_chart({'type': 'column'})
    ch_prod_cat.add_series({
        'name': 'Revenue',
        'categories': "='Product Analysis'!$B$26:$B$35",
        'values': "='Product Analysis'!$E$26:$E$35",
        'fill': {'color': NAVY}
    })
    ch_prod_cat.add_series({
        'name': 'Gross Profit',
        'categories': "='Product Analysis'!$B$26:$B$35",
        'values': "='Product Analysis'!$G$26:$G$35",
        'fill': {'color': CYAN}
    })
    ch_prod_cat.set_title({'name': 'Top 10 Categories: Revenue vs Gross Profit', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch_prod_cat.set_x_axis({'num_font': {'size': 7, 'rotation': -30}})
    ch_prod_cat.set_legend({'position': 'top'})
    ch_prod_cat.set_size({'width': 560, 'height': 240})
    ws_prod.insert_chart('B12', ch_prod_cat)

    # Chart 2: Category Margin Ranking
    ch_prod_margin = wb.add_chart({'type': 'bar'})
    ch_prod_margin.add_series({
        'name': 'Gross Margin %',
        'categories': "='Product Analysis'!$B$26:$B$35",
        'values': "='Product Analysis'!$H$26:$H$35",
        'fill': {'color': ROYAL_BLUE}
    })
    ch_prod_margin.set_title({'name': 'Top 10 Categories Gross Margin %', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch_prod_margin.set_legend({'none': True})
    ch_prod_margin.set_size({'width': 440, 'height': 240})
    ws_prod.insert_chart('H12', ch_prod_margin)

    # Category Matrix Table
    ws_prod.merge_range('B24:L24', "COMPLETE CATEGORY PERFORMANCE MATRIX (ALL 22 CATEGORIES)", f_sec_hdr)
    prod_tbl_headers = ["Category", "SKU Count", "Units Sold", "Total Revenue", "Rev Share %", "Total COGS", "Gross Profit", "Margin %", "Returns Count", "Refunds (INR)", "Return Rate %"]
    for c_idx, h in enumerate(prod_tbl_headers):
        ws_prod.write(24, 1 + c_idx, h, f_th)

    for r_idx, row in cat_all.iterrows():
        r_num = 25 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        fmt_n = f_td_num_alt if r_idx % 2 == 1 else f_td_num
        fmt_c = f_td_curr_alt if r_idx % 2 == 1 else f_td_curr
        fmt_p = f_td_pct_alt if r_idx % 2 == 1 else f_td_pct

        ws_prod.write(r_num, 1, row['category'], fmt_t)
        ws_prod.write(r_num, 2, int(row['sku_count']), fmt_n)
        ws_prod.write(r_num, 3, int(row['units_sold']), fmt_n)
        ws_prod.write(r_num, 4, float(row['revenue']), fmt_c)
        ws_prod.write(r_num, 5, float(row['rev_share']), fmt_p)
        ws_prod.write(r_num, 6, float(row['cost']), fmt_c)
        ws_prod.write(r_num, 7, float(row['profit']), fmt_c)
        ws_prod.write(r_num, 8, float(row['margin_pct']), fmt_p)
        ws_prod.write(r_num, 9, int(row['return_count']), fmt_n)
        ws_prod.write(r_num, 10, float(row['refund_amount']), fmt_c)
        ws_prod.write(r_num, 11, float(row['return_rate']), fmt_p)

    tot_p_r = 25 + len(cat_all)
    ws_prod.write(tot_p_r, 1, "Total / Enterprise", f_tot_label)
    ws_prod.write(tot_p_r, 2, "=SUM(C26:C47)", f_tot_num)
    ws_prod.write(tot_p_r, 3, "=SUM(D26:D47)", f_tot_num)
    ws_prod.write(tot_p_r, 4, "=SUM(E26:E47)", f_tot_curr)
    ws_prod.write(tot_p_r, 5, 1.0, f_tot_pct)
    ws_prod.write(tot_p_r, 6, "=SUM(G26:G47)", f_tot_curr)
    ws_prod.write(tot_p_r, 7, "=SUM(H26:H47)", f_tot_curr)
    ws_prod.write(tot_p_r, 8, "=H48/E48", f_tot_pct)
    ws_prod.write(tot_p_r, 9, "=SUM(J26:J47)", f_tot_num)
    ws_prod.write(tot_p_r, 10, "=SUM(K26:K47)", f_tot_curr)
    ws_prod.write(tot_p_r, 11, "=J48/D48", f_tot_pct)

    # Top 20 Best Sellers Table Header (Row 50+)
    ws_prod.merge_range('B50:K50', "TOP 20 BEST-SELLING PRODUCTS IN CATALOG", f_sec_hdr)
    top_prod_headers = ["Product ID", "Product Name", "Category", "Units Sold", "Selling Price", "COGS Cost", "Revenue", "Gross Profit", "Margin %", "Rank"]
    for c_idx, h in enumerate(top_prod_headers):
        ws_prod.write(50, 1 + c_idx, h, f_th)

    for r_idx, row in prod_top20.iterrows():
        r_num = 51 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        fmt_n = f_td_num_alt if r_idx % 2 == 1 else f_td_num
        fmt_c = f_td_curr_alt if r_idx % 2 == 1 else f_td_curr
        fmt_p = f_td_pct_alt if r_idx % 2 == 1 else f_td_pct

        ws_prod.write(r_num, 1, row['product_id'], fmt_t)
        ws_prod.write(r_num, 2, row['product_name'], fmt_t)
        ws_prod.write(r_num, 3, row['category'], fmt_t)
        ws_prod.write(r_num, 4, int(row['units_sold']), fmt_n)
        ws_prod.write(r_num, 5, float(row['avg_selling_price']), fmt_c)
        ws_prod.write(r_num, 6, float(row['cost'] / row['units_sold']), fmt_c)
        ws_prod.write(r_num, 7, float(row['revenue']), fmt_c)
        ws_prod.write(r_num, 8, float(row['profit']), fmt_c)
        ws_prod.write(r_num, 9, float(row['margin_pct']), fmt_p)
        ws_prod.write(r_num, 10, r_idx + 1, fmt_n)

    ws_prod.freeze_panes(4, 0)

    # ==========================================
    # 4. CUSTOMER ANALYSIS
    # ==========================================
    print("Building Customer Analysis sheet...")
    ws_cust = wb.add_worksheet('Customer Analysis')
    ws_cust.set_tab_color(ROYAL_BLUE)
    ws_cust.hide_gridlines(0)
    ws_cust.set_column('A:A', 3)
    ws_cust.set_column('B:K', 16)

    render_header_nav(
        ws_cust,
        "INDIAN E-COMMERCE CUSTOMER DEMOGRAPHICS & LIFECYCLE ANALYTICS",
        "Customer Segmentation, Churn Analysis, RFM Behavioral Profiles, and Loyalty Tiers",
        3, 'K'
    )

    # KPI Cards
    ws_cust.write('B7', "REGISTERED USERS", f_card_top)
    ws_cust.write('B8', 25000, f_card_val_num)
    ws_cust.write('B9', "Pan-India Accounts", f_card_bot)

    ws_cust.write('C7', "ACTIVE PURCHASERS", f_card_top)
    ws_cust.write('C8', 22820, f_card_val_num)
    ws_cust.write('C9', "91.3% Conversion Rate", f_card_bot)

    ws_cust.write('D7', "REPEAT PURCHASERS", f_card_top)
    ws_cust.write('D8', 0.8287, f_card_val_pct)
    ws_cust.write('D9', "18,912 Multiple Orders", f_card_bot)

    ws_cust.write('E7', "CHURNED CUSTOMERS", f_card_top)
    ws_cust.write('E8', 13365, f_card_val_num)
    ws_cust.write('E9', "53.5% Churn Rate (>180d)", f_card_bot)

    ws_cust.write('F7', "AT RISK CUSTOMERS", f_card_top)
    ws_cust.write('F8', 2101, f_card_val_num)
    ws_cust.write('F9', "8.4% Winback Priority", f_card_bot)

    ws_cust.write('G7', "AVG ANNUAL SPEND", f_card_top)
    ws_cust.write('G8', 107357.68, f_card_val_curr)
    ws_cust.write('G9', "Per Purchasing User", f_card_bot)

    ws_cust.set_row(6, 14)
    ws_cust.set_row(7, 26)
    ws_cust.set_row(8, 14)

    # Customer Charts Header
    ws_cust.merge_range('B11:K11', "CUSTOMER SEGMENTATION & BEHAVIORAL CHARTS", f_sec_hdr)

    # Chart 1: Customer Segmentation
    ch_cust_seg = wb.add_chart({'type': 'doughnut'})
    ch_cust_seg.add_series({
        'name': 'Customer Segments',
        'categories': "='Customer Analysis'!$B$26:$B$29",
        'values': "='Customer Analysis'!$C$26:$C$29",
        'points': [
            {'fill': {'color': NAVY}},
            {'fill': {'color': ROYAL_BLUE}},
            {'fill': {'color': CYAN}},
            {'fill': {'color': '#95A5A6'}}
        ]
    })
    ch_cust_seg.set_title({'name': 'Customer Distribution by Segment', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch_cust_seg.set_legend({'position': 'right'})
    ch_cust_seg.set_size({'width': 440, 'height': 240})
    ws_cust.insert_chart('B12', ch_cust_seg)

    # Chart 2: Customer Status / Lifecycle
    ch_cust_stat = wb.add_chart({'type': 'column'})
    ch_cust_stat.add_series({
        'name': 'Customer Count',
        'categories': "='Customer Analysis'!$B$33:$B$36",
        'values': "='Customer Analysis'!$C$33:$C$36",
        'fill': {'color': ROYAL_BLUE}
    })
    ch_cust_stat.set_title({'name': 'Customer Lifecycle & Churn Distribution', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch_cust_stat.set_x_axis({'num_font': {'size': 8}})
    ch_cust_stat.set_legend({'none': True})
    ch_cust_stat.set_size({'width': 480, 'height': 240})
    ws_cust.insert_chart('G12', ch_cust_stat)

    # Table 1: Customer Segment Profiles
    ws_cust.merge_range('B24:H24', "CUSTOMER SEGMENT VALUE MATRIX", f_sec_hdr)
    seg_headers = ["Segment", "Customer Count", "Total Orders", "Total Spend (INR)", "Spend Share %", "Avg AOV (INR)", "Orders / Cust"]
    for c_idx, h in enumerate(seg_headers):
        ws_cust.write(24, 1 + c_idx, h, f_th)

    for r_idx, row in cust_seg.iterrows():
        r_num = 25 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        fmt_n = f_td_num_alt if r_idx % 2 == 1 else f_td_num
        fmt_c = f_td_curr_alt if r_idx % 2 == 1 else f_td_curr
        fmt_p = f_td_pct_alt if r_idx % 2 == 1 else f_td_pct
        fmt_d = f_td_dec_alt if r_idx % 2 == 1 else f_td_dec

        ws_cust.write(r_num, 1, row['customer_segment'], fmt_t)
        ws_cust.write(r_num, 2, int(row['customer_count']), fmt_n)
        ws_cust.write(r_num, 3, int(row['total_orders']), fmt_n)
        ws_cust.write(r_num, 4, float(row['total_spend']), fmt_c)
        ws_cust.write(r_num, 5, float(row['spend_share']), fmt_p)
        ws_cust.write(r_num, 6, float(row['avg_aov']), fmt_c)
        ws_cust.write(r_num, 7, float(row['total_orders'] / row['customer_count']), fmt_d)

    # Table 2: Customer Lifecycle Status
    ws_cust.merge_range('B31:E31', "CUSTOMER LIFECYCLE & CHURN STATUS", f_sec_hdr)
    stat_headers = ["Status", "Customer Count", "Total Spend (INR)", "Customer Share %"]
    for c_idx, h in enumerate(stat_headers):
        ws_cust.write(31, 1 + c_idx, h, f_th)

    for r_idx, row in cust_stat.iterrows():
        r_num = 32 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        fmt_n = f_td_num_alt if r_idx % 2 == 1 else f_td_num
        fmt_c = f_td_curr_alt if r_idx % 2 == 1 else f_td_curr
        fmt_p = f_td_pct_alt if r_idx % 2 == 1 else f_td_pct

        ws_cust.write(r_num, 1, row['customer_status'], fmt_t)
        ws_cust.write(r_num, 2, int(row['customer_count']), fmt_n)
        ws_cust.write(r_num, 3, float(row['total_spend']), fmt_c)
        ws_cust.write(r_num, 4, float(row['cust_share']), fmt_p)

    # Table 3: Loyalty Tiers
    ws_cust.merge_range('G31:J31', "LOYALTY TIER BREAKDOWN", f_sec_hdr)
    tier_headers = ["Tier", "Customer Count", "Total Spend (INR)", "Spend Share %"]
    for c_idx, h in enumerate(tier_headers):
        ws_cust.write(31, 6 + c_idx, h, f_th)

    for r_idx, row in cust_tier.iterrows():
        r_num = 32 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        fmt_n = f_td_num_alt if r_idx % 2 == 1 else f_td_num
        fmt_c = f_td_curr_alt if r_idx % 2 == 1 else f_td_curr
        fmt_p = f_td_pct_alt if r_idx % 2 == 1 else f_td_pct

        ws_cust.write(r_num, 6, row['loyalty_tier'], fmt_t)
        ws_cust.write(r_num, 7, int(row['customer_count']), fmt_n)
        ws_cust.write(r_num, 8, float(row['total_spend']), fmt_c)
        ws_cust.write(r_num, 9, float(row['spend_share']), fmt_p)

    ws_cust.freeze_panes(4, 0)

    # ==========================================
    # 5. GEOGRAPHIC ANALYSIS
    # ==========================================
    print("Building Geographic Analysis sheet...")
    ws_geo = wb.add_worksheet('Geographic Analysis')
    ws_geo.set_tab_color(NAVY)
    ws_geo.hide_gridlines(0)
    ws_geo.set_column('A:A', 3)
    ws_geo.set_column('B:L', 15)

    render_header_nav(
        ws_geo,
        "INDIAN E-COMMERCE PAN-INDIA GEOGRAPHIC SALES INTELLIGENCE",
        "State-Level Distribution, Tier-1 vs Tier-2 Metro Penetration, Logistics Speed, and Regional Return Rates",
        4, 'L'
    )

    # KPI Cards
    ws_geo.write('B7', "STATES COVERED", f_card_top)
    ws_geo.write('B8', 15, f_card_val_num)
    ws_geo.write('B9', "Pan-India Footprint", f_card_bot)

    ws_geo.write('C7', "ACTIVE CITIES", f_card_top)
    ws_geo.write('C8', 31, f_card_val_num)
    ws_geo.write('C9', "Major Metros & Hubs", f_card_bot)

    ws_geo.write('D7', "TOP STATE (REV)", f_card_top)
    ws_geo.write('D8', 342721950.80, f_card_val_curr)
    ws_geo.write('D9', "Maharashtra (14.0% Share)", f_card_bot)

    ws_geo.write('E7', "TOP CITY (REV)", f_card_top)
    ws_geo.write('E8', 117364280.20, f_card_val_curr)
    ws_geo.write('E9', "Mumbai (4.8% Share)", f_card_bot)

    ws_geo.write('F7', "TOP 3 STATES SHARE", f_card_top)
    ws_geo.write('F8', 0.3545, f_card_val_pct)
    ws_geo.write('F9', "MH, UP, GJ (35.5% Rev)", f_card_bot)

    ws_geo.write('G7', "REGIONAL DELAY RATE", f_card_top)
    ws_geo.write('G8', 0.4176, f_card_val_pct)
    ws_geo.write('G9', "37,448 Late Shipments", f_card_bot)

    ws_geo.set_row(6, 14)
    ws_geo.set_row(7, 26)
    ws_geo.set_row(8, 14)

    # Charts Header
    ws_geo.merge_range('B11:L11', "GEOGRAPHIC DISTRIBUTION & METRO RANKINGS", f_sec_hdr)

    # Chart 1: State Revenue Bar
    ch_geo_state = wb.add_chart({'type': 'bar'})
    ch_geo_state.add_series({
        'name': 'Revenue',
        'categories': "='Geographic Analysis'!$B$26:$B$35",
        'values': "='Geographic Analysis'!$D$26:$D$35",
        'fill': {'color': NAVY}
    })
    ch_geo_state.set_title({'name': 'Top 10 Indian States by E-Commerce Revenue', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch_geo_state.set_legend({'none': True})
    ch_geo_state.set_size({'width': 500, 'height': 240})
    ws_geo.insert_chart('B12', ch_geo_state)

    # Chart 2: Top 10 Cities Revenue
    ch_geo_city = wb.add_chart({'type': 'column'})
    ch_geo_city.add_series({
        'name': 'Revenue',
        'categories': "='Geographic Analysis'!$B$45:$B$54",
        'values': "='Geographic Analysis'!$D$45:$D$54",
        'fill': {'color': ROYAL_BLUE}
    })
    ch_geo_city.set_title({'name': 'Top 10 Cities by Revenue', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch_geo_city.set_x_axis({'num_font': {'size': 7, 'rotation': -30}})
    ch_geo_city.set_legend({'none': True})
    ch_geo_city.set_size({'width': 500, 'height': 240})
    ws_geo.insert_chart('H12', ch_geo_city)

    # State Performance Table
    ws_geo.merge_range('B24:H24', "STATE-LEVEL PERFORMANCE REGISTER (ALL 15 STATES)", f_sec_hdr)
    state_headers = ["State", "Orders Count", "Gross Revenue (INR)", "Revenue Share %", "Delivered Orders", "Returned Orders", "Return Rate %"]
    for c_idx, h in enumerate(state_headers):
        ws_geo.write(24, 1 + c_idx, h, f_th)

    for r_idx, row in state_all.iterrows():
        r_num = 25 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        fmt_n = f_td_num_alt if r_idx % 2 == 1 else f_td_num
        fmt_c = f_td_curr_alt if r_idx % 2 == 1 else f_td_curr
        fmt_p = f_td_pct_alt if r_idx % 2 == 1 else f_td_pct

        ws_geo.write(r_num, 1, row['delivery_state'], fmt_t)
        ws_geo.write(r_num, 2, int(row['orders_count']), fmt_n)
        ws_geo.write(r_num, 3, float(row['revenue']), fmt_c)
        ws_geo.write(r_num, 4, float(row['rev_share']), fmt_p)
        ws_geo.write(r_num, 5, int(row['delivered_count']), fmt_n)
        ws_geo.write(r_num, 6, int(row['returned_count']), fmt_n)
        ws_geo.write(r_num, 7, float(row['return_rate']), fmt_p)

    # Top 20 Cities Table
    ws_geo.merge_range('B43:G43', "TOP 20 METRO & TIER-2 CITIES PERFORMANCE", f_sec_hdr)
    city_headers = ["City", "State", "Orders Count", "Gross Revenue (INR)", "Revenue Share %", "Delivered Orders"]
    for c_idx, h in enumerate(city_headers):
        ws_geo.write(43, 1 + c_idx, h, f_th)

    for r_idx, row in city_all.head(20).iterrows():
        r_num = 44 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        fmt_n = f_td_num_alt if r_idx % 2 == 1 else f_td_num
        fmt_c = f_td_curr_alt if r_idx % 2 == 1 else f_td_curr
        fmt_p = f_td_pct_alt if r_idx % 2 == 1 else f_td_pct

        ws_geo.write(r_num, 1, row['delivery_city'], fmt_t)
        ws_geo.write(r_num, 2, row['delivery_state'], fmt_t)
        ws_geo.write(r_num, 3, int(row['orders_count']), fmt_n)
        ws_geo.write(r_num, 4, float(row['revenue']), fmt_c)
        ws_geo.write(r_num, 5, float(row['rev_share']), fmt_p)
        ws_geo.write(r_num, 6, int(row['delivered_count']), fmt_n)

    ws_geo.freeze_panes(4, 0)

    # ==========================================
    # 6. PAYMENT ANALYSIS
    # ==========================================
    print("Building Payment Analysis sheet...")
    ws_pay = wb.add_worksheet('Payment Analysis')
    ws_pay.set_tab_color(ROYAL_BLUE)
    ws_pay.hide_gridlines(0)
    ws_pay.set_column('A:A', 3)
    ws_pay.set_column('B:K', 15)

    render_header_nav(
        ws_pay,
        "INDIAN E-COMMERCE PAYMENT GATEWAY & TRANSACTION ANALYTICS",
        "Payment Instrument Breakdown, UPI Dominance, Cash on Delivery Share, and Refund Reconciliations",
        5, 'K'
    )

    # KPI Cards
    ws_pay.write('B7', "UPI VOLUME", f_card_top)
    ws_pay.write('B8', 31971, f_card_val_num)
    ws_pay.write('B9', "32.0% #1 Instrument", f_card_bot)

    ws_pay.write('C7', "COD SHARE", f_card_top)
    ws_pay.write('C8', 0.1727, f_card_val_pct)
    ws_pay.write('C9', "17,274 COD Orders", f_card_bot)

    ws_pay.write('D7', "DIGITAL PAY SHARE", f_card_top)
    ws_pay.write('D8', 0.8273, f_card_val_pct)
    ws_pay.write('D9', "82.7% Prepaid / Digital", f_card_bot)

    ws_pay.write('E7', "PAYMENT SUCCESS %", f_card_top)
    ws_pay.write('E8', 0.8214, f_card_val_pct)
    ws_pay.write('E9', "82,137 Success Orders", f_card_bot)

    ws_pay.write('F7', "TOTAL REFUNDS", f_card_top)
    ws_pay.write('F8', 268740811.64, f_card_val_curr)
    ws_pay.write('F9', "10,690 Refund Records", f_card_bot)

    ws_pay.write('G7', "TOTAL AMOUNT PAID", f_card_top)
    ws_pay.write('G8', 2391049619.92, f_card_val_curr)
    ws_pay.write('G9', "Gross Inflows Collected", f_card_bot)

    ws_pay.set_row(6, 14)
    ws_pay.set_row(7, 26)
    ws_pay.set_row(8, 14)

    # Charts Header
    ws_pay.merge_range('B11:K11', "PAYMENT INSTRUMENTS & SETTLEMENT VISUALIZATIONS", f_sec_hdr)

    # Chart 1: Payment Method Value
    ch_pay_meth = wb.add_chart({'type': 'column'})
    ch_pay_meth.add_series({
        'name': 'Amount Paid',
        'categories': "='Payment Analysis'!$B$26:$B$35",
        'values': "='Payment Analysis'!$D$26:$D$35",
        'fill': {'color': NAVY}
    })
    ch_pay_meth.set_title({'name': 'Payment Methods by Value (INR)', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch_pay_meth.set_x_axis({'num_font': {'size': 7, 'rotation': -30}})
    ch_pay_meth.set_legend({'none': True})
    ch_pay_meth.set_size({'width': 500, 'height': 240})
    ws_pay.insert_chart('B12', ch_pay_meth)

    # Chart 2: Payment Status Outcomes
    ch_pay_stat = wb.add_chart({'type': 'doughnut'})
    ch_pay_stat.add_series({
        'name': 'Status Count',
        'categories': "='Payment Analysis'!$B$40:$B$44",
        'values': "='Payment Analysis'!$C$40:$C$44",
        'points': [
            {'fill': {'color': NAVY}},
            {'fill': {'color': '#E67E22'}},
            {'fill': {'color': RED}},
            {'fill': {'color': '#95A5A6'}},
            {'fill': {'color': CYAN}}
        ]
    })
    ch_pay_stat.set_title({'name': 'Payment Status Outcome Breakdown', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch_pay_stat.set_legend({'position': 'right'})
    ch_pay_stat.set_size({'width': 440, 'height': 240})
    ws_pay.insert_chart('G12', ch_pay_stat)

    # Table 1: Payment Methods Matrix
    ws_pay.merge_range('B24:I24', "PAYMENT INSTRUMENT COMPREHENSIVE SCORECARD", f_sec_hdr)
    pay_headers = ["Payment Method", "Transactions", "Amount Paid (INR)", "Value Share %", "Success Count", "Success Rate %", "Refunds (INR)", "Fee (INR)"]
    for c_idx, h in enumerate(pay_headers):
        ws_pay.write(24, 1 + c_idx, h, f_th)

    for r_idx, row in pay_meth.iterrows():
        r_num = 25 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        fmt_n = f_td_num_alt if r_idx % 2 == 1 else f_td_num
        fmt_c = f_td_curr_alt if r_idx % 2 == 1 else f_td_curr
        fmt_p = f_td_pct_alt if r_idx % 2 == 1 else f_td_pct

        ws_pay.write(r_num, 1, row['payment_method'], fmt_t)
        ws_pay.write(r_num, 2, int(row['transactions']), fmt_n)
        ws_pay.write(r_num, 3, float(row['amount_paid']), fmt_c)
        ws_pay.write(r_num, 4, float(row['share_pct']), fmt_p)
        ws_pay.write(r_num, 5, int(row['success_count']), fmt_n)
        ws_pay.write(r_num, 6, float(row['success_rate']), fmt_p)
        ws_pay.write(r_num, 7, float(row['refund_amount']), fmt_c)
        ws_pay.write(r_num, 8, float(row['fee_est']), fmt_c)

    # Table 2: Payment Status Summary
    ws_pay.merge_range('B38:E38', "PAYMENT STATUS BREAKDOWN", f_sec_hdr)
    pstat_headers = ["Payment Status", "Transaction Count", "Total Amount (INR)", "Share %"]
    for c_idx, h in enumerate(pstat_headers):
        ws_pay.write(38, 1 + c_idx, h, f_th)

    for r_idx, row in pay_stat.iterrows():
        r_num = 39 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        fmt_n = f_td_num_alt if r_idx % 2 == 1 else f_td_num
        fmt_c = f_td_curr_alt if r_idx % 2 == 1 else f_td_curr
        fmt_p = f_td_pct_alt if r_idx % 2 == 1 else f_td_pct

        ws_pay.write(r_num, 1, row['payment_status'], fmt_t)
        ws_pay.write(r_num, 2, int(row['count']), fmt_n)
        ws_pay.write(r_num, 3, float(row['amount']), fmt_c)
        ws_pay.write(r_num, 4, float(row['count_share']), fmt_p)

    ws_pay.freeze_panes(4, 0)

    # ==========================================
    # 7. ORDERS & RETURNS
    # ==========================================
    print("Building Orders & Returns sheet...")
    ws_ret = wb.add_worksheet('Orders & Returns')
    ws_ret.set_tab_color(NAVY)
    ws_ret.hide_gridlines(0)
    ws_ret.set_column('A:A', 3)
    ws_ret.set_column('B:K', 15)

    render_header_nav(
        ws_ret,
        "INDIAN E-COMMERCE ORDER FULFILLMENT, SHIPMENTS & RETURNS ANALYSIS",
        "Logistics SLA Performance, Delivery Delay Breakdown, Return Reasons Root-Cause, and Reverse Logistics Cost",
        6, 'K'
    )

    # KPI Cards
    ws_ret.write('B7', "TOTAL ORDERS", f_card_top)
    ws_ret.write('B8', 100000, f_card_val_num)
    ws_ret.write('B9', "Gross Orders Placed", f_card_bot)

    ws_ret.write('C7', "DELIVERED ORDERS", f_card_top)
    ws_ret.write('C8', 80679, f_card_val_num)
    ws_ret.write('C9', "80.7% Fulfill Rate", f_card_bot)

    ws_ret.write('D7', "ORDER RETURN RATE", f_card_top)
    ws_ret.write('D8', 0.0668, f_card_val_pct)
    ws_ret.write('D9', "6,681 Orders Returned", f_card_bot)

    ws_ret.write('E7', "ITEM RETURNS", f_card_top)
    ws_ret.write('E8', 12075, f_card_val_num)
    ws_ret.write('E9', "4.75% of Total Items", f_card_bot)

    ws_ret.write('F7', "RETURN REFUNDS", f_card_top)
    ws_ret.write('F8', 105898381.80, f_card_val_curr)
    ws_ret.write('F9', "₹10.59 Cr Returned", f_card_bot)

    ws_ret.write('G7', "ON-TIME DELIVERY %", f_card_top)
    ws_ret.write('G8', 0.5824, f_card_val_pct)
    ws_ret.write('G9', "52,233 On-Time Orders", f_card_bot)

    ws_ret.set_row(6, 14)
    ws_ret.set_row(7, 26)
    ws_ret.set_row(8, 14)

    # Charts Header
    ws_ret.merge_range('B11:K11', "REVERSE LOGISTICS & DELIVERY PERFORMANCE VISUALIZATIONS", f_sec_hdr)

    # Chart 1: Return Reasons Breakdown
    ch_ret_reas = wb.add_chart({'type': 'bar'})
    ch_ret_reas.add_series({
        'name': 'Returns Count',
        'categories': "='Orders & Returns'!$B$26:$B$33",
        'values': "='Orders & Returns'!$C$26:$C$33",
        'fill': {'color': RED}
    })
    ch_ret_reas.set_title({'name': 'Top Reasons for Product Returns', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch_ret_reas.set_legend({'none': True})
    ch_ret_reas.set_size({'width': 500, 'height': 240})
    ws_ret.insert_chart('B12', ch_ret_reas)

    # Chart 2: Shipping Method Volume & Delays
    ch_ship_delay = wb.add_chart({'type': 'column'})
    ch_ship_delay.add_series({
        'name': 'Total Shipments',
        'categories': "='Orders & Returns'!$B$37:$B$40",
        'values': "='Orders & Returns'!$C$37:$C$40",
        'fill': {'color': NAVY}
    })
    ch_ship_delay.add_series({
        'name': 'Delayed Shipments',
        'categories': "='Orders & Returns'!$B$37:$B$40",
        'values': "='Orders & Returns'!$E$37:$E$40",
        'fill': {'color': '#E67E22'}
    })
    ch_ship_delay.set_title({'name': 'Shipments vs Delays by Shipping Method', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch_ship_delay.set_legend({'position': 'top'})
    ch_ship_delay.set_size({'width': 440, 'height': 240})
    ws_ret.insert_chart('G12', ch_ship_delay)

    # Table 1: Return Reasons
    ws_ret.merge_range('B24:F24', "RETURN REASONS & REFUND EXPOSURE", f_sec_hdr)
    ret_headers = ["Return Reason", "Return Items Count", "Refund Amount (INR)", "Share of Returns %", "Avg Refund / Item"]
    for c_idx, h in enumerate(ret_headers):
        ws_ret.write(24, 1 + c_idx, h, f_th)

    for r_idx, row in ret_reason.iterrows():
        r_num = 25 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        fmt_n = f_td_num_alt if r_idx % 2 == 1 else f_td_num
        fmt_c = f_td_curr_alt if r_idx % 2 == 1 else f_td_curr
        fmt_p = f_td_pct_alt if r_idx % 2 == 1 else f_td_pct

        ws_ret.write(r_num, 1, row['return_reason'], fmt_t)
        ws_ret.write(r_num, 2, int(row['return_count']), fmt_n)
        ws_ret.write(r_num, 3, float(row['refund_amount']), fmt_c)
        ws_ret.write(r_num, 4, float(row['share_pct']), fmt_p)
        ws_ret.write(r_num, 5, float(row['refund_amount'] / row['return_count']), fmt_c)

    # Table 2: Shipping Method Performance
    ws_ret.merge_range('B35:G35', "SHIPPING METHOD & LOGISTICS SLA MATRIX", f_sec_hdr)
    ship_headers = ["Shipping Method", "Shipments", "Delivered Orders", "Delayed Count", "Delay Rate %", "Avg Delivery Days"]
    for c_idx, h in enumerate(ship_headers):
        ws_ret.write(35, 1 + c_idx, h, f_th)

    for r_idx, row in ship_meth.iterrows():
        r_num = 36 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        fmt_n = f_td_num_alt if r_idx % 2 == 1 else f_td_num
        fmt_p = f_td_pct_alt if r_idx % 2 == 1 else f_td_pct
        fmt_d = f_td_dec_alt if r_idx % 2 == 1 else f_td_dec

        ws_ret.write(r_num, 1, row['shipping_method'], fmt_t)
        ws_ret.write(r_num, 2, int(row['shipment_count']), fmt_n)
        ws_ret.write(r_num, 3, int(row['delivered_count']), fmt_n)
        ws_ret.write(r_num, 4, int(row['delayed_count']), fmt_n)
        ws_ret.write(r_num, 5, float(row['delay_rate']), fmt_p)
        ws_ret.write(r_num, 6, float(row['avg_delivery_days']), fmt_d)

    ws_ret.freeze_panes(4, 0)

    # ==========================================
    # 8. MARKETING ANALYSIS
    # ==========================================
    print("Building Marketing Analysis sheet...")
    ws_mkt = wb.add_worksheet('Marketing Analysis')
    ws_mkt.set_tab_color(ROYAL_BLUE)
    ws_mkt.hide_gridlines(0)
    ws_mkt.set_column('A:A', 3)
    ws_mkt.set_column('B:L', 15)

    render_header_nav(
        ws_mkt,
        "INDIAN E-COMMERCE MARKETING CAMPAIGN & ATTRIBUTION ANALYTICS",
        "Multi-Channel ROI, Ad Spend Efficiency, Customer Acquisition Cost (CAC/CPA), and Conversion Funnel",
        7, 'L'
    )

    # KPI Cards
    ws_mkt.write('B7', "CAMPAIGNS", f_card_top)
    ws_mkt.write('B8', 151, f_card_val_num)
    ws_mkt.write('B9', "Across 10 Channels", f_card_bot)

    ws_mkt.write('C7', "TOTAL AD SPEND", f_card_top)
    ws_mkt.write('C8', 187120911.26, f_card_val_curr)
    ws_mkt.write('C9', "₹18.71 Cr Invested", f_card_bot)

    ws_mkt.write('D7', "ATTRIBUTED REVENUE", f_card_top)
    ws_mkt.write('D8', 436296290.91, f_card_val_curr)
    ws_mkt.write('D9', "₹43.63 Cr Generated", f_card_bot)

    ws_mkt.write('E7', "BLENDED ROI", f_card_top)
    ws_mkt.write('E8', 2.33, f_card_val_dec)
    ws_mkt.write('E9', "₹2.33 Rev / ₹1 Spend", f_card_bot)

    ws_mkt.write('F7', "TOTAL CONVERSIONS", f_card_top)
    ws_mkt.write('F8', 17179, f_card_val_num)
    ws_mkt.write('F9', "Acquired Orders", f_card_bot)

    ws_mkt.write('G7', "BLENDED CPA / CAC", f_card_top)
    ws_mkt.write('G8', 10892.42, f_card_val_curr)
    ws_mkt.write('G9', "Cost Per Acquired Order", f_card_bot)

    ws_mkt.set_row(6, 14)
    ws_mkt.set_row(7, 26)
    ws_mkt.set_row(8, 14)

    # Charts Header
    ws_mkt.merge_range('B11:L11', "MARKETING CHANNEL EFFICIENCY & FUNNEL VISUALIZATIONS", f_sec_hdr)

    # Chart 1: ROI by Channel
    ch_mkt_roi = wb.add_chart({'type': 'bar'})
    ch_mkt_roi.add_series({
        'name': 'ROI (Multiple)',
        'categories': "='Marketing Analysis'!$B$26:$B$35",
        'values': "='Marketing Analysis'!$I$26:$I$35",
        'fill': {'color': NAVY}
    })
    ch_mkt_roi.set_title({'name': 'Marketing Campaign ROI by Channel (Revenue / Spend)', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch_mkt_roi.set_legend({'none': True})
    ch_mkt_roi.set_size({'width': 500, 'height': 240})
    ws_mkt.insert_chart('B12', ch_mkt_roi)

    # Chart 2: Spend vs Revenue by Channel
    ch_mkt_spend = wb.add_chart({'type': 'column'})
    ch_mkt_spend.add_series({
        'name': 'Ad Spend (INR)',
        'categories': "='Marketing Analysis'!$B$26:$B$35",
        'values': "='Marketing Analysis'!$D$26:$D$35",
        'fill': {'color': ROYAL_BLUE}
    })
    ch_mkt_spend.add_series({
        'name': 'Attributed Revenue (INR)',
        'categories': "='Marketing Analysis'!$B$26:$B$35",
        'values': "='Marketing Analysis'!$H$26:$H$35",
        'fill': {'color': CYAN}
    })
    ch_mkt_spend.set_title({'name': 'Ad Spend vs Attributed Revenue by Channel', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch_mkt_spend.set_x_axis({'num_font': {'size': 7, 'rotation': -30}})
    ch_mkt_spend.set_legend({'position': 'top'})
    ch_mkt_spend.set_size({'width': 500, 'height': 240})
    ws_mkt.insert_chart('H12', ch_mkt_spend)

    # Table 1: Channel Performance Matrix
    ws_mkt.merge_range('B24:L24', "MARKETING CHANNEL ATTRIBUTION SCORECARD (ALL 10 CHANNELS)", f_sec_hdr)
    mkt_headers = ["Channel", "Campaigns", "Ad Spend (INR)", "Impressions", "Clicks", "Conversions", "Attributed Rev (INR)", "ROI (x)", "CPA (INR)", "CVR %", "CTR %"]
    for c_idx, h in enumerate(mkt_headers):
        ws_mkt.write(24, 1 + c_idx, h, f_th)

    for r_idx, row in mkt_chan.iterrows():
        r_num = 25 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        fmt_n = f_td_num_alt if r_idx % 2 == 1 else f_td_num
        fmt_c = f_td_curr_alt if r_idx % 2 == 1 else f_td_curr
        fmt_p = f_td_pct_alt if r_idx % 2 == 1 else f_td_pct
        fmt_d = f_td_dec_alt if r_idx % 2 == 1 else f_td_dec

        ws_mkt.write(r_num, 1, row['channel'], fmt_t)
        ws_mkt.write(r_num, 2, int(row['campaigns_count']), fmt_n)
        ws_mkt.write(r_num, 3, float(row['ad_spend']), fmt_c)
        ws_mkt.write(r_num, 4, int(row['impressions']), fmt_n)
        ws_mkt.write(r_num, 5, int(row['clicks']), fmt_n)
        ws_mkt.write(r_num, 6, int(row['conversions']), fmt_n)
        ws_mkt.write(r_num, 7, float(row['revenue_gen']), fmt_c)
        ws_mkt.write(r_num, 8, float(row['roi']), fmt_d)
        ws_mkt.write(r_num, 9, float(row['cpa']), fmt_c)
        ws_mkt.write(r_num, 10, float(row['cvr']), fmt_p)
        ws_mkt.write(r_num, 11, float(row['ctr']), fmt_p)

    tot_m_r = 25 + len(mkt_chan)
    ws_mkt.write(tot_m_r, 1, "Total / Enterprise", f_tot_label)
    ws_mkt.write(tot_m_r, 2, "=SUM(C26:C35)", f_tot_num)
    ws_mkt.write(tot_m_r, 3, "=SUM(D26:D35)", f_tot_curr)
    ws_mkt.write(tot_m_r, 4, "=SUM(E26:E35)", f_tot_num)
    ws_mkt.write(tot_m_r, 5, "=SUM(F26:F35)", f_tot_num)
    ws_mkt.write(tot_m_r, 6, "=SUM(G26:G35)", f_tot_num)
    ws_mkt.write(tot_m_r, 7, "=SUM(H26:H35)", f_tot_curr)
    ws_mkt.write(tot_m_r, 8, "=H36/D36", f_tot_pct)
    ws_mkt.write(tot_m_r, 9, "=D36/G36", f_tot_curr)
    ws_mkt.write(tot_m_r, 10, "=G36/F36", f_tot_pct)
    ws_mkt.write(tot_m_r, 11, "=F36/E36", f_tot_pct)

    ws_mkt.freeze_panes(4, 0)

    # ==========================================
    # 9. REVIEWS ANALYSIS
    # ==========================================
    print("Building Reviews Analysis sheet...")
    ws_rev = wb.add_worksheet('Reviews Analysis')
    ws_rev.set_tab_color(NAVY)
    ws_rev.hide_gridlines(0)
    ws_rev.set_column('A:A', 3)
    ws_rev.set_column('B:K', 15)

    render_header_nav(
        ws_rev,
        "INDIAN E-COMMERCE CUSTOMER REVIEWS & SENTIMENT ANALYTICS",
        "Star Ratings Distribution, Customer Sentiment Breakdown, Verified Purchase Trust, and Category Satisfaction",
        8, 'K'
    )

    # KPI Cards
    ws_rev.write('B7', "TOTAL REVIEWS", f_card_top)
    ws_rev.write('B8', 77530, f_card_val_num)
    ws_rev.write('B9', "Customer Feedbacks", f_card_bot)

    ws_rev.write('C7', "AVERAGE RATING", f_card_top)
    ws_rev.write('C8', 3.69, f_card_val_dec)
    ws_rev.write('C9', "Out of 5.0 Stars", f_card_bot)

    ws_rev.write('D7', "POSITIVE SENTIMENT", f_card_top)
    ws_rev.write('D8', 0.5971, f_card_val_pct)
    ws_rev.write('D9', "46,168 Positive Reviews", f_card_bot)

    ws_rev.write('E7', "NEGATIVE SENTIMENT", f_card_top)
    ws_rev.write('E8', 0.0905, f_card_val_pct)
    ws_rev.write('E9', "6,997 Negative Reviews", f_card_bot)

    ws_rev.write('F7', "VERIFIED BUYERS", f_card_top)
    ws_rev.write('F8', 0.9200, f_card_val_pct)
    ws_rev.write('F9', "71,328 Verified Purchases", f_card_bot)

    ws_rev.write('G7', "5-STAR REVIEWS", f_card_top)
    ws_rev.write('G8', 15046, f_card_val_num)
    ws_rev.write('G9', "19.4% Top Rating Share", f_card_bot)

    ws_rev.set_row(6, 14)
    ws_rev.set_row(7, 26)
    ws_rev.set_row(8, 14)

    # Charts Header
    ws_rev.merge_range('B11:K11', "CUSTOMER SATISFACTION & SENTIMENT VISUALIZATIONS", f_sec_hdr)

    # Chart 1: Star Rating Distribution
    ch_rev_stars = wb.add_chart({'type': 'column'})
    ch_rev_stars.add_series({
        'name': 'Reviews Count',
        'categories': "='Reviews Analysis'!$B$26:$B$30",
        'values': "='Reviews Analysis'!$C$26:$C$30",
        'fill': {'color': NAVY}
    })
    ch_rev_stars.set_title({'name': 'Star Rating Distribution (1 to 5 Stars)', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch_rev_stars.set_x_axis({'name': 'Rating Stars', 'name_font': {'size': 8}})
    ch_rev_stars.set_legend({'none': True})
    ch_rev_stars.set_size({'width': 480, 'height': 240})
    ws_rev.insert_chart('B12', ch_rev_stars)

    # Chart 2: Sentiment Distribution
    ch_rev_sent = wb.add_chart({'type': 'doughnut'})
    ch_rev_sent.add_series({
        'name': 'Review Sentiment',
        'categories': "='Reviews Analysis'!$B$35:$B$37",
        'values': "='Reviews Analysis'!$C$35:$C$37",
        'points': [
            {'fill': {'color': GREEN}},
            {'fill': {'color': '#95A5A6'}},
            {'fill': {'color': RED}}
        ]
    })
    ch_rev_sent.set_title({'name': 'Customer Review Sentiment Breakdown', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch_rev_sent.set_legend({'position': 'right'})
    ch_rev_sent.set_size({'width': 440, 'height': 240})
    ws_rev.insert_chart('G12', ch_rev_sent)

    # Table 1: Star Ratings Breakdown
    ws_rev.merge_range('B24:H24', "STAR RATING & SENTIMENT CORRELATION", f_sec_hdr)
    rev_headers = ["Star Rating", "Review Count", "Share %", "Verified Purchases", "Positive Sentiment", "Neutral Sentiment", "Negative Sentiment"]
    for c_idx, h in enumerate(rev_headers):
        ws_rev.write(24, 1 + c_idx, h, f_th)

    for r_idx, row in rev_stars.iterrows():
        r_num = 25 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        fmt_n = f_td_num_alt if r_idx % 2 == 1 else f_td_num
        fmt_p = f_td_pct_alt if r_idx % 2 == 1 else f_td_pct

        ws_rev.write(r_num, 1, f"{int(row['rating'])} Stars", fmt_t)
        ws_rev.write(r_num, 2, int(row['review_count']), fmt_n)
        ws_rev.write(r_num, 3, float(row['share_pct']), fmt_p)
        ws_rev.write(r_num, 4, int(row['verified_count']), fmt_n)
        ws_rev.write(r_num, 5, int(row['positive_count']), fmt_n)
        ws_rev.write(r_num, 6, int(row['neutral_count']), fmt_n)
        ws_rev.write(r_num, 7, int(row['negative_count']), fmt_n)

    # Table 2: Sentiment Summary
    ws_rev.merge_range('B33:D33', "OVERALL SENTIMENT BREAKDOWN", f_sec_hdr)
    sent_headers = ["Sentiment", "Review Count", "Share %"]
    for c_idx, h in enumerate(sent_headers):
        ws_rev.write(33, 1 + c_idx, h, f_th)

    for r_idx, row in rev_sent.head(3).iterrows():
        r_num = 34 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        fmt_n = f_td_num_alt if r_idx % 2 == 1 else f_td_num
        fmt_p = f_td_pct_alt if r_idx % 2 == 1 else f_td_pct

        ws_rev.write(r_num, 1, row['sentiment'], fmt_t)
        ws_rev.write(r_num, 2, int(row['count']), fmt_n)
        ws_rev.write(r_num, 3, float(row['share_pct']), fmt_p)

    ws_rev.freeze_panes(4, 0)

    # ==========================================
    # 10. DATA DICTIONARY & DOCUMENTATION
    # ==========================================
    print("Building Data Dictionary & Documentation sheet...")
    ws_doc = wb.add_worksheet('Data Dictionary & Documentation')
    ws_doc.set_tab_color(ROYAL_BLUE)
    ws_doc.hide_gridlines(0)
    ws_doc.set_column('A:A', 3)
    ws_doc.set_column('B:B', 18)
    ws_doc.set_column('C:C', 22)
    ws_doc.set_column('D:D', 14)
    ws_doc.set_column('E:E', 45)
    ws_doc.set_column('F:F', 30)
    ws_doc.set_column('G:G', 12)
    ws_doc.set_column('H:H', 20)

    render_header_nav(
        ws_doc,
        "INDIAN E-COMMERCE ENTERPRISE DATA DICTIONARY & TECHNICAL DOCUMENTATION",
        "Data Architecture, Entity Relationships, Metric Formulas, and Automated Data Quality Audit Report",
        9, 'H'
    )

    # Project Overview Box
    ws_doc.merge_range('B6:H11',
        "PROJECT & ARCHITECTURAL OVERVIEW:\n"
        "• Academic Alignment: Designed for the B.Tech CSE Data Science & Big Data Analytics Presentation (IILM University).\n"
        "• Data Source: Complete Kaggle Indian E-Commerce Sales & Customer Intelligence Dataset (shiyalkishan01), covering 2023-01-01 to 2025-12-31.\n"
        "• Scale: 100,000 Transactions | 254,331 Line Items | 25,000 Customers | 550 Products | 15 Indian States | 31 Major Cities.\n"
        "• Relational Integrity: Star/Snowflake schema linked by consistent surrogate keys: customer_id, product_id, order_id, order_item_id, payment_id, shipment_id, return_id, review_id, campaign_id.\n"
        "• Visual Design: Aligned with the 'DataScience_BigData_Presentation.pptx' design language (#174A78 Navy, #216DA8 Royal Blue, #08A9D0 Cyan, #EAF5FB Pale Blue, #F3F4F6 Light Grey).",
        wb.add_format({'font_size': 9, 'font_name': 'Calibri', 'font_color': NAVY, 'bg_color': PALE_BLUE, 'border': 1, 'valign': 'top'})
    )

    # Core Metric Definitions Table
    ws_doc.merge_range('B13:H13', "CORE KPI DEFINITIONS & MATHEMATICAL FORMULAS", f_sec_hdr)
    kpi_def_headers = ["Metric Name", "Mathematical Formula", "Granularity", "Business Purpose & Reconciliation Rule", "Dataset Source"]
    ws_doc.write('B14', "Metric Name", f_th_left)
    ws_doc.write('C14', "Mathematical Formula", f_th_left)
    ws_doc.write('D14', "Granularity", f_th_left)
    ws_doc.merge_range('E14:G14', "Business Purpose & Reconciliation Rule", f_th_left)
    ws_doc.write('H14', "Source Table", f_th_left)

    kpi_defs = [
        ("Gross Revenue", "SUM(order_items.item_revenue) == SUM(orders.subtotal)", "Transaction / Line Item", "Total value of items ordered before tax and shipping fees. Reconciles exactly across tables.", "orders / order_items"),
        ("Net Revenue", "SUM(Delivered subtotal) - SUM(returns.refund_amount)", "Fulfilled Orders", "True retained top-line after subtracting product returns and customer refunds.", "orders / returns"),
        ("Gross Profit", "SUM(order_items.profit) == SUM(item_revenue - item_cost)", "Line Item", "Manufacturing/merchant margin before operating expenses. Never fabricated without cost data.", "order_items"),
        ("Gross Margin %", "Gross Profit / Gross Revenue", "Percentage Ratio", "Direct unit economics profitability indicator.", "order_items"),
        ("Average Order Value (AOV)", "Gross Revenue / Total Orders Placed", "Order Level", "Average basket size per purchasing transaction.", "orders"),
        ("Return Rate (Order)", "Returned Orders / Total Orders Placed", "Order Level", "Percentage of orders where fulfillment resulted in a return.", "orders"),
        ("Return Rate (Item)", "Returned Items / Total Units Sold", "Unit Level", "Specific product defect or dissatisfaction rate.", "returns / order_items"),
        ("Repeat Customer Rate", "Customers with >1 Order / Purchasing Customers", "Customer Level", "Customer loyalty and retention velocity metric.", "customers / orders"),
        ("Campaign ROI", "Attributed Revenue / Campaign Ad Spend", "Campaign Channel", "Marketing capital efficiency multiplier.", "marketing_campaigns"),
        ("On-Time Delivery Rate", "1 - (Delayed Shipments / Total Shipments)", "Logistics Shipment", "SLA compliance for delivery logistics.", "shipments")
    ]

    for r_idx, row in enumerate(kpi_defs):
        r_num = 15 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        ws_doc.write(r_num, 1, row[0], fmt_t)
        ws_doc.write(r_num, 2, row[1], fmt_t)
        ws_doc.write(r_num, 3, row[2], fmt_t)
        ws_doc.merge_range(r_num, 4, r_num, 6, row[3], fmt_t)
        ws_doc.write(r_num, 7, row[4], fmt_t)

    # Complete Data Dictionary Table
    dict_start_r = 27
    ws_doc.merge_range(f'B{dict_start_r}:H{dict_start_r}', "DATA DICTIONARY (SCHEMA & COLUMN DEFINITIONS)", f_sec_hdr)
    dict_headers = ["Table Name", "Column Name", "Data Type", "Description", "Business Meaning", "Nullable", "Sample Value"]
    for c_idx, h in enumerate(dict_headers):
        ws_doc.write(dict_start_r, 1 + c_idx, h, f_th)

    for r_idx, row in data_dict.iterrows():
        r_num = dict_start_r + 1 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        ws_doc.write(r_num, 1, str(row['table_name']), fmt_t)
        ws_doc.write(r_num, 2, str(row['column_name']), fmt_t)
        ws_doc.write(r_num, 3, str(row['data_type']), fmt_t)
        ws_doc.write(r_num, 4, str(row['description']), fmt_t)
        ws_doc.write(r_num, 5, str(row['business_meaning']), fmt_t)
        ws_doc.write(r_num, 6, str(row['nullable']), fmt_t)
        ws_doc.write(r_num, 7, str(row['example']), fmt_t)

    # Validation Report Section
    val_start_r = dict_start_r + len(data_dict) + 3
    ws_doc.merge_range(f'B{val_start_r}:F{val_start_r}', "AUTOMATED DATA QUALITY VALIDATION AUDIT (43 CHECKS PASS)", f_sec_hdr)
    val_headers = ["Validation Check", "Expected Condition", "Actual Result", "Status", "Notes"]
    for c_idx, h in enumerate(val_headers):
        ws_doc.write(val_start_r, 1 + c_idx, h, f_th)

    for r_idx, row in val_rep.iterrows():
        r_num = val_start_r + 1 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        ws_doc.write(r_num, 1, str(row['validation_check']), fmt_t)
        ws_doc.write(r_num, 2, str(row['expected_condition']), fmt_t)
        ws_doc.write(r_num, 3, str(row['actual_result']), fmt_t)
        ws_doc.write(r_num, 4, str(row['status']), fmt_t)
        ws_doc.write(r_num, 5, str(row['notes']) if pd.notnull(row['notes']) else "-", fmt_t)

    ws_doc.freeze_panes(4, 0)

    # ==========================================
    # 11. SUPPORTING DATA
    # ==========================================
    print("Building Supporting Data sheet...")
    ws_sup = wb.add_worksheet('Supporting Data')
    ws_sup.set_tab_color(ROYAL_BLUE)
    ws_sup.hide_gridlines(0)
    ws_sup.set_column('A:AO', 14)

    render_header_nav(
        ws_sup,
        "INDIAN E-COMMERCE ANALYTICAL CUBES & REFERENCE MASTER DATA",
        "Pre-aggregated multidimensional data feeding charts, dynamic lookups, and executive scorecards",
        10, 'N'
    )

    # Table 1: Monthly Master (Cols A to G)
    ws_sup.merge_range('A5:G5', "MONTHLY MASTER CUBE", f_th)
    sup_m_headers = ["Year-Month", "Year", "Month", "Gross Revenue", "Units Sold", "Orders Count", "Gross Profit"]
    for c, h in enumerate(sup_m_headers):
        ws_sup.write(5, c, h, f_th)
    for r, row in monthly_all.iterrows():
        ws_sup.write(6 + r, 0, row['year_month'], f_td_text)
        ws_sup.write(6 + r, 1, int(row['year']), f_td_num)
        ws_sup.write(6 + r, 2, int(row['month']), f_td_num)
        ws_sup.write(6 + r, 3, float(row['subtotal_rev']), f_td_curr)
        ws_sup.write(6 + r, 4, int(row['units_sold']), f_td_num)
        ws_sup.write(6 + r, 5, int(row['orders_count']), f_td_num)
        ws_sup.write(6 + r, 6, float(row['profit']), f_td_curr)

    # Table 2: Category Master (Cols I to N)
    ws_sup.merge_range('I5:N5', "CATEGORY MASTER CUBE", f_th)
    sup_cat_headers = ["Category", "SKU Count", "Revenue", "Units Sold", "Profit", "Margin %"]
    for c, h in enumerate(sup_cat_headers):
        ws_sup.write(5, 8 + c, h, f_th)
    for r, row in cat_all.head(10).iterrows():
        ws_sup.write(6 + r, 8, row['category'], f_td_text)
        ws_sup.write(6 + r, 9, int(row['sku_count']), f_td_num)
        ws_sup.write(6 + r, 10, float(row['revenue']), f_td_curr)
        ws_sup.write(6 + r, 11, int(row['units_sold']), f_td_num)
        ws_sup.write(6 + r, 12, float(row['profit']), f_td_curr)
        ws_sup.write(6 + r, 13, float(row['margin_pct']), f_td_pct)

    # Table 3: Order Status (Cols Q to S)
    ws_sup.merge_range('Q5:S5', "ORDER STATUS MASTER", f_th)
    sup_stat_headers = ["Order Status", "Order Count", "Share %"]
    for c, h in enumerate(sup_stat_headers):
        ws_sup.write(5, 16 + c, h, f_th)
    status_summary = orders['order_status'].value_counts().reset_index()
    status_summary.columns = ['status', 'count']
    status_summary['share'] = status_summary['count'] / status_summary['count'].sum()
    for r, row in status_summary.iterrows():
        ws_sup.write(6 + r, 16, row['status'], f_td_text)
        ws_sup.write(6 + r, 17, int(row['count']), f_td_num)
        ws_sup.write(6 + r, 18, float(row['share']), f_td_pct)

    # Table 4: Payment Methods (Cols U to X)
    ws_sup.merge_range('U5:X5', "PAYMENT METHOD MASTER", f_th)
    sup_pay_headers = ["Payment Method", "Transactions", "Amount Paid", "Share %"]
    for c, h in enumerate(sup_pay_headers):
        ws_sup.write(5, 20 + c, h, f_th)
    for r, row in pay_meth.head(10).iterrows():
        ws_sup.write(6 + r, 20, row['payment_method'], f_td_text)
        ws_sup.write(6 + r, 21, int(row['transactions']), f_td_num)
        ws_sup.write(6 + r, 22, float(row['amount_paid']), f_td_curr)
        ws_sup.write(6 + r, 23, float(row['share_pct']), f_td_pct)

    # Table 5: State Master (Cols AA to AD)
    ws_sup.merge_range('AA5:AD5', "STATE MASTER CUBE", f_th)
    sup_st_headers = ["State", "Orders Count", "Revenue", "Share %"]
    for c, h in enumerate(sup_st_headers):
        ws_sup.write(5, 26 + c, h, f_th)
    for r, row in state_all.head(15).iterrows():
        ws_sup.write(6 + r, 26, row['delivery_state'], f_td_text)
        ws_sup.write(6 + r, 27, int(row['orders_count']), f_td_num)
        ws_sup.write(6 + r, 28, float(row['revenue']), f_td_curr)
        ws_sup.write(6 + r, 29, float(row['rev_share']), f_td_pct)

    # Table 6: Quarterly Master (Cols AE to AH)
    ws_sup.merge_range('AE5:AH5', "QUARTERLY MASTER CUBE", f_th)
    sup_q_headers = ["Quarter", "Orders", "Revenue", "Profit"]
    for c, h in enumerate(sup_q_headers):
        ws_sup.write(5, 30 + c, h, f_th)
    for r, row in quarterly_all.iterrows():
        ws_sup.write(6 + r, 30, row['year_quarter'], f_td_text)
        ws_sup.write(6 + r, 31, int(row['orders_count']), f_td_num)
        ws_sup.write(6 + r, 32, float(row['subtotal_rev']), f_td_curr)
        ws_sup.write(6 + r, 33, float(row['profit']), f_td_curr)

    # Granular Orders Fact Extract Table (5,000 Rows Sample for Audit & Direct Exploration)
    ws_sup.merge_range('A46:N46', "AUDIT SAMPLE: 5,000 GRANULAR ORDERS FACT RECORDS", f_sec_hdr)
    audit_headers = ["Order ID", "Customer ID", "Order Date", "Year", "Month", "Order Status", "Delivery City", "Delivery State", "Subtotal (INR)", "Tax (INR)", "Final Amount (INR)", "Payment Method", "Delivery Status", "Delayed Flag"]
    for c, h in enumerate(audit_headers):
        ws_sup.write(46, c, h, f_th)

    # Sample 5,000 orders
    orders_sample = orders.head(5000).merge(
        payments[['order_id', 'payment_method']].drop_duplicates('order_id'), on='order_id', how='left'
    ).merge(
        shipments[['order_id', 'delivery_status', 'delayed_flag']].drop_duplicates('order_id'), on='order_id', how='left'
    )
    orders_sample['delayed_flag'] = orders_sample['delayed_flag'].fillna(0).astype(int)
    orders_sample['delivery_status'] = orders_sample['delivery_status'].fillna('Not Shipped')

    for r_idx, row in orders_sample.iterrows():
        r_num = 47 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        fmt_n = f_td_num_alt if r_idx % 2 == 1 else f_td_num
        fmt_c = f_td_curr_alt if r_idx % 2 == 1 else f_td_curr

        ws_sup.write(r_num, 0, str(row['order_id']), fmt_t)
        ws_sup.write(r_num, 1, str(row['customer_id']), fmt_t)
        ws_sup.write(r_num, 2, str(row['order_date'])[:10], fmt_t)
        ws_sup.write(r_num, 3, int(row['year']), fmt_n)
        ws_sup.write(r_num, 4, int(row['month']), fmt_n)
        ws_sup.write(r_num, 5, str(row['order_status']), fmt_t)
        ws_sup.write(r_num, 6, str(row['delivery_city']), fmt_t)
        ws_sup.write(r_num, 7, str(row['delivery_state']), fmt_t)
        ws_sup.write(r_num, 8, float(row['subtotal']), fmt_c)
        ws_sup.write(r_num, 9, float(row['tax_amount']), fmt_c)
        ws_sup.write(r_num, 10, float(row['final_amount']), fmt_c)
        ws_sup.write(r_num, 11, str(row['payment_method']), fmt_t)
        ws_sup.write(r_num, 12, str(row['delivery_status']), fmt_t)
        ws_sup.write(r_num, 13, int(row['delayed_flag']), fmt_n)

    ws_sup.freeze_panes(4, 0)

    # Ensure Executive Dashboard is active when workbook opens
    ws_exec.activate()

    print("Closing workbook and saving to disk...")
    wb.close()
    print("Workbook successfully created at:", output_file)
    print("File size:", round(os.path.getsize(output_file) / (1024 * 1024), 2), "MB")

if __name__ == '__main__':
    build_dashboard()
