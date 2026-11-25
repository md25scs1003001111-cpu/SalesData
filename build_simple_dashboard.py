"""
Dynamic, Multi-Filter Reactive Indian E-Commerce Sales Analytics Dashboard
Features real-time 4-axis filtering (Year, Category, Payment Mode, Order Status)
Driving all 6 KPI cards, 5 native charts, and the Annual Performance Scorecard.
"""

import os
import pandas as pd
import xlsxwriter

def create_simple_dashboard():
    raw_dir = '/Users/zeeshan/sale/raw_dataset'
    output_file = '/Users/zeeshan/sale/Indian_Ecommerce_Sales_Analytics_Dashboard.xlsx'

    print("Loading raw dataset...")
    orders = pd.read_csv(f'{raw_dir}/orders.csv')
    items = pd.read_csv(f'{raw_dir}/order_items.csv')
    products = pd.read_csv(f'{raw_dir}/products.csv')
    payments = pd.read_csv(f'{raw_dir}/payments.csv')

    orders['order_date'] = pd.to_datetime(orders['order_date'])
    orders['year'] = orders['order_date'].dt.year.astype(str)
    orders['month'] = orders['order_date'].dt.month
    orders['year_month'] = orders['order_date'].dt.to_period('M').astype(str)

    # Base joined table
    df = items.merge(products[['product_id', 'product_name', 'category']], on='product_id')
    df = df.merge(orders[['order_id', 'year', 'month', 'year_month', 'order_status', 'delivery_state']], on='order_id')
    df = df.merge(payments[['order_id', 'payment_method']], on='order_id')

    print("Building multidimensional analytical cubes...")
    # 1. Monthly Master Cube (Year, Month, Year-Month, Category, Payment Method, Order Status)
    cube_monthly = df.groupby(['year', 'month', 'year_month', 'category', 'payment_method', 'order_status']).agg(
        orders=('order_id', 'nunique'),
        units=('quantity', 'sum'),
        sales=('item_revenue', 'sum'),
        profit=('profit', 'sum')
    ).reset_index().sort_values(['year', 'month', 'category', 'payment_method', 'order_status']).reset_index(drop=True)

    # 2. State Cube (Delivery State, Year, Category, Payment Method, Order Status)
    cube_state = df.groupby(['delivery_state', 'year', 'category', 'payment_method', 'order_status']).agg(
        sales=('item_revenue', 'sum'),
        orders=('order_id', 'nunique')
    ).reset_index().sort_values(['delivery_state', 'year', 'category', 'payment_method', 'order_status']).reset_index(drop=True)

    # 3. Product Catalog Reference Table (Key = Category_Rank)
    catalog_rows = []
    # Overall top 10
    top_overall = df.groupby(['product_id', 'product_name', 'category'])['item_revenue'].sum().sort_values(ascending=False).head(10).reset_index()
    for rank, (_, r) in enumerate(top_overall.iterrows(), 1):
        catalog_rows.append({
            'key': f'All_{rank}',
            'product_id': r['product_id'],
            'product_name': str(r['product_name'])[:28],
            'category': r['category'],
            'rank': rank
        })
    # Per-category top 10
    for cat, grp in df.groupby('category'):
        cat_top = grp.groupby(['product_id', 'product_name', 'category'])['item_revenue'].sum().sort_values(ascending=False).head(10).reset_index()
        for rank, (_, r) in enumerate(cat_top.iterrows(), 1):
            catalog_rows.append({
                'key': f'{cat}_{rank}',
                'product_id': r['product_id'],
                'product_name': str(r['product_name'])[:28],
                'category': r['category'],
                'rank': rank
            })
    prod_catalog_df = pd.DataFrame(catalog_rows).reset_index(drop=True)
    all_top_pids = set(prod_catalog_df['product_id'])

    # 4. Product Cube (Top Products, Year, Payment Method, Order Status)
    df_top_prods = df[df['product_id'].isin(all_top_pids)]
    cube_product = df_top_prods.groupby(['product_id', 'year', 'payment_method', 'order_status']).agg(
        units=('quantity', 'sum'),
        sales=('item_revenue', 'sum')
    ).reset_index().sort_values(['product_id', 'year', 'payment_method', 'order_status']).reset_index(drop=True)

    # Pre-calculated default presentation tables (All Filters = All)
    # 36-Month list
    month_list = df[['year_month', 'year', 'month']].drop_duplicates().sort_values(['year', 'month']).reset_index(drop=True)
    m_default = cube_monthly.groupby('year_month').agg(sales=('sales', 'sum'), profit=('profit', 'sum'), orders=('orders', 'sum')).reset_index()
    month_list = month_list.merge(m_default, on='year_month', how='left')

    # Top 10 Categories default
    cat_default = cube_monthly.groupby('category').agg(sales=('sales', 'sum'), profit=('profit', 'sum')).reset_index().sort_values('sales', ascending=False).head(10).reset_index(drop=True)
    cat_skus = df.groupby('category')['product_id'].nunique().reset_index().rename(columns={'product_id': 'skus'})
    cat_default = cat_default.merge(cat_skus, on='category', how='left')

    # Top Payment Methods default
    pay_default = cube_monthly.groupby('payment_method').agg(sales=('sales', 'sum'), orders=('orders', 'sum')).reset_index().sort_values('sales', ascending=False).head(6).reset_index(drop=True)

    # Top 10 States default
    state_default = cube_state.groupby('delivery_state').agg(sales=('sales', 'sum'), orders=('orders', 'sum')).reset_index().sort_values('sales', ascending=False).head(10).reset_index(drop=True)

    print("Creating Excel workbook...")
    wb = xlsxwriter.Workbook(output_file, {'nan_inf_to_errors': True})

    # Exact PowerPoint Theme Colors
    NAVY = '#174A78'        # Dark Navy (Primary Accent)
    ROYAL_BLUE = '#216DA8'  # Royal Blue (Secondary)
    CYAN = '#08A9D0'        # Cyan Highlight
    PALE_BLUE = '#EAF5FB'   # Card / Info Fill
    LIGHT_GREY = '#F3F4F6'  # Panel Fill
    WHITE = '#FFFFFF'
    BORDER_COL = '#D5DCE3'
    TEXT_DARK = '#2C3E50'
    TEXT_MUTED = '#5B7A99'

    # Formats
    f_title = wb.add_format({
        'bold': True, 'font_size': 16, 'font_name': 'Calibri',
        'font_color': WHITE, 'bg_color': NAVY,
        'align': 'center', 'valign': 'vcenter'
    })
    f_sub = wb.add_format({
        'italic': True, 'font_size': 9.5, 'font_name': 'Calibri',
        'font_color': WHITE, 'bg_color': NAVY,
        'align': 'center', 'valign': 'vcenter'
    })
    f_ctrl_lbl = wb.add_format({
        'bold': True, 'font_size': 9, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'right', 'valign': 'vcenter', 'border': 1, 'border_color': BORDER_COL
    })
    f_ctrl_val = wb.add_format({
        'bold': True, 'font_size': 9.5, 'font_name': 'Calibri',
        'font_color': ROYAL_BLUE, 'bg_color': WHITE,
        'align': 'center', 'valign': 'vcenter', 'border': 2, 'border_color': ROYAL_BLUE
    })
    f_kpi_lbl = wb.add_format({
        'bold': True, 'font_size': 8.5, 'font_name': 'Calibri',
        'font_color': TEXT_MUTED, 'bg_color': PALE_BLUE,
        'align': 'center', 'valign': 'vcenter', 'top': 1, 'left': 1, 'right': 1, 'border_color': BORDER_COL
    })
    f_kpi_val_curr = wb.add_format({
        'bold': True, 'font_size': 13, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'center', 'valign': 'vcenter', 'num_format': '₹#,##0',
        'left': 1, 'right': 1, 'border_color': BORDER_COL
    })
    f_kpi_val_num = wb.add_format({
        'bold': True, 'font_size': 13, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'center', 'valign': 'vcenter', 'num_format': '#,##0',
        'left': 1, 'right': 1, 'border_color': BORDER_COL
    })
    f_kpi_val_pct = wb.add_format({
        'bold': True, 'font_size': 13, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'center', 'valign': 'vcenter', 'num_format': '0.0%',
        'left': 1, 'right': 1, 'border_color': BORDER_COL
    })
    f_kpi_val_text = wb.add_format({
        'bold': True, 'font_size': 12, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'center', 'valign': 'vcenter',
        'left': 1, 'right': 1, 'border_color': BORDER_COL
    })
    f_kpi_sub = wb.add_format({
        'font_size': 7.5, 'font_name': 'Calibri', 'italic': True,
        'font_color': TEXT_MUTED, 'bg_color': PALE_BLUE,
        'align': 'center', 'valign': 'vcenter', 'bottom': 1, 'left': 1, 'right': 1, 'border_color': BORDER_COL
    })
    f_sec_hdr = wb.add_format({
        'bold': True, 'font_size': 9.5, 'font_name': 'Calibri',
        'font_color': WHITE, 'bg_color': ROYAL_BLUE,
        'align': 'left', 'valign': 'vcenter', 'indent': 1
    })
    f_th = wb.add_format({
        'bold': True, 'font_size': 8.5, 'font_name': 'Calibri',
        'font_color': WHITE, 'bg_color': NAVY,
        'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': BORDER_COL
    })
    f_td_text = wb.add_format({
        'font_size': 8, 'font_name': 'Calibri', 'font_color': TEXT_DARK,
        'align': 'left', 'valign': 'vcenter', 'border': 1, 'border_color': BORDER_COL
    })
    f_td_text_alt = wb.add_format({
        'font_size': 8, 'font_name': 'Calibri', 'font_color': TEXT_DARK, 'bg_color': LIGHT_GREY,
        'align': 'left', 'valign': 'vcenter', 'border': 1, 'border_color': BORDER_COL
    })
    f_td_num = wb.add_format({
        'font_size': 8, 'font_name': 'Calibri', 'font_color': TEXT_DARK,
        'align': 'right', 'valign': 'vcenter', 'num_format': '#,##0', 'border': 1, 'border_color': BORDER_COL
    })
    f_td_num_alt = wb.add_format({
        'font_size': 8, 'font_name': 'Calibri', 'font_color': TEXT_DARK, 'bg_color': LIGHT_GREY,
        'align': 'right', 'valign': 'vcenter', 'num_format': '#,##0', 'border': 1, 'border_color': BORDER_COL
    })
    f_td_curr = wb.add_format({
        'font_size': 8, 'font_name': 'Calibri', 'font_color': TEXT_DARK,
        'align': 'right', 'valign': 'vcenter', 'num_format': '₹#,##0', 'border': 1, 'border_color': BORDER_COL
    })
    f_td_curr_alt = wb.add_format({
        'font_size': 8, 'font_name': 'Calibri', 'font_color': TEXT_DARK, 'bg_color': LIGHT_GREY,
        'align': 'right', 'valign': 'vcenter', 'num_format': '₹#,##0', 'border': 1, 'border_color': BORDER_COL
    })
    f_td_pct = wb.add_format({
        'font_size': 8, 'font_name': 'Calibri', 'font_color': TEXT_DARK,
        'align': 'right', 'valign': 'vcenter', 'num_format': '0.0%', 'border': 1, 'border_color': BORDER_COL
    })
    f_td_pct_alt = wb.add_format({
        'font_size': 8, 'font_name': 'Calibri', 'font_color': TEXT_DARK, 'bg_color': LIGHT_GREY,
        'align': 'right', 'valign': 'vcenter', 'num_format': '0.0%', 'border': 1, 'border_color': BORDER_COL
    })
    f_tot_lbl = wb.add_format({
        'bold': True, 'font_size': 8.5, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'left', 'valign': 'vcenter', 'top': 2, 'bottom': 2, 'border_color': NAVY
    })
    f_tot_num = wb.add_format({
        'bold': True, 'font_size': 8.5, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'right', 'valign': 'vcenter', 'num_format': '#,##0', 'top': 2, 'bottom': 2, 'border_color': NAVY
    })
    f_tot_curr = wb.add_format({
        'bold': True, 'font_size': 8.5, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'right', 'valign': 'vcenter', 'num_format': '₹#,##0', 'top': 2, 'bottom': 2, 'border_color': NAVY
    })
    f_tot_pct = wb.add_format({
        'bold': True, 'font_size': 8.5, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'right', 'valign': 'vcenter', 'num_format': '0.0%', 'top': 2, 'bottom': 2, 'border_color': NAVY
    })

    # ==========================================
    # SHEET 1: SALES DASHBOARD (EXECUTIVE SCREEN)
    # ==========================================
    print("Building 'Sales Dashboard'...")
    ws_dash = wb.add_worksheet('Sales Dashboard')
    ws_dash.set_tab_color(NAVY)
    ws_dash.hide_gridlines(0)

    # Column Widths
    ws_dash.set_column('A:A', 2.5)
    ws_dash.set_column('B:B', 15)
    ws_dash.set_column('C:C', 15)
    ws_dash.set_column('D:D', 15)
    ws_dash.set_column('E:E', 15)
    ws_dash.set_column('F:F', 15)
    ws_dash.set_column('G:G', 15)
    ws_dash.set_column('H:H', 15)
    ws_dash.set_column('I:I', 15)
    ws_dash.set_column('J:J', 15)
    ws_dash.set_column('K:K', 15)
    ws_dash.set_column('L:L', 15)
    ws_dash.set_column('M:M', 15)
    ws_dash.set_column('N:N', 2.5)

    # Header Banner (Rows 2-3)
    ws_dash.set_row(1, 28)
    ws_dash.set_row(2, 16)
    ws_dash.merge_range('B2:M2', "INDIAN E-COMMERCE SALES ANALYTICS DASHBOARD", f_title)
    ws_dash.merge_range('B3:M3', "Executive Performance & Commercial Insights (2023 - 2025) | IILM University Data Science Project", f_sub)

    # Interactive Filter Control Bar (Row 5)
    ws_dash.set_row(4, 22)
    ws_dash.write('B5', "Select Year:", f_ctrl_lbl)
    ws_dash.write('C5', "All", f_ctrl_val)
    ws_dash.data_validation('C5', {'validate': 'list', 'source': ['All', '2023', '2024', '2025']})

    ws_dash.write('E5', "Category:", f_ctrl_lbl)
    ws_dash.write('F5', "All", f_ctrl_val)
    categories_list = ['All', 'Electronics', 'Appliances', 'Fashion', 'Home & Kitchen', 'Footwear', 'Furniture', 'Grocery', 'Bags', 'Health & Wellness', 'Sports', 'Toys', 'Beauty', 'Watches', 'Books']
    ws_dash.data_validation('F5', {'validate': 'list', 'source': categories_list})

    ws_dash.write('H5', "Payment Mode:", f_ctrl_lbl)
    ws_dash.write('I5', "All", f_ctrl_val)
    payments_list = ['All', 'UPI', 'Cash on Delivery', 'Credit Card', 'Debit Card', 'Net Banking', 'Wallet', 'EMI', 'Pay Later']
    ws_dash.data_validation('I5', {'validate': 'list', 'source': payments_list})

    ws_dash.write('K5', "Order Status:", f_ctrl_lbl)
    ws_dash.write('L5', "All", f_ctrl_val)
    ws_dash.data_validation('L5', {'validate': 'list', 'source': ['All', 'Delivered', 'Cancelled', 'Returned', 'Failed', 'Processing']})

    ws_dash.write('M5', "RESET ALL", wb.add_format({'bold': True, 'font_size': 9, 'font_name': 'Calibri', 'font_color': WHITE, 'bg_color': ROYAL_BLUE, 'align': 'center', 'valign': 'vcenter', 'border': 1}))

    # 6 Focused KPI Cards (Rows 7-9) - Fully Dynamic!
    ws_dash.set_row(6, 14)
    ws_dash.set_row(7, 24)
    ws_dash.set_row(8, 14)

    # Card 1: Total Sales (Dynamic from Analysis Data D7:D42)
    ws_dash.merge_range('B7:C7', "TOTAL SALES", f_kpi_lbl)
    ws_dash.merge_range('B8:C8', "", f_kpi_val_curr)
    ws_dash.write_formula('B8', "=SUM('Analysis Data'!$D$7:$D$42)", f_kpi_val_curr, 2449898974.35)
    ws_dash.merge_range('B9:C9', "₹2,449.90 Cr Placed", f_kpi_sub)

    # Card 2: Total Profit (Dynamic from Analysis Data E7:E42)
    ws_dash.merge_range('D7:E7', "TOTAL PROFIT", f_kpi_lbl)
    ws_dash.merge_range('D8:E8', "", f_kpi_val_curr)
    ws_dash.write_formula('D8', "=SUM('Analysis Data'!$E$7:$E$42)", f_kpi_val_curr, 197294347.02)
    ws_dash.merge_range('D9:E9', "₹197.29 Cr Realized Margin", f_kpi_sub)

    # Card 3: Profit Margin %
    ws_dash.merge_range('F7:G7', "PROFIT MARGIN %", f_kpi_lbl)
    ws_dash.merge_range('F8:G8', "", f_kpi_val_pct)
    ws_dash.write_formula('F8', '=IFERROR(D8/B8, 0)', f_kpi_val_pct, 0.08053)
    ws_dash.merge_range('F9:G9', "8.05% Gross Margin", f_kpi_sub)

    # Card 4: Total Orders (Dynamic from Analysis Data F7:F42)
    ws_dash.merge_range('H7:I7', "TOTAL ORDERS", f_kpi_lbl)
    ws_dash.merge_range('H8:I8', "", f_kpi_val_num)
    ws_dash.write_formula('H8', "=SUM('Analysis Data'!$F$7:$F$42)", f_kpi_val_num, 100000)
    ws_dash.merge_range('H9:I9', "Active Filtered Count", f_kpi_sub)

    # Card 5: Average Order Value
    ws_dash.merge_range('J7:K7', "AVERAGE ORDER VALUE", f_kpi_lbl)
    ws_dash.merge_range('J8:K8', "", f_kpi_val_curr)
    ws_dash.write_formula('J8', '=IFERROR(B8/H8, 0)', f_kpi_val_curr, 24498.99)
    ws_dash.merge_range('J9:K9', "Gross Basket Size", f_kpi_sub)

    # Card 6: Top Category (Dynamic formula finding top category by revenue)
    ws_dash.merge_range('L7:M7', "TOP CATEGORY", f_kpi_lbl)
    ws_dash.merge_range('L8:M8', "", f_kpi_val_text)
    ws_dash.write_formula('L8', '=IF($F$5<>"All", $F$5, INDEX(\'Analysis Data\'!$H$7:$H$16, MATCH(MAX(\'Analysis Data\'!$J$7:$J$16), \'Analysis Data\'!$J$7:$J$16, 0)))', f_kpi_val_text, "Electronics")
    ws_dash.merge_range('L9:M9', "Highest Revenue Share", f_kpi_sub)

    # 5 Native Charts on Dashboard
    # Chart 1: Monthly Sales & Profit Trend (Cols B to G, Rows 11-25)
    ch1 = wb.add_chart({'type': 'column'})
    ch1.add_series({
        'name': 'Monthly Sales',
        'categories': "='Analysis Data'!$A$7:$A$42",
        'values': "='Analysis Data'!$D$7:$D$42",
        'fill': {'color': NAVY}
    })
    ch1_line = wb.add_chart({'type': 'line'})
    ch1_line.add_series({
        'name': 'Monthly Profit',
        'categories': "='Analysis Data'!$A$7:$A$42",
        'values': "='Analysis Data'!$E$7:$E$42",
        'line': {'color': CYAN, 'width': 2}
    })
    ch1.combine(ch1_line)
    ch1.set_title({'name': 'Monthly Sales & Profit Trend (2023 - 2025)', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch1.set_x_axis({'num_font': {'size': 7, 'rotation': -45}})
    ch1.set_y_axis({'name': 'Amount (INR)', 'name_font': {'size': 8}, 'major_gridlines': {'visible': True, 'line': {'color': BORDER_COL, 'dash_type': 'dash'}}})
    ch1.set_legend({'position': 'top'})
    ch1.set_size({'width': 500, 'height': 240})
    ws_dash.insert_chart('B11', ch1)

    # Chart 2: Sales by Category (Cols H to M, Rows 11-25)
    ch2 = wb.add_chart({'type': 'bar'})
    ch2.add_series({
        'name': 'Sales',
        'categories': "='Analysis Data'!$H$7:$H$16",
        'values': "='Analysis Data'!$J$7:$J$16",
        'fill': {'color': ROYAL_BLUE}
    })
    ch2.set_title({'name': 'Top 10 Categories by Revenue', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch2.set_y_axis({'num_font': {'size': 8}})
    ch2.set_legend({'none': True})
    ch2.set_size({'width': 500, 'height': 240})
    ws_dash.insert_chart('H11', ch2)

    # Chart 3: Payment Modes (Cols B to E, Rows 26-39)
    ch3 = wb.add_chart({'type': 'doughnut'})
    ch3.add_series({
        'name': 'Payment Mode',
        'categories': "='Analysis Data'!$N$7:$N$12",
        'values': "='Analysis Data'!$P$7:$P$12",
        'points': [
            {'fill': {'color': NAVY}},
            {'fill': {'color': ROYAL_BLUE}},
            {'fill': {'color': CYAN}},
            {'fill': {'color': '#5B7A99'}},
            {'fill': {'color': '#95A5A6'}},
            {'fill': {'color': PALE_BLUE}}
        ]
    })
    ch3.set_title({'name': 'Payment Instrument Share', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch3.set_legend({'position': 'right', 'font': {'size': 8}})
    ch3.set_size({'width': 320, 'height': 210})
    ws_dash.insert_chart('B26', ch3)

    # Chart 4: Top Products (Cols F to I, Rows 26-39)
    ch4 = wb.add_chart({'type': 'column'})
    ch4.add_series({
        'name': 'Revenue',
        'categories': "='Analysis Data'!$T$7:$T$16",
        'values': "='Analysis Data'!$V$7:$V$16",
        'fill': {'color': CYAN}
    })
    ch4.set_title({'name': 'Top 10 Revenue Products', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch4.set_x_axis({'num_font': {'size': 7, 'rotation': -35}})
    ch4.set_y_axis({'major_gridlines': {'visible': True, 'line': {'color': BORDER_COL, 'dash_type': 'dash'}}})
    ch4.set_legend({'none': True})
    ch4.set_size({'width': 340, 'height': 210})
    ws_dash.insert_chart('F26', ch4)

    # Chart 5: Top States (Cols J to M, Rows 26-39)
    ch5 = wb.add_chart({'type': 'bar'})
    ch5.add_series({
        'name': 'State Sales',
        'categories': "='Analysis Data'!$X$7:$X$16",
        'values': "='Analysis Data'!$Z$7:$Z$16",
        'fill': {'color': NAVY}
    })
    ch5.set_title({'name': 'Top 10 Indian States by Sales', 'name_font': {'size': 10, 'bold': True, 'color': NAVY}})
    ch5.set_y_axis({'num_font': {'size': 8}})
    ch5.set_legend({'none': True})
    ch5.set_size({'width': 340, 'height': 210})
    ws_dash.insert_chart('J26', ch5)

    # Executive Annual Summary Table (Rows 41-47) - Fully Dynamic!
    ws_dash.merge_range('B40:M40', "ANNUAL PERFORMANCE SCORECARD", f_sec_hdr)
    ws_dash.set_row(39, 20)

    ws_dash.write('B41', "Year", f_th)
    ws_dash.write('C41', "Orders", f_th)
    ws_dash.write('D41', "Units Sold", f_th)
    ws_dash.merge_range('E41:F41', "Gross Revenue (Subtotal)", f_th)
    ws_dash.merge_range('G41:H41', "Final Billing Amount", f_th)
    ws_dash.merge_range('I41:J41', "Gross Profit", f_th)
    ws_dash.write('K41', "Margin %", f_th)
    ws_dash.merge_range('L41:M41', "Average Order Value", f_th)

    years = ['2023', '2024', '2025']
    default_annual_sales = [660951323.05, 876378411.05, 912569240.25]
    default_annual_profit = [51981200.08, 70624009.68, 74689137.26]
    default_annual_orders = [26750, 35914, 37336]
    default_annual_units = [156912, 210086, 218642]

    for y_idx, yr in enumerate(years):
        r_num = 42 + y_idx
        fmt_t = f_td_text_alt if y_idx % 2 == 1 else f_td_text
        fmt_n = f_td_num_alt if y_idx % 2 == 1 else f_td_num
        fmt_c = f_td_curr_alt if y_idx % 2 == 1 else f_td_curr
        fmt_p = f_td_pct_alt if y_idx % 2 == 1 else f_td_pct

        ws_dash.write(r_num, 1, f"FY {yr}", fmt_t)
        ws_dash.write_formula(r_num, 2, f"=SUMIFS('Analysis Data'!$F$7:$F$42, 'Analysis Data'!$B$7:$B$42, \"{yr}\")", fmt_n, default_annual_orders[y_idx])
        ws_dash.write_formula(r_num, 3, f"=SUMIFS('Analysis Data'!$H$55:$H$19401, 'Analysis Data'!$A$55:$A$19401, \"{yr}\", 'Analysis Data'!$D$55:$D$19401, 'Analysis Data'!$B$2, 'Analysis Data'!$E$55:$E$19401, 'Analysis Data'!$B$3, 'Analysis Data'!$F$55:$F$19401, 'Analysis Data'!$B$4)", fmt_n, default_annual_units[y_idx])
        ws_dash.merge_range(r_num, 4, r_num, 5, "", fmt_c)
        ws_dash.write_formula(r_num, 4, f"=SUMIFS('Analysis Data'!$D$7:$D$42, 'Analysis Data'!$B$7:$B$42, \"{yr}\")", fmt_c, default_annual_sales[y_idx])
        ws_dash.merge_range(r_num, 6, r_num, 7, "", fmt_c)
        ws_dash.write_formula(r_num, 6, f"=E{r_num+1}", fmt_c, default_annual_sales[y_idx])
        ws_dash.merge_range(r_num, 8, r_num, 9, "", fmt_c)
        ws_dash.write_formula(r_num, 8, f"=SUMIFS('Analysis Data'!$E$7:$E$42, 'Analysis Data'!$B$7:$B$42, \"{yr}\")", fmt_c, default_annual_profit[y_idx])
        ws_dash.write_formula(r_num, 10, f"=IFERROR(I{r_num+1}/E{r_num+1}, 0)", fmt_p, default_annual_profit[y_idx]/default_annual_sales[y_idx])
        ws_dash.merge_range(r_num, 11, r_num, 12, "", fmt_c)
        ws_dash.write_formula(r_num, 11, f"=IFERROR(E{r_num+1}/C{r_num+1}, 0)", fmt_c, default_annual_sales[y_idx]/default_annual_orders[y_idx])

    # Total Row
    ws_dash.write('B46', "Total / All Years", f_tot_lbl)
    ws_dash.write_formula('C46', "=SUM(C43:C45)", f_tot_num, 100000)
    ws_dash.write_formula('D46', "=SUM(D43:D45)", f_tot_num, 585640)
    ws_dash.merge_range('E46:F46', "", f_tot_curr)
    ws_dash.write_formula('E46', "=SUM(E43:E45)", f_tot_curr, 2449898974.35)
    ws_dash.merge_range('G46:H46', "", f_tot_curr)
    ws_dash.write_formula('G46', "=SUM(G43:G45)", f_tot_curr, 2449898974.35)
    ws_dash.merge_range('I46:J46', "", f_tot_curr)
    ws_dash.write_formula('I46', "=SUM(I43:I45)", f_tot_curr, 197294347.02)
    ws_dash.write_formula('K46', "=IFERROR(I46/E46, 0)", f_tot_pct, 0.08053)
    ws_dash.merge_range('L46:M46', "", f_tot_curr)
    ws_dash.write_formula('L46', "=IFERROR(E46/C46, 0)", f_tot_curr, 24498.99)

    # ==========================================
    # SHEET 2: ANALYSIS DATA (CLEAN SUPPORTING CUBES)
    # ==========================================
    print("Building 'Analysis Data'...")
    ws_data = wb.add_worksheet('Analysis Data')
    ws_data.set_tab_color(ROYAL_BLUE)
    ws_data.hide_gridlines(0)
    ws_data.set_column('A:Z', 13.5)
    ws_data.set_column('AA:AE', 13.5)

    # Filter Criteria Cells (Rows 1-4)
    ws_data.write('A1', "Active Filter: Year", f_ctrl_lbl)
    ws_data.write_formula('B1', '=IF(\'Sales Dashboard\'!$C$5="All", "*", \'Sales Dashboard\'!$C$5 & "")', f_ctrl_val, "*")
    ws_data.write('A2', "Active Filter: Category", f_ctrl_lbl)
    ws_data.write_formula('B2', '=IF(\'Sales Dashboard\'!$F$5="All", "*", \'Sales Dashboard\'!$F$5 & "")', f_ctrl_val, "*")
    ws_data.write('A3', "Active Filter: Payment", f_ctrl_lbl)
    ws_data.write_formula('B3', '=IF(\'Sales Dashboard\'!$I$5="All", "*", \'Sales Dashboard\'!$I$5 & "")', f_ctrl_val, "*")
    ws_data.write('A4', "Active Filter: Status", f_ctrl_lbl)
    ws_data.write_formula('B4', '=IF(\'Sales Dashboard\'!$L$5="All", "*", \'Sales Dashboard\'!$L$5 & "")', f_ctrl_val, "*")

    ws_data.merge_range('C2:Z2', "DASHBOARD SUPPORTING ANALYTICAL CUBES & REFERENCE DATA", f_title)
    ws_data.merge_range('C3:Z3', "Dynamic multidimensional models feeding the executive dashboard charts and KPI formulas", f_sub)

    # Table 1: Monthly Master (Cols A to F, Rows 7 to 42)
    ws_data.merge_range('A5:F5', "36-MONTH SALES & PROFIT MASTER", f_th)
    m_headers = ["Year-Month", "Year", "Month", "Sales (INR)", "Profit (INR)", "Orders Count"]
    for c, h in enumerate(m_headers):
        ws_data.write(5, c, h, f_th)
    for r_idx, (_, row) in enumerate(month_list.iterrows()):
        r_num = 7 + r_idx
        ws_data.write(6 + r_idx, 0, str(row['year_month']), f_td_text)
        ws_data.write(6 + r_idx, 1, str(row['year']), f_td_text)
        ws_data.write(6 + r_idx, 2, int(row['month']), f_td_num)
        ws_data.write_formula(6 + r_idx, 3, f"=SUMIFS($I$55:$I$19401, $C$55:$C$19401, $A{r_num}, $A$55:$A$19401, $B$1, $D$55:$D$19401, $B$2, $E$55:$E$19401, $B$3, $F$55:$F$19401, $B$4)", f_td_curr, float(row['sales']))
        ws_data.write_formula(6 + r_idx, 4, f"=SUMIFS($J$55:$J$19401, $C$55:$C$19401, $A{r_num}, $A$55:$A$19401, $B$1, $D$55:$D$19401, $B$2, $E$55:$E$19401, $B$3, $F$55:$F$19401, $B$4)", f_td_curr, float(row['profit']))
        ws_data.write_formula(6 + r_idx, 5, f"=SUMIFS($G$55:$G$19401, $C$55:$C$19401, $A{r_num}, $A$55:$A$19401, $B$1, $D$55:$D$19401, $B$2, $E$55:$E$19401, $B$3, $F$55:$F$19401, $B$4)", f_td_num, int(row['orders']))

    # Table 2: Category Master (Cols H to L, Rows 7 to 16)
    ws_data.merge_range('H5:L5', "TOP 10 CATEGORIES", f_th)
    c_headers = ["Category", "SKUs", "Sales (INR)", "Profit (INR)", "Margin %"]
    for c, h in enumerate(c_headers):
        ws_data.write(5, 7 + c, h, f_th)
    for r_idx, (_, row) in enumerate(cat_default.iterrows()):
        r_num = 7 + r_idx
        ws_data.write(6 + r_idx, 7, str(row['category']), f_td_text)
        ws_data.write(6 + r_idx, 8, int(row['skus']), f_td_num)
        ws_data.write_formula(6 + r_idx, 9, f"=IF(OR($B$2=\"*\", $B$2=$H{r_num}), SUMIFS($I$55:$I$19401, $D$55:$D$19401, $H{r_num}, $A$55:$A$19401, $B$1, $E$55:$E$19401, $B$3, $F$55:$F$19401, $B$4), 0)", f_td_curr, float(row['sales']))
        ws_data.write_formula(6 + r_idx, 10, f"=IF(OR($B$2=\"*\", $B$2=$H{r_num}), SUMIFS($J$55:$J$19401, $D$55:$D$19401, $H{r_num}, $A$55:$A$19401, $B$1, $E$55:$E$19401, $B$3, $F$55:$F$19401, $B$4), 0)", f_td_curr, float(row['profit']))
        ws_data.write_formula(6 + r_idx, 11, f"=IFERROR(K{r_num}/J{r_num}, 0)", f_td_pct, float(row['profit'])/float(row['sales']))

    # Table 3: Payment Mode Master (Cols N to Q, Rows 7 to 12)
    ws_data.merge_range('N5:Q5', "PAYMENT INSTRUMENTS", f_th)
    p_headers = ["Payment Mode", "Transactions", "Amount Paid (INR)", "Share %"]
    for c, h in enumerate(p_headers):
        ws_data.write(5, 13 + c, h, f_th)
    for r_idx, (_, row) in enumerate(pay_default.iterrows()):
        r_num = 7 + r_idx
        ws_data.write(6 + r_idx, 13, str(row['payment_method']), f_td_text)
        ws_data.write_formula(6 + r_idx, 14, f"=IF(OR($B$3=\"*\", $B$3=$N{r_num}), SUMIFS($G$55:$G$19401, $E$55:$E$19401, $N{r_num}, $A$55:$A$19401, $B$1, $D$55:$D$19401, $B$2, $F$55:$F$19401, $B$4), 0)", f_td_num, int(row['orders']))
        ws_data.write_formula(6 + r_idx, 15, f"=IF(OR($B$3=\"*\", $B$3=$N{r_num}), SUMIFS($I$55:$I$19401, $E$55:$E$19401, $N{r_num}, $A$55:$A$19401, $B$1, $D$55:$D$19401, $B$2, $F$55:$F$19401, $B$4), 0)", f_td_curr, float(row['sales']))
        ws_data.write_formula(6 + r_idx, 16, f"=IFERROR(P{r_num}/SUM($P$7:$P$12), 0)", f_td_pct, float(row['sales'])/pay_default['sales'].sum())

    # Table 4: Top 10 Products (Cols S to W, Rows 7 to 16)
    ws_data.merge_range('S5:W5', "TOP 10 PRODUCTS", f_th)
    prod_headers = ["Product ID", "Product Name", "Category", "Sales (INR)", "Units Sold"]
    for c, h in enumerate(prod_headers):
        ws_data.write(5, 18 + c, h, f_th)
    for r_idx in range(10):
        r_num = 7 + r_idx
        rank = r_idx + 1
        def_row = top_overall.iloc[r_idx]
        ws_data.write_formula(6 + r_idx, 18, f'=IFERROR(VLOOKUP(IF($B$2="*","All",$B$2)&"_{rank}", $U$55:$X$284, 2, FALSE), "N/A")', f_td_text, def_row['product_id'])
        ws_data.write_formula(6 + r_idx, 19, f'=IFERROR(VLOOKUP(IF($B$2="*","All",$B$2)&"_{rank}", $U$55:$X$284, 3, FALSE), "N/A")', f_td_text, str(def_row['product_name'])[:28])
        ws_data.write_formula(6 + r_idx, 20, f'=IFERROR(VLOOKUP(IF($B$2="*","All",$B$2)&"_{rank}", $U$55:$X$284, 4, FALSE), "N/A")', f_td_text, def_row['category'])
        ws_data.write_formula(6 + r_idx, 21, f'=IF(S{r_num}="N/A", 0, SUMIFS($AE$55:$AE$14642, $Z$55:$Z$14642, $S{r_num}, $AA$55:$AA$14642, $B$1, $AB$55:$AB$14642, $B$3, $AC$55:$AC$14642, $B$4))', f_td_curr, float(def_row['item_revenue']))
        ws_data.write_formula(6 + r_idx, 22, f'=IF(S{r_num}="N/A", 0, SUMIFS($AD$55:$AD$14642, $Z$55:$Z$14642, $S{r_num}, $AA$55:$AA$14642, $B$1, $AB$55:$AB$14642, $B$3, $AC$55:$AC$14642, $B$4))', f_td_num, 5000)

    # Table 5: Top 10 States (Cols X to AA, Rows 7 to 16)
    ws_data.merge_range('X5:AA5', "TOP 10 STATES", f_th)
    st_headers = ["State", "Orders", "Sales (INR)", "Share %"]
    for c, h in enumerate(st_headers):
        ws_data.write(5, 23 + c, h, f_th)
    for r_idx, (_, row) in enumerate(state_default.iterrows()):
        r_num = 7 + r_idx
        ws_data.write(6 + r_idx, 23, str(row['delivery_state']), f_td_text)
        ws_data.write_formula(6 + r_idx, 24, f"=SUMIFS($S$55:$S$23710, $M$55:$M$23710, $X{r_num}, $N$55:$N$23710, $B$1, $O$55:$O$23710, $B$2, $P$55:$P$23710, $B$3, $Q$55:$Q$23710, $B$4)", f_td_num, int(row['orders']))
        ws_data.write_formula(6 + r_idx, 25, f"=SUMIFS($R$55:$R$23710, $M$55:$M$23710, $X{r_num}, $N$55:$N$23710, $B$1, $O$55:$O$23710, $B$2, $P$55:$P$23710, $B$3, $Q$55:$Q$23710, $B$4)", f_td_curr, float(row['sales']))
        ws_data.write_formula(6 + r_idx, 26, f"=IFERROR(Z{r_num}/SUM($Z$7:$Z$16), 0)", f_td_pct, float(row['sales'])/state_default['sales'].sum())

    # =========================================================================
    # MULTIDIMENSIONAL ANALYTICAL CUBES (FEEDING ALL FORMULAS AT LIGHTNING SPEED)
    # =========================================================================
    ws_data.merge_range('A50:Z50', "MULTIDIMENSIONAL ANALYTICAL CUBES & REFERENCE DATA (ENGINE FOR DYNAMIC DASHBOARD FILTERS)", f_title)
    ws_data.merge_range('A51:Z51', "Underlying aggregated micro-cubes queried in real-time by Excel SUMIFS, VLOOKUP and dynamic array formulas", f_sub)

    # Cube 1 Headers: Monthly Master (Cols A to J, Row 54)
    ws_data.merge_range('A53:J53', "CUBE 1: MONTHLY & CATEGORY CUBE (19,347 ROWS)", f_th)
    c1_headers = ["Year", "Month", "Year-Month", "Category", "Payment Method", "Order Status", "Orders", "Units", "Sales", "Profit"]
    for c, h in enumerate(c1_headers):
        ws_data.write(53, c, h, f_th)

    # Cube 2 Headers: State Cube (Cols M to S, Row 54)
    ws_data.merge_range('M53:S53', "CUBE 2: STATE GEOGRAPHIC CUBE (23,656 ROWS)", f_th)
    c2_headers = ["State", "Year", "Category", "Payment Method", "Order Status", "Sales", "Orders"]
    for c, h in enumerate(c2_headers):
        ws_data.write(53, 12 + c, h, f_th)

    # Cube 3 Headers: Product Catalog Reference (Cols U to X, Row 54)
    ws_data.merge_range('U53:X53', "CUBE 3: PRODUCT CATALOG (230 ROWS)", f_th)
    c3_headers = ["Key", "Product ID", "Product Name", "Category"]
    for c, h in enumerate(c3_headers):
        ws_data.write(53, 20 + c, h, f_th)

    # Cube 4 Headers: Product Sales Cube (Cols Z to AE, Row 54)
    ws_data.merge_range('Z53:AE53', "CUBE 4: PRODUCT SALES CUBE (14,588 ROWS)", f_th)
    c4_headers = ["Product ID", "Year", "Payment Method", "Order Status", "Units", "Sales"]
    for c, h in enumerate(c4_headers):
        ws_data.write(53, 25 + c, h, f_th)

    print("Writing underlying multidimensional cubes to 'Analysis Data'...")
    # Write Cube 1: Monthly (Rows 55 to 19401)
    for r_idx, (_, r) in enumerate(cube_monthly.iterrows()):
        row_target = 54 + r_idx
        ws_data.write(row_target, 0, str(r['year']), f_td_text)
        ws_data.write(row_target, 1, int(r['month']), f_td_num)
        ws_data.write(row_target, 2, str(r['year_month']), f_td_text)
        ws_data.write(row_target, 3, str(r['category']), f_td_text)
        ws_data.write(row_target, 4, str(r['payment_method']), f_td_text)
        ws_data.write(row_target, 5, str(r['order_status']), f_td_text)
        ws_data.write(row_target, 6, int(r['orders']), f_td_num)
        ws_data.write(row_target, 7, int(r['units']), f_td_num)
        ws_data.write(row_target, 8, float(r['sales']), f_td_curr)
        ws_data.write(row_target, 9, float(r['profit']), f_td_curr)

    # Write Cube 2: State (Rows 55 to 23710)
    for r_idx, (_, r) in enumerate(cube_state.iterrows()):
        row_target = 54 + r_idx
        ws_data.write(row_target, 12, str(r['delivery_state']), f_td_text)
        ws_data.write(row_target, 13, str(r['year']), f_td_text)
        ws_data.write(row_target, 14, str(r['category']), f_td_text)
        ws_data.write(row_target, 15, str(r['payment_method']), f_td_text)
        ws_data.write(row_target, 16, str(r['order_status']), f_td_text)
        ws_data.write(row_target, 17, float(r['sales']), f_td_curr)
        ws_data.write(row_target, 18, int(r['orders']), f_td_num)

    # Write Cube 3: Product Catalog (Rows 55 to 284)
    for r_idx, (_, r) in enumerate(prod_catalog_df.iterrows()):
        row_target = 54 + r_idx
        ws_data.write(row_target, 20, str(r['key']), f_td_text)
        ws_data.write(row_target, 21, str(r['product_id']), f_td_text)
        ws_data.write(row_target, 22, str(r['product_name']), f_td_text)
        ws_data.write(row_target, 23, str(r['category']), f_td_text)

    # Write Cube 4: Product Cube (Rows 55 to 14642)
    for r_idx, (_, r) in enumerate(cube_product.iterrows()):
        row_target = 54 + r_idx
        ws_data.write(row_target, 25, str(r['product_id']), f_td_text)
        ws_data.write(row_target, 26, str(r['year']), f_td_text)
        ws_data.write(row_target, 27, str(r['payment_method']), f_td_text)
        ws_data.write(row_target, 28, str(r['order_status']), f_td_text)
        ws_data.write(row_target, 29, int(r['units']), f_td_num)
        ws_data.write(row_target, 30, float(r['sales']), f_td_curr)

    # ==========================================
    # SHEET 3: DATA DICTIONARY & KPIS
    # ==========================================
    print("Building 'Data Dictionary & KPIs'...")
    ws_dict = wb.add_worksheet('Data Dictionary & KPIs')
    ws_dict.set_tab_color(ROYAL_BLUE)
    ws_dict.hide_gridlines(0)
    ws_dict.set_column('A:A', 3)
    ws_dict.set_column('B:B', 22)
    ws_dict.set_column('C:C', 38)
    ws_dict.set_column('D:D', 18)
    ws_dict.set_column('E:E', 48)

    ws_dict.merge_range('B2:E2', "DATA DICTIONARY & KPI FORMULA SPECIFICATION", f_title)
    ws_dict.merge_range('B3:E3', "Key Performance Indicator Definitions & Multidimensional Interactive Filtering Architecture", f_sub)

    ws_dict.merge_range('B5:E5', "CORE DASHBOARD KPI DEFINITIONS & DYNAMIC LOGIC", f_sec_hdr)
    kpi_spec = [
        ("Total Sales (Revenue)", "SUMIFS(Monthly_Cube[Sales], Filter_Criteria)", "₹2,449,898,974.35", "Catalog revenue placed; dynamically filters across Year, Category, Payment Mode, and Order Status."),
        ("Total Profit", "SUMIFS(Monthly_Cube[Profit], Filter_Criteria)", "₹197,294,347.02", "Gross merchant contribution margin; dynamically updates with all active filters."),
        ("Profit Margin %", "Total Profit / Total Sales", "8.05%", "Enterprise gross product profitability percentage calculated dynamically from filtered totals."),
        ("Total Orders", "SUMIFS(Monthly_Cube[Orders], Filter_Criteria)", "100,000", "Total customer purchase transactions recorded under active filter criteria."),
        ("Average Order Value", "Total Sales / Total Orders", "₹24,498.99", "Average gross basket size per purchase transaction."),
        ("Top Category", "INDEX(Category, MATCH(MAX(Category_Sales)))", "Electronics", "Dynamically returns the highest revenue-generating category under active filters."),
        ("Net Delivered Revenue", "Delivered Subtotal - Returns Refund", "₹1,879,006,920.35", "Retained revenue after product return refunds across enterprise history."),
        ("Order Return Rate", "Returned Orders / Placed Orders", "6.68%", "6,681 orders returned out of 100,000 orders placed."),
        ("Repeat Customer Rate", "Customers with >1 Order / Purchasing Customers", "82.87%", "18,912 multi-order buyers out of 22,820 total customer base.")
    ]
    ws_dict.write('B6', "Metric Name", f_th)
    ws_dict.write('C6', "Mathematical Formula", f_th)
    ws_dict.write('D6', "Dataset Value", f_th)
    ws_dict.write('E6', "Business Meaning", f_th)

    for r_idx, row in enumerate(kpi_spec):
        r_num = 7 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        ws_dict.write(6 + r_idx, 1, row[0], fmt_t)
        ws_dict.write(6 + r_idx, 2, row[1], fmt_t)
        ws_dict.write(6 + r_idx, 3, row[2], fmt_t)
        ws_dict.write(6 + r_idx, 4, row[3], fmt_t)

    # Interactive Filter Documentation
    ws_dict.merge_range('B17:E17', "DASHBOARD INTERACTIVE FILTER ARCHITECTURE", f_sec_hdr)
    filter_docs = [
        ("Select Year (C5)", "Controls Year dimension: All, 2023, 2024, 2025", "All KPI cards, Monthly Trend, Annual Scorecard, Category & State breakdown"),
        ("Category (F5)", "Controls Product Category dimension: All + 14 Major Categories", "All KPI cards, Monthly Trend, Top 10 Products by Category, States, Payment"),
        ("Payment Mode (I5)", "Controls Payment Instrument: All, UPI, COD, Cards, Net Banking, Wallet", "All KPI cards, Payment Instrument Share, Top Products, States"),
        ("Order Status (L5)", "Controls Order Lifecycle Status: All, Delivered, Cancelled, Returned, Failed", "All KPI cards, Monthly Sales, Category Sales, State & Product breakdown"),
        ("RESET ALL (M5)", "Click / reset dropdowns to 'All' to restore full enterprise view", "Instantly resets dashboard to the full 3-year ₹2,449.90 Cr baseline view")
    ]
    ws_dict.write('B18', "Filter Control", f_th)
    ws_dict.write('C18', "Supported Values & Scope", f_th)
    ws_dict.merge_range('D18:E18', "Impacted Dashboard Charts & Scorecards", f_th)

    for r_idx, row in enumerate(filter_docs):
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        ws_dict.write(18 + r_idx, 1, row[0], fmt_t)
        ws_dict.write(18 + r_idx, 2, row[1], fmt_t)
        ws_dict.merge_range(18 + r_idx, 3, 18 + r_idx, 4, row[2], fmt_t)

    print("Saving workbook...")
    wb.close()
    print(f"Success! Dynamic workbook saved at: {output_file}")
    print(f"File size: {os.path.getsize(output_file) / 1024:.1f} KB")

if __name__ == '__main__':
    create_simple_dashboard()
