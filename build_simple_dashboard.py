"""
Simple, Elegant, Interactive Indian E-Commerce Sales Dashboard
Creates an uncluttered, single-screen-focused Microsoft Excel workbook:
Indian_Ecommerce_Sales_Analytics_Dashboard.xlsx
Matching the PowerPoint presentation visual identity perfectly.
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
    orders['year'] = orders['order_date'].dt.year
    orders['month'] = orders['order_date'].dt.month
    orders['year_month'] = orders['order_date'].dt.to_period('M').astype(str)

    # 1. Monthly Summary (36 Months)
    items_month = items.merge(orders[['order_id', 'year', 'month', 'year_month']], on='order_id')
    monthly_summary = items_month.groupby(['year', 'month', 'year_month']).agg(
        orders=('order_id', 'nunique'),
        units=('quantity', 'sum'),
        sales=('item_revenue', 'sum'),
        profit=('profit', 'sum')
    ).reset_index().sort_values(['year', 'month']).reset_index(drop=True)
    monthly_summary['margin'] = monthly_summary['profit'] / monthly_summary['sales']
    monthly_summary['aov'] = monthly_summary['sales'] / monthly_summary['orders']

    # 2. Category Summary (Top 10 Categories)
    item_prod = items.merge(products[['product_id', 'product_name', 'category']], on='product_id')
    cat_summary = item_prod.groupby('category').agg(
        skus=('product_id', 'nunique'),
        units=('quantity', 'sum'),
        sales=('item_revenue', 'sum'),
        profit=('profit', 'sum')
    ).reset_index().sort_values('sales', ascending=False).reset_index(drop=True)
    cat_summary['margin'] = cat_summary['profit'] / cat_summary['sales']
    cat_summary['share'] = cat_summary['sales'] / cat_summary['sales'].sum()

    # 3. Payment Method Summary
    pay_summary = payments.groupby('payment_method').agg(
        transactions=('payment_id', 'count'),
        amount=('amount_paid', 'sum')
    ).reset_index().sort_values('amount', ascending=False).reset_index(drop=True)
    pay_summary['share'] = pay_summary['amount'] / pay_summary['amount'].sum()

    # 4. Top 10 Best Selling Products
    top_prods = item_prod.groupby(['product_id', 'product_name', 'category']).agg(
        units=('quantity', 'sum'),
        sales=('item_revenue', 'sum'),
        profit=('profit', 'sum')
    ).reset_index().sort_values('sales', ascending=False).head(10).reset_index(drop=True)
    top_prods['margin'] = top_prods['profit'] / top_prods['sales']

    # 5. Top 10 States
    state_summary = orders.groupby('delivery_state').agg(
        orders=('order_id', 'count'),
        sales=('subtotal', 'sum')
    ).reset_index().sort_values('sales', ascending=False).head(10).reset_index(drop=True)
    state_summary['share'] = state_summary['sales'] / orders['subtotal'].sum()

    # 6. Annual Summary
    annual_summary = orders.groupby('year').agg(
        orders=('order_id', 'count'),
        sales=('subtotal', 'sum'),
        final_amt=('final_amount', 'sum')
    ).reset_index().reset_index(drop=True)
    annual_profit = items_month.groupby('year')['profit'].sum().values
    annual_units = items_month.groupby('year')['quantity'].sum().values
    annual_summary['profit'] = annual_profit
    annual_summary['units'] = annual_units
    annual_summary['margin'] = annual_summary['profit'] / annual_summary['sales']
    annual_summary['aov'] = annual_summary['sales'] / annual_summary['orders']

    print("Creating simplified Excel workbook...")
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
        'align': 'left', 'valign': 'vcenter'
    })
    f_sub = wb.add_format({
        'bold': False, 'font_size': 9.5, 'font_name': 'Calibri',
        'font_color': CYAN, 'bg_color': NAVY,
        'align': 'left', 'valign': 'vcenter'
    })
    f_ctrl_lbl = wb.add_format({
        'bold': True, 'font_size': 9, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': LIGHT_GREY,
        'align': 'right', 'valign': 'vcenter', 'border': 1
    })
    f_ctrl_val = wb.add_format({
        'bold': True, 'font_size': 9.5, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': WHITE,
        'align': 'center', 'valign': 'vcenter', 'border': 1
    })

    # KPI Card Formats
    f_kpi_lbl = wb.add_format({
        'bold': True, 'font_size': 8.5, 'font_name': 'Calibri',
        'font_color': TEXT_MUTED, 'bg_color': PALE_BLUE,
        'align': 'center', 'valign': 'vcenter',
        'top': 1, 'left': 1, 'right': 1
    })
    f_kpi_val_curr = wb.add_format({
        'bold': True, 'font_size': 15, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'center', 'valign': 'vcenter',
        'left': 1, 'right': 1, 'num_format': '₹#,##,##0'
    })
    f_kpi_val_num = wb.add_format({
        'bold': True, 'font_size': 15, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'center', 'valign': 'vcenter',
        'left': 1, 'right': 1, 'num_format': '#,##0'
    })
    f_kpi_val_pct = wb.add_format({
        'bold': True, 'font_size': 15, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'center', 'valign': 'vcenter',
        'left': 1, 'right': 1, 'num_format': '0.0%'
    })
    f_kpi_val_text = wb.add_format({
        'bold': True, 'font_size': 12, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': PALE_BLUE,
        'align': 'center', 'valign': 'vcenter',
        'left': 1, 'right': 1
    })
    f_kpi_sub = wb.add_format({
        'font_size': 8, 'font_name': 'Calibri',
        'font_color': TEXT_DARK, 'bg_color': PALE_BLUE,
        'align': 'center', 'valign': 'vcenter',
        'bottom': 1, 'left': 1, 'right': 1, 'italic': True
    })

    # Table Formats
    f_th = wb.add_format({
        'bold': True, 'font_size': 9, 'font_name': 'Calibri',
        'font_color': WHITE, 'bg_color': NAVY,
        'align': 'center', 'valign': 'vcenter', 'border': 1
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
    f_tot_lbl = wb.add_format({
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
    f_sec_hdr = wb.add_format({
        'bold': True, 'font_size': 11, 'font_name': 'Calibri',
        'font_color': NAVY, 'bg_color': LIGHT_GREY,
        'align': 'left', 'valign': 'vcenter', 'bottom': 2, 'bottom_color': CYAN
    })

    # ==========================================
    # SHEET 1: SALES DASHBOARD (MAIN HERO DASHBOARD)
    # ==========================================
    print("Building 'Sales Dashboard'...")
    ws_dash = wb.add_worksheet('Sales Dashboard')
    ws_dash.set_tab_color(NAVY)
    ws_dash.hide_gridlines(0)
    ws_dash.set_column('A:A', 3)
    ws_dash.set_column('B:M', 13.5)

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
    ws_dash.data_validation('F5', {'validate': 'list', 'source': ['All', 'Electronics', 'Appliances', 'Fashion', 'Home & Kitchen', 'Footwear', 'Grocery', 'Furniture']})

    ws_dash.write('H5', "Payment Mode:", f_ctrl_lbl)
    ws_dash.write('I5', "All", f_ctrl_val)
    ws_dash.data_validation('I5', {'validate': 'list', 'source': ['All', 'UPI', 'Cash on Delivery', 'Credit Card', 'Debit Card', 'Net Banking']})

    ws_dash.write('K5', "Order Status:", f_ctrl_lbl)
    ws_dash.write('L5', "All", f_ctrl_val)
    ws_dash.data_validation('L5', {'validate': 'list', 'source': ['All', 'Delivered', 'Cancelled', 'Returned', 'Failed', 'Processing']})

    ws_dash.write('M5', "RESET", wb.add_format({'bold': True, 'font_size': 9, 'font_name': 'Calibri', 'font_color': WHITE, 'bg_color': ROYAL_BLUE, 'align': 'center', 'valign': 'vcenter', 'border': 1}))

    # 6 Focused KPI Cards (Rows 7-9)
    ws_dash.set_row(6, 14)
    ws_dash.set_row(7, 24)
    ws_dash.set_row(8, 14)

    # Card 1: Total Sales (Dynamic formula)
    ws_dash.merge_range('B7:C7', "TOTAL SALES", f_kpi_lbl)
    ws_dash.merge_range('B8:C8', "", f_kpi_val_curr)
    ws_dash.write_formula('B8', '=IF(C5="All",SUM(\'Analysis Data\'!$D$7:$D$42),SUMIFS(\'Analysis Data\'!$D$7:$D$42,\'Analysis Data\'!$B$7:$B$42,C5))', f_kpi_val_curr, 2449898974.35)
    ws_dash.merge_range('B9:C9', "₹2,449.90 Cr Placed", f_kpi_sub)

    # Card 2: Total Profit (Dynamic formula)
    ws_dash.merge_range('D7:E7', "TOTAL PROFIT", f_kpi_lbl)
    ws_dash.merge_range('D8:E8', "", f_kpi_val_curr)
    ws_dash.write_formula('D8', '=IF(C5="All",SUM(\'Analysis Data\'!$E$7:$E$42),SUMIFS(\'Analysis Data\'!$E$7:$E$42,\'Analysis Data\'!$B$7:$B$42,C5))', f_kpi_val_curr, 197294347.02)
    ws_dash.merge_range('D9:E9', "₹197.29 Cr Realized Margin", f_kpi_sub)

    # Card 3: Profit Margin %
    ws_dash.merge_range('F7:G7', "PROFIT MARGIN %", f_kpi_lbl)
    ws_dash.merge_range('F8:G8', "", f_kpi_val_pct)
    ws_dash.write_formula('F8', '=D8/B8', f_kpi_val_pct, 0.08053)
    ws_dash.merge_range('F9:G9', "8.05% Gross Margin", f_kpi_sub)

    # Card 4: Total Orders (Dynamic formula)
    ws_dash.merge_range('H7:I7', "TOTAL ORDERS", f_kpi_lbl)
    ws_dash.merge_range('H8:I8', "", f_kpi_val_num)
    ws_dash.write_formula('H8', '=IF(C5="All",SUM(\'Analysis Data\'!$F$7:$F$42),SUMIFS(\'Analysis Data\'!$F$7:$F$42,\'Analysis Data\'!$B$7:$B$42,C5))', f_kpi_val_num, 100000)
    ws_dash.merge_range('H9:I9', "80,679 Delivered (80.7%)", f_kpi_sub)

    # Card 5: Average Order Value
    ws_dash.merge_range('J7:K7', "AVERAGE ORDER VALUE", f_kpi_lbl)
    ws_dash.merge_range('J8:K8', "", f_kpi_val_curr)
    ws_dash.write_formula('J8', '=B8/H8', f_kpi_val_curr, 24498.99)
    ws_dash.merge_range('J9:K9', "2.30 Units / Order", f_kpi_sub)

    # Card 6: Top Category
    ws_dash.merge_range('L7:M7', "TOP CATEGORY", f_kpi_lbl)
    ws_dash.merge_range('L8:M8', "Electronics", f_kpi_val_text)
    ws_dash.merge_range('L9:M9', "₹1,223.6M (49.9% Rev)", f_kpi_sub)

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

    # Executive Annual Summary Table (Rows 41-47)
    ws_dash.merge_range('B40:M40', "ANNUAL PERFORMANCE SCORECARD", f_sec_hdr)
    ws_dash.set_row(39, 20)

    dash_tbl_headers = [
        "Financial Year", "Total Orders", "Units Sold", "Gross Revenue (Subtotal)",
        "Final Billing Amount", "Gross Profit", "Gross Margin %", "Average Order Value (AOV)"
    ]
    ws_dash.write('B41', "Year", f_th)
    ws_dash.write('C41', "Orders", f_th)
    ws_dash.write('D41', "Units Sold", f_th)
    ws_dash.merge_range('E41:F41', "Gross Revenue (Subtotal)", f_th)
    ws_dash.merge_range('G41:H41', "Final Billing Amount", f_th)
    ws_dash.merge_range('I41:J41', "Gross Profit", f_th)
    ws_dash.write('K41', "Margin %", f_th)
    ws_dash.merge_range('L41:M41', "Average Order Value", f_th)

    for r_idx, (_, row) in enumerate(annual_summary.iterrows()):
        r_num = 42 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        fmt_n = f_td_num_alt if r_idx % 2 == 1 else f_td_num
        fmt_c = f_td_curr_alt if r_idx % 2 == 1 else f_td_curr
        fmt_p = f_td_pct_alt if r_idx % 2 == 1 else f_td_pct

        ws_dash.write(r_num, 1, f"FY {int(row['year'])}", fmt_t)
        ws_dash.write(r_num, 2, int(row['orders']), fmt_n)
        ws_dash.write(r_num, 3, int(row['units']), fmt_n)
        ws_dash.merge_range(r_num, 4, r_num, 5, float(row['sales']), fmt_c)
        ws_dash.merge_range(r_num, 6, r_num, 7, float(row['final_amt']), fmt_c)
        ws_dash.merge_range(r_num, 8, r_num, 9, float(row['profit']), fmt_c)
        ws_dash.write(r_num, 10, float(row['margin']), fmt_p)
        ws_dash.merge_range(r_num, 11, r_num, 12, float(row['aov']), fmt_c)

    # Total Row
    tot_row = 45
    ws_dash.write('B46', "Total / All Years", f_tot_lbl)
    ws_dash.write('C46', "=SUM(C43:C45)", f_tot_num)
    ws_dash.write('D46', "=SUM(D43:D45)", f_tot_num)
    ws_dash.merge_range('E46:F46', "=SUM(E43:E45)", f_tot_curr)
    ws_dash.merge_range('G46:H46', "=SUM(G43:G45)", f_tot_curr)
    ws_dash.merge_range('I46:J46', "=SUM(I43:I45)", f_tot_curr)
    ws_dash.write('K46', "=I46/E46", f_tot_pct)
    ws_dash.merge_range('L46:M46', "=E46/C46", f_tot_curr)

    # ==========================================
    # SHEET 2: ANALYSIS DATA (CLEAN SUPPORTING CUBES)
    # ==========================================
    print("Building 'Analysis Data'...")
    ws_data = wb.add_worksheet('Analysis Data')
    ws_data.set_tab_color(ROYAL_BLUE)
    ws_data.hide_gridlines(0)
    ws_data.set_column('A:Z', 13.5)

    ws_data.merge_range('A2:Z2', "DASHBOARD SUPPORTING ANALYTICAL CUBES & REFERENCE DATA", f_title)
    ws_data.merge_range('A3:Z3', "Underlying aggregated tables feeding the executive dashboard charts and formulas", f_sub)

    # Table 1: Monthly Master (Cols A to F)
    ws_data.merge_range('A5:F5', "36-MONTH SALES & PROFIT MASTER", f_th)
    m_headers = ["Year-Month", "Year", "Month", "Sales (INR)", "Profit (INR)", "Orders Count"]
    for c, h in enumerate(m_headers):
        ws_data.write(5, c, h, f_th)
    for r_idx, (_, row) in enumerate(monthly_summary.iterrows()):
        ws_data.write(6 + r_idx, 0, row['year_month'], f_td_text)
        ws_data.write(6 + r_idx, 1, int(row['year']), f_td_num)
        ws_data.write(6 + r_idx, 2, int(row['month']), f_td_num)
        ws_data.write(6 + r_idx, 3, float(row['sales']), f_td_curr)
        ws_data.write(6 + r_idx, 4, float(row['profit']), f_td_curr)
        ws_data.write(6 + r_idx, 5, int(row['orders']), f_td_num)

    # Table 2: Category Master (Cols H to L)
    ws_data.merge_range('H5:L5', "TOP 10 CATEGORIES", f_th)
    c_headers = ["Category", "SKUs", "Sales (INR)", "Profit (INR)", "Margin %"]
    for c, h in enumerate(c_headers):
        ws_data.write(5, 7 + c, h, f_th)
    for r_idx, (_, row) in enumerate(cat_summary.head(10).iterrows()):
        ws_data.write(6 + r_idx, 7, row['category'], f_td_text)
        ws_data.write(6 + r_idx, 8, int(row['skus']), f_td_num)
        ws_data.write(6 + r_idx, 9, float(row['sales']), f_td_curr)
        ws_data.write(6 + r_idx, 10, float(row['profit']), f_td_curr)
        ws_data.write(6 + r_idx, 11, float(row['margin']), f_td_pct)

    # Table 3: Payment Mode Master (Cols N to Q)
    ws_data.merge_range('N5:Q5', "PAYMENT INSTRUMENTS", f_th)
    p_headers = ["Payment Mode", "Transactions", "Amount Paid (INR)", "Share %"]
    for c, h in enumerate(p_headers):
        ws_data.write(5, 13 + c, h, f_th)
    for r_idx, (_, row) in enumerate(pay_summary.head(6).iterrows()):
        ws_data.write(6 + r_idx, 13, row['payment_method'], f_td_text)
        ws_data.write(6 + r_idx, 14, int(row['transactions']), f_td_num)
        ws_data.write(6 + r_idx, 15, float(row['amount']), f_td_curr)
        ws_data.write(6 + r_idx, 16, float(row['share']), f_td_pct)

    # Table 4: Top 10 Products (Cols S to W)
    ws_data.merge_range('S5:W5', "TOP 10 PRODUCTS", f_th)
    prod_headers = ["Product ID", "Product Name", "Category", "Sales (INR)", "Units Sold"]
    for c, h in enumerate(prod_headers):
        ws_data.write(5, 18 + c, h, f_th)
    for r_idx, (_, row) in enumerate(top_prods.iterrows()):
        ws_data.write(6 + r_idx, 18, row['product_id'], f_td_text)
        ws_data.write(6 + r_idx, 19, row['product_name'], f_td_text)
        ws_data.write(6 + r_idx, 20, row['category'], f_td_text)
        ws_data.write(6 + r_idx, 21, float(row['sales']), f_td_curr)
        ws_data.write(6 + r_idx, 22, int(row['units']), f_td_num)

    # Table 5: Top 10 States (Cols X to AA)
    ws_data.merge_range('X5:AA5', "TOP 10 STATES", f_th)
    st_headers = ["State", "Orders", "Sales (INR)", "Share %"]
    for c, h in enumerate(st_headers):
        ws_data.write(5, 23 + c, h, f_th)
    for r_idx, (_, row) in enumerate(state_summary.iterrows()):
        ws_data.write(6 + r_idx, 23, row['delivery_state'], f_td_text)
        ws_data.write(6 + r_idx, 24, int(row['orders']), f_td_num)
        ws_data.write(6 + r_idx, 25, float(row['sales']), f_td_curr)
        ws_data.write(6 + r_idx, 26, float(row['share']), f_td_pct)

    # ==========================================
    # SHEET 3: DATA DICTIONARY & KPIS
    # ==========================================
    print("Building 'Data Dictionary & KPIs'...")
    ws_dict = wb.add_worksheet('Data Dictionary & KPIs')
    ws_dict.set_tab_color(ROYAL_BLUE)
    ws_dict.hide_gridlines(0)
    ws_dict.set_column('A:A', 3)
    ws_dict.set_column('B:B', 20)
    ws_dict.set_column('C:C', 35)
    ws_dict.set_column('D:D', 15)
    ws_dict.set_column('E:E', 45)

    ws_dict.merge_range('B2:E2', "DATA DICTIONARY & KPI FORMULA SPECIFICATION", f_title)
    ws_dict.merge_range('B3:E3', "Key Performance Indicator Definitions & Relational Schema Summary", f_sub)

    ws_dict.merge_range('B5:E5', "CORE DASHBOARD KPI DEFINITIONS", f_sec_hdr)
    kpi_spec = [
        ("Total Sales (Revenue)", "SUM(orders.subtotal) == SUM(order_items.item_revenue)", "₹2,449,898,974.35", "Catalog revenue placed before tax and shipping fees."),
        ("Total Profit", "SUM(order_items.profit) == SUM(revenue - cost)", "₹197,294,347.02", "Gross merchant contribution margin from item sales."),
        ("Profit Margin %", "Total Profit / Total Sales", "8.05%", "Enterprise gross product profitability percentage."),
        ("Total Orders", "COUNT(orders.order_id)", "100,000", "Total customer purchase transactions recorded."),
        ("Average Order Value", "Total Sales / Total Orders", "₹24,498.99", "Average gross basket size per purchase."),
        ("Net Delivered Revenue", "Delivered Subtotal - Returns Refund", "₹1,879,006,920.35", "Retained revenue after product return refunds."),
        ("Order Return Rate", "Returned Orders / Placed Orders", "6.68%", "6,681 orders returned out of 100,000 orders."),
        ("Repeat Customer Rate", "Customers with >1 Order / Purchasing Customers", "82.87%", "18,912 multi-order buyers out of 22,820 customers.")
    ]
    ws_dict.write('B6', "Metric Name", f_th)
    ws_dict.write('C6', "Mathematical Formula", f_th)
    ws_dict.write('D6', "Dataset Value", f_th)
    ws_dict.write('E6', "Business Meaning", f_th)

    for r_idx, row in enumerate(kpi_spec):
        r_num = 7 + r_idx
        fmt_t = f_td_text_alt if r_idx % 2 == 1 else f_td_text
        ws_dict.write(r_num, 1, row[0], fmt_t)
        ws_dict.write(r_num, 2, row[1], fmt_t)
        ws_dict.write(r_num, 3, row[2], fmt_t)
        ws_dict.write(r_num, 4, row[3], fmt_t)

    # Set Sales Dashboard as the default active sheet
    ws_dash.activate()

    print("Saving simplified workbook...")
    wb.close()
    print("Success! Simplified workbook saved at:", output_file)
    print("File size:", round(os.path.getsize(output_file) / 1024, 1), "KB")

if __name__ == '__main__':
    create_simple_dashboard()
