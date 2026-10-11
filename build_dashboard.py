import pandas as pd
import numpy as np
import xlsxwriter

SRC = r'C:\Users\onash\Documents\Project\Zomato\Zomato  Order Data.xlsx'
OUT = r'C:\Users\onash\Documents\Project\Zomato\Zomato_Customer_Analytics.xlsx'

# ---------------- Load & clean ----------------
xl = pd.ExcelFile(SRC)
cust = xl.parse('Customer'); cust.columns = cust.columns.str.strip()
rest = xl.parse('Restaurants'); rest.columns = rest.columns.str.strip()
orders = xl.parse('Orders'); orders.columns = orders.columns.str.strip()

cust['City'] = cust['City'].str.strip()
rest['city'] = rest['city'].str.strip()

# Customer C00001 exists in Orders but is missing from the Customer master -> keep as Unknown
if 'C00001' not in set(cust['Customer_id']):
    cust = pd.concat([pd.DataFrame([{'Customer_id': 'C00001', 'Customer_name': 'Unknown',
                                   'City': 'Unknown', 'Signup_Time': pd.NaT,
                                   'Acquisition_channel': 'Unknown'}]), cust], ignore_index=True)

orders['order_date'] = pd.to_datetime(orders['order_timestamp'])
orders['month'] = orders['order_date'].dt.strftime('%b-%y')
orders['net_revenue'] = (orders['order_amount'] - orders['discount_amount']).round(2)
orders['has_coupon'] = orders['discount_amount'] > 0

cmap = cust.set_index('Customer_id')
orders['customer_city'] = orders['customer_id'].map(cmap['City'])
orders['acq_channel'] = orders['customer_id'].map(cmap['Acquisition_channel'])
rmap = rest.set_index('restaurant_id')
orders['cuisine'] = orders['restaurant_id'].map(rmap['cuisine'])
orders['rest_name'] = orders['restaurant_id'].map(rmap['restaurant_name'])
orders['rest_city'] = orders['restaurant_id'].map(rmap['city'])
orders['avg_rating'] = orders['restaurant_id'].map(rmap['avg_rating'])

DATA_END = orders['order_date'].max()          # 2026-04-30
CHURN_DAYS = 90
CHURN_CUTOFF = DATA_END - pd.Timedelta(days=CHURN_DAYS)

delivered = orders[orders['order_status'] == 'Delivered']
cancelled = orders[orders['order_status'] == 'Cancelled']
refunded  = orders[orders['order_status'] == 'Refunded']

M = '"₹"#,##0'
M2 = '"₹"#,##0.00'
P = '0.0%'
N = '#,##0'

# ---------------- Core metrics ----------------
total_orders = len(orders)
total_revenue = round(delivered['net_revenue'].sum(), 2)
gross_gov = round(orders['order_amount'].sum(), 2)
aov = round(total_revenue / len(delivered), 2)
cancel_rate = len(cancelled) / total_orders
refund_rate = len(refunded) / total_orders
delivered_rate = len(delivered) / total_orders
coupon_orders = int(orders['has_coupon'].sum())
coupon_pct = coupon_orders / total_orders
lost_cancel = round(cancelled['net_revenue'].sum(), 2)
lost_refund = round(refunded['net_revenue'].sum(), 2)

# customer level
cust_orders = orders.groupby('customer_id').agg(orders=('order_id', 'count'),
                                                revenue=('net_revenue', 'sum'),
                                                gross=('order_amount', 'sum'),
                                                last_order=('order_date', 'max')).reset_index()
cust_orders = cust_orders.merge(cust[['Customer_id', 'Customer_name', 'City', 'Acquisition_channel']],
                                left_on='customer_id', right_on='Customer_id', how='left')
cust_orders['churned'] = cust_orders['last_order'] < CHURN_CUTOFF
n_churned = int(cust_orders['churned'].sum())
churn_rate = n_churned / len(cust_orders)
revenue_lost_churn = round(cust_orders.loc[cust_orders['churned'], 'revenue'].sum(), 2)
repeat_customers = int((cust_orders['orders'] >= 2).sum())
repeat_rate = repeat_customers / len(cust_orders)

# coupon customer groups
ever_coupon = orders.groupby('customer_id')['has_coupon'].any()
cust_orders['coupon_user'] = cust_orders['customer_id'].map(ever_coupon)

# restaurant level
rest_orders = orders.groupby('restaurant_id').agg(orders=('order_id', 'count'),
                                                  revenue=('net_revenue', 'sum'),
                                                  cancels=('order_status', lambda s: (s == 'Cancelled').sum())).reset_index()
rest_orders = rest_orders.merge(rest, on='restaurant_id', how='left')
rest_orders['cancel_rate'] = rest_orders['cancels'] / rest_orders['orders']

# ---------------- Workbook & formats ----------------
wb = xlsxwriter.Workbook(OUT, {'nan_inf_to_errors': True})
wb.set_properties({'title': 'Zomato Customer Analytics', 'author': 'Data Analyst'})

F = {}
F['title'] = wb.add_format({'bold': True, 'font_size': 18, 'font_color': '#E23744'})
F['section'] = wb.add_format({'bold': True, 'font_size': 12, 'font_color': 'white',
                              'bg_color': '#E23744', 'align': 'left', 'valign': 'vcenter'})
F['q'] = wb.add_format({'italic': True, 'font_size': 10, 'text_wrap': True, 'valign': 'vcenter',
                        'bg_color': '#F7F7F7', 'border': 1, 'border_color': '#DDDDDD'})
F['a'] = wb.add_format({'bold': True, 'font_size': 11, 'valign': 'vcenter',
                        'bg_color': '#FFF3F4', 'border': 1, 'border_color': '#E23744'})
F['kpi_lbl'] = wb.add_format({'font_size': 10, 'font_color': '#555555', 'align': 'center'})
F['kpi_val'] = wb.add_format({'bold': True, 'font_size': 20, 'font_color': '#E23744', 'align': 'center'})
F['kpi_money'] = wb.add_format({'bold': True, 'font_size': 15, 'font_color': '#E23744', 'align': 'center', 'num_format': M})
F['kpi_num'] = wb.add_format({'bold': True, 'font_size': 15, 'font_color': '#E23744', 'align': 'center', 'num_format': N})
F['kpi_pct'] = wb.add_format({'bold': True, 'font_size': 15, 'font_color': '#E23744', 'align': 'center', 'num_format': P})
F['note'] = wb.add_format({'font_size': 9, 'italic': True, 'font_color': '#777777', 'text_wrap': True})
F['txt'] = wb.add_format({'font_size': 10, 'text_wrap': True, 'valign': 'top'})
F['hdr'] = wb.add_format({'bold': True, 'bg_color': '#E23744', 'font_color': 'white', 'border': 1})

TH = 'Table Style Medium 2'

_fmt_cache = {}
_table_counter = [0]

def get_fmt(spec):
    if spec is None:
        return None
    if spec not in _fmt_cache:
        _fmt_cache[spec] = wb.add_format({'num_format': spec})
    return _fmt_cache[spec]

def fmt_list(df, fmts):
    return [{'header': c, 'format': get_fmt(fmts.get(i))} for i, c in enumerate(df.columns)]

def write_table(ws, df, row, col, fmts=None, style=TH):
    data = df.replace({np.nan: None}).values.tolist()
    _table_counter[0] += 1
    ws.add_table(row, col, row + len(df), col + len(df.columns) - 1,
                 {'data': data, 'columns': fmt_list(df, fmts or {}), 'style': style,
                  'name': f'Table{_table_counter[0]}'})
    return (row + 1, col, row + len(df), col)  # data range of first col

def rng(ws_name, r, c):
    return [ws_name, r[0], c, r[2], c]

def col_chart(wb, ws, ws_name, title, ylab, cats, vals, ser_name, color='#E23744',
              anchor=(0, 0), y_fmt=None, width=620, height=330, stacked=False):
    ch = wb.add_chart({'type': 'column'})
    ch.add_series({'name': ser_name, 'categories': cats, 'values': vals,
                   'fill': {'color': color}, 'gap': 60,
                   'data_labels': {'value': False}})
    ch.set_title({'name': title, 'name_font': {'size': 12, 'bold': True}})
    ch.set_x_axis({'num_font': {'size': 8, 'rotation': -45}})
    ch.set_y_axis({'name': ylab, 'num_format': y_fmt or '', 'num_font': {'size': 8},
                   'name_font': {'size': 9}})
    ch.set_legend({'none': True})
    ch.set_chartarea({'border': {'color': '#CCCCCC'}})
    if stacked:
        ch.set_x_axis({'num_font': {'size': 8, 'rotation': -45}})
    ws.insert_chart(anchor[0], anchor[1], ch,
                    {'x_scale': width / 480, 'y_scale': height / 288})
    return ch

def bar_chart(wb, ws, ws_name, title, ylab, cats, vals, ser_name, color='#E23744',
              anchor=(0, 0), y_fmt=None, width=620, height=330):
    ch = wb.add_chart({'type': 'bar'})
    ch.add_series({'name': ser_name, 'categories': cats, 'values': vals,
                   'fill': {'color': color}, 'gap': 60})
    ch.set_title({'name': title, 'name_font': {'size': 12, 'bold': True}})
    ch.set_x_axis({'num_format': y_fmt or '', 'num_font': {'size': 8}})
    ch.set_y_axis({'name': ylab, 'num_font': {'size': 8}, 'name_font': {'size': 9}})
    ch.set_legend({'none': True})
    ws.insert_chart(anchor[0], anchor[1], ch,
                    {'x_scale': width / 480, 'y_scale': height / 288})
    return ch

def pie_chart(wb, ws, ws_name, title, cats, vals, ser_name, anchor=(0, 0),
              width=420, height=330):
    ch = wb.add_chart({'type': 'pie'})
    ch.add_series({'name': ser_name, 'categories': cats, 'values': vals,
                   'data_labels': {'percentage': True, 'font': {'size': 9}},
                   'points': [{'fill': {'color': c}} for c in
                              ['#E23744', '#2E8B57', '#4A90D9', '#F5A623', '#7B61FF', '#999999'][:1]]})
    ch.set_title({'name': title, 'name_font': {'size': 12, 'bold': True}})
    ch.set_legend({'position': 'right', 'font': {'size': 9}})
    ws.insert_chart(anchor[0], anchor[1], ch,
                    {'x_scale': width / 480, 'y_scale': height / 288})
    return ch

def line_chart(wb, ws, ws_name, title, ylab, cats, vals, ser_name, color='#E23744',
               anchor=(0, 0), y_fmt=None, width=760, height=330):
    ch = wb.add_chart({'type': 'line'})
    ch.add_series({'name': ser_name, 'categories': cats, 'values': vals,
                   'line': {'color': color, 'width': 2.25},
                   'marker': {'type': 'circle', 'size': 4, 'fill': {'color': color}}})
    ch.set_title({'name': title, 'name_font': {'size': 12, 'bold': True}})
    ch.set_x_axis({'num_font': {'size': 8, 'rotation': -45}})
    ch.set_y_axis({'name': ylab, 'num_format': y_fmt or '', 'num_font': {'size': 8},
                   'name_font': {'size': 9}})
    ch.set_legend({'none': True})
    ws.insert_chart(anchor[0], anchor[1], ch,
                    {'x_scale': width / 480, 'y_scale': height / 288})
    return ch

def scatter_chart(wb, ws, ws_name, title, xlab, ylab, xs, ys, ser_name,
                  anchor=(0, 0), width=620, height=380, x_fmt=None, y_fmt=None):
    ch = wb.add_chart({'type': 'scatter'})
    ch.add_series({'name': ser_name, 'categories': xs, 'values': ys,
                   'marker': {'type': 'circle', 'size': 5, 'fill': {'color': '#E23744'},
                              'border': {'none': True}},
                   'line': {'none': True},
                   'trendline': {'type': 'linear', 'line': {'color': '#333333',
                                                             'width': 1.5, 'dash_type': 'dash'}}})
    ch.set_title({'name': title, 'name_font': {'size': 12, 'bold': True}})
    ch.set_x_axis({'name': xlab, 'num_format': x_fmt or '0.0', 'num_font': {'size': 8},
                   'name_font': {'size': 9}, 'min': 3.4, 'max': 5.05})
    ch.set_y_axis({'name': ylab, 'num_format': y_fmt or '', 'num_font': {'size': 8},
                   'name_font': {'size': 9}})
    ch.set_legend({'none': True})
    ws.insert_chart(anchor[0], anchor[1], ch,
                    {'x_scale': width / 480, 'y_scale': height / 288})
    return ch

def sheet_head(ws, title, subtitle, ncols=13):
    ws.hide_gridlines(2)
    ws.merge_range(0, 0, 0, ncols - 1, title, F['title'])
    ws.merge_range(1, 0, 1, ncols - 1, subtitle, F['note'])
    ws.set_row(0, 26)
    ws.set_column(0, 6, 15)
    ws.set_column(7, 15, 10)
    return 3

def answers_block(ws, row, pairs, ncols=13):
    ws.merge_range(row, 0, row, ncols - 1, 'KEY ANSWERS', F['section'])
    row += 1
    for q, a in pairs:
        ws.merge_range(row, 0, row, 4, q, F['q'])
        ws.merge_range(row, 5, row, ncols - 1, a, F['a'])
        ws.set_row(row, 30)
        row += 1
    return row + 1

def kpi_row(ws, row, kpis, col=0):
    kpi_fmt = {M: F['kpi_money'], N: F['kpi_num'], P: F['kpi_pct']}
    for i, (lbl, val, fmt) in enumerate(kpis):
        c = col + i * 3
        ws.merge_range(row, c, row, c + 1, lbl, F['kpi_lbl'])
        ws.merge_range(row + 1, c, row + 1, c + 1, val, kpi_fmt.get(fmt, F['kpi_val']))
        ws.set_row(row + 1, 28)
    return row + 3

def section(ws, row, text, ncols=13):
    ws.merge_range(row, 0, row, ncols - 1, text, F['section'])
    return row + 1

# ============================================================
# SHEET: Dashboard
# ============================================================
ws = wb.add_worksheet('Dashboard')
ws.set_tab_color('#E23744')
r = sheet_head(ws, 'ZOMATO — CUSTOMER ANALYTICS DASHBOARD',
               f'Analysis period: Jan-24 to Apr-26  |  Source: Zomato Order Data (Customer, Restaurants, Orders)')
r = kpi_row(ws, r, [
    ('TOTAL REVENUE', total_revenue, M),
    ('TOTAL ORDERS', total_orders, N),
    ('AVG ORDER VALUE', aov, M),
    ('CANCELLATION RATE', cancel_rate, P),
])
r = kpi_row(ws, r, [
    ('REFUND RATE', refund_rate, P),
    ('COUPON USAGE', coupon_pct, P),
    ('CHURN RATE (90d)', churn_rate, P),
    ('REPEAT CUSTOMERS', repeat_rate, P),
])
r = section(ws, r, 'MONTHLY REVENUE TREND (Delivered orders, net of discounts)')
m = delivered.groupby('month', sort=False).agg(orders=('order_id', 'count'),
                                             revenue=('net_revenue', 'sum')).reset_index()
mo = pd.DataFrame({'month': [f'{d:%b-%y}' for d in pd.period_range('2024-01', '2026-04', freq='M').to_timestamp()],
                   'rev': [0.0] * 28, 'ord': [0] * 28})
mo['month'] = mo['month'].astype(str)
mm = m.copy(); mm['month'] = mm['month'].astype(str)
mo = mo.merge(mm, on='month', how='left').fillna({'revenue': 0, 'orders': 0})
mo['revenue'] = mo['revenue'].round(2)
mo = mo.rename(columns={'orders': 'Orders', 'revenue': 'Revenue'})
mo_rng = write_table(ws, mo, r, 0, {0: None, 1: N, 2: M})
line_chart(wb, ws, 'Dashboard', 'Monthly Revenue Trend', 'Revenue (₹)',
           rng('Dashboard', mo_rng, 0), rng('Dashboard', mo_rng, 2), 'Revenue',
           anchor=(r, 8), y_fmt=M)
r += max(len(mo) + 2, 22)

r = section(ws, r, 'REVENUE BY CITY')
ct = delivered.merge(cust[['Customer_id', 'City']], left_on='customer_id', right_on='Customer_id') \
             .groupby('City').agg(orders=('order_id', 'count'), revenue=('net_revenue', 'sum')) \
             .reset_index().sort_values('revenue', ascending=False)
ct['share'] = ct['revenue'] / ct['revenue'].sum()
ct['aov'] = ct['revenue'] / ct['orders']
ct = ct.round({'revenue': 2, 'share': 4, 'aov': 2})
ct_rng = write_table(ws, ct, r, 0, {0: None, 1: N, 2: M, 3: P, 4: M})
col_chart(wb, ws, 'Dashboard', 'Revenue by City', 'Revenue (₹)',
          rng('Dashboard', ct_rng, 0), rng('Dashboard', ct_rng, 2), 'Revenue',
          anchor=(r, 8), y_fmt=M)
r += max(len(ct) + 2, 22)

r = section(ws, r, 'ORDER STATUS MIX')
st = orders.groupby('order_status').agg(orders=('order_id', 'count')).reset_index()
st['pct'] = st['orders'] / total_orders
st = st.round({'pct': 4})
st_rng = write_table(ws, st, r, 0, {0: None, 1: N, 2: P})
pie_chart(wb, ws, 'Dashboard', 'Order Status Mix',
          rng('Dashboard', st_rng, 0), rng('Dashboard', st_rng, 1), 'Orders',
          anchor=(r, 8))
r += max(len(st) + 2, 20)

r = section(ws, r, 'CUSTOMER CHURN (no order in last 90 days)')
ch = cust_orders.groupby('churned').agg(customers=('customer_id', 'count')).reset_index()
ch['churned'] = ch['churned'].map({True: 'Churned', False: 'Active'})
ch['pct'] = ch['customers'] / len(cust_orders)
ch = ch.round({'pct': 4})
ch_rng = write_table(ws, ch, r, 0, {0: None, 1: N, 2: P})
pie_chart(wb, ws, 'Dashboard', 'Active vs Churned Customers',
          rng('Dashboard', ch_rng, 0), rng('Dashboard', ch_rng, 1), 'Customers',
          anchor=(r, 8))

# ============================================================
# SHEET: Revenue Analysis
# ============================================================
ws = wb.add_worksheet('Revenue Analysis')
ws.set_tab_color('#E23744')
r = sheet_head(ws, 'REVENUE ANALYSIS',
               'Revenue = order_amount − discount_amount, counted for Delivered orders only.')
top_city = ct.iloc[0]
top_pay = delivered.groupby('payment_mode')['net_revenue'].sum().idxmax()
top_pay_rev = delivered.groupby('payment_mode')['net_revenue'].sum().max()
peak_mo = mo.loc[mo['Revenue'].idxmax()]
r = answers_block(ws, r, [
    ('Q1. What is the total revenue?',
     f'₹{total_revenue:,.0f} net revenue from {len(delivered):,} delivered orders '
     f'(gross order value across all {total_orders:,} orders: ₹{gross_gov:,.0f})'),
    ('Q2. What is the monthly revenue trend?',
     f'Peak month: {peak_mo["month"]} with ₹{peak_mo["Revenue"]:,.0f}. Full trend in the chart below.'),
    ('Q3. Which city contributes the highest revenue?',
     f'{top_city["City"]} — ₹{top_city["revenue"]:,.0f} ({top_city["share"]:.1%} of total revenue)'),
    ('Q4. Which payment mode generates the most revenue?',
     f'{top_pay} — ₹{top_pay_rev:,.0f} ({top_pay_rev / total_revenue:.1%} of revenue)'),
    ('Q5. What is the Average Order Value (AOV)?',
     f'₹{aov:,.0f} per delivered order (net of discounts)'),
])
r = kpi_row(ws, r, [('TOTAL REVENUE', total_revenue, M),
                    ('GROSS ORDER VALUE', gross_gov, M),
                    ('DELIVERED ORDERS', len(delivered), N),
                    ('AOV', aov, M)])

r = section(ws, r, 'Q2. MONTHLY REVENUE TREND')
write_table(ws, mo, r, 0, {0: None, 1: N, 2: M})
line_chart(wb, ws, 'Revenue Analysis', 'Monthly Revenue Trend', 'Revenue (₹)',
           rng('Revenue Analysis', (r + 1, 0, r + len(mo), 0), 0),
           rng('Revenue Analysis', (r + 1, 2, r + len(mo), 2), 2), 'Revenue',
           anchor=(r, 8), y_fmt=M)
r += max(len(mo) + 2, 22)

r = section(ws, r, 'Q3. REVENUE BY CITY')
write_table(ws, ct, r, 0, {0: None, 1: N, 2: M, 3: P, 4: M})
col_chart(wb, ws, 'Revenue Analysis', 'Revenue by City', 'Revenue (₹)',
          rng('Revenue Analysis', (r + 1, 0, r + len(ct), 0), 0),
          rng('Revenue Analysis', (r + 1, 2, r + len(ct), 2), 2), 'Revenue',
          anchor=(r, 8), y_fmt=M)
r += max(len(ct) + 2, 22)

r = section(ws, r, 'Q4 & Q5. REVENUE BY PAYMENT MODE')
pm = delivered.groupby('payment_mode').agg(orders=('order_id', 'count'),
                                          revenue=('net_revenue', 'sum')).reset_index()
pm = pm.sort_values('revenue', ascending=False)
pm['share'] = pm['revenue'] / pm['revenue'].sum()
pm['aov'] = pm['revenue'] / pm['orders']
pm = pm.round({'revenue': 2, 'share': 4, 'aov': 2})
pm_rng = write_table(ws, pm, r, 0, {0: None, 1: N, 2: M, 3: P, 4: M})
pie_chart(wb, ws, 'Revenue Analysis', 'Revenue Share by Payment Mode',
          rng('Revenue Analysis', pm_rng, 0), rng('Revenue Analysis', pm_rng, 2), 'Revenue',
          anchor=(r, 8))

# ============================================================
# SHEET: Customer Analysis
# ============================================================
ws = wb.add_worksheet('Customer Analysis')
ws.set_tab_color('#E23744')
r = sheet_head(ws, 'CUSTOMER ANALYSIS',
               'Customer revenue = net revenue of delivered orders placed by the customer.')
top20 = cust_orders.sort_values('revenue', ascending=False).head(20).reset_index(drop=True)
top20_rev = top20['revenue'].sum()
top20_share = top20_rev / total_revenue
chan = cust_orders.groupby('Acquisition_channel').agg(
    customers=('customer_id', 'count'), revenue=('revenue', 'sum'),
    orders=('orders', 'sum')).reset_index()
chan['avg_rev_per_cust'] = chan['revenue'] / chan['customers']
chan = chan.sort_values('avg_rev_per_cust', ascending=False).round(
    {'revenue': 2, 'avg_rev_per_cust': 2})
top_chan = chan.iloc[0]
freq = cust_orders['orders']
bins = [0, 1, 3, 6, 9, 12, np.inf]
labels = ['1 order', '2–3 orders', '4–6 orders', '7–9 orders', '10–12 orders', '13+ orders']
freq_dist = pd.DataFrame({'bucket': labels,
                          'customers': pd.cut(freq, bins=bins, labels=labels, right=True).value_counts().reindex(labels).values})
r = answers_block(ws, r, [
    ('Q1. Who are the top 20 customers by revenue?',
     f'Top customer: {top20.iloc[0]["Customer_name"]} ({top20.iloc[0]["customer_id"]}, {top20.iloc[0]["City"]}) '
     f'— ₹{top20.iloc[0]["revenue"]:,.0f} across {int(top20.iloc[0]["orders"])} orders. Full list below.'),
    ('Q2. What percentage of revenue comes from top customers?',
     f'Top 20 customers contribute ₹{top20_rev:,.0f} = {top20_share:.1%} of total revenue '
     f'(they are just {20/len(cust_orders):.1%} of the customer base)'),
    ('Q3. Which acquisition channel brings the highest-value customers?',
     f'{top_chan["Acquisition_channel"]} — avg ₹{top_chan["avg_rev_per_cust"]:,.0f} revenue per customer '
     f'({len(chan)} channels compared; see table below)'),
    ('Q4. How many repeat customers do we have?',
     f'{repeat_customers:,} of {len(cust_orders):,} customers placed 2+ orders ({repeat_rate:.1%} repeat rate)'),
])
r = kpi_row(ws, r, [('TOTAL CUSTOMERS', len(cust_orders), N),
                    ('REPEAT CUSTOMERS', repeat_customers, N),
                    ('REPEAT RATE', repeat_rate, P),
                    ('AVG REVENUE / CUSTOMER', total_revenue / len(cust_orders), M)])

r = section(ws, r, 'Q1. TOP 20 CUSTOMERS BY REVENUE')
t20 = top20.copy()
t20['rank'] = range(1, 21)
t20['share'] = t20['revenue'] / total_revenue
t20 = t20[['rank', 'customer_id', 'Customer_name', 'City', 'orders', 'revenue', 'share']]
t20.columns = ['Rank', 'Customer ID', 'Name', 'City', 'Orders', 'Revenue', '% of Revenue']
t20 = t20.round({'Revenue': 2, '% of Revenue': 4})
t20_rng = write_table(ws, t20, r, 0, {0: N, 1: None, 2: None, 3: None, 4: N, 5: M, 6: P})
col_chart(wb, ws, 'Customer Analysis', 'Top 20 Customers by Revenue', 'Revenue (₹)',
          rng('Customer Analysis', (t20_rng[0], 2, t20_rng[2], 2), 2),
          rng('Customer Analysis', (t20_rng[0], 5, t20_rng[2], 5), 5), 'Revenue',
          anchor=(r, 8), y_fmt=M)
r += max(len(t20) + 2, 24)

r = section(ws, r, 'Q2. REVENUE CONCENTRATION — TOP 20 VS REST')
conc = pd.DataFrame({'segment': ['Top 20 customers', 'All other customers'],
                     'customers': [20, len(cust_orders) - 20],
                     'revenue': [round(top20_rev, 2), round(total_revenue - top20_rev, 2)]})
conc['pct'] = conc['revenue'] / conc['revenue'].sum()
conc = conc.round({'pct': 4})
conc_rng = write_table(ws, conc, r, 0, {0: None, 1: N, 2: M, 3: P})
pie_chart(wb, ws, 'Customer Analysis', 'Revenue: Top 20 vs Rest',
          rng('Customer Analysis', conc_rng, 0), rng('Customer Analysis', conc_rng, 2), 'Revenue',
          anchor=(r, 8))
r += max(len(conc) + 2, 20)

r = section(ws, r, 'Q3. CUSTOMER VALUE BY ACQUISITION CHANNEL')
chan2 = chan.copy()
ch_rep = cust_orders.groupby('Acquisition_channel').apply(
    lambda d: (d['orders'] >= 2).mean(), include_groups=False).reset_index()
ch_rep.columns = ['Acquisition_channel', 'repeat_rate']
chan2 = chan2.merge(ch_rep, on='Acquisition_channel')
chan2 = chan2[['Acquisition_channel', 'customers', 'orders', 'revenue',
               'avg_rev_per_cust', 'repeat_rate']].round({'repeat_rate': 4})
chan2.columns = ['Channel', 'Customers', 'Orders', 'Revenue', 'Avg Revenue / Customer', 'Repeat Rate']
chan_rng = write_table(ws, chan2, r, 0, {0: None, 1: N, 2: N, 3: M, 4: M, 5: P})
col_chart(wb, ws, 'Customer Analysis', 'Avg Revenue per Customer by Channel', '₹ / customer',
          rng('Customer Analysis', (chan_rng[0], 0, chan_rng[2], 0), 0),
          rng('Customer Analysis', (chan_rng[0], 4, chan_rng[2], 4), 4), 'Avg Revenue / Customer',
          anchor=(r, 8), y_fmt=M)
r += max(len(chan2) + 2, 20)

r = section(ws, r, 'Q4. CUSTOMER ORDER-FREQUENCY DISTRIBUTION')
fd = freq_dist.copy()
fd['pct'] = fd['customers'] / fd['customers'].sum()
fd = fd.round({'pct': 4})
fd_rng = write_table(ws, fd, r, 0, {0: None, 1: N, 2: P})
col_chart(wb, ws, 'Customer Analysis', 'Customers by Order Frequency', 'Customers',
          rng('Customer Analysis', (fd_rng[0], 0, fd_rng[2], 0), 0),
          rng('Customer Analysis', (fd_rng[0], 1, fd_rng[2], 1), 1), 'Customers',
          anchor=(r, 8), y_fmt=N)

# ============================================================
# SHEET: Restaurant Performance
# ============================================================
ws = wb.add_worksheet('Restaurant Performance')
ws.set_tab_color('#E23744')
r = sheet_head(ws, 'RESTAURANT PERFORMANCE',
               'Restaurant revenue/order counts are based on all orders; ratings are from the Restaurants master.')
top10_rev = rest_orders.sort_values('revenue', ascending=False).head(10)
top10_ord = rest_orders.sort_values('orders', ascending=False).head(10)
cui = orders.groupby('cuisine').agg(restaurants=('restaurant_id', 'nunique'),
                                     orders=('order_id', 'count'),
                                     revenue=('net_revenue', 'sum')).reset_index()
cui = cui.sort_values('orders', ascending=False)
cui['pct_orders'] = cui['orders'] / total_orders
cui = cui.round({'revenue': 2, 'pct_orders': 4})
top_cui = cui.iloc[0]
# rating buckets
rest_orders['rating_bucket'] = pd.cut(rest_orders['avg_rating'],
                                      bins=[3.4, 3.99, 4.49, 5.01],
                                      labels=['3.5 – 3.9', '4.0 – 4.4', '4.5 – 5.0'])
rb = rest_orders.groupby('rating_bucket', observed=True).agg(
    restaurants=('restaurant_id', 'count'), revenue=('revenue', 'sum'),
    orders=('orders', 'sum')).reset_index()
rb['avg_rev_per_rest'] = rb['revenue'] / rb['restaurants']
rb['avg_rating'] = rest_orders.groupby('rating_bucket', observed=True)['avg_rating'].mean().values
rb = rb.round({'revenue': 2, 'avg_rev_per_rest': 2, 'avg_rating': 2})
hi_rev = rb[rb['rating_bucket'] == '4.5 – 5.0']['avg_rev_per_rest'].iloc[0]
lo_rev = rb[rb['rating_bucket'] == '3.5 – 3.9']['avg_rev_per_rest'].iloc[0]
bot5 = rest_orders.sort_values('revenue').head(5)
r = answers_block(ws, r, [
    ('Q1. Which restaurants generate the highest revenue?',
     f'#{1}: {top10_rev.iloc[0]["restaurant_name"]} ({top10_rev.iloc[0]["restaurant_id"]}, {top10_rev.iloc[0]["city"]}) '
     f'— ₹{top10_rev.iloc[0]["revenue"]:,.0f}. Top 10 listed below.'),
    ('Q2. Which restaurants receive the most orders?',
     f'#{1}: {top10_ord.iloc[0]["restaurant_name"]} ({top10_ord.iloc[0]["restaurant_id"]}) '
     f'— {int(top10_ord.iloc[0]["orders"]):,} orders. Top 10 listed below.'),
    ('Q3. Which cuisines are most popular?',
     f'{top_cui["cuisine"]} — {int(top_cui["orders"]):,} orders ({top_cui["pct_orders"]:.1%} of all orders), '
     f'₹{top_cui["revenue"]:,.0f} revenue'),
    ('Q4. Do highly-rated restaurants generate more revenue?',
     f'Yes — restaurants rated 4.5–5.0 average ₹{hi_rev:,.0f} revenue each vs ₹{lo_rev:,.0f} for 3.5–3.9 rated '
     f'({(hi_rev/lo_rev - 1):+.0%} higher). See scatter & bucket table below.'),
    ('Q5. Last 5 (bottom) restaurants by revenue',
     f'Weakest performer: {bot5.iloc[0]["restaurant_name"]} ({bot5.iloc[0]["restaurant_id"]}) — '
     f'₹{bot5.iloc[0]["revenue"]:,.0f} from {int(bot5.iloc[0]["orders"])} orders'),
])
r = kpi_row(ws, r, [('RESTAURANTS', len(rest), N),
                    ('TOTAL ORDERS', total_orders, N),
                    ('AVG ORDERS / RESTAURANT', total_orders / len(rest), N),
                    ('AVG REVENUE / RESTAURANT', total_revenue / len(rest), M)])

r = section(ws, r, 'Q1. TOP 10 RESTAURANTS BY REVENUE')
def rest_table(df, rank_col=True):
    d = df.copy().reset_index(drop=True)
    d['rank'] = range(1, len(d) + 1)
    d = d[['rank', 'restaurant_id', 'restaurant_name', 'city', 'cuisine', 'avg_rating', 'orders', 'revenue']]
    d.columns = ['Rank', 'ID', 'Restaurant', 'City', 'Cuisine', 'Rating', 'Orders', 'Revenue']
    return d.round({'Revenue': 2, 'Rating': 1})
t1 = rest_table(top10_rev)
t1_rng = write_table(ws, t1, r, 0, {0: N, 1: None, 2: None, 3: None, 4: None, 5: '0.0', 6: N, 7: M})
col_chart(wb, ws, 'Restaurant Performance', 'Top 10 Restaurants by Revenue', 'Revenue (₹)',
          rng('Restaurant Performance', (t1_rng[0], 2, t1_rng[2], 2), 2),
          rng('Restaurant Performance', (t1_rng[0], 7, t1_rng[2], 7), 7), 'Revenue',
          anchor=(r, 8), y_fmt=M)
r += max(len(t1) + 2, 20)

r = section(ws, r, 'Q2. TOP 10 RESTAURANTS BY ORDER COUNT')
t2 = rest_table(top10_ord)
t2_rng = write_table(ws, t2, r, 0, {0: N, 1: None, 2: None, 3: None, 4: None, 5: '0.0', 6: N, 7: M})
col_chart(wb, ws, 'Restaurant Performance', 'Top 10 Restaurants by Orders', 'Orders',
          rng('Restaurant Performance', (t2_rng[0], 2, t2_rng[2], 2), 2),
          rng('Restaurant Performance', (t2_rng[0], 6, t2_rng[2], 6), 6), 'Orders',
          anchor=(r, 8), y_fmt=N)
r += max(len(t2) + 2, 20)

r = section(ws, r, 'Q3. CUISINE POPULARITY')
cui2 = cui.copy()
cui2.columns = ['Cuisine', 'Restaurants', 'Orders', 'Revenue', '% of Orders']
cui_rng = write_table(ws, cui2, r, 0, {0: None, 1: N, 2: N, 3: M, 4: P})
col_chart(wb, ws, 'Restaurant Performance', 'Orders by Cuisine', 'Orders',
          rng('Restaurant Performance', (cui_rng[0], 0, cui_rng[2], 0), 0),
          rng('Restaurant Performance', (cui_rng[0], 2, cui_rng[2], 2), 2), 'Orders',
          anchor=(r, 8), y_fmt=N)
r += max(len(cui2) + 2, 20)

r = section(ws, r, 'Q4. RATING vs REVENUE (per restaurant)')
sc = rest_orders[['avg_rating', 'revenue']].rename(columns={'avg_rating': 'Rating', 'revenue': 'Revenue'})
sc_rng = write_table(ws, sc.round(2), r, 0, {0: '0.0', 1: M})
scatter_chart(wb, ws, 'Restaurant Performance', 'Restaurant Rating vs Revenue',
              'Avg Rating', 'Revenue (₹)',
              rng('Restaurant Performance', (sc_rng[0], 0, sc_rng[2], 0), 0),
              rng('Restaurant Performance', (sc_rng[0], 1, sc_rng[2], 1), 1),
              'Restaurants', anchor=(r, 8), y_fmt=M)
r += max(len(sc) + 2, 22)

r = section(ws, r, 'Q4 (contd.). REVENUE BY RATING BUCKET')
rb2 = rb.copy()
rb2.columns = ['Rating Bucket', 'Restaurants', 'Revenue', 'Orders', 'Avg Revenue / Restaurant', 'Avg Rating']
rb2 = rb2[['Rating Bucket', 'Restaurants', 'Avg Rating', 'Revenue', 'Orders', 'Avg Revenue / Restaurant']]
rb_rng = write_table(ws, rb2, r, 0, {0: None, 1: N, 2: '0.00', 3: M, 4: N, 5: M})
col_chart(wb, ws, 'Restaurant Performance', 'Avg Revenue per Restaurant by Rating Bucket', '₹ / restaurant',
          rng('Restaurant Performance', (rb_rng[0], 0, rb_rng[2], 0), 0),
          rng('Restaurant Performance', (rb_rng[0], 5, rb_rng[2], 5), 5), 'Avg Revenue / Restaurant',
          anchor=(r, 8), y_fmt=M)
r += max(len(rb2) + 2, 20)

r = section(ws, r, 'Q5. BOTTOM 5 RESTAURANTS BY REVENUE (UNDERPERFORMERS)')
t5 = rest_table(bot5)
t5_rng = write_table(ws, t5, r, 0, {0: N, 1: None, 2: None, 3: None, 4: None, 5: '0.0', 6: N, 7: M})
col_chart(wb, ws, 'Restaurant Performance', 'Bottom 5 Restaurants by Revenue', 'Revenue (₹)',
          rng('Restaurant Performance', (t5_rng[0], 2, t5_rng[2], 2), 2),
          rng('Restaurant Performance', (t5_rng[0], 7, t5_rng[2], 7), 7), 'Revenue',
          anchor=(r, 8), y_fmt=M)

# ============================================================
# SHEET: Coupon Analysis
# ============================================================
ws = wb.add_worksheet('Coupon Analysis')
ws.set_tab_color('#E23744')
r = sheet_head(ws, 'COUPON ANALYSIS',
               'A coupon order = any order with discount_amount > 0 (≈10% off the order amount).')
cg = orders.groupby('has_coupon').agg(orders=('order_id', 'count'),
                                       gross=('order_amount', 'sum'),
                                       revenue=('net_revenue', 'sum')).reset_index()
cg['avg_gross'] = cg['gross'] / cg['orders']
cg['avg_net'] = cg['revenue'] / cg['orders']
cg['avg_disc'] = (cg['gross'] - cg['revenue']) / cg['orders']
c_orders = cg.loc[cg['has_coupon'] == True].iloc[0]
nc_orders = cg.loc[cg['has_coupon'] == False].iloc[0]
coupon_city = orders.groupby('customer_city').agg(total=('order_id', 'count'),
                                                  cp=('has_coupon', 'sum')).reset_index()
coupon_city['usage'] = coupon_city['cp'] / coupon_city['total']
coupon_city = coupon_city.sort_values('usage', ascending=False)
coupon_city = coupon_city.round({'usage': 4})
top_cp_city = coupon_city.iloc[0]
rep_cp = cust_orders.groupby('coupon_user').agg(customers=('customer_id', 'count'),
                                                repeat=('orders', lambda s: (s >= 2).sum()),
                                                avg_orders=('orders', 'mean')).reset_index()
rep_cp['repeat_rate'] = rep_cp['repeat'] / rep_cp['customers']
rep_cp = rep_cp.round({'repeat_rate': 4, 'avg_orders': 2})
rep_yes = rep_cp.loc[rep_cp['coupon_user'] == True].iloc[0]
rep_no = rep_cp.loc[rep_cp['coupon_user'] == False].iloc[0]
r = answers_block(ws, r, [
    ('Q1. What percentage of orders use coupons?',
     f'{coupon_pct:.1%} of orders ({coupon_orders:,} of {total_orders:,}) use a coupon'),
    ('Q2. Do coupon users spend more than non-coupon users?',
     f'Yes — avg basket ₹{c_orders["avg_gross"]:,.0f} (coupon) vs ₹{nc_orders["avg_gross"]:,.0f} (no coupon); '
     f'after discount, net spend ₹{c_orders["avg_net"]:,.0f} vs ₹{nc_orders["avg_net"]:,.0f}'),
    ('Q3. Which city uses the most coupons?',
     f'{top_cp_city["customer_city"]} — {top_cp_city["usage"]:.1%} of its orders use coupons '
     f'({int(top_cp_city["cp"]):,} coupon orders)'),
    ('Q4. Are coupons helping customer retention?',
     f'Yes — {rep_yes["repeat_rate"]:.1%} of coupon users are repeat customers vs {rep_no["repeat_rate"]:.1%} '
     f'of non-users; coupon users also average {rep_yes["avg_orders"]:.1f} orders vs {rep_no["avg_orders"]:.1f}'),
])
r = kpi_row(ws, r, [('COUPON ORDERS', coupon_orders, N),
                    ('COUPON USAGE RATE', coupon_pct, P),
                    ('DISCOUNT VALUE GIVEN', round(orders['discount_amount'].sum(), 2), M),
                    ('AVG DISCOUNT / COUPON ORDER', c_orders['avg_disc'], M)])

r = section(ws, r, 'Q1. COUPON USAGE — ORDER SPLIT')
cu = pd.DataFrame({'group': ['Coupon orders', 'Non-coupon orders'],
                   'orders': [int(c_orders['orders']), int(nc_orders['orders'])]})
cu['pct'] = cu['orders'] / total_orders
cu = cu.round({'pct': 4})
cu_rng = write_table(ws, cu, r, 0, {0: None, 1: N, 2: P})
pie_chart(wb, ws, 'Coupon Analysis', 'Share of Orders Using Coupons',
          rng('Coupon Analysis', cu_rng, 0), rng('Coupon Analysis', cu_rng, 1), 'Orders',
          anchor=(r, 8))
r += max(len(cu) + 2, 20)

r = section(ws, r, 'Q2. SPEND: COUPON vs NON-COUPON ORDERS')
sp = cg.copy()
sp['group'] = sp['has_coupon'].map({True: 'Coupon users', False: 'Non-coupon users'})
sp = sp[['group', 'orders', 'avg_gross', 'avg_disc', 'avg_net', 'revenue']]
sp.columns = ['Group', 'Orders', 'Avg Gross Basket', 'Avg Discount', 'Avg Net Spend', 'Total Net Revenue']
sp = sp.round(2)
sp_rng = write_table(ws, sp, r, 0, {0: None, 1: N, 2: M, 3: M, 4: M, 5: M})
col_chart(wb, ws, 'Coupon Analysis', 'Average Spend per Order', '₹ / order',
          rng('Coupon Analysis', (sp_rng[0], 0, sp_rng[2], 0), 0),
          rng('Coupon Analysis', (sp_rng[0], 4, sp_rng[2], 4), 4), 'Avg Net Spend',
          anchor=(r, 8), y_fmt=M)
r += max(len(sp) + 2, 20)

r = section(ws, r, 'Q3. COUPON USAGE BY CITY')
cc = coupon_city.copy()
cc.columns = ['City', 'Coupon Orders', 'Total Orders', 'Coupon Usage %']
cc_rng = write_table(ws, cc, r, 0, {0: None, 1: N, 2: N, 3: P})
col_chart(wb, ws, 'Coupon Analysis', 'Coupon Usage % by City', '% of orders',
          rng('Coupon Analysis', (cc_rng[0], 0, cc_rng[2], 0), 0),
          rng('Coupon Analysis', (cc_rng[0], 3, cc_rng[2], 3), 3), 'Coupon Usage %',
          anchor=(r, 8), y_fmt=P)
r += max(len(cc) + 2, 20)

r = section(ws, r, 'Q4. COUPONS & CUSTOMER RETENTION')
rp = rep_cp.copy()
rp['group'] = rp['coupon_user'].map({True: 'Coupon users', False: 'Non-coupon users'})
rp = rp[['group', 'customers', 'repeat', 'repeat_rate', 'avg_orders']]
rp.columns = ['Group', 'Customers', 'Repeat Customers', 'Repeat Rate', 'Avg Orders / Customer']
rp_rng = write_table(ws, rp, r, 0, {0: None, 1: N, 2: N, 3: P, 4: '0.0'})
col_chart(wb, ws, 'Coupon Analysis', 'Repeat Customer Rate', '% repeat',
          rng('Coupon Analysis', (rp_rng[0], 0, rp_rng[2], 0), 0),
          rng('Coupon Analysis', (rp_rng[0], 3, rp_rng[2], 3), 3), 'Repeat Rate',
          anchor=(r, 8), y_fmt=P)

# ============================================================
# SHEET: Cancellation & Refund Analysis
# ============================================================
ws = wb.add_worksheet('Cancellation & Refund')
ws.set_tab_color('#E23744')
r = sheet_head(ws, 'CANCELLATION & REFUND ANALYSIS',
               'Cancellation/Refund rates are shares of all orders. Revenue lost = net revenue of affected orders.')
cancel_rest = rest_orders.sort_values('cancel_rate', ascending=False).head(10)
ml = orders[orders['order_status'].isin(['Cancelled', 'Refunded'])].groupby('month', sort=False) \
        .agg(lost=('net_revenue', 'sum')).reset_index()
mlo = pd.DataFrame({'month': mo['month'].tolist()})
mlm = ml.copy(); mlm['month'] = mlm['month'].astype(str)
mlo = mlo.merge(mlm, on='month', how='left').fillna({'lost': 0})
mlo['lost'] = mlo['lost'].round(2)
mlo.columns = ['Month', 'Revenue Lost']
top_can_rest = cancel_rest.iloc[0]
r = answers_block(ws, r, [
    ('Q1. What is the cancellation rate?',
     f'{cancel_rate:.1%} — {len(cancelled):,} of {total_orders:,} orders were cancelled'),
    ('Q2. What is the refund rate?',
     f'{refund_rate:.1%} — {len(refunded):,} of {total_orders:,} orders were refunded'),
    ('Q3. How much revenue is lost due to cancellations?',
     f'₹{lost_cancel:,.0f} lost to cancellations (+ ₹{lost_refund:,.0f} more to refunds = '
     f'₹{lost_cancel + lost_refund:,.0f} total, {(lost_cancel + lost_refund) / gross_gov:.1%} of gross order value)'),
    ('Q4. Which restaurants have the highest cancellation rate?',
     f'{top_can_rest["restaurant_name"]} ({top_can_rest["restaurant_id"]}, {top_can_rest["city"]}) — '
     f'{top_can_rest["cancel_rate"]:.1%} of its {int(top_can_rest["orders"])} orders cancelled. Top 10 below.'),
])
r = kpi_row(ws, r, [('CANCELLATION RATE', cancel_rate, P),
                    ('REFUND RATE', refund_rate, P),
                    ('REVENUE LOST — CANCELLATIONS', lost_cancel, M),
                    ('REVENUE LOST — REFUNDS', lost_refund, M)])

r = section(ws, r, 'Q1 & Q2. ORDER STATUS BREAKDOWN')
st2 = orders.groupby('order_status').agg(orders=('order_id', 'count'),
                                        revenue=('net_revenue', 'sum')).reset_index()
st2['pct'] = st2['orders'] / total_orders
st2 = st2.round({'pct': 4, 'revenue': 2})
st2.columns = ['Status', 'Orders', 'Net Revenue', '% of Orders']
st2_rng = write_table(ws, st2, r, 0, {0: None, 1: N, 2: M, 3: P})
pie_chart(wb, ws, 'Cancellation & Refund', 'Order Status Breakdown',
          rng('Cancellation & Refund', st2_rng, 0), rng('Cancellation & Refund', st2_rng, 1), 'Orders',
          anchor=(r, 8))
r += max(len(st2) + 2, 20)

r = section(ws, r, 'Q3. MONTHLY REVENUE LOST (CANCELLATIONS + REFUNDS)')
ml_rng = write_table(ws, mlo, r, 0, {0: None, 1: M})
line_chart(wb, ws, 'Cancellation & Refund', 'Monthly Revenue Lost', 'Revenue lost (₹)',
           rng('Cancellation & Refund', (ml_rng[0], 0, ml_rng[2], 0), 0),
           rng('Cancellation & Refund', (ml_rng[0], 1, ml_rng[2], 1), 1), 'Revenue Lost',
           anchor=(r, 8), y_fmt=M, color='#F5A623')
r += max(len(mlo) + 2, 22)

r = section(ws, r, 'Q4. TOP 10 RESTAURANTS BY CANCELLATION RATE')
cr = cancel_rest.copy().reset_index(drop=True)
cr['rank'] = range(1, 11)
cr = cr[['rank', 'restaurant_id', 'restaurant_name', 'city', 'orders', 'cancels', 'cancel_rate']]
cr.columns = ['Rank', 'ID', 'Restaurant', 'City', 'Orders', 'Cancelled Orders', 'Cancellation Rate']
cr = cr.round({'Cancellation Rate': 4})
cr_rng = write_table(ws, cr, r, 0, {0: N, 1: None, 2: None, 3: None, 4: N, 5: N, 6: P})
col_chart(wb, ws, 'Cancellation & Refund', 'Top 10 Restaurants by Cancellation Rate', 'Cancellation rate',
          rng('Cancellation & Refund', (cr_rng[0], 2, cr_rng[2], 2), 2),
          rng('Cancellation & Refund', (cr_rng[0], 6, cr_rng[2], 6), 6), 'Cancellation Rate',
          anchor=(r, 8), y_fmt=P, color='#F5A623')

# ============================================================
# SHEET: Customer Churn Analysis
# ============================================================
ws = wb.add_worksheet('Customer Churn')
ws.set_tab_color('#E23744')
r = sheet_head(ws, 'CUSTOMER CHURN ANALYSIS',
               f'Churned = customer whose last order is more than {CHURN_DAYS} days before the data end '
               f'({DATA_END:%d-%b-%Y}). Cutoff: {CHURN_CUTOFF:%d-%b-%Y}.')
ch_city = cust_orders.groupby('City').agg(customers=('customer_id', 'count'),
                                          churned=('churned', 'sum'),
                                          revenue=('revenue', 'sum')).reset_index()
ch_city['churn_rate'] = ch_city['churned'] / ch_city['customers']
ch_city['rev_lost'] = cust_orders[cust_orders['churned']].groupby('City')['revenue'].sum().reindex(ch_city['City']).fillna(0).values
ch_city = ch_city.sort_values('churn_rate', ascending=False).round({'churn_rate': 4, 'rev_lost': 2})
top_ch_city = ch_city.iloc[0]
hv = cust_orders[cust_orders['churned']].sort_values('revenue', ascending=False).head(10).reset_index(drop=True)
hv_top = hv.iloc[0]
r = answers_block(ws, r, [
    ('Q1. How many customers have churned?',
     f'{n_churned:,} of {len(cust_orders):,} customers have not ordered since {CHURN_CUTOFF:%d-%b-%Y}'),
    ('Q2. What is the churn rate?',
     f'{churn_rate:.1%} of the customer base is churned (90-day window)'),
    ('Q3. Which city has the highest churn?',
     f'{top_ch_city["City"]} — {top_ch_city["churn_rate"]:.1%} churn rate '
     f'({int(top_ch_city["churned"]):,} of {int(top_ch_city["customers"]):,} customers)'),
    ('Q4. How much revenue is lost due to churn?',
     f'₹{revenue_lost_churn:,.0f} — lifetime revenue historically generated by churned customers '
     f'({revenue_lost_churn / total_revenue:.1%} of all revenue)'),
    ('Q5. Who are the high-value churned customers?',
     f'Top: {hv_top["Customer_name"]} ({hv_top["customer_id"]}, {hv_top["City"]}) — '
     f'₹{hv_top["revenue"]:,.0f} lifetime revenue over {int(hv_top["orders"])} orders; last order {hv_top["last_order"]:%d-%b-%Y}. Top 10 below.'),
])
r = kpi_row(ws, r, [('CHURNED CUSTOMERS', n_churned, N),
                    ('CHURN RATE', churn_rate, P),
                    ('ACTIVE CUSTOMERS', len(cust_orders) - n_churned, N),
                    ('REVENUE LOST TO CHURN', revenue_lost_churn, M)])

r = section(ws, r, 'Q1 & Q2. ACTIVE vs CHURNED CUSTOMERS')
ac = pd.DataFrame({'status': ['Active customers', 'Churned customers'],
                   'customers': [len(cust_orders) - n_churned, n_churned]})
ac['pct'] = ac['customers'] / len(cust_orders)
ac = ac.round({'pct': 4})
ac_rng = write_table(ws, ac, r, 0, {0: None, 1: N, 2: P})
pie_chart(wb, ws, 'Customer Churn', 'Active vs Churned Customers',
          rng('Customer Churn', ac_rng, 0), rng('Customer Churn', ac_rng, 1), 'Customers',
          anchor=(r, 8))
r += max(len(ac) + 2, 20)

r = section(ws, r, 'Q3. CHURN BY CITY')
cc2 = ch_city[['City', 'customers', 'churned', 'churn_rate', 'rev_lost']].copy()
cc2.columns = ['City', 'Customers', 'Churned', 'Churn Rate', 'Revenue Lost']
cc2_rng = write_table(ws, cc2, r, 0, {0: None, 1: N, 2: N, 3: P, 4: M})
col_chart(wb, ws, 'Customer Churn', 'Churn Rate by City', 'Churn rate',
          rng('Customer Churn', (cc2_rng[0], 0, cc2_rng[2], 0), 0),
          rng('Customer Churn', (cc2_rng[0], 3, cc2_rng[2], 3), 3), 'Churn Rate',
          anchor=(r, 8), y_fmt=P, color='#7B61FF')
r += max(len(cc2) + 2, 20)

r = section(ws, r, 'Q4. REVENUE LOST TO CHURN BY CITY')
col_chart(wb, ws, 'Customer Churn', 'Revenue Lost to Churn by City', 'Revenue lost (₹)',
          rng('Customer Churn', (cc2_rng[0], 0, cc2_rng[2], 0), 0),
          rng('Customer Churn', (cc2_rng[0], 4, cc2_rng[2], 4), 4), 'Revenue Lost',
          anchor=(r, 8), y_fmt=M, color='#7B61FF')
r += 20

r = section(ws, r, 'Q5. TOP 10 HIGH-VALUE CHURNED CUSTOMERS')
hv2 = hv.copy()
hv2['rank'] = range(1, 11)
hv2 = hv2[['rank', 'customer_id', 'Customer_name', 'City', 'orders', 'revenue', 'last_order']]
hv2.columns = ['Rank', 'Customer ID', 'Name', 'City', 'Orders', 'Lifetime Revenue', 'Last Order']
hv2 = hv2.round({'Lifetime Revenue': 2})
hv_rng = write_table(ws, hv2, r, 0, {0: N, 1: None, 2: None, 3: None, 4: N, 5: M, 6: None})
col_chart(wb, ws, 'Customer Churn', 'Top 10 High-Value Churned Customers', 'Lifetime revenue (₹)',
          rng('Customer Churn', (hv_rng[0], 2, hv_rng[2], 2), 2),
          rng('Customer Churn', (hv_rng[0], 5, hv_rng[2], 5), 5), 'Lifetime Revenue',
          anchor=(r, 8), y_fmt=M, color='#7B61FF')

# ============================================================
# SHEET: Assumptions
# ============================================================
ws = wb.add_worksheet('Assumptions')
ws.hide_gridlines(2)
ws.merge_range(0, 0, 0, 9, 'ASSUMPTIONS, DEFINITIONS & DATA NOTES', F['title'])
notes = [
    ('Data source', 'Zomato  Order Data.xlsx — sheets: Customer (4,999 rows), Restaurants (200 rows), Orders (50,000 rows).'),
    ('Analysis period', f'Order data spans {orders["order_date"].min():%d-%b-%Y} to {DATA_END:%d-%b-%Y}.'),
    ('Net revenue', 'order_amount − discount_amount. "Total revenue" counts DELIVERED orders only, since cancelled/refunded orders generate no realised revenue.'),
    ('Gross order value (GOV)', 'Sum of order_amount across all 50,000 orders regardless of status.'),
    ('AOV', 'Total net revenue ÷ number of delivered orders.'),
    ('Coupon order', 'Any order with discount_amount > 0 (≈10% off). 20,128 of 50,000 orders (40.3%) use a coupon.'),
    ('Repeat customer', 'Customer with 2 or more orders in the period.'),
    ('Churn definition', f'Customer whose most recent order is older than {CHURN_DAYS} days before the last date in the dataset ({DATA_END:%d-%b-%Y}); i.e. last order before {CHURN_CUTOFF:%d-%b-%Y}. Churn rate = churned ÷ all customers with orders.'),
    ('Revenue lost to churn', 'Lifetime net revenue historically generated by churned customers (used as a proxy for lost future value).'),
    ('Revenue lost to cancellations/refunds', 'Net revenue of orders with status Cancelled / Refunded.'),
    ('Data quality — city', "One customer record had city ' pune' (leading space); trimmed to 'Pune'."),
    ('Data quality — missing customer', "10 orders reference customer C00001, which is missing from the Customer master. C00001 is included in all metrics with name/city/channel shown as 'Unknown'."),
    ('Data quality — duplicates', 'No duplicate order IDs; no null values found in any sheet.'),
    ('Currency', 'All monetary values are in ₹ (INR) as present in the source data.'),
    ('Charts', 'All visuals are native Excel charts — they update automatically if the underlying table data is changed.'),
]
rr = 2
for k, v in notes:
    ws.merge_range(rr, 0, rr, 2, k, F['hdr'])
    ws.merge_range(rr, 3, rr, 9, v, F['txt'])
    ws.set_row(rr, 34)
    rr += 1
ws.set_column(0, 2, 26)
ws.set_column(3, 9, 22)

wb.close()
print('Workbook written:', OUT)
print('Total revenue:', round(total_revenue), '| AOV:', round(aov), '| Cancel:', round(cancel_rate,4), '| Refund:', round(refund_rate,4), '| Coupon:', round(coupon_pct,4), '| Churn:', round(churn_rate,4))
