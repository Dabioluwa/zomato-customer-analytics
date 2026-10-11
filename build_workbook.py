import pandas as pd
import numpy as np
import xlsxwriter

SRC = r'C:\Users\onash\Documents\Project\Zomato\Zomato  Order Data.xlsx'
FINAL = r'C:\Users\onash\Documents\Project\Zomato\Zomato_Customer_Analytics.xlsx'
OUT = r'C:\Users\onash\AppData\Local\Temp\opencode\Zomato_Analytics_build.xlsx'

# ============================================================
# 1. LOAD & PREPARE DATA
# ============================================================
xl = pd.ExcelFile(SRC)
cust = xl.parse('Customer'); cust.columns = cust.columns.str.strip()
rest = xl.parse('Restaurants'); rest.columns = rest.columns.str.strip()
orders = xl.parse('Orders'); orders.columns = orders.columns.str.strip()

cust['City'] = cust['City'].str.strip()
rest['city'] = rest['city'].str.strip()
if 'C00001' not in set(cust['Customer_id']):
    cust = pd.concat([pd.DataFrame([{'Customer_id': 'C00001', 'Customer_name': 'Unknown',
                                     'City': 'Unknown', 'Signup_Time': pd.NaT,
                                     'Acquisition_channel': 'Unknown'}]), cust], ignore_index=True)

orders['order_date'] = pd.to_datetime(orders['order_timestamp'])
orders['month'] = orders['order_date'].dt.strftime('%Y-%m')
orders['net_revenue'] = (orders['order_amount'] - orders['discount_amount']).round(2)
orders['coupon'] = (orders['discount_amount'] > 0).astype(int)

cmap = cust.set_index('Customer_id')
orders['customer_name'] = orders['customer_id'].map(cmap['Customer_name']).fillna('Unknown')
orders['customer_city'] = orders['customer_id'].map(cmap['City']).fillna('Unknown')
orders['city'] = orders['customer_city']
orders['channel'] = orders['customer_id'].map(cmap['Acquisition_channel']).fillna('Unknown')
rmap = rest.set_index('restaurant_id')
orders['restaurant_name'] = orders['restaurant_id'].map(rmap['restaurant_name']).fillna('Unknown')
orders['cuisine'] = orders['restaurant_id'].map(rmap['cuisine']).fillna('Unknown')
orders['rest_city'] = orders['restaurant_id'].map(rmap['city']).fillna('Unknown')
orders['rating'] = orders['restaurant_id'].map(rmap['avg_rating'])
orders['rev_delivered'] = np.where(orders['order_status'] == 'Delivered', orders['net_revenue'], 0.0)
orders['rev_lost'] = np.where(orders['order_status'] != 'Delivered', orders['net_revenue'], 0.0)
orders['payment'] = orders['payment_mode']
orders['status'] = orders['order_status']

cities = sorted(cust['City'].dropna().unique().tolist())
channels = sorted(cust['Acquisition_channel'].dropna().unique().tolist())
payments = sorted(orders['payment_mode'].dropna().unique().tolist())
statuses = ['Delivered', 'Cancelled', 'Refunded']
cuisines = sorted(rest['cuisine'].dropna().unique().tolist())
month_labels = list(pd.period_range('2024-01', '2026-04', freq='M').strftime('%Y-%m'))
DATA_END = orders['order_date'].max()
CHURN_DAYS = 90
CHURN_CUTOFF = DATA_END - pd.Timedelta(days=CHURN_DAYS)

cust_master = cust[['Customer_id', 'Customer_name', 'City', 'Acquisition_channel']].drop_duplicates(
    subset=['Customer_id']).reset_index(drop=True)
cust_master.columns = ['Customer_id', 'Customer_name', 'City', 'Channel']
rest_master = rest.drop_duplicates(subset=['restaurant_id']).reset_index(drop=True)

NDATA = len(orders); DATA_LAST = NDATA + 1
NREST = len(rest_master); CR_LAST = NREST + 1
NCUST = len(cust_master); CC_LAST = NCUST + 1

# ============================================================
# 2. WORKBOOK & THEME (cream + light brown)
# ============================================================
wb = xlsxwriter.Workbook(OUT, {'nan_inf_to_errors': True})
wb.set_properties({'title': 'Zomato Customer Analytics Dashboard', 'author': 'Data Analyst'})

RED = '#8B5A2B'      # dark warm brown  (primary accent)
DARK = '#5C3A21'     # near-black brown (headers / labels)
GREEN = '#A97B50'    # mid brown
AMBER = '#C9A277'    # light brown
PURPLE = '#B08D57'   # golden tan
BLUE = '#7E5A3A'
LIGHT = '#FBF3E4'    # cream panel
GREY = '#F5EDD9'     # soft cream (alt rows)
BORD = '#E0CDA8'
CREAM_PLOT = '#FDF7EA'   # plot area
CREAM_CHART = '#FBF3E3'  # chart area
PIE = ['#8B5A2B', '#A97B50', '#C9A277', '#6E4422', '#BD8E5C', '#DFC09F', '#935F3A',
       '#5C3A21', '#D6B98C', '#E8D8BD']
M = '"₹"#,##0'; M2 = '"₹"#,##0.00'; P = '0.0%'; N0 = '#,##0'; D0 = 'dd-mmm-yy'; DEC = '0.0'
FN = 'Segoe UI'

F = {}
F['title'] = wb.add_format({'bold': True, 'font_size': 20, 'font_color': 'white', 'bg_color': RED,
                            'font_name': FN, 'valign': 'vcenter'})
F['sub'] = wb.add_format({'font_size': 9, 'italic': True, 'font_color': 'white', 'bg_color': RED,
                          'font_name': FN, 'valign': 'vcenter'})
F['banner'] = wb.add_format({'bold': True, 'font_size': 11, 'font_color': 'white', 'bg_color': RED,
                             'indent': 1, 'valign': 'vcenter', 'font_name': FN})
F['qhdr'] = wb.add_format({'bold': True, 'font_size': 10, 'font_color': RED, 'bg_color': LIGHT,
                           'valign': 'vcenter', 'border': 1, 'border_color': BORD, 'text_wrap': True,
                           'font_name': FN})
F['ans'] = wb.add_format({'font_size': 10, 'valign': 'vcenter', 'bg_color': LIGHT, 'border': 1,
                          'border_color': RED, 'italic': True, 'font_name': FN})
F['thead'] = wb.add_format({'bold': True, 'font_size': 10, 'font_color': 'white', 'bg_color': DARK,
                            'border': 1, 'border_color': 'white', 'align': 'center', 'valign': 'vcenter',
                            'font_name': FN})
F['kpi_lbl'] = wb.add_format({'font_size': 9, 'font_color': '#A08464', 'align': 'center',
                              'valign': 'vcenter', 'font_name': FN})
F['note'] = wb.add_format({'font_size': 9, 'italic': True, 'font_color': '#A08464',
                           'valign': 'top', 'font_name': FN})
F['txt'] = wb.add_format({'font_size': 10, 'valign': 'top', 'font_name': FN})
F['flt_lbl'] = wb.add_format({'bold': True, 'font_size': 9, 'font_color': 'white', 'bg_color': DARK,
                              'align': 'center', 'valign': 'vcenter', 'font_name': FN})
F['flt_val'] = wb.add_format({'bold': True, 'font_size': 11, 'font_color': RED, 'bg_color': 'white',
                              'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': RED,
                              'font_name': FN})
F['filt_note'] = wb.add_format({'font_size': 9, 'italic': True, 'font_color': '#A08464', 'font_name': FN})
F['sl_banner'] = wb.add_format({'bold': True, 'font_size': 10, 'font_color': DARK, 'bg_color': '#F3E6CE',
                                'align': 'left', 'valign': 'vcenter', 'border': 1, 'border_color': BORD,
                                'indent': 1, 'font_name': FN})

def kpi_fmt(numfmt, color=RED):
    return wb.add_format({'bold': True, 'font_size': 15, 'font_color': color, 'align': 'center',
                          'valign': 'vcenter', 'num_format': numfmt, 'font_name': FN})

_body = {}
def bfmt(numfmt, alt, bold=False, color=None):
    key = (numfmt, alt, bold, color)
    if key not in _body:
        d = {'font_size': 10, 'border': 1, 'border_color': BORD, 'font_name': FN, 'valign': 'vcenter'}
        if numfmt: d['num_format'] = numfmt
        if alt: d['bg_color'] = GREY
        if bold: d['bold'] = True
        if color: d['font_color'] = color
        _body[key] = wb.add_format(d)
    return _body[key]

def rfmt(numfmt=None):
    d = {'font_size': 10, 'font_name': FN}
    if numfmt: d['num_format'] = numfmt
    return wb.add_format(d)

# ============================================================
# 3. WORKSHEETS  (tables and charts on separate sheets)
# ============================================================
ws_dash = wb.add_worksheet('Dashboard')
ws_revT = wb.add_worksheet('Revenue - Tables')
ws_revC = wb.add_worksheet('Revenue - Charts')
ws_custT = wb.add_worksheet('Customer - Tables')
ws_custC = wb.add_worksheet('Customer - Charts')
ws_restT = wb.add_worksheet('Restaurant - Tables')
ws_restC = wb.add_worksheet('Restaurant - Charts')
ws_coupT = wb.add_worksheet('Coupon - Tables')
ws_coupC = wb.add_worksheet('Coupon - Charts')
ws_canT = wb.add_worksheet('Cancel - Tables')
ws_canC = wb.add_worksheet('Cancel - Charts')
ws_churnT = wb.add_worksheet('Churn - Tables')
ws_churnC = wb.add_worksheet('Churn - Charts')
ws_assum = wb.add_worksheet('Assumptions')
ws_data = wb.add_worksheet('Data')
ws_list = wb.add_worksheet('Lists')
ws_cc = wb.add_worksheet('Calc_Customers')
ws_cr = wb.add_worksheet('Calc_Restaurants')

for w in (ws_dash, ws_revC, ws_custC, ws_restC, ws_coupC, ws_canC, ws_churnC):
    w.hide_gridlines(2); w.set_tab_color(RED)
for w in (ws_revT, ws_custT, ws_restT, ws_coupT, ws_canT, ws_churnT, ws_assum):
    w.hide_gridlines(2); w.set_tab_color('#D9C7A7')
ws_data.set_tab_color(DARK)
for w in (ws_list, ws_cc, ws_cr):
    w.hide(); w.set_tab_color('#BFBFBF')

# ============================================================
# 4. DATA SHEET
# ============================================================
data_cols = ['order_id', 'order_date', 'month', 'customer_id', 'customer_name', 'city', 'channel',
             'restaurant_id', 'restaurant_name', 'cuisine', 'rest_city', 'rating', 'payment',
             'status', 'order_amount', 'discount_amount', 'delivery_fee', 'net_revenue', 'coupon',
             'match', 'rev_delivered', 'rev_lost']
headers = ['Order ID', 'Order Date', 'Month', 'Customer ID', 'Customer Name', 'City', 'Channel',
           'Restaurant ID', 'Restaurant Name', 'Cuisine', 'Restaurant City', 'Rating', 'Payment',
           'Status', 'Order Amount', 'Discount', 'Delivery Fee', 'Net Revenue', 'Coupon',
           'Match', 'Rev Delivered', 'Rev Lost']
for j, h in enumerate(headers):
    ws_data.write(0, j, h, F['thead'])

write_plain = rfmt()
write_date = rfmt(D0)
write_rate = rfmt('0.0')
str_cols = {'order_id', 'month', 'customer_id', 'customer_name', 'city', 'channel',
            'restaurant_id', 'restaurant_name', 'cuisine', 'rest_city', 'payment', 'status'}
num_cols = {'order_amount', 'discount_amount', 'delivery_fee', 'net_revenue', 'coupon',
            'rev_delivered', 'rev_lost'}
for j, c in enumerate(data_cols):
    if c == 'order_date':
        ws_data.write_column(1, j, orders['order_date'].dt.to_pydatetime(), write_date)
    elif c == 'rating':
        ws_data.write_column(1, j, orders['rating'].values, write_rate)
    elif c in str_cols:
        ws_data.write_column(1, j, orders[c].values, write_plain)
    elif c in num_cols:
        ws_data.write_column(1, j, orders[c].values, write_plain)

match_fmt = rfmt('0')
for i in range(NDATA):
    r = i + 2
    ws_data.write_formula(i + 1, 19,
        f'=(((Filter_City="(All)")+($F{r}=Filter_City))>0)'
        f'*(((Filter_Channel="(All)")+($G{r}=Filter_Channel))>0)'
        f'*(((Filter_Payment="(All)")+($M{r}=Filter_Payment))>0)'
        f'*(((Filter_Status="(All)")+($N{r}=Filter_Status))>0)'
        f'*(((Filter_Cuisine="(All)")+($J{r}=Filter_Cuisine))>0)'
        f'*(((Filter_Month="(All)")+($C{r}=Filter_Month))>0)', match_fmt)

for j, wdt in enumerate([11, 12, 8, 11, 16, 11, 16, 11, 18, 12, 13, 7, 9, 11, 12, 11, 11, 12, 8, 8, 12, 11]):
    ws_data.set_column(j, j, wdt)
ws_data.set_column(19, 19, None, None, {'hidden': True})
ws_data.freeze_panes(1, 0)

def dname(name, col_letter):
    wb.define_name(name, f"=Data!${col_letter}$2:${col_letter}${DATA_LAST}")
for nm, col in [('D_Date', 'B'), ('D_Month', 'C'), ('D_CustId', 'D'), ('D_City', 'F'), ('D_Channel', 'G'),
                ('D_RestId', 'H'), ('D_Cuisine', 'J'), ('D_Rating', 'L'), ('D_Pay', 'M'),
                ('D_Status', 'N'), ('D_Amount', 'O'), ('D_Discount', 'P'), ('D_Net', 'R'),
                ('D_Coupon', 'S'), ('D_Match', 'T'), ('D_Del', 'U'), ('D_Lost', 'V')]:
    dname(nm, col)

# ============================================================
# 5. LISTS
# ============================================================
lists = [('City', cities, 'A'), ('Channel', channels, 'B'), ('Payment', payments, 'C'),
         ('Status', statuses, 'D'), ('Cuisine', cuisines, 'E'), ('Month', month_labels, 'F')]
for j, (nm, vals, col) in enumerate(lists):
    ws_list.write(0, j, nm, F['thead'])
    ws_list.write(1, j, '(All)')
    for i, v in enumerate(vals):
        ws_list.write(2 + i, j, v)
ws_list.write(0, 7, 'ChurnCutoff', F['thead']); ws_list.write_datetime(1, 7, CHURN_CUTOFF.to_pydatetime(), rfmt(D0))
ws_list.write(0, 8, 'DataEnd', F['thead']); ws_list.write_datetime(1, 8, DATA_END.to_pydatetime(), rfmt(D0))
ws_list.write(0, 9, 'ChurnDays', F['thead']); ws_list.write(1, 9, CHURN_DAYS)
wb.define_name('Churn_Cutoff', "=Lists!$H$2")

# ============================================================
# 6. CALC_CUSTOMERS
# ============================================================
for j, h in enumerate(['Customer ID', 'Name', 'City', 'Channel', 'Orders', 'Revenue', 'CouponUser',
                       'Churned', 'Repeat']):
    ws_cc.write(0, j, h, F['thead'])
ws_cc.write_column(1, 0, cust_master['Customer_id'].values, write_plain)
ws_cc.write_column(1, 1, cust_master['Customer_name'].values, write_plain)
ws_cc.write_column(1, 2, cust_master['City'].values, write_plain)
ws_cc.write_column(1, 3, cust_master['Channel'].values, write_plain)
ccf_n = rfmt('0'); ccf_m = rfmt(M)
for i in range(NCUST):
    r = i + 2
    ws_cc.write_formula(i + 1, 4, f'=COUNTIFS(D_CustId,$A{r},D_Match,1)', ccf_n)
    ws_cc.write_formula(i + 1, 5, f'=SUMIFS(D_Del,D_CustId,$A{r},D_Match,1)', ccf_m)
    ws_cc.write_formula(i + 1, 6, f'=IF(COUNTIFS(D_CustId,$A{r},D_Match,1,D_Coupon,1)>0,1,0)', ccf_n)
    ws_cc.write_formula(i + 1, 7,
        f'=IF(COUNTIFS(D_CustId,$A{r},D_Match,1)<1,"",'
        f'IF(COUNTIFS(D_CustId,$A{r},D_Match,1,D_Date,">="&Churn_Cutoff)=0,1,0))', ccf_n)
    ws_cc.write_formula(i + 1, 8, f'=IF($E{r}<1,"",IF($E{r}>=2,1,0))', ccf_n)
ws_cc.write_dynamic_array_formula(1, 11, 1, 16,
    '=IFERROR(SORT(FILTER($A$2:$F$'+str(CC_LAST)+',$E$2:$E$'+str(CC_LAST)+'>0),6,-1),"")')
ws_cc.write_dynamic_array_formula(1, 19, 1, 24,
    '=IFERROR(SORT(FILTER($A$2:$F$'+str(CC_LAST)+',$H$2:$H$'+str(CC_LAST)+'=1),6,-1),"")')
for j, wdt in enumerate([11, 18, 11, 16, 9, 12, 11, 9, 9]):
    ws_cc.set_column(j, j, wdt)

def ccname(name, col):
    wb.define_name(name, f"=Calc_Customers!${col}$2:${col}${CC_LAST}")
for nm, col in [('CC_Id', 'A'), ('CC_Name', 'B'), ('CC_City', 'C'), ('CC_Channel', 'D'),
                ('CC_Orders', 'E'), ('CC_Revenue', 'F'), ('CC_Coupon', 'G'), ('CC_Churned', 'H'),
                ('CC_Repeat', 'I')]:
    ccname(nm, col)

# ============================================================
# 7. CALC_RESTAURANTS
# ============================================================
for j, h in enumerate(['Restaurant ID', 'Name', 'City', 'Cuisine', 'Rating', 'Orders', 'Revenue',
                       'Cancels', 'CancelRate']):
    ws_cr.write(0, j, h, F['thead'])
ws_cr.write_column(1, 0, rest_master['restaurant_id'].values, write_plain)
ws_cr.write_column(1, 1, rest_master['restaurant_name'].values, write_plain)
ws_cr.write_column(1, 2, rest_master['city'].values, write_plain)
ws_cr.write_column(1, 3, rest_master['cuisine'].values, write_plain)
ws_cr.write_column(1, 4, rest_master['avg_rating'].values, write_rate)
crf_n = rfmt('0'); crf_m = rfmt(M); crf_p = rfmt(P)
for i in range(NREST):
    r = i + 2
    ws_cr.write_formula(i + 1, 5, f'=COUNTIFS(D_RestId,$A{r},D_Match,1)', crf_n)
    ws_cr.write_formula(i + 1, 6, f'=SUMIFS(D_Del,D_RestId,$A{r},D_Match,1)', crf_m)
    ws_cr.write_formula(i + 1, 7, f'=COUNTIFS(D_RestId,$A{r},D_Match,1,D_Status,"Cancelled")', crf_n)
    ws_cr.write_formula(i + 1, 8, f'=IF($F{r}<1,"",$H{r}/$F{r})', crf_p)
ws_cr.write_dynamic_array_formula(1, 10, 1, 18,
    '=IFERROR(SORT(FILTER($A$2:$I$'+str(CR_LAST)+',$F$2:$F$'+str(CR_LAST)+'>0),7,-1),"")')   # revenue desc K
ws_cr.write_dynamic_array_formula(1, 20, 1, 28,
    '=IFERROR(SORT(FILTER($A$2:$I$'+str(CR_LAST)+',$F$2:$F$'+str(CR_LAST)+'>0),6,-1),"")')   # orders desc U
ws_cr.write_dynamic_array_formula(1, 30, 1, 38,
    '=IFERROR(SORT(FILTER($A$2:$I$'+str(CR_LAST)+',$F$2:$F$'+str(CR_LAST)+'>0),9,-1),"")')   # cancel desc AE
ws_cr.write_dynamic_array_formula(1, 40, 1, 48,
    '=IFERROR(SORT(FILTER($A$2:$I$'+str(CR_LAST)+',$F$2:$F$'+str(CR_LAST)+'>0),7,1),"")')    # revenue asc AO
for j, wdt in enumerate([11, 18, 11, 12, 7, 9, 12, 9, 11]):
    ws_cr.set_column(j, j, wdt)

def crname(name, col):
    wb.define_name(name, f"=Calc_Restaurants!${col}$2:${col}${CR_LAST}")
for nm, col in [('CR_Id', 'A'), ('CR_Name', 'B'), ('CR_City', 'C'), ('CR_Cuisine', 'D'),
                ('CR_Rating', 'E'), ('CR_Orders', 'F'), ('CR_Revenue', 'G'), ('CR_Cancels', 'H'),
                ('CR_CancelRate', 'I')]:
    crname(nm, col)

# ============================================================
# 8. ANALYSIS HELPERS
# ============================================================
RG = {}
CH_ROW = 7; CH_STEP = 27

def sheet_head(ws, title, subtitle, ncols=16):
    ws.merge_range(0, 0, 0, ncols - 1, title, F['title'])
    ws.set_row(0, 30)
    if subtitle:
        ws.merge_range(1, 0, 1, ncols - 1, subtitle, F['sub'])
        ws.set_row(1, 16)
    return 3

def banner(ws, r, text, ncols=16):
    ws.merge_range(r, 0, r, ncols - 1, text, F['banner'])
    ws.set_row(r, 20)
    return r + 1

def write_answers(ws, start, pairs, ncols=16):
    ws.merge_range(start, 0, start, ncols - 1,
                   'KEY ANSWERS',
                   wb.add_format({'bold': True, 'font_size': 11, 'font_color': 'white', 'bg_color': RED,
                                  'indent': 1, 'valign': 'vcenter', 'font_name': FN}))
    ws.set_row(start, 20)
    r = start + 1
    for q, a in pairs:
        ws.merge_range(r, 0, r, 4, q, F['qhdr'])
        ws.merge_range(r, 5, r, ncols - 1, a, F['ans'])
        ws.set_row(r, 30)
        r += 1
    return r

def kpi_cards(ws, r, kpis):
    fmts = {'M': kpi_fmt(M), 'N': kpi_fmt(N0), 'P': kpi_fmt(P), 'D': kpi_fmt(D0)}
    for i, (lbl, formula, kind) in enumerate(kpis):
        c = i * 4
        ws.merge_range(r, c, r, c + 3, lbl, F['kpi_lbl'])
        ws.merge_range(r + 1, c, r + 1, c + 3, formula, fmts[kind])
        ws.set_row(r, 16); ws.set_row(r + 1, 26)
    return r + 3

def put_table(ws, r, c, key, headers, rows, numfmts=None, widths=None):
    numfmts = numfmts or {}
    for j, h in enumerate(headers):
        ws.write(r, c + j, h, F['thead'])
    for i, rowvals in enumerate(rows):
        alt = (i % 2 == 1)
        for j, val in enumerate(rowvals):
            fmt = bfmt(numfmts.get(j), alt)
            if isinstance(val, str) and val.startswith('='):
                ws.write_formula(r + 1 + i, c + j, val, fmt)
            elif val is None:
                ws.write_blank(r + 1 + i, c + j, None, fmt)
            else:
                ws.write(r + 1 + i, c + j, val, fmt)
    RG[key] = (r + 1, r + len(rows), c)
    if widths:
        for j, w in widths.items():
            ws.set_column(c + j, c + j, w)
    return r + len(rows) + 1

def cref(sheet, key, colidx):
    r1, r2, c = RG[key]
    return [sheet, r1, c + colidx, r2, c + colidx]

def arng(sheet, key, colidx):
    r1, r2, c = RG[key]
    L = chr(65 + c + colidx)
    return f"'{sheet}'!${L}${r1+1}:${L}${r2+1}"

def acell(sheet, key, row0, colidx):
    r1, r2, c = RG[key]
    L = chr(65 + c + colidx)
    return f"'{sheet}'!${L}${r1+row0+1}"

def make_chart(kind, title, cats, vals, ser, color=RED, ylab='', yfmt='', anchor=(0, 0),
               xlab='', xfmt='', trend=False, w=620, h=360, target=None):
    ws_t = target or ws
    ch = wb.add_chart({'type': kind})
    s = {'name': ser, 'categories': cats, 'values': vals}
    if kind == 'line':
        s['line'] = {'color': color, 'width': 2.25}
        s['marker'] = {'type': 'circle', 'size': 4, 'fill': {'color': color}}
    elif kind == 'scatter':
        s['marker'] = {'type': 'circle', 'size': 5, 'fill': {'color': color}, 'border': {'none': True}}
        s['line'] = {'none': True}
        if trend:
            s['trendline'] = {'type': 'linear', 'line': {'color': '#96764B', 'width': 1.5, 'dash_type': 'dash'}}
    elif kind == 'pie':
        s['data_labels'] = {'percentage': True, 'font': {'size': 9}, 'color': '#3A2410', 'bold': True}
        s['points'] = [{'fill': {'color': c}} for c in PIE]
    else:
        s['fill'] = {'color': color}
        s['gap'] = 60
    ch.add_series(s)
    ch.set_title({'name': title, 'name_font': {'size': 11, 'bold': True, 'name': FN, 'color': DARK}})
    if kind in ('column', 'line', 'bar'):
        ch.set_x_axis({'num_font': {'size': 8, 'rotation': -45, 'font_color': DARK}})
        ch.set_y_axis({'name': ylab, 'num_format': yfmt, 'num_font': {'size': 8, 'font_color': DARK},
                       'name_font': {'size': 8, 'color': DARK},
                       'major_gridlines': {'visible': kind == 'line',
                                           'line': {'color': '#E8D9BD'}}})
    elif kind == 'scatter':
        ch.set_x_axis({'name': xlab, 'num_format': xfmt or '0.0', 'num_font': {'size': 8, 'font_color': DARK},
                       'name_font': {'size': 8, 'color': DARK}})
        ch.set_y_axis({'name': ylab, 'num_format': yfmt, 'num_font': {'size': 8, 'font_color': DARK},
                       'name_font': {'size': 8, 'color': DARK}})
    if kind == 'pie':
        ch.set_legend({'position': 'right', 'font': {'size': 9}})
    else:
        ch.set_legend({'none': True})
    ch.set_chartarea({'border': {'color': BORD}, 'fill': {'color': CREAM_CHART}})
    ch.set_plotarea({'fill': {'color': CREAM_PLOT}})
    ws_t.insert_chart(anchor[0], anchor[1], ch, {'x_scale': w / 480, 'y_scale': h / 288})
    return ch

def place_chart(ws_ch, idx, kind, title, cats, vals, ser, **kw):
    col = 0 if (idx % 2 == 0) else 11
    row = CH_ROW + (idx // 2) * CH_STEP
    kw.setdefault('w', 620); kw.setdefault('h', 360)
    return make_chart(kind, title, cats, vals, ser, anchor=(row, col), target=ws_ch, **kw)

def drill_strip(ws, drills):
    for (lbl, opts, col) in drills:
        ws.merge_range(3, col, 3, col + 1, lbl, F['flt_lbl'])
        ws.merge_range(4, col, 4, col + 1, '(All)', F['flt_val'])
        ws.data_validation(4, col, 4, col, {'validate': 'list', 'source': ['(All)'] + opts,
                                            'dropdown': True})
    ws.set_row(3, 16); ws.set_row(4, 24)

def tcell(src, guard, n, colidx):
    return (f'=IFERROR(IF({n}<={guard},INDEX({src},{n},{colidx}),""),"")')

def chaindef(name, refs):
    expr = refs[-1]
    for r in reversed(refs[:-1]):
        expr = f'IF({r}="(All)",{expr},{r})'
    wb.define_name(name, '=' + expr)

# ============================================================
# 9. REVENUE ANALYSIS — TABLES
# ============================================================
ws = ws_revT
r = sheet_head(ws, 'REVENUE ANALYSIS — TABLES',
               'Revenue = order_amount − discount_amount on DELIVERED orders.')
ans = r
r = r + 1 + 5 + 1
r = kpi_cards(ws, r, [
    ('TOTAL REVENUE', '=SUMIFS(D_Del,D_Match,1)', 'M'),
    ('GROSS ORDER VALUE', '=SUMIFS(D_Amount,D_Match,1)', 'M'),
    ('DELIVERED ORDERS', '=COUNTIFS(D_Match,1,D_Status,"Delivered")', 'N'),
    ('AVG ORDER VALUE', '=IFERROR(SUMIFS(D_Del,D_Match,1)/COUNTIFS(D_Match,1,D_Status,"Delivered"),0)', 'M'),
])
r = banner(ws, r, 'Q2.   MONTHLY REVENUE TREND')
rows = [[m,
         f'=COUNTIFS(D_Match,1,D_Month,"{m}",D_Status,"Delivered")',
         f'=SUMIFS(D_Del,D_Match,1,D_Month,"{m}")',
         f'=IFERROR(SUMIFS(D_Del,D_Match,1,D_Month,"{m}")/COUNTIFS(D_Match,1,D_Month,"{m}",D_Status,"Delivered"),"")']
        for m in month_labels]
r = put_table(ws, r, 0, 'rev_month', ['Month', 'Delivered Orders', 'Net Revenue', 'AOV'], rows,
              {1: N0, 2: M, 3: M}, {0: 9, 1: 14, 2: 15, 3: 12})
r = banner(ws, r, 'Q3.   REVENUE BY CITY')
rows = [[c,
         f'=COUNTIFS(D_Match,1,D_City,"{c}",D_Status,"Delivered")',
         f'=SUMIFS(D_Del,D_Match,1,D_City,"{c}")',
         f'=IFERROR(SUMIFS(D_Del,D_Match,1,D_City,"{c}")/SUMIFS(D_Del,D_Match,1),0)',
         f'=IFERROR(SUMIFS(D_Del,D_Match,1,D_City,"{c}")/COUNTIFS(D_Match,1,D_City,"{c}",D_Status,"Delivered"),"")']
        for c in cities]
r = put_table(ws, r, 0, 'rev_city', ['City', 'Delivered Orders', 'Net Revenue', '% of Revenue', 'AOV'],
              rows, {1: N0, 2: M, 3: P, 4: M}, {0: 12, 1: 14, 2: 15, 3: 12, 4: 12})
r = banner(ws, r, 'Q4 & Q5.   REVENUE & AOV BY PAYMENT MODE')
rows = [[p,
         f'=COUNTIFS(D_Match,1,D_Pay,"{p}",D_Status,"Delivered")',
         f'=SUMIFS(D_Del,D_Match,1,D_Pay,"{p}")',
         f'=IFERROR(SUMIFS(D_Del,D_Match,1,D_Pay,"{p}")/SUMIFS(D_Del,D_Match,1),0)',
         f'=IFERROR(SUMIFS(D_Del,D_Match,1,D_Pay,"{p}")/COUNTIFS(D_Match,1,D_Pay,"{p}",D_Status,"Delivered"),"")']
        for p in payments]
r = put_table(ws, r, 0, 'rev_pay', ['Payment Mode', 'Delivered Orders', 'Net Revenue', '% of Revenue', 'AOV'],
              rows, {1: N0, 2: M, 3: P, 4: M}, {0: 15, 1: 14, 2: 15, 3: 12, 4: 12})
MR = arng('Revenue - Tables', 'rev_month', 2); MM = arng('Revenue - Tables', 'rev_month', 0)
CR = arng('Revenue - Tables', 'rev_city', 2); CM = arng('Revenue - Tables', 'rev_city', 0)
PR = arng('Revenue - Tables', 'rev_pay', 2); PM = arng('Revenue - Tables', 'rev_pay', 0)
write_answers(ws, ans, [
    ('Q1. What is the total revenue?',
     '="₹"&TEXT(SUMIFS(D_Del,D_Match,1),"#,##0")&" net revenue from "&TEXT(COUNTIFS(D_Match,1,D_Status,"Delivered"),"#,##0")'
     '&" delivered orders   (gross order value achieved: ₹"&TEXT(SUMIFS(D_Amount,D_Match,1),"#,##0")&")"'),
    ('Q2. What is the monthly revenue trend?',
     f'="Peak month: "&INDEX({MM},MATCH(MAX({MR}),{MR},0))&" with ₹"&TEXT(MAX({MR}),"#,##0")'
     f'&"   |   Lowest: "&INDEX({MM},MATCH(MIN({MR}),{MR},0))&" with ₹"&TEXT(MIN({MR}),"#,##0")&"   — see chart"'),
    ('Q3. Which city contributes the highest revenue?',
     f'=INDEX({CM},MATCH(MAX({CR}),{CR},0))&" — ₹"&TEXT(MAX({CR}),"#,##0")&" ("&TEXT(IFERROR(MAX({CR})/SUMIFS(D_Del,D_Match,1),0),"0.0%")&" of total revenue)"'),
    ('Q4. Which payment mode generates the most revenue?',
     f'=INDEX({PM},MATCH(MAX({PR}),{PR},0))&" — ₹"&TEXT(MAX({PR}),"#,##0")&" ("&TEXT(IFERROR(MAX({PR})/SUMIFS(D_Del,D_Match,1),0),"0.0%")&" of total revenue)"'),
    ('Q5. What is the Average Order Value (AOV)?',
     '="₹"&TEXT(IFERROR(SUMIFS(D_Del,D_Match,1)/COUNTIFS(D_Match,1,D_Status,"Delivered"),0),"#,##0")&" per delivered order (net of discounts)"'),
])

# ============================================================
# 10. REVENUE ANALYSIS — CHARTS
# ============================================================
ws = ws_revC
r = sheet_head(ws, 'REVENUE — CHARTS', '')
drill_strip(ws_revC, [('City', cities, 4), ('Month', month_labels, 8)])
place_chart(ws_revC, 0, 'line', 'Monthly Revenue Trend', cref('Revenue - Tables', 'rev_month', 0),
            cref('Revenue - Tables', 'rev_month', 2), 'Net Revenue', ylab='Revenue (₹)', yfmt=M, w=640, h=360)
place_chart(ws_revC, 1, 'column', 'Revenue by City', cref('Revenue - Tables', 'rev_city', 0),
            cref('Revenue - Tables', 'rev_city', 2), 'Net Revenue', ylab='Revenue (₹)', yfmt=M, w=640, h=360)
place_chart(ws_revC, 2, 'pie', 'Revenue Share by Payment Mode', cref('Revenue - Tables', 'rev_pay', 0),
            cref('Revenue - Tables', 'rev_pay', 2), 'Net Revenue', w=520, h=360)

# ============================================================
# 11. CUSTOMER ANALYSIS — TABLES
# ============================================================
ws = ws_custT
r = sheet_head(ws, 'CUSTOMER ANALYSIS — TABLES',
               'Customer revenue = net revenue of delivered orders.')
ans = r
r = r + 1 + 4 + 1
r = kpi_cards(ws, r, [
    ('CUSTOMERS (with orders)', '=COUNTIFS(CC_Orders,">0")', 'N'),
    ('REPEAT CUSTOMERS', '=COUNTIFS(CC_Repeat,1)', 'N'),
    ('REPEAT RATE', '=IFERROR(COUNTIFS(CC_Repeat,1)/COUNTIFS(CC_Orders,">0"),0)', 'P'),
    ('AVG REVENUE / CUSTOMER', '=IFERROR(SUMIFS(D_Del,D_Match,1)/COUNTIFS(CC_Orders,">0"),0)', 'M'),
])
srcC = f'Calc_Customers!$L$2:$Q${CC_LAST}'
guardC = 'COUNTIFS(CC_Orders,">0")'
r = banner(ws, r, 'Q1.   TOP 20 CUSTOMERS BY REVENUE')
rows = [[i] + [tcell(srcC, guardC, i, k) for k in (1, 2, 3, 5, 6)] for i in range(1, 21)]
r = put_table(ws, r, 0, 'top20', ['#', 'Customer ID', 'Name', 'City', 'Orders', 'Revenue'], rows,
              {0: N0, 4: N0, 5: M}, {0: 5, 1: 12, 2: 18, 3: 12, 4: 9, 5: 14})
r = banner(ws, r, 'Q2.   REVENUE CONCENTRATION — TOP 20 vs REST')
conc_rows = [['Top 20 customers', '=20', f'=SUM({arng("Customer - Tables","top20",5)})',
              f'=IFERROR(SUM({arng("Customer - Tables","top20",5)})/SUMIFS(D_Del,D_Match,1),0)'],
             ['All other customers', f'=COUNTIFS(CC_Orders,">0")-20',
              f'=SUMIFS(D_Del,D_Match,1)-SUM({arng("Customer - Tables","top20",5)})',
              f'=IFERROR(1-SUM({arng("Customer - Tables","top20",5)})/SUMIFS(D_Del,D_Match,1),0)']]
r = put_table(ws, r, 0, 'conc', ['Segment', 'Customers', 'Revenue', '% of Revenue'], conc_rows,
              {1: N0, 2: M, 3: P}, {0: 18, 1: 12, 2: 16, 3: 12})
r = banner(ws, r, 'Q3.   CUSTOMER VALUE BY ACQUISITION CHANNEL')
rows = [[ch,
         f'=COUNTIFS(CC_Channel,"{ch}",CC_Orders,">0")',
         f'=SUMIFS(CC_Orders,CC_Channel,"{ch}",CC_Orders,">0")',
         f'=SUMIFS(CC_Revenue,CC_Channel,"{ch}")',
         f'=IFERROR(SUMIFS(CC_Revenue,CC_Channel,"{ch}")/COUNTIFS(CC_Channel,"{ch}",CC_Orders,">0"),0)',
         f'=IFERROR(COUNTIFS(CC_Channel,"{ch}",CC_Repeat,1)/COUNTIFS(CC_Channel,"{ch}",CC_Orders,">0"),0)']
        for ch in channels]
r = put_table(ws, r, 0, 'channel', ['Channel', 'Customers', 'Orders', 'Revenue', 'Avg Revenue / Customer', 'Repeat Rate'],
              rows, {1: N0, 2: N0, 3: M, 4: M, 5: P}, {0: 16, 1: 11, 2: 10, 3: 15, 4: 19, 5: 12})
r = banner(ws, r, 'Q4.   CUSTOMER ORDER-FREQUENCY DISTRIBUTION')
freq = [('1 order', '=COUNTIFS(CC_Orders,1)'), ('2–3 orders', '=COUNTIFS(CC_Orders,">=2",CC_Orders,"<=3")'),
        ('4–6 orders', '=COUNTIFS(CC_Orders,">=4",CC_Orders,"<=6")'),
        ('7–9 orders', '=COUNTIFS(CC_Orders,">=7",CC_Orders,"<=9")'),
        ('10–12 orders', '=COUNTIFS(CC_Orders,">=10",CC_Orders,"<=12")'),
        ('13+ orders', '=COUNTIFS(CC_Orders,">=13")')]
rows = [[lbl, f, f'=IFERROR({f[1:]}/COUNTIFS(CC_Orders,">0"),0)'] for lbl, f in freq]
r = put_table(ws, r, 0, 'freq', ['Order Bucket', 'Customers', '% of Customers'], rows,
              {1: N0, 2: P}, {0: 16, 1: 12, 2: 13})
T20N = acell('Customer - Tables', 'top20', 0, 2); T20I = acell('Customer - Tables', 'top20', 0, 1)
T20C = acell('Customer - Tables', 'top20', 0, 3); T20O = acell('Customer - Tables', 'top20', 0, 4)
T20R = acell('Customer - Tables', 'top20', 0, 5); T20RS = arng('Customer - Tables', 'top20', 5)
CHN = arng('Customer - Tables', 'channel', 0); CHA = arng('Customer - Tables', 'channel', 4)
write_answers(ws, ans, [
    ('Q1. Who are the top 20 customers by revenue?',
     f'="Top customer: "&{T20N}&" ("&{T20I}&", "&{T20C}&") — ₹"&TEXT({T20R},"#,##0")&" across "&TEXT({T20O},"#,##0")&" orders   — full list below"'),
    ('Q2. What percentage of revenue comes from top customers?',
     f'="Top 20 customers account for ₹"&TEXT(SUM({T20RS}),"#,##0")&" = "&TEXT(IFERROR(SUM({T20RS})/SUMIFS(D_Del,D_Match,1),0),"0.0%")'
     f'&" of total revenue, from "&TEXT(COUNTIFS(CC_Orders,">0"),"#,##0")&" active customers"'),
    ('Q3. Which acquisition channel brings the highest-value customers?',
     f'=INDEX({CHN},MATCH(MAX({CHA}),{CHA},0))&" — avg ₹"&TEXT(MAX({CHA}),"#,##0")&" revenue per customer"'),
    ('Q4. How many repeat customers do we have?',
     '=TEXT(COUNTIFS(CC_Repeat,1),"#,##0")&" of "&TEXT(COUNTIFS(CC_Orders,">0"),"#,##0")&" customers placed 2+ orders ("'
     '&TEXT(IFERROR(COUNTIFS(CC_Repeat,1)/COUNTIFS(CC_Orders,">0"),0),"0.0%")&" repeat rate)"'),
])

# ============================================================
# 12. CUSTOMER ANALYSIS — CHARTS
# ============================================================
ws = ws_custC
r = sheet_head(ws, 'CUSTOMER — CHARTS', '')
drill_strip(ws_custC, [('City', cities, 4), ('Channel', channels, 8)])
place_chart(ws_custC, 0, 'column', 'Top 20 Customers by Revenue', cref('Customer - Tables', 'top20', 2),
            cref('Customer - Tables', 'top20', 5), 'Revenue', ylab='Revenue (₹)', yfmt=M, w=640, h=380)
place_chart(ws_custC, 1, 'pie', 'Revenue: Top 20 Customers vs Rest', cref('Customer - Tables', 'conc', 0),
            cref('Customer - Tables', 'conc', 2), 'Revenue', w=520, h=360)
place_chart(ws_custC, 2, 'column', 'Avg Revenue per Customer by Channel', cref('Customer - Tables', 'channel', 0),
            cref('Customer - Tables', 'channel', 4), 'Avg Revenue / Customer', color=GREEN,
            ylab='₹ / customer', yfmt=M, w=640, h=360)
place_chart(ws_custC, 3, 'column', 'Customers by Order Frequency', cref('Customer - Tables', 'freq', 0),
            cref('Customer - Tables', 'freq', 1), 'Customers', color=PURPLE, ylab='Customers', yfmt=N0, w=640, h=360)

# ============================================================
# 13. RESTAURANT PERFORMANCE — TABLES
# ============================================================
ws = ws_restT
r = sheet_head(ws, 'RESTAURANT PERFORMANCE — TABLES',
               'Revenue / order counts use all orders; rankings respond to the filters.')
ans = r
r = r + 1 + 5 + 1
r = kpi_cards(ws, r, [
    ('RESTAURANTS', f'={NREST}', 'N'),
    ('ACTIVE RESTAURANTS', '=COUNTIFS(CR_Orders,">0")', 'N'),
    ('AVG ORDERS / RESTAURANT', '=IFERROR(COUNTIFS(D_Match,1)/COUNTIFS(CR_Orders,">0"),0)', 'N'),
    ('AVG REVENUE / RESTAURANT', '=IFERROR(SUMIFS(D_Del,D_Match,1)/COUNTIFS(CR_Orders,">0"),0)', 'M'),
])
srcRev = f'Calc_Restaurants!$K$2:$S${CR_LAST}'
srcOrd = f'Calc_Restaurants!$U$2:$AC${CR_LAST}'
srcCan = f'Calc_Restaurants!$AE$2:$AM${CR_LAST}'
srcBot = f'Calc_Restaurants!$AO$2:$AW${CR_LAST}'
guardR = 'COUNTIFS(CR_Orders,">0")'
r = banner(ws, r, 'Q1.   TOP 10 RESTAURANTS BY REVENUE')
rows = [[i] + [tcell(srcRev, guardR, i, k) for k in (1, 2, 3, 4, 5, 6, 7)] for i in range(1, 11)]
r = put_table(ws, r, 0, 'top_rest_rev', ['#', 'ID', 'Restaurant', 'City', 'Cuisine', 'Rating', 'Orders', 'Revenue'],
              rows, {0: N0, 5: DEC, 6: N0, 7: M}, {0: 4, 1: 8, 2: 20, 3: 12, 4: 12, 5: 7, 6: 9, 7: 14})
r = banner(ws, r, 'Q2.   TOP 10 RESTAURANTS BY ORDER COUNT')
rows = [[i] + [tcell(srcOrd, guardR, i, k) for k in (1, 2, 3, 4, 5, 6, 7)] for i in range(1, 11)]
r = put_table(ws, r, 0, 'top_rest_ord', ['#', 'ID', 'Restaurant', 'City', 'Cuisine', 'Rating', 'Orders', 'Revenue'],
              rows, {0: N0, 5: DEC, 6: N0, 7: M}, {0: 4, 1: 8, 2: 20, 3: 12, 4: 12, 5: 7, 6: 9, 7: 14})
r = banner(ws, r, 'Q3.   CUISINE POPULARITY')
rows = [[cu,
         f'=COUNTIFS(D_Match,1,D_Cuisine,"{cu}")',
         f'=SUMIFS(D_Del,D_Match,1,D_Cuisine,"{cu}")',
         f'=IFERROR(COUNTIFS(D_Match,1,D_Cuisine,"{cu}")/COUNTIFS(D_Match,1),0)']
        for cu in cuisines]
r = put_table(ws, r, 0, 'cuisine', ['Cuisine', 'Orders', 'Revenue', '% of Orders'], rows,
              {1: N0, 2: M, 3: P}, {0: 14, 1: 11, 2: 15, 3: 12})
r = banner(ws, r, 'Q4.   REVENUE BY RATING BUCKET')
buckets = [('3.5 – 3.9', 3.4, 4.0), ('4.0 – 4.4', 4.0, 4.5), ('4.5 – 5.0', 4.5, 5.1)]
rows = []
for lbl, lo, hi in buckets:
    rows.append([lbl,
                 f'=COUNTIFS(CR_Rating,">="&{lo},CR_Rating,"<"&{hi},CR_Orders,">0")',
                 f'=SUMIFS(D_Del,D_Match,1,D_Rating,">="&{lo},D_Rating,"<"&{hi})',
                 f'=COUNTIFS(D_Match,1,D_Rating,">="&{lo},D_Rating,"<"&{hi})',
                 f'=IFERROR(SUMIFS(D_Del,D_Match,1,D_Rating,">="&{lo},D_Rating,"<"&{hi})/COUNTIFS(CR_Rating,">="&{lo},CR_Rating,"<"&{hi},CR_Orders,">0"),0)'])
r = put_table(ws, r, 0, 'rbucket', ['Rating Bucket', 'Restaurants', 'Revenue', 'Orders', 'Avg Revenue / Restaurant'],
              rows, {1: N0, 2: M, 3: N0, 4: M}, {0: 14, 1: 12, 2: 15, 3: 11, 4: 20})
r = banner(ws, r, 'Q5.   BOTTOM 5 RESTAURANTS BY REVENUE  (underperformers)')
rows = [[i] + [tcell(srcBot, guardR, i, k) for k in (1, 2, 3, 4, 5, 6, 7)] for i in range(1, 6)]
r = put_table(ws, r, 0, 'bottom_rest', ['#', 'ID', 'Restaurant', 'City', 'Cuisine', 'Rating', 'Orders', 'Revenue'],
              rows, {0: N0, 5: DEC, 6: N0, 7: M}, {0: 4, 1: 8, 2: 20, 3: 12, 4: 12, 5: 7, 6: 9, 7: 14})
RR_N = acell('Restaurant - Tables', 'top_rest_rev', 0, 2); RR_R = acell('Restaurant - Tables', 'top_rest_rev', 0, 7)
RO_N = acell('Restaurant - Tables', 'top_rest_ord', 0, 2); RO_O = acell('Restaurant - Tables', 'top_rest_ord', 0, 6)
CU_N = arng('Restaurant - Tables', 'cuisine', 0); CU_O = arng('Restaurant - Tables', 'cuisine', 1)
RB_A = arng('Restaurant - Tables', 'rbucket', 4)
BN_N = acell('Restaurant - Tables', 'bottom_rest', 0, 2); BN_R = acell('Restaurant - Tables', 'bottom_rest', 0, 7)
write_answers(ws, ans, [
    ('Q1. Which restaurants generate the highest revenue?',
     f'={RR_N}&" — ₹"&TEXT({RR_R},"#,##0")&"   (top 10 in the table below)"'),
    ('Q2. Which restaurants receive the most orders?',
     f'={RO_N}&" — "&TEXT({RO_O},"#,##0")&" orders   (top 10 below)"'),
    ('Q3. Which cuisines are most popular?',
     f'=INDEX({CU_N},MATCH(MAX({CU_O}),{CU_O},0))&" — "&TEXT(MAX({CU_O}),"#,##0")&" orders ("'
     f'&TEXT(IFERROR(MAX({CU_O})/COUNTIFS(D_Match,1),0),"0.0%")&" of all orders)"'),
    ('Q4. Do highly-rated restaurants generate more revenue?',
     f'="Yes — restaurants rated 4.5–5.0 average ₹"&TEXT(INDEX({RB_A},3),"#,##0")&" revenue each vs ₹"'
     f'&TEXT(INDEX({RB_A},1),"#,##0")&" for 3.5–3.9 rated   (see scatter + bucket table)"'),
    ('Q5. Last 5 restaurants by revenue (underperformers)',
     f'="Weakest: "&{BN_N}&" — ₹"&TEXT({BN_R},"#,##0")&"   (bottom 5 in the table below)"'),
])

# ============================================================
# 14. RESTAURANT PERFORMANCE — CHARTS
# ============================================================
ws = ws_restC
r = sheet_head(ws, 'RESTAURANT — CHARTS', '')
drill_strip(ws_restC, [('Cuisine', cuisines, 4), ('City', cities, 8)])
place_chart(ws_restC, 0, 'column', 'Top 10 Restaurants by Revenue', cref('Restaurant - Tables', 'top_rest_rev', 2),
            cref('Restaurant - Tables', 'top_rest_rev', 7), 'Revenue', ylab='Revenue (₹)', yfmt=M, w=640, h=360)
place_chart(ws_restC, 1, 'column', 'Top 10 Restaurants by Orders', cref('Restaurant - Tables', 'top_rest_ord', 2),
            cref('Restaurant - Tables', 'top_rest_ord', 6), 'Orders', color=BLUE, ylab='Orders', yfmt=N0, w=640, h=360)
place_chart(ws_restC, 2, 'column', 'Orders by Cuisine', cref('Restaurant - Tables', 'cuisine', 0),
            cref('Restaurant - Tables', 'cuisine', 1), 'Orders', color=AMBER, ylab='Orders', yfmt=N0, w=640, h=360)
place_chart(ws_restC, 3, 'scatter', 'Restaurant Rating vs Revenue', ['Calc_Restaurants', 1, 4, CR_LAST - 1, 4],
            ['Calc_Restaurants', 1, 6, CR_LAST - 1, 6], 'Restaurants', xlab='Avg Rating', xfmt='0.0',
            ylab='Revenue (₹)', yfmt=M, trend=True, w=640, h=380)
place_chart(ws_restC, 4, 'column', 'Avg Revenue per Restaurant by Rating', cref('Restaurant - Tables', 'rbucket', 0),
            cref('Restaurant - Tables', 'rbucket', 4), 'Avg Revenue / Restaurant', color=GREEN,
            ylab='₹ / restaurant', yfmt=M, w=640, h=360)
place_chart(ws_restC, 5, 'column', 'Bottom 5 Restaurants by Revenue', cref('Restaurant - Tables', 'bottom_rest', 2),
            cref('Restaurant - Tables', 'bottom_rest', 7), 'Revenue', color=AMBER, ylab='Revenue (₹)', yfmt=M, w=640, h=360)

# ============================================================
# 15. COUPON ANALYSIS — TABLES
# ============================================================
ws = ws_coupT
r = sheet_head(ws, 'COUPON ANALYSIS — TABLES',
               'A coupon order = any order with discount_amount > 0 (≈10% off).')
ans = r
r = r + 1 + 4 + 1
r = kpi_cards(ws, r, [
    ('COUPON ORDERS', '=SUMIFS(D_Coupon,D_Match,1)', 'N'),
    ('COUPON USAGE RATE', '=IFERROR(SUMIFS(D_Coupon,D_Match,1)/COUNTIFS(D_Match,1),0)', 'P'),
    ('DISCOUNT VALUE GIVEN', '=SUMIFS(D_Discount,D_Match,1)', 'M'),
    ('AVG DISCOUNT / COUPON ORDER', '=IFERROR(SUMIFS(D_Discount,D_Match,1,D_Coupon,1)/COUNTIFS(D_Match,1,D_Coupon,1),0)', 'M'),
])
r = banner(ws, r, 'Q1.   COUPON USAGE — ORDER SPLIT')
rows = [['Coupon orders', '=SUMIFS(D_Coupon,D_Match,1)', '=IFERROR(SUMIFS(D_Coupon,D_Match,1)/COUNTIFS(D_Match,1),0)'],
        ['Non-coupon orders', '=COUNTIFS(D_Match,1)-SUMIFS(D_Coupon,D_Match,1)',
         '=IFERROR(1-SUMIFS(D_Coupon,D_Match,1)/COUNTIFS(D_Match,1),0)']]
r = put_table(ws, r, 0, 'coup_split', ['Group', 'Orders', '% of Orders'], rows, {1: N0, 2: P},
              {0: 16, 1: 12, 2: 12})
r = banner(ws, r, 'Q2.   SPEND: COUPON vs NON-COUPON ORDERS')
def grp(nm, flag):
    return [nm,
            f'=COUNTIFS(D_Match,1,D_Coupon,{flag})',
            f'=IFERROR(SUMIFS(D_Amount,D_Match,1,D_Coupon,{flag})/COUNTIFS(D_Match,1,D_Coupon,{flag}),0)',
            f'=IFERROR(SUMIFS(D_Discount,D_Match,1,D_Coupon,{flag})/COUNTIFS(D_Match,1,D_Coupon,{flag}),0)',
            f'=IFERROR(SUMIFS(D_Net,D_Match,1,D_Coupon,{flag})/COUNTIFS(D_Match,1,D_Coupon,{flag}),0)',
            f'=SUMIFS(D_Net,D_Match,1,D_Coupon,{flag})']
rows = [grp('Coupon users', 1), grp('Non-coupon users', 0)]
r = put_table(ws, r, 0, 'coup_spend', ['Group', 'Orders', 'Avg Gross Basket', 'Avg Discount', 'Avg Net Spend', 'Total Net Revenue'],
              rows, {1: N0, 2: M, 3: M, 4: M, 5: M}, {0: 16, 1: 10, 2: 15, 3: 13, 4: 13, 5: 15})
r = banner(ws, r, 'Q3.   COUPON USAGE BY CITY')
rows = [[c, f'=COUNTIFS(D_Match,1,D_City,"{c}",D_Coupon,1)', f'=COUNTIFS(D_Match,1,D_City,"{c}")',
         f'=IFERROR(COUNTIFS(D_Match,1,D_City,"{c}",D_Coupon,1)/COUNTIFS(D_Match,1,D_City,"{c}"),0)']
        for c in cities]
r = put_table(ws, r, 0, 'coup_city', ['City', 'Coupon Orders', 'Total Orders', 'Coupon Usage %'], rows,
              {1: N0, 2: N0, 3: P}, {0: 12, 1: 13, 2: 12, 3: 14})
r = banner(ws, r, 'Q4.   COUPONS & CUSTOMER RETENTION')
rows = [['Coupon users',
         '=COUNTIFS(CC_Coupon,1,CC_Orders,">0")', '=COUNTIFS(CC_Coupon,1,CC_Repeat,1)',
         '=IFERROR(COUNTIFS(CC_Coupon,1,CC_Repeat,1)/COUNTIFS(CC_Coupon,1,CC_Orders,">0"),0)',
         '=IFERROR(SUMIFS(CC_Orders,CC_Coupon,1,CC_Orders,">0")/COUNTIFS(CC_Coupon,1,CC_Orders,">0"),0)'],
        ['Non-coupon users',
         '=COUNTIFS(CC_Coupon,0,CC_Orders,">0")', '=COUNTIFS(CC_Coupon,0,CC_Repeat,1)',
         '=IFERROR(COUNTIFS(CC_Coupon,0,CC_Repeat,1)/COUNTIFS(CC_Coupon,0,CC_Orders,">0"),0)',
         '=IFERROR(SUMIFS(CC_Orders,CC_Coupon,0,CC_Orders,">0")/COUNTIFS(CC_Coupon,0,CC_Orders,">0"),0)']]
r = put_table(ws, r, 0, 'coup_ret', ['Group', 'Customers', 'Repeat Customers', 'Repeat Rate', 'Avg Orders / Customer'],
              rows, {1: N0, 2: N0, 3: P, 4: DEC}, {0: 16, 1: 12, 2: 15, 3: 13, 4: 18})
CC_N = arng('Coupon - Tables', 'coup_city', 0); CC_U = arng('Coupon - Tables', 'coup_city', 3)
CS_GV = arng('Coupon - Tables', 'coup_spend', 2); CS_N = arng('Coupon - Tables', 'coup_spend', 4)
CR_R = arng('Coupon - Tables', 'coup_ret', 3); CR_O = arng('Coupon - Tables', 'coup_ret', 4)
write_answers(ws, ans, [
    ('Q1. What percentage of orders use coupons?',
     '=TEXT(IFERROR(SUMIFS(D_Coupon,D_Match,1)/COUNTIFS(D_Match,1),0),"0.0%")&" of orders ("'
     '&TEXT(SUMIFS(D_Coupon,D_Match,1),"#,##0")&" of "&TEXT(COUNTIFS(D_Match,1),"#,##0")&") use a coupon"'),
    ('Q2. Do coupon users spend more than non-coupon users?',
     f'="Avg basket ₹"&TEXT(INDEX({CS_GV},1),"#,##0")&" (coupon) vs ₹"&TEXT(INDEX({CS_GV},2),"#,##0")'
     f'&" (no coupon); net spend ₹"&TEXT(INDEX({CS_N},1),"#,##0")&" vs ₹"&TEXT(INDEX({CS_N},2),"#,##0")'),
    ('Q3. Which city uses the most coupons?',
     f'=INDEX({CC_N},MATCH(MAX({CC_U}),{CC_U},0))&" — "&TEXT(MAX({CC_U}),"0.0%")&" of its orders use coupons"'),
    ('Q4. Are coupons helping customer retention?',
     f'="Yes — "&TEXT(INDEX({CR_R},1),"0.0%")&" of coupon users are repeat customers vs "&TEXT(INDEX({CR_R},2),"0.0%")'
     f'&" of non-users; coupon users average "&TEXT(INDEX({CR_O},1),"0.0")&" orders vs "&TEXT(INDEX({CR_O},2),"0.0")'),
])

# ============================================================
# 16. COUPON ANALYSIS — CHARTS
# ============================================================
ws = ws_coupC
r = sheet_head(ws, 'COUPON — CHARTS', '')
drill_strip(ws_coupC, [('Payment Mode', payments, 4), ('City', cities, 8)])
place_chart(ws_coupC, 0, 'pie', 'Orders Using Coupons vs Not', cref('Coupon - Tables', 'coup_split', 0),
            cref('Coupon - Tables', 'coup_split', 1), 'Orders', w=520, h=360)
place_chart(ws_coupC, 1, 'column', 'Average Spend per Order', cref('Coupon - Tables', 'coup_spend', 0),
            cref('Coupon - Tables', 'coup_spend', 4), 'Avg Net Spend', color=BLUE, ylab='₹ / order', yfmt=M, w=640, h=360)
place_chart(ws_coupC, 2, 'column', 'Coupon Usage % by City', cref('Coupon - Tables', 'coup_city', 0),
            cref('Coupon - Tables', 'coup_city', 3), 'Coupon Usage %', color=GREEN, ylab='% of orders',
            yfmt=P, w=640, h=360)
place_chart(ws_coupC, 3, 'column', 'Repeat Customer Rate', cref('Coupon - Tables', 'coup_ret', 0),
            cref('Coupon - Tables', 'coup_ret', 3), 'Repeat Rate', color=PURPLE, ylab='% repeat', yfmt=P, w=640, h=360)

# ============================================================
# 17. CANCELLATION & REFUND — TABLES
# ============================================================
ws = ws_canT
r = sheet_head(ws, 'CANCELLATION & REFUND — TABLES',
               'Rates are shares of all orders; revenue lost = net revenue of affected orders.')
ans = r
r = r + 1 + 4 + 1
r = kpi_cards(ws, r, [
    ('CANCELLATION RATE', '=IFERROR(COUNTIFS(D_Match,1,D_Status,"Cancelled")/COUNTIFS(D_Match,1),0)', 'P'),
    ('REFUND RATE', '=IFERROR(COUNTIFS(D_Match,1,D_Status,"Refunded")/COUNTIFS(D_Match,1),0)', 'P'),
    ('REVENUE LOST — CANCELLATIONS', '=SUMIFS(D_Lost,D_Match,1,D_Status,"Cancelled")', 'M'),
    ('REVENUE LOST — REFUNDS', '=SUMIFS(D_Lost,D_Match,1,D_Status,"Refunded")', 'M'),
])
r = banner(ws, r, 'Q1 & Q2.   ORDER STATUS BREAKDOWN')
rows = [[s, f'=COUNTIFS(D_Match,1,D_Status,"{s}")', f'=SUMIFS(D_Net,D_Match,1,D_Status,"{s}")',
         f'=IFERROR(COUNTIFS(D_Match,1,D_Status,"{s}")/COUNTIFS(D_Match,1),0)'] for s in statuses]
r = put_table(ws, r, 0, 'status', ['Status', 'Orders', 'Net Revenue', '% of Orders'], rows,
              {1: N0, 2: M, 3: P}, {0: 12, 1: 12, 2: 15, 3: 12})
r = banner(ws, r, 'Q3.   MONTHLY REVENUE LOST  (cancellations + refunds)')
rows = [[m, f'=SUMIFS(D_Lost,D_Match,1,D_Month,"{m}")'] for m in month_labels]
r = put_table(ws, r, 0, 'monthly_lost', ['Month', 'Revenue Lost'], rows, {1: M}, {0: 9, 1: 15})
r = banner(ws, r, 'Q4.   TOP 10 RESTAURANTS BY CANCELLATION RATE')
rows = [[i] + [tcell(srcCan, guardR, i, k) for k in (1, 2, 3, 6, 8, 9)] for i in range(1, 11)]
r = put_table(ws, r, 0, 'can_rest', ['#', 'ID', 'Restaurant', 'City', 'Orders', 'Cancelled', 'Cancellation Rate'],
              rows, {0: N0, 4: N0, 5: N0, 6: P}, {0: 4, 1: 8, 2: 20, 3: 12, 4: 9, 5: 11, 6: 16})
CRN = acell('Cancel - Tables', 'can_rest', 0, 2); CRR = acell('Cancel - Tables', 'can_rest', 0, 6)
CRO = acell('Cancel - Tables', 'can_rest', 0, 4)
write_answers(ws, ans, [
    ('Q1. What is the cancellation rate?',
     '=TEXT(IFERROR(COUNTIFS(D_Match,1,D_Status,"Cancelled")/COUNTIFS(D_Match,1),0),"0.0%")&" — "'
     '&TEXT(COUNTIFS(D_Match,1,D_Status,"Cancelled"),"#,##0")&" of "&TEXT(COUNTIFS(D_Match,1),"#,##0")&" orders cancelled"'),
    ('Q2. What is the refund rate?',
     '=TEXT(IFERROR(COUNTIFS(D_Match,1,D_Status,"Refunded")/COUNTIFS(D_Match,1),0),"0.0%")&" — "'
     '&TEXT(COUNTIFS(D_Match,1,D_Status,"Refunded"),"#,##0")&" of "&TEXT(COUNTIFS(D_Match,1),"#,##0")&" orders refunded"'),
    ('Q3. How much revenue is lost due to cancellations?',
     '="₹"&TEXT(SUMIFS(D_Lost,D_Match,1,D_Status,"Cancelled"),"#,##0")&" lost to cancellations (+ ₹"'
     '&TEXT(SUMIFS(D_Lost,D_Match,1,D_Status,"Refunded"),"#,##0")&" refunds = ₹"'
     '&TEXT(SUMIFS(D_Lost,D_Match,1),"#,##0")&" total, "&TEXT(IFERROR(SUMIFS(D_Lost,D_Match,1)/SUMIFS(D_Amount,D_Match,1),0),"0.0%")&" of gross order value)"'),
    ('Q4. Which restaurants have the highest cancellation rate?',
     f'={CRN}&" — "&TEXT({CRR},"0.0%")&" of its "&TEXT({CRO},"#,##0")&" orders cancelled   (top 10 below)"'),
])

# ============================================================
# 18. CANCELLATION & REFUND — CHARTS
# ============================================================
ws = ws_canC
r = sheet_head(ws, 'CANCELLATION & REFUND — CHARTS', '')
drill_strip(ws_canC, [('Status', statuses, 4), ('City', cities, 8)])
place_chart(ws_canC, 0, 'pie', 'Order Status Breakdown', cref('Cancel - Tables', 'status', 0),
            cref('Cancel - Tables', 'status', 1), 'Orders', w=520, h=360)
place_chart(ws_canC, 1, 'line', 'Monthly Revenue Lost', cref('Cancel - Tables', 'monthly_lost', 0),
            cref('Cancel - Tables', 'monthly_lost', 1), 'Revenue Lost', color=AMBER,
            ylab='Revenue lost (₹)', yfmt=M, w=640, h=360)
place_chart(ws_canC, 2, 'column', 'Top 10 Restaurants by Cancellation Rate',
            cref('Cancel - Tables', 'can_rest', 2), cref('Cancel - Tables', 'can_rest', 6),
            'Cancellation Rate', color=AMBER, ylab='Cancellation rate', yfmt=P, w=640, h=380)

# ============================================================
# 19. CUSTOMER CHURN — TABLES
# ============================================================
ws = ws_churnT
r = sheet_head(ws, 'CUSTOMER CHURN — TABLES',
               f'Churned = active customer with no order in the last {CHURN_DAYS} days (cut-off {CHURN_CUTOFF:%d-%b-%Y}).')
ans = r
r = r + 1 + 5 + 1
r = kpi_cards(ws, r, [
    ('CHURNED CUSTOMERS', '=COUNTIFS(CC_Churned,1)', 'N'),
    ('CHURN RATE', '=IFERROR(COUNTIFS(CC_Churned,1)/COUNTIFS(CC_Orders,">0"),0)', 'P'),
    ('ACTIVE CUSTOMERS', '=COUNTIFS(CC_Orders,">0")-COUNTIFS(CC_Churned,1)', 'N'),
    ('REVENUE LOST TO CHURN', '=SUMIFS(CC_Revenue,CC_Churned,1)', 'M'),
])
r = banner(ws, r, 'Q1 & Q2.   ACTIVE vs CHURNED CUSTOMERS')
rows = [['Active customers', '=COUNTIFS(CC_Orders,">0")-COUNTIFS(CC_Churned,1)',
         '=IFERROR((COUNTIFS(CC_Orders,">0")-COUNTIFS(CC_Churned,1))/COUNTIFS(CC_Orders,">0"),0)'],
        ['Churned customers', '=COUNTIFS(CC_Churned,1)',
         '=IFERROR(COUNTIFS(CC_Churned,1)/COUNTIFS(CC_Orders,">0"),0)']]
r = put_table(ws, r, 0, 'churn_split', ['Status', 'Customers', '% of Customers'], rows, {1: N0, 2: P},
              {0: 16, 1: 12, 2: 13})
r = banner(ws, r, 'Q3 & Q4.   CHURN BY CITY')
rows = [[c, f'=COUNTIFS(CC_City,"{c}",CC_Orders,">0")', f'=COUNTIFS(CC_City,"{c}",CC_Churned,1)',
         f'=IFERROR(COUNTIFS(CC_City,"{c}",CC_Churned,1)/COUNTIFS(CC_City,"{c}",CC_Orders,">0"),0)',
         f'=SUMIFS(CC_Revenue,CC_City,"{c}",CC_Churned,1)'] for c in cities]
r = put_table(ws, r, 0, 'churn_city', ['City', 'Customers', 'Churned', 'Churn Rate', 'Revenue Lost'], rows,
              {1: N0, 2: N0, 3: P, 4: M}, {0: 12, 1: 12, 2: 10, 3: 12, 4: 14})
r = banner(ws, r, 'Q5.   TOP 10 HIGH-VALUE CHURNED CUSTOMERS')
srcH = f'Calc_Customers!$T$2:$Y${CC_LAST}'
guardH = 'COUNTIFS(CC_Churned,1)'
rows = [[i] + [tcell(srcH, guardH, i, k) for k in (1, 2, 3, 5, 6)] for i in range(1, 11)]
r = put_table(ws, r, 0, 'hv_churn', ['#', 'Customer ID', 'Name', 'City', 'Orders', 'Lifetime Revenue'],
              rows, {0: N0, 4: N0, 5: M}, {0: 4, 1: 12, 2: 20, 3: 12, 4: 9, 5: 16})
CITN = arng('Churn - Tables', 'churn_city', 0); CITR = arng('Churn - Tables', 'churn_city', 3)
CITL = arng('Churn - Tables', 'churn_city', 4)
HVN = acell('Churn - Tables', 'hv_churn', 0, 2); HVR = acell('Churn - Tables', 'hv_churn', 0, 5)
HVO = acell('Churn - Tables', 'hv_churn', 0, 4)
write_answers(ws, ans, [
    ('Q1. How many customers have churned?',
     f'=TEXT(COUNTIFS(CC_Churned,1),"#,##0")&" of "&TEXT(COUNTIFS(CC_Orders,">0"),"#,##0")'
     f'&" customers have not ordered in the last {CHURN_DAYS} days (since {CHURN_CUTOFF:%d-%b-%Y})"'),
    ('Q2. What is the churn rate?',
     '=TEXT(IFERROR(COUNTIFS(CC_Churned,1)/COUNTIFS(CC_Orders,">0"),0),"0.0%")&" of the active customer base is churned"'),
    ('Q3. Which city has the highest churn?',
     f'=INDEX({CITN},MATCH(MAX({CITR}),{CITR},0))&" — "&TEXT(MAX({CITR}),"0.0%")&" churn rate"'),
    ('Q4. How much revenue is lost due to churn?',
     '="₹"&TEXT(SUMIFS(CC_Revenue,CC_Churned,1),"#,##0")&" — lifetime revenue generated by churned customers ("'
     '&TEXT(IFERROR(SUMIFS(CC_Revenue,CC_Churned,1)/SUMIFS(D_Del,D_Match,1),0),"0.0%")&" of all revenue)"'),
    ('Q5. Who are the high-value churned customers?',
     f'="Top: "&{HVN}&" — ₹"&TEXT({HVR},"#,##0")&" lifetime revenue over "&TEXT({HVO},"#,##0")&" orders   (top 10 below)"'),
])

# ============================================================
# 20. CUSTOMER CHURN — CHARTS
# ============================================================
ws = ws_churnC
r = sheet_head(ws, 'CUSTOMER CHURN — CHARTS', '')
drill_strip(ws_churnC, [('City', cities, 4), ('Status', statuses, 8)])
place_chart(ws_churnC, 0, 'pie', 'Active vs Churned Customers', cref('Churn - Tables', 'churn_split', 0),
            cref('Churn - Tables', 'churn_split', 1), 'Customers', w=520, h=360)
place_chart(ws_churnC, 1, 'column', 'Churn Rate by City', cref('Churn - Tables', 'churn_city', 0),
            cref('Churn - Tables', 'churn_city', 3), 'Churn Rate', color=PURPLE, ylab='Churn rate', yfmt=P, w=640, h=360)
place_chart(ws_churnC, 2, 'column', 'Top 10 High-Value Churned Customers', cref('Churn - Tables', 'hv_churn', 2),
            cref('Churn - Tables', 'hv_churn', 5), 'Lifetime Revenue', color=PURPLE, ylab='Lifetime revenue (₹)',
            yfmt=M, w=640, h=380)

# ============================================================
# 21. DASHBOARD (filter panel + KPIs + charts)
# ============================================================
ws = ws_dash
ws.merge_range(0, 0, 0, 15, 'ZOMATO — CUSTOMER ANALYTICS DASHBOARD', F['title'])
ws.merge_range(1, 0, 1, 15, f'Data period {orders["order_date"].min():%b-%Y} to {DATA_END:%b-%Y}', F['sub'])
ws.set_row(0, 30); ws.set_row(1, 16)

ws.merge_range(3, 0, 3, 15, 'FILTERS',
               wb.add_format({'bold': True, 'font_size': 11, 'font_color': 'white', 'bg_color': DARK,
                              'indent': 1, 'valign': 'vcenter', 'font_name': FN}))
ws.set_row(3, 22)

filter_defs = [('City', cities, 0), ('Channel', channels, 2), ('Payment Mode', payments, 4),
               ('Order Status', statuses, 6), ('Cuisine', cuisines, 8), ('Month', month_labels, 10)]
for (lbl, opts, col) in filter_defs:
    ws.merge_range(4, col, 4, col + 1, lbl, F['flt_lbl'])
    ws.merge_range(5, col, 5, col + 1, '(All)', F['flt_val'])
    ws.data_validation(5, col, 5, col, {'validate': 'list', 'source': ['(All)'] + opts,
                                        'dropdown': True})
ws.set_row(4, 16); ws.set_row(5, 22)

kpi_cards(ws, 8, [
    ('TOTAL REVENUE', '=SUMIFS(D_Del,D_Match,1)', 'M'),
    ('TOTAL ORDERS', '=COUNTIFS(D_Match,1)', 'N'),
    ('AVG ORDER VALUE', '=IFERROR(SUMIFS(D_Del,D_Match,1)/COUNTIFS(D_Match,1,D_Status,"Delivered"),0)', 'M'),
    ('CANCELLATION RATE', '=IFERROR(COUNTIFS(D_Match,1,D_Status,"Cancelled")/COUNTIFS(D_Match,1),0)', 'P'),
])
kpi_cards(ws, 11, [
    ('REFUND RATE', '=IFERROR(COUNTIFS(D_Match,1,D_Status,"Refunded")/COUNTIFS(D_Match,1),0)', 'P'),
    ('COUPON USAGE', '=IFERROR(SUMIFS(D_Coupon,D_Match,1)/COUNTIFS(D_Match,1),0)', 'P'),
    ('CHURN RATE', '=IFERROR(COUNTIFS(CC_Churned,1)/COUNTIFS(CC_Orders,">0"),0)', 'P'),
    ('REPEAT CUSTOMER RATE', '=IFERROR(COUNTIFS(CC_Repeat,1)/COUNTIFS(CC_Orders,">0"),0)', 'P'),
])

r = banner(ws, r=14, text='REVENUE  —  trend & geography')
make_chart('line', 'Monthly Revenue Trend', cref('Revenue - Tables', 'rev_month', 0),
           cref('Revenue - Tables', 'rev_month', 2), 'Net Revenue', ylab='Revenue (₹)', yfmt=M,
           anchor=(15, 0), w=740, h=360)
make_chart('column', 'Revenue by City', cref('Revenue - Tables', 'rev_city', 0),
           cref('Revenue - Tables', 'rev_city', 2), 'Net Revenue', ylab='Revenue (₹)', yfmt=M,
           anchor=(15, 11), w=680, h=360)
banner(ws, 34, 'CUSTOMERS & OPERATIONS  —  status, churn & coupons')
make_chart('pie', 'Order Status Mix', cref('Cancel - Tables', 'status', 0),
           cref('Cancel - Tables', 'status', 1), 'Orders', anchor=(35, 0), w=520, h=340)
make_chart('pie', 'Active vs Churned Customers', cref('Churn - Tables', 'churn_split', 0),
           cref('Churn - Tables', 'churn_split', 1), 'Customers', anchor=(35, 11), w=520, h=340)
banner(ws, 54, 'TOP PERFORMERS  —  restaurants & customers')
make_chart('column', 'Top 10 Restaurants by Revenue', cref('Restaurant - Tables', 'top_rest_rev', 2),
           cref('Restaurant - Tables', 'top_rest_rev', 7), 'Revenue', ylab='Revenue (₹)', yfmt=M,
           anchor=(55, 0), w=680, h=360)
make_chart('column', 'Top 20 Customers by Revenue', cref('Customer - Tables', 'top20', 2),
           cref('Customer - Tables', 'top20', 5), 'Revenue', color=AMBER, ylab='Revenue (₹)', yfmt=M,
           anchor=(55, 11), w=680, h=360)
ws.set_column(0, 15, 11)

# ============================================================
# 22. FILTER NAME CHAINS  (Dashboard panel + per-chart drill cells)
# ============================================================
chaindef('Filter_City', ["Dashboard!$A$6", "'Revenue - Charts'!$E$5", "'Customer - Charts'!$E$5",
                         "'Restaurant - Charts'!$I$5", "'Cancel - Charts'!$I$5", "'Churn - Charts'!$E$5"])
chaindef('Filter_Channel', ["Dashboard!$C$6", "'Customer - Charts'!$I$5"])
chaindef('Filter_Payment', ["Dashboard!$E$6", "'Coupon - Charts'!$E$5"])
chaindef('Filter_Status', ["Dashboard!$G$6", "'Cancel - Charts'!$E$5", "'Churn - Charts'!$I$5"])
chaindef('Filter_Cuisine', ["Dashboard!$I$6", "'Restaurant - Charts'!$E$5"])
chaindef('Filter_Month', ["Dashboard!$K$6", "'Revenue - Charts'!$I$5"])

# ============================================================
# 23. ASSUMPTIONS
# ============================================================
ws = ws_assum
ws.merge_range(0, 0, 0, 8, 'ASSUMPTIONS, DEFINITIONS & DATA NOTES', F['title'])
notes = [
    ('Data source', 'Zomato  Order Data.xlsx — sheets: Customer (4,999), Restaurants (200), Orders (50,000).'),
    ('Analysis period', f'{orders["order_date"].min():%d-%b-%Y} to {DATA_END:%d-%b-%Y}.'),
    ('Net revenue', 'order_amount − discount_amount, counted for DELIVERED orders (cancelled/refunded generate no realised revenue).'),
    ('Gross order value', 'Sum of order_amount across all matching orders regardless of status.'),
    ('AOV', 'Total net revenue ÷ delivered orders.'),
    ('Coupon order', 'discount_amount > 0 (≈10% off). 40.3% of all orders use a coupon.'),
    ('Repeat customer', 'Customer with 2 or more orders in the matching period.'),
    ('Churn definition', f'Active customer (≥1 order overall) with NO order in the last {CHURN_DAYS} days before {DATA_END:%d-%b-%Y} (cut-off {CHURN_CUTOFF:%d-%b-%Y}).'),
    ('Revenue lost to churn', 'Lifetime net revenue historically generated by churned customers.'),
    ('Revenue lost to cancellations/refunds', 'Net revenue of orders with status Cancelled / Refunded.'),
    ('Data quality — city', "One customer had city ' pune' (leading space) — trimmed to 'Pune'."),
    ('Data quality — missing customer', "10 orders reference C00001, absent from the Customer master; included with name/city/channel 'Unknown'."),
]
rr = 2
for k, v in notes:
    ws.write(rr, 0, k, F['thead'])
    ws.merge_range(rr, 1, rr, 8, v, F['txt'])
    ws.set_row(rr, 30)
    rr += 1
ws.set_column(0, 0, 26); ws.set_column(1, 8, 22)

print('analysis tables', len(RG), '| charts', 6 + 3 + 4 + 6 + 4 + 3 + 3)
wb.close()
print('BUILD OK ->', OUT)