import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from functions import get_df, port_flows, plot_returns

### ======== ###
### CREDITOS ###
### ======== ###
future_usd = pd.read_excel('C:/Users/DATOS-INVERSIONES/OneDrive/Escritorio/FWD de Tasas.xlsx', sheet_name = 'SOFR', parse_dates = ['Date'])
future_mxn = pd.read_excel('C:/Users/DATOS-INVERSIONES/OneDrive/Escritorio/FWD de Tasas.xlsx', sheet_name = 'TIIE', parse_dates = ['Date'])
drange = pd.date_range(future_usd.Date.min(), future_usd.Date.max(), freq = 'D')
data_usd = pd.DataFrame(drange, columns = ['Date'])
data_usd = pd.merge_asof(data_usd, future_usd, on = 'Date')
data_mxn = pd.DataFrame(drange, columns = ['Date'])
data_mxn = pd.merge_asof(data_mxn, future_mxn, on = 'Date')

cambio = pd.read_excel('C:/Users/DATOS-INVERSIONES/OneDrive/Escritorio/grid2.xlsx', sheet_name = 'cambios', parse_dates = ['Date'])
ddrange = pd.date_range(cambio.Date.min(), cambio.Date.max(), freq = 'D')
fx = pd.DataFrame(ddrange, columns = ['Date'])
fx = pd.merge_asof(fx, cambio, on = 'Date')

data_usd['ssz'] = (data_usd['Spot'] + 0.9) / 36000 * 3007203
data_mxn['ssz'] = ((data_mxn['Spot'] + 0.9) * fx.fx) / 36000 * 199580412.3
data_usd['jpm_7536'] = (data_usd['Spot'] + 1.5) / 36000 * 10537448.5
data_usd['jpm_6512'] = (data_usd['Spot'] + 0.85) / 36000 * 13989665.2
data_usd['msy'] = (data_usd['Spot'] + 0.92) / 36000 * 5484301.9
data_usd['dbk'] = (data_usd['3 Mo'] + 0.9) / 36000 * 1500000

data_usd['ssz'] = data_usd['ssz'].cumsum() + 3007203
data_mxn['ssz'] = data_mxn['ssz'].cumsum() + (199580412.3 * 0.05422426)
data_usd['jpm_7536'] = data_usd['jpm_7536'].cumsum() + 10537448.5
data_usd['jpm_6512'] = data_usd['jpm_6512'].cumsum() + 13989665.2
data_usd['msy'] = data_usd['msy'].cumsum() + 5484301.9
data_usd['dbk'] = data_usd['dbk'].cumsum() + 1500000

data_usd['credits'] = data_usd[['ssz', 'jpm_7536', 'jpm_6512', 'msy', 'dbk']].sum(axis = 1) + data_mxn['ssz']


### ======== ###
###   NOTAS  ###
### ======== ###
asset = pd.read_excel('C:/Users/DATOS-INVERSIONES/OneDrive/Escritorio/FWD de Tasas.xlsx', sheet_name = 'notas', parse_dates = ['start_date', 'end_date'])
notes = pd.DataFrame({
    'Dates_1': pd.date_range(asset.start_date.iloc[0], asset.end_date.iloc[0], freq = 'D'),
    'Note_1': asset.coupon.iloc[0] / 36000 * asset.quantity.iloc[0],
    'Dates_2': pd.date_range(asset.start_date.iloc[1], asset.end_date.iloc[1], freq = 'D'),
    'Note_2': asset.coupon.iloc[1] / 36000 * asset.quantity.iloc[1],
    'Dates_3': pd.date_range(asset.start_date.iloc[2], asset.end_date.iloc[2], freq = 'D'),
    'Note_3': asset.coupon.iloc[2] / 36000 * asset.quantity.iloc[2],
    'Dates_4': pd.date_range(asset.start_date.iloc[3], asset.end_date.iloc[3], freq = 'D'),
    'Note_4': asset.coupon.iloc[3] / 36000 * asset.quantity.iloc[3],
    'Dates_5': pd.date_range(asset.start_date.iloc[4], asset.end_date.iloc[4], freq = 'D'),
    'Note_5': asset.coupon.iloc[4] / 36000 * asset.quantity.iloc[4],
    'Dates_6': pd.date_range(asset.start_date.iloc[5], asset.end_date.iloc[5], freq = 'D'),
    'Note_6': asset.coupon.iloc[5] / 36000 * asset.quantity.iloc[5],
    'Dates_7': pd.date_range(asset.start_date.iloc[6], asset.end_date.iloc[6], freq = 'D'),
    'Note_7': asset.coupon.iloc[6] / 36000 * asset.quantity.iloc[6],
    'Dates_8': pd.date_range(asset.start_date.iloc[7], asset.end_date.iloc[7], freq = 'D'),
    'Note_8': asset.coupon.iloc[7] / 36000 * asset.quantity.iloc[7],
    'Dates_9': pd.date_range(asset.start_date.iloc[8], asset.end_date.iloc[8], freq = 'D'),
    'Note_9': asset.coupon.iloc[8] / 36000 * asset.quantity.iloc[8],
    'Dates_10': pd.date_range(asset.start_date.iloc[9], asset.end_date.iloc[9], freq = 'D'),
    'Note_10': asset.coupon.iloc[9] / 36000 * asset.quantity.iloc[9],
    'Dates_11': pd.date_range(asset.start_date.iloc[10], asset.end_date.iloc[10], freq = 'D'),
    'Note_11': asset.coupon.iloc[10] / 36000 * asset.quantity.iloc[10],
    'Dates': pd.date_range('2025-10-22', '2027-10-22', freq = 'D'),
})

notes['Note_1'] = notes.Note_1.cumsum() + asset.quantity.iloc[0]
notes['Note_2'] = notes.Note_2.cumsum() + asset.quantity.iloc[1]
notes['Note_3'] = notes.Note_3.cumsum() + asset.quantity.iloc[2]
notes['Note_4'] = notes.Note_4.cumsum() + asset.quantity.iloc[3]
notes['Note_5'] = notes.Note_5.cumsum() + asset.quantity.iloc[4]
notes['Note_6'] = notes.Note_6.cumsum() + asset.quantity.iloc[5]
notes['Note_7'] = notes.Note_7.cumsum() + asset.quantity.iloc[6]
notes['Note_8'] = notes.Note_8.cumsum() + asset.quantity.iloc[7]
notes['Note_9'] = notes.Note_9.cumsum() + asset.quantity.iloc[8]
notes['Note_10'] = notes.Note_10.cumsum() + asset.quantity.iloc[9]
notes['Note_11'] = notes.Note_11.cumsum() + asset.quantity.iloc[10]


### ======== ###
###   LQDTY  ###
### ======== ###
fed = pd.read_excel('C:/Users/DATOS-INVERSIONES/OneDrive/Escritorio/FED Rate Fwd.xlsx', sheet_name = 'FED', parse_dates = ['Date'])
lqdty_ammount = pd.read_excel('C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Contributtion/2025-10-17-ASSET_ALLOC.xlsx', sheet_name = 'CONSOLIDADO')
lqdty_ammount = lqdty_ammount.groupby(['PORTFOLIO', 'ASSET_CLASS'])['MKT_VALUE_USD'].sum()

dddrange = pd.date_range('2025-10-22', '2027-10-22', freq = 'D')
lqdty = pd.DataFrame(dddrange, columns = ['Date'])
lqdty = pd.merge_asof(lqdty, fed, on = 'Date')
lqdty['lqdty'] = lqdty.Rate / 36000 * lqdty_ammount.loc[('SCH', 'LIQUIDITY')]
lqdty['lqdty'] = lqdty.lqdty.cumsum() + lqdty_ammount.loc[('SCH', 'LIQUIDITY')]

### ======== ###
### FIX. INC ###
### ======== ###
path_db = 'C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/DATOS/Bases de datos'

fx_ytd = get_df(path_db, 'SigCap.db', 'FX', where = "WHERE DATE >= '2024-12-31'")
mxnusd_ytd = fx_ytd.MXNUSD.values.tolist()
fx_ytd = fx_ytd.EURUSD

hcity_ytd = get_df(path_db, 'SigCap.db', 'HCITY', where = 'WHERE DATE >= "2024-12-31"')
hcity_ytd['TRANSACTIONS'] = hcity_ytd.TRANSACTIONS * mxnusd_ytd

csch_ytd = 'WHERE DEPS IS NOT NULL AND DATE >= "2024-12-31"'
cschl_ytd = 'WHERE OP_LOAN IS NOT NULL AND DATE >= "2024-12-31"'
cols = ['DATE','Retorno Efectivo No Apalancado','Retorno Efectivo Apalancado','Retorno AUMS', 'Retorno NAV','Retorno Efectivo No Apalancado sin HCITY','Retorno Efectivo Apalancado sin HCITY']
sch = port_flows(['SSZ', 'DBK', 'GSS', 'UBS', 'JPM', 'MSY'], path_db, 'SCH', csch_ytd, cschl_ytd, hcity_ytd, fx_ytd, cols, dt = '"2024-12-31"')

sch = sch.sort_values(by = 'DATE', ignore_index = True)[['DATE', 'Retorno AUMS', 'Retorno NAV','Retorno Efectivo No Apalancado sin HCITY','Retorno Efectivo Apalancado sin HCITY']]
x = sch[sch.columns[1:]].iloc[-1].sort_values(ascending = True)

ret_cols_hcity = ['Retorno Efectivo No Apalancado sin HCITY', 'Retorno Efectivo Apalancado sin HCITY']
ret_cols = ['Retorno Efectivo No Apalancado', 'Retorno Efectivo Apalancado']
colors = pd.read_excel('C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/DATOS/INDICES.xlsx', sheet_name = 'COLORES')

plot_returns('C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/DATOS', 'PRUEBA', sch, x, ret_cols_hcity, ret_cols, colors, 30, 60, 8)
# cols = ['DATE', 'EQTS', 'FIX_INC', 'ALTS']

# dbk = get_df(path_db, 'SigCap_SCH.db', 'DBK', 'WHERE DATE > "2019-12-31"')
# gss = get_df(path_db, 'SigCap_SCH.db', 'GSS', 'WHERE DATE > "2019-12-31"')
# jpm = get_df(path_db, 'SigCap_SCH.db', 'JPM', 'WHERE DATE > "2019-12-31"')
# msy = get_df(path_db, 'SigCap_SCH.db', 'MSY', 'WHERE DATE > "2019-12-31"')
# ssz = get_df(path_db, 'SigCap_SCH.db', 'SSZ', 'WHERE DATE > "2019-12-31"')
# ubs = get_df(path_db, 'SigCap_SCH.db', 'UBS', 'WHERE DATE > "2019-12-31"')
# nav = get_df(path_db, 'SigCap.db', 'PORT_SCH_NAV', 'WHERE DATE > "2019-12-31"')

# fix_inc = dbk.FIX_INC + gss.FIX_INC + jpm.FIX_INC + msy.FIX_INC + ssz.FIX_INC + ubs.FIX_INC
# eqts = dbk.EQTS + gss.EQTS + jpm.EQTS + msy.EQTS + ssz.EQTS + ubs.EQTS
# alts = dbk.ALTS + gss.ALTS + jpm.ALTS + msy.ALTS + ssz.ALTS + ubs.ALTS

# fi = np.log(fix_inc / fix_inc.shift(+1)).fillna(0)
# rend_fi = nav.NAV.iloc[0] * np.exp(fi.cumsum())

# rend_fixinc = np.log(fix_inc / fix_inc.shift(+1)).fillna(0) * nav.NAV
# rend_fixinc = rend_fixinc.cumsum()

# rend_eqts = np.log(eqts / eqts.shift(+1)).fillna(0)
# rend_alts = np.log(alts / alts.shift(+1)).fillna(0)

# series = pd.read_excel('C:/Users/DATOS-INVERSIONES/OneDrive/Escritorio/AGG SPX ACWI Series.xlsx')
# agg = series[series.Date > '2019-12-31'][['Date', 'AGG']]


plt.figure(figsize = (12, 6))

### ======== ###
### CREDITOS ###
### ======== ###
# plt.plot(data_usd.Date, data_usd.ssz, c = 'g', label = 'SSZ_USD')
# plt.plot(data_mxn.Date, data_mxn.ssz, c = 'k', label = 'SSZ_MXN')
# plt.plot(data_usd.Date, data_usd.jpm_7536, c = 'b', label = 'JPM_7536')
# plt.plot(data_usd.Date, data_usd.jpm_6512, c = 'r', label = 'JPM_6512')
# plt.plot(data_usd.Date, data_usd.msy, c = 'orange', label = 'MSY')
# plt.plot(data_usd.Date, data_usd.dbk, c = 'm', label = 'DBK')
# plt.plot(data_usd.Date, data_usd.credits, 'k-.', lw = 1.5, label = 'Credits')

### ======== ###
###   NOTAS  ###
### ======== ###
# plt.plot(notes.Dates, notes.Note_1, label = 'SPX 9.16%')
# plt.plot(notes.Dates, notes.Note_2, label = 'SPX 8.5%')
# plt.plot(notes.Dates, notes.Note_3, label = 'SPX|NDX|RTY 11.70%')
# plt.plot(notes.Dates, notes.Note_4, label = 'SPX|NDX|RTY 9.70%')
# plt.plot(notes.Dates, notes.Note_5, label = 'SPX|NDX|RTY 8.50%')
# plt.plot(notes.Dates, notes.Note_6, label = 'SPX|NDX|NKY 10.02%')
# plt.plot(notes.Dates, notes.Note_7, label = 'SPX|RTY|NKY 10.14%')
# plt.plot(notes.Dates, notes.Note_8, label = 'SPX|NDX|RTY 10.61%')
# plt.plot(notes.Dates, notes.Note_9, label = 'SPX|NDX|RTY 11.62%')
# plt.plot(notes.Dates, notes.Note_10, label = 'SPX|NDX|RTY 9.50%')
# plt.plot(notes.Dates, notes.Note_11, label = 'SPX|SX5E|RTY 14.60%')

### ======== ###
###   LQDTY  ###
### ======== ###
# plt.plot(lqdty.Date, lqdty.lqdty, 'g--', lw = 1.5, label = 'LQDTY')

### ======== ###
### FIX. INC ###
### ======== ###
# plt.plot(dbk.DATE, rend_fi, 'g--', label = 'Fixed Income')
# plt.plot(dbk.DATE, fi, 'b--', label = 'Fixed Income')
# plt.plot(dbk.DATE, rend_eqts, 'b--', label = 'Equities')
# plt.plot(dbk.DATE, rend_alts, 'r--', label = 'Alternatives')
# plt.plot(agg.Date, agg.AGG, 'b--', label = 'AGG')


plt.legend(loc = 'best')
plt.xlabel('Date')
plt.ylabel('Credit')
plt.grid(True)

plt.tight_layout()
plt.show()