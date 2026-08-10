### ============================================= ###
### Notas:
###     - Archivos a actualizar antes de ejecutar
###       * .../Nube/DATOS/INDICES.xlsx
### ============================================= ###

import pandas as pd
from config import PARENT_PATH, DATE
from debugg.functionsV1 import get_df, rends_benchs, port_flows, plot_returns, ret_data

path_pc = PARENT_PATH / 'DATOS'
path_to_db = path_pc / 'Bases de Datos'

colors = pd.read_excel(path_pc / 'INDICES.xlsx', sheet_name = 'COLORES')

### ============================================================================= ###
###                                   MSY WK/SCH                                  ###       
### ============================================================================= ###
ncols = ['DATE', 'CASH', 'LIQUIDITY', 'EQUITIES', 'FIXED_INCOME', 'ALTERNATIVES', 'STRUCTURED_NOTES', 'DERIVATIVES', 'LOAN', 'AUMS', 'NAV', 'DEPS', 'WDRS', 'LABEL', 'VLINE']

msy_sch = get_df(path_to_db, 'SigCap_SCH.db', 'MSY', where = "WHERE DATE >= '2023-12-31'")
msy_sch.columns = ncols

msy_wk = get_df(path_to_db, 'SigCap_WK.db', 'MSY', where = "WHERE DATE >= '2023-12-31'")
msy_wk.columns = ncols

### ============================================================================= ###
###                              TABLA CON FX Y HCITY                             ###
### ============================================================================= ###
#### FX ####
fx = get_df(path_to_db, 'SigCap.db', 'FX', where = "WHERE DATE >= '2023-12-31'")
mxnusd = [0] + fx['MXNUSD'].values.tolist()
fx = fx['EURUSD']

#### FX YTD ####
fx_ytd = get_df(path_to_db, 'SigCap.db', 'FX', where = "WHERE DATE >= '2025-12-31'")
mxnusd_ytd = fx_ytd.MXNUSD.values.tolist()
fx_ytd = fx_ytd.EURUSD

#### HCITY ####
hcity = get_df(path_to_db, 'SigCap.db', 'HCITY', where = 'WHERE DATE >= "2023-12-31"')
hcity['TRANSACTIONS'] = hcity.TRANSACTIONS * mxnusd

#### HCITY YTD ####
hcity_ytd = get_df(path_to_db, 'SigCap.db', 'HCITY', where = 'WHERE DATE >= "2025-12-31"')
hcity_ytd['TRANSACTIONS'] = hcity_ytd.TRANSACTIONS * mxnusd_ytd

# with pd.ExcelWriter(PARENT_PATH / 'Contributtion' / 'hcity.xlsx', engine = 'openpyxl') as writer:
#     hcity.to_excel(writer, sheet_name = 'hcity', index = False)
#     hcity_ytd.to_excel(writer, sheet_name = 'hcity_ytd', index = False)
# exit()

### ============================================================================= ###
###                              TABLAS POR PORTAFOLIO                            ###
### ============================================================================= ###
#### SCH ####
csch = 'WHERE DEPS IS NOT NULL AND DATE >= "2023-12-31"'
cschl = 'WHERE OP_LOAN IS NOT NULL AND DATE >= "2023-12-31"'
cols = ['DATE','Retorno Efectivo No Apalancado','Retorno Efectivo Apalancado','Retorno para Sharp AUMS', 'Retorno para Sharp NAV','Retorno Efectivo No Apalancado sin HCITY','Retorno Efectivo Apalancado sin HCITY']
sch = port_flows(['SSZ', 'DBK', 'GSS', 'UBS', 'JPM', 'MSY'], path_to_db, 'SCH', csch, cschl, hcity, fx, cols)

#### SCH YTD ####
csch_ytd = 'WHERE DEPS IS NOT NULL AND DATE >= "2025-12-31"'
cschl_ytd = 'WHERE OP_LOAN IS NOT NULL AND DATE >= "2025-12-31"'
sch_ytd = port_flows(['SSZ', 'DBK', 'GSS', 'UBS', 'JPM', 'MSY'], path_to_db, 'SCH', csch_ytd, cschl_ytd, hcity_ytd, fx_ytd, cols, dt = '"2025-12-31"')

#### NFD ####
cnfd = 'WHERE DATE > "2023-12-31"'
cols = ['DATE','Retorno Efectivo No Apalancado','Retorno Efectivo Apalancado','Retorno para SHARP AUMS','Retorno para SHARP NAV','Retorno Efectivo No Apalancado sin Equities','Retorno Efectivo Apalancado sin Equities']
nfd = port_flows(['JBR', 'CTI', 'SSZ', 'UBS'], path_to_db, 'NFD', cnfd, cnfd, hcity, fx, cols, googl = False)

#### NFD YTD ####
cnfd_ytd = 'WHERE DATE > "2025-12-31"'
nfd_ytd = port_flows(['JBR', 'CTI', 'SSZ', 'UBS'], path_to_db, 'NFD', cnfd_ytd, cnfd_ytd, hcity_ytd, fx_ytd, cols, googl = False)

with pd.ExcelWriter(PARENT_PATH / 'Contributtion' / 'RENDIMIENTOS.xlsx', engine = 'openpyxl') as writer:
    sch.to_excel(writer, sheet_name = 'Rendimientos_SCH', index = False)
    nfd.to_excel(writer, sheet_name = 'Rendimientos_NFD', index = False)
    sch_ytd.to_excel(writer, sheet_name = 'Rendimientos_SCH_YTD', index = False)
    nfd_ytd.to_excel(writer, sheet_name = 'Rendimientos_NFD_YTD', index = False)

# exit() #---------------------------------- DBK - 1778854.71 / 27/03/2026

### ============================================================================= ###
###                              TABLAS DE RENDIMIENTOS                           ###
### ============================================================================= ###
rends, benchs = rends_benchs(path_pc, 'INDICES')
rends_ytd, benchs_ytd = rends_benchs(path_pc, 'INDICES_2025')

sch = sch[['DATE','Retorno Efectivo No Apalancado','Retorno Efectivo Apalancado','Retorno Efectivo No Apalancado sin HCITY','Retorno Efectivo Apalancado sin HCITY']]
nfd = nfd[['DATE','Retorno Efectivo No Apalancado','Retorno Efectivo Apalancado','Retorno Efectivo No Apalancado sin Equities','Retorno Efectivo Apalancado sin Equities']]
sch_ytd = sch_ytd[['DATE','Retorno Efectivo No Apalancado','Retorno Efectivo Apalancado','Retorno Efectivo No Apalancado sin HCITY','Retorno Efectivo Apalancado sin HCITY']]
nfd_ytd = nfd_ytd[['DATE','Retorno Efectivo No Apalancado','Retorno Efectivo Apalancado','Retorno Efectivo No Apalancado sin Equities','Retorno Efectivo Apalancado sin Equities']]

nfdr = pd.merge(nfd, rends[['DATE', 'SPX', 'ACWI', 'SHV', 'AGG', 'TLT']], on = 'DATE', how = 'left')
schr = pd.merge(sch, rends[['DATE', 'SPX', 'BRK', 'ACWI', 'SHV', 'AGG', 'TLT', 'SPX/TLT(60/40)', 'SPX/AGG(60/40)', 'ACWI/TLT(60/40)',
                            'ACWI/AGG(60/40)', 'BRK/TLT(60/40)', 'BRK/AGG(60/40)']])

nfdr_ytd = pd.merge(nfd_ytd, rends_ytd[['DATE', 'SPX', 'ACWI', 'SHV', 'AGG', 'TLT']], on = 'DATE', how = 'left')
schr_ytd = pd.merge(sch_ytd, rends_ytd[['DATE', 'SPX', 'BRK', 'ACWI', 'SHV', 'AGG', 'TLT', 'SPX/TLT(60/40)', 'SPX/AGG(60/40)', 'ACWI/TLT(60/40)',
                            'ACWI/AGG(60/40)', 'BRK/TLT(60/40)', 'BRK/AGG(60/40)']])

### ============================================================================= ###
###                                  GRAFICAS SCH                                 ###
### ============================================================================= ###
schr.loc[len(schr)] = [pd.Timestamp(2023, 12, 31, 0)] + [0] * len(schr.columns[1:])
schr = schr.sort_values(by = 'DATE', ignore_index = True)

x = schr[schr.columns[1:]].iloc[-1].sort_values(ascending = False)
ret_cols_hcity = ['Retorno Efectivo No Apalancado sin HCITY', 'Retorno Efectivo Apalancado sin HCITY']
ret_cols = ['Retorno Efectivo No Apalancado', 'Retorno Efectivo Apalancado']

# GRAFICA SCH
plot_returns(path_pc, 'Graficas_PPT/REND_SCH', schr, x, ret_cols_hcity, ret_cols, colors, ncols = 6, bbox = (0.5, 1.11), top = 0.926)

# GRAFICA SCH sin HCITY
plot_returns(path_pc, 'Graficas_PPT/REND_SCH_SIN_HCITY', schr, x, ret_cols, ret_cols_hcity, colors, bbox = (0.5, 1.13), top = 0.912)

# Grafica diferencia RE-SPX/TLT
schr['RE_SHV'] = schr['Retorno Efectivo Apalancado'] - schr['SPX/TLT(60/40)']
schr['DIF_POS'] = schr.RE_SHV.clip(lower = 0)
schr['DIF_NEG'] = schr.RE_SHV.clip(upper = 0)
schr = schr[['DATE', 'SPX/TLT(60/40)', 'SPX/AGG(60/40)', 'ACWI/AGG(60/40)', 'Retorno Efectivo Apalancado', 'DIF_POS', 'DIF_NEG']]
x = schr[schr.columns[1:]].iloc[-1].sort_values(ascending = False)

plot_returns(path_pc, 'Graficas_PPT/REND_SCHvsSHV', schr, x, [], ['Retorno Efectivo Apalancado'], colors, diffs = True)

### ========================= ###
###            YTD            ###
### ========================= ###
# GRAFICA SCH
schr_ytd.loc[len(schr_ytd)] = [pd.Timestamp(2025, 12, 31, 0)] + [0] * len(schr_ytd.columns[1:])
schr_ytd = schr_ytd.sort_values(by = 'DATE', ignore_index = True)

x_ytd = schr_ytd[schr_ytd.columns[1:]].iloc[-1].sort_values(ascending = False)

# GRAFICA SCH
plot_returns(path_pc, 'Graficas_PPT/REND_SCH_YTD', schr_ytd, x_ytd, ret_cols_hcity, ret_cols, colors, ncols = 6, bbox = (0.5, 1.11), top = 0.926)

# GRAFICA SCH sin HCITY
plot_returns(path_pc, 'Graficas_PPT/REND_SCH_SIN_HCITY_YTD', schr_ytd, x_ytd, ret_cols, ret_cols_hcity, colors, bbox = (0.5, 1.11), top = 0.926)

# Grafica diferencia RE-SPX/TLT
schr_ytd['RE_SHV'] = schr_ytd['Retorno Efectivo Apalancado'] - schr_ytd['SPX/TLT(60/40)']
schr_ytd['DIF_POS'] = schr_ytd.RE_SHV.clip(lower = 0)
schr_ytd['DIF_NEG'] = schr_ytd.RE_SHV.clip(upper = 0)
schr_ytd = schr_ytd[['DATE', 'SPX/TLT(60/40)', 'SPX/AGG(60/40)', 'ACWI/AGG(60/40)', 'Retorno Efectivo Apalancado', 'DIF_POS', 'DIF_NEG']]
x = schr_ytd[schr_ytd.columns[1:]].iloc[-1].sort_values(ascending = False)

plot_returns(path_pc, 'Graficas_PPT/REND_SCHvsSHV_YTD', schr_ytd, x, [], ['Retorno Efectivo Apalancado'], colors, diffs = True)

### ============================================================================= ###
###                                  GRAFICAS NFD                                 ###
### ============================================================================= ###
nfdr.loc[len(nfdr)] = [pd.Timestamp(2023, 12, 31, 0)] + [0] * (len(nfdr.columns) - 1)
nfdr = nfdr.sort_values(by = 'DATE', ignore_index = True)

x = nfdr.drop('DATE', axis = 1).iloc[-1].sort_values(ascending = False)
ret_cols_eqts = ['Retorno Efectivo No Apalancado sin Equities', 'Retorno Efectivo Apalancado sin Equities']
ret_cols = ['Retorno Efectivo No Apalancado', 'Retorno Efectivo Apalancado']

# GRAFICA NFD
plot_returns(path_pc, 'Graficas_PPT/REND_NFD', nfdr, x, ret_cols_eqts, ret_cols, colors, ncols = 7)

# GRAFICA NFD sin equities
# plot_returns(path_pc, 'Graficas_PPT/REND_NFD_SIN_EQTS', nfdr, x, ret_cols, ret_cols_eqts, colors, ncols = 7)

# Grafica diferencia RE-SHV
nfdr['RE_SHV'] = nfdr['Retorno Efectivo Apalancado'] - nfdr.SHV
nfdr['DIF_POS'] = nfdr.RE_SHV.clip(lower = 0)
nfdr['DIF_NEG'] = nfdr.RE_SHV.clip(upper = 0)
nfdr = nfdr[['DATE', 'SHV', 'Retorno Efectivo Apalancado', 'DIF_POS', 'DIF_NEG']]
x = nfdr.drop('DATE', axis = 1).iloc[-1].sort_values(ascending = False)

plot_returns(path_pc, 'Graficas_PPT/REND_NFDvsSHV', nfdr, x, [], ['Retorno Efectivo Apalancado'], colors, diffs = True)

### ========================= ###
###            YTD            ###
### ========================= ###
nfdr_ytd.loc[len(nfdr_ytd)] = [pd.Timestamp(2025, 12, 31, 0)] + [0] * (len(nfdr_ytd.columns) - 1)
nfdr_ytd = nfdr_ytd.sort_values(by = 'DATE', ignore_index = True)

x_ytd = nfdr_ytd.drop('DATE', axis = 1).iloc[-1].sort_values(ascending = False)

# GRAFICA NFD
plot_returns(path_pc, 'Graficas_PPT/REND_NFD_YTD', nfdr_ytd, x_ytd, ret_cols_eqts, ret_cols, colors, ncols = 7)

# GRAFICA NFD sin equities
# plot_returns(path_pc, 'Graficas_PPT/REND_NFD_SIN_EQTS_YTD', nfdr_ytd, x_ytd, ret_cols, ret_cols_eqts, colors, ncols = 7)

### ============================================================================= ###
###                         TABLA DE RENDIMIENTOS DE DEUDA                        ###
### ============================================================================= ###
rends = rends[rends.DATE >= pd.Timestamp.today() - pd.Timedelta(days = 365)].reset_index(drop = True)
benchs = benchs[benchs.DATE >= pd.Timestamp.today() - pd.Timedelta(days = 365)].reset_index(drop = True)
rend_ncols = ['Global HY', 'EM Debt', 'US Floaters', 'US Corp Bonds', 'US TIPS', 'Treasury Short Term', 'Treasury Inter Term', 'Treasury Long Term']
rend_acols = ['LG30TRUU Index', 'EMUSTRUU Index', 'BTFLTRUU Index', 'LUACTRUU Index', 'LBUTTRUU Index', 'VGSH US Equity', 'VGIT US Equity', 'VGLT US Equity']
rends_, x = ret_data(rends, benchs, rend_ncols, rend_acols)

plot_returns(path_pc, 'Graficas_PPT/REND_DEBT', rends_, x, [], [], colors, yTicks = 2, annots = True, ncols = 6, bbox = (0.5, 1.08), top = 0.951)

### ============================================================================= ###
###                         TABLA DE RENDIMIENTOS SECTORES                        ###
### ============================================================================= ###
rend_ncols = ['Technology', 'C. Discretional', 'Financials', 'Industrial', 'Energy', 'Materials', 'Com. Services', 'Health Care', 'C. Stapples', 'Utilities', 'Nasdaq', 'Rusell 2000', 'Dow Jones',
              'Equal Weighted', 'Mag 7', 'SPX']
rend_acols = ['XLK US EQUITY', 'XLY US EQUITY', 'XLF US EQUITY', 'XLI US EQUITY', 'XLE US EQUITY', 'XLB US EQUITY', 'XLC US EQUITY', 'VHT US EQUITY', 'XLP US EQUITY', 'XLU US EQUITY', 'NDX Index',
              'RTY Index', 'DJI Index', 'SPW Index', 'BM7T Index', 'SPX Index']
rends_, x = ret_data(rends, benchs, rend_ncols, rend_acols, diffs = True, indexes = ['Mag 7', 'SPX'])

plot_returns(path_pc, 'Graficas_PPT/REND_SECTORS', rends_, x, [], [], colors, yTicks = 5, annots = True, w = 2, ncols = 8, bbox = (0.5, 1.08), top = 0.951, diffs = True)

### ============================================================================= ###
###                        TABLA DE RENDIMIENTOS ESTRATEGIA                       ###
### ============================================================================= ###
rend_ncols = ['Global', 'Emerging Market', 'Commodities', 'VIX', 'Development Markets']
rend_acols = ['MXWD Index', 'MXEF Index', 'BCOM Index', 'VIX Index', 'DM Index']
rends_, x = ret_data(rends, benchs, rend_ncols, rend_acols)

plot_returns(path_pc, 'Graficas_PPT/REND_STRATEGY', rends_, x, [], [], colors, annots = True, sec_y = True)

### ============================================================================= ###
###                        TABLA DE RENDIMIENTOS JPM FUNDS                        ###
### ============================================================================= ###
rend_ncols = ['SPX', 'JPMUSTC', 'JPGCACU', 'BGBCEAU', 'ROBUSLD', 'JANGLII', 'NASDAQ']
rend_acols = ['SPX Index', 'JPMUSTC LX EQUITY', 'JPGCACU LX EQUITY', 'BGBCEAU LX EQUITY', 'ROBUSLD LX EQUITY', 'JANGLII ID EQUITY', 'NDX Index']
rends_, x = ret_data(rends, benchs, rend_ncols, rend_acols)

plot_returns(path_pc, 'Graficas_PPT/REND_JPM_FUNDS', rends_, x, [], [], colors, yTicks = 5, annots = True, w = 2, ncols = 7)

### ============================================================================= ###
###                        TABLA DE RENDIMIENTOS DE PAISES                        ###
### ============================================================================= ###
rend_ncols = ['Spain', 'Australia', 'Europe', 'United Kingdom', 'Germany', 'France', 'Canada', 'Japan', 'US', 'Taiwan', 'India', 'South Africa', 'Singapore', 'South Korea', 'China', 'Thailand',
              'Peru', 'Chile', 'Mexico', 'Brazil']
rend_acols = ['IBEX Index', 'AS51 Index', 'SX5E Index', 'UKX Index', 'Dax Index', 'CAC Index', 'SPTSX Index', 'NKY Index', 'SPX Index', 'TWSE Index', 'NIFTY Index', 'TOP40 Index', 'STI Index',
              'KOSPI Index', 'SHCOMP Index', 'SET Index', 'MXPE Index', 'IPSA Index', 'MEXBOL Index', 'IBOV Index']
rends_, x = ret_data(rends, benchs, rend_ncols, rend_acols)

plot_returns(path_pc, 'Graficas_PPT/REND_COUNTRIES', rends_, x, [], [], colors, yTicks = 10, annots = True, w = 2, ncols = 9, bbox = (0.5, 1.14), top = 0.905)