import pandas as pd
from functions import get_df, plot_returns, plot_diffs, Bench60_40, flows

path_pc = 'C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/DATOS'
path_to_db = f'{path_pc}/Bases de Datos'

colors = pd.read_excel(f'{path_pc}/INDICES_2025.xlsx', sheet_name = 'COLORES')

# =============================================================================
#           TABLA CON FX POR SI SE NECESITA CALCULAR ALGO EN DLRS
# =============================================================================
fx = get_df(path_to_db, 'SigCap.db', 'FX', where = "WHERE DATE >= '2024-12-31'")
hcity = get_df(path_to_db, 'SigCap.db', 'HCITY', where = 'WHERE DATE >= "2024-12-31"')

mxnusd = fx['MXNUSD'].values.tolist()
fx = fx['EURUSD']
hcity['TRANSACTIONS'] = hcity.TRANSACTIONS * mxnusd

# =============================================================================
#  Tablas de SCH
# =============================================================================
# Lectura de tablas DBK, GSS, JPM, MSY, SSZ
ssz = get_df(path_to_db, 'SigCap_SCH.db', 'SSZ', where = "WHERE DEPS IS NOT NULL AND DATE>='2024-12-31'")
# ssz.loc[len(ssz)-1, 'EQTS'] = ssz.loc[len(ssz)-1, 'EQTS'] - 787857.23
sszl = get_df(path_to_db, 'SigCap_SCH.db', 'SSZLOAN', where = "WHERE OP_LOAN_USD IS NOT NULL AND DATE>='2024-12-31'")

sch = pd.DataFrame({})
sch['DATE'] = ssz['DATE']
sch[['CASH', 'LQDTY', 'EQTS', 'FIX_INC', 'ALTS', 'DRVTS', 'STR_NTS', 'LOAN_TOTAL', 'CASH_FLOW', 'LOAN_FLOW']] = 0
sch = flows(sszl, ssz, sch)

dbk = get_df(path_to_db, 'SigCap_SCH.db', 'DBK', where = "WHERE DEPS IS NOT NULL AND DATE>='2024-12-31'") 
dbkl = get_df(path_to_db, 'SigCap_SCH.db', 'DBKLOAN', where = "WHERE OP_LOAN IS NOT NULL AND DATE>='2024-12-31'")
sch = flows(dbkl, dbk, sch)

gss = get_df(path_to_db, 'SigCap_SCH.db', 'GSS', where = "WHERE DEPS IS NOT NULL AND DATE>='2024-12-31'")
# gss.loc[len(gss)-1, 'EQTS'] = gss.loc[len(gss)-1, 'EQTS'] - 553597.90
gssl = get_df(path_to_db, 'SigCap_SCH.db', 'GSSLOAN', where = "WHERE OP_LOAN IS NOT NULL AND DATE>='2024-12-31'")
sch = flows(gssl, gss, sch)

ubs = get_df(path_to_db, 'SigCap_SCH.db', 'UBS', where = "WHERE DEPS IS NOT NULL AND DATE>='2024-12-31'")
# ubs.loc[len(ubs)-1, 'EQTS'] = ubs.loc[len(ubs)-1, 'EQTS'] - 147002.66
ubsl = get_df(path_to_db, 'SigCap_SCH.db', 'UBSLOAN', where = "WHERE OP_LOAN IS NOT NULL AND DATE>='2024-12-31'")
sch = flows(ubsl, ubs, sch)

jpm = get_df(path_to_db, 'SigCap_SCH.db', 'JPM', where = "WHERE DEPS IS NOT NULL AND DATE>='2024-12-31'")
# jpm.loc[len(jpm)-1, 'EQTS'] = jpm.loc[len(jpm)-1, 'EQTS'] - 337175.96
jpml = get_df(path_to_db, 'SigCap_SCH.db', 'JPMLOAN', where = "WHERE OP_LOAN IS NOT NULL AND DATE>='2024-12-31'")
sch = flows(jpml, jpm, sch)

msy = get_df(path_to_db, 'SigCap_SCH.db', 'MSY', where = "WHERE DEPS IS NOT NULL AND DATE>='2024-12-31'")
msyl = get_df(path_to_db, 'SigCap_SCH.db', 'MSYLOAN', where = "WHERE OP_LOAN IS NOT NULL AND DATE>='2024-12-31'")
sch = flows(msyl, msy, sch)

sch['AUMS'] = sch[['CASH', 'LQDTY', 'EQTS', 'FIX_INC', 'ALTS', 'DRVTS', 'STR_NTS']].sum(axis = 1)
sch['NAV'] = sch['AUMS'] + sch['LOAN_TOTAL']

sch['HOLD_CASH_FLOW'] = sch.CASH_FLOW.cumsum()
sch['HOLD_LOAN_FLOW'] = sch.LOAN_FLOW.cumsum()
hcity['TRANCUM'] = hcity.TRANSACTIONS.cumsum()

sch['AUM_BS_ZR'] = sch['AUMS'] - sch['HOLD_CASH_FLOW'] - sch['HOLD_LOAN_FLOW']
sch['NAV_BS_ZR'] = sch['NAV'] - sch['HOLD_CASH_FLOW']
 
sch['AUM_SH'] = sch['AUMS'] - sch['CASH_FLOW'] - sch['LOAN_FLOW']
sch['NAV_SH'] = sch['NAV'] - sch['CASH_FLOW']

sch['AUM_BS_ZR_WO_EQTS'] = sch['AUM_BS_ZR'] - hcity['HCITY_SCH'] - hcity['TRANCUM']
sch['NAV_BS_ZR_WO_EQTS'] = sch['NAV_BS_ZR'] - hcity['HCITY_SCH'] - hcity['TRANCUM']

cols_rend, cols_sharp = ['REND_AUM', 'REND_NAV', 'REND_AUM_WO_EQTS', 'REND_NAV_WO_EQTS'], ['REND_AUM_SHARP', 'REND_NAV_SHARP']
act_rend, act_sharp = ['AUM_BS_ZR', 'NAV_BS_ZR', 'AUM_BS_ZR_WO_EQTS', 'NAV_BS_ZR_WO_EQTS'], ['AUM_SH', 'NAV_SH']
sch[cols_rend] = sch[act_rend] / sch[act_rend].iloc[0] - 1
sch[cols_sharp] = (sch[act_sharp] / sch[act_sharp].shift(+1) - 1).fillna(0)

sch = sch[['DATE','REND_AUM','REND_NAV','REND_AUM_SHARP','REND_NAV_SHARP','REND_AUM_WO_EQTS','REND_NAV_WO_EQTS']]
sch.columns = ['DATE','Retorno Efectivo No Apalancado','Retorno Efectivo Apalancado','Retorno para Sharp AUMS', 'Retorno para Sharp NAV','Retorno Efectivo No Apalancado sin HCITY','Retorno Efectivo Apalancado sin HCITY']




benchs = pd.read_excel(f'{path_pc}/INDICES_2025.xlsx', sheet_name = 'INDICES', skiprows = 8)
benchs = benchs.loc[1:].reset_index(drop = True)
benchs['DATE'] = benchs['DATE'].apply(pd.to_datetime)
benchs[benchs.columns[1:]] = benchs[benchs.columns[1:]].apply(pd.to_numeric)

# Calcula los rendimientos de acuerdo a las series
rends = pd.DataFrame({})
rends['DATE'] = benchs['DATE']

ncols = ['SPX', 'ACWI', 'BRK', 'AGG', 'TLT', 'SHV']
acols = ['SPX Index', 'MXWO Index', 'BRK/B US Equity', 'AGG US EQUITY', 'TLT US EQUITY', 'SHV US EQUITY']
rends[ncols] = benchs[acols] / benchs[acols].iloc[0] - 1

rends['SPX/TLT(60/40)'] = Bench60_40(benchs[['DATE', 'SPX Index', 'TLT US EQUITY']], 'SPX Index', 'TLT US EQUITY')
rends['SPX/AGG(60/40)'] = Bench60_40(benchs[['DATE', 'SPX Index', 'AGG US EQUITY']], 'SPX Index', 'AGG US EQUITY')
rends['ACWI/TLT(60/40)'] = Bench60_40(benchs[['DATE', 'MXWO Index', 'TLT US EQUITY']], 'MXWO Index', 'TLT US EQUITY')
rends['ACWI/AGG(60/40)'] = Bench60_40(benchs[['DATE', 'MXWO Index', 'AGG US EQUITY']], 'MXWO Index', 'AGG US EQUITY')
rends['BRK/TLT(60/40)'] = Bench60_40(benchs[['DATE', 'BRK/B US Equity', 'TLT US EQUITY']], 'BRK/B US Equity', 'TLT US EQUITY')
rends['BRK/AGG(60/40)'] = Bench60_40(benchs[['DATE', 'BRK/B US Equity', 'AGG US EQUITY']], 'BRK/B US Equity', 'AGG US EQUITY')

sch = sch[['DATE','Retorno Efectivo No Apalancado','Retorno Efectivo Apalancado','Retorno Efectivo No Apalancado sin HCITY','Retorno Efectivo Apalancado sin HCITY']]

schr = pd.merge(sch, rends[['DATE', 'SPX', 'BRK', 'ACWI', 'SHV', 'AGG', 'TLT', 'SPX/TLT(60/40)', 'SPX/AGG(60/40)', 'ACWI/TLT(60/40)',
                            'ACWI/AGG(60/40)', 'BRK/TLT(60/40)', 'BRK/AGG(60/40)']])

# =============================================================================
#  GRAFICAS SCH
# =============================================================================

# GRAFICA SCH
schr.loc[len(schr)] = [pd.Timestamp(2024, 12, 31, 0)] + [0] * len(schr.columns[1:])
schr = schr.sort_values(by = 'DATE', ignore_index = True)

x = schr[schr.columns[1:]].iloc[-1].sort_values(ascending = True)
ret_cols_hcity = ['Retorno Efectivo No Apalancado sin HCITY', 'Retorno Efectivo Apalancado sin HCITY']
ret_cols = ['Retorno Efectivo No Apalancado', 'Retorno Efectivo Apalancado']

# GRAFICA SCH
plot_returns(path_pc, 'REND_SCH_YTD', schr, x, ret_cols_hcity, ret_cols, colors, 80, 85)

#  GRAFICA SCH sin HCITY
plot_returns(path_pc, 'REND_SCH_SIN_HCITY_YTD', schr, x, ret_cols, ret_cols_hcity, colors, 30, 60)

# Grafica diferencia RE-SPX/TLT
schr['RE_SHV'] = schr['Retorno Efectivo Apalancado'] - schr['SPX/TLT(60/40)']
schr['DIF_POS'] = schr.RE_SHV.clip(lower = 0)
schr['DIF_NEG'] = schr.RE_SHV.clip(upper = 0)
schr = schr[['DATE', 'SPX/TLT(60/40)', 'SPX/AGG(60/40)', 'ACWI/AGG(60/40)', 'Retorno Efectivo Apalancado', 'DIF_POS', 'DIF_NEG']]
axy = [0, 0]
axys = [(0, 15), (0, 40), (0, 5), (0, 0)]

plot_diffs(path_pc, 'REND_SCHvsSHV_YTD', schr, colors , axys, axy)






# =============================================================================
#  Tablas de NFD
# =============================================================================
# Lectura de tablas CTI, JBR, SSZ, UBS
jbr = get_df(path_to_db, 'SigCap_NFD.db', 'JBR', 'WHERE DATE>"2024-12-31"')
jbrl = get_df(path_to_db, 'SigCap_NFD.db', 'JBRLOAN', 'WHERE DATE>"2024-12-31"')
nfd = pd.DataFrame({})
nfd['DATE'] = jbr['DATE']
nfd[['CASH', 'LQDTY', 'EQTS', 'FIX_INC', 'ALTS', 'DRVTS', 'STR_NTS', 'LOAN_TOTAL', 'CASH_FLOW', 'LOAN_FLOW']] = 0
nfd = flows(jbrl, jbr, nfd)

cti = get_df(path_to_db, 'SigCap_NFD.db', 'CTI', 'WHERE DATE>"2024-12-31"')
ctil = get_df(path_to_db, 'SigCap_NFD.db', 'CTILOAN', 'WHERE DATE>"2024-12-31"')
nfd = flows(ctil, cti, nfd)

ssz = get_df(path_to_db, 'SigCap_NFD.db', 'SSZ', 'WHERE DATE>"2024-12-31"')
# ssz.loc[ssz.DATE == '2025-09-19', 'DEPS'] = 241.195 * 1398

sszl = get_df(path_to_db, 'SigCap_NFD.db', 'SSZLOAN', 'WHERE DATE>"2024-12-31"')
nfd = flows(sszl, ssz, nfd)

ubs = get_df(path_to_db, 'SigCap_NFD.db', 'UBS', 'WHERE DATE>"2024-12-31"')
ubsl = get_df(path_to_db, 'SigCap_NFD.db', 'UBSLOAN', 'WHERE DATE>"2024-12-31"')
nfd = flows(ubsl, ubs, nfd)

nts_str = get_df(path_to_db, 'SigCap.db', 'NOTAS_NOMINAL_NFD', 'WHERE DATE>"2024-12-31"')
nfd['STR_NTS'] = nts_str[['STRNTS_CTI_USD', 'STRNTS_JBR_USD', 'STRNTS_SSZ_USD', 'STRNTS_UBS_USD']].sum(axis = 1) + nts_str['STRNTS_JBR_EUR'] * fx
nfd['LOAN_FLOW'] = ctil['OP_LOAN']+ctil['ADV_PAY_LOAN']+jbrl['OP_LOAN_USD']+jbrl['ADV_PAY_LOAN_USD']+jbrl['OP_LOAN_EUR']*fx+jbrl['ADV_PAY_LOAN_EUR']*fx+sszl['OP_LOAN']+sszl['ADV_PAY_LOAN']+ubsl['OP_LOAN']+ubsl['ADV_PAY_LOAN']

nfd['AUMS'] = nfd[['CASH', 'LQDTY', 'EQTS', 'FIX_INC', 'ALTS', 'DRVTS', 'STR_NTS']].sum(axis = 1)
nfd['NAV'] = nfd['AUMS'] + nfd['LOAN_TOTAL']

nfd['HOLD_CASH_FLOW'] = nfd.CASH_FLOW.cumsum()
nfd['HOLD_LOAN_FLOW'] = nfd.LOAN_FLOW.cumsum()

nfd['AUM_BS_ZR'] = nfd['AUMS'] - nfd['HOLD_CASH_FLOW'] - nfd['HOLD_LOAN_FLOW']
nfd['NAV_BS_ZR'] = nfd['NAV'] - nfd['HOLD_CASH_FLOW']

nfd['AUM_SH'] = nfd['AUMS'] - nfd['CASH_FLOW'] - nfd['LOAN_FLOW']
nfd['NAV_SH'] = nfd['NAV'] - nfd['CASH_FLOW']

nfd['AUM_BS_ZR_WO_EQTS'] = nfd['AUM_BS_ZR'] - nfd['EQTS']
nfd['NAV_BS_ZR_WO_EQTS'] = nfd['NAV_BS_ZR'] - nfd['EQTS']

nfd[cols_rend] = nfd[act_rend] / nfd[act_rend].iloc[0] - 1
nfd[cols_sharp] = (nfd[act_sharp] / nfd[act_sharp].shift(+1) - 1).fillna(0)

nfd = nfd[['DATE','REND_AUM','REND_NAV','REND_AUM_SHARP','REND_NAV_SHARP','REND_AUM_WO_EQTS','REND_NAV_WO_EQTS']]
nfd.columns = ['DATE','Retorno Efectivo No Apalancado','Retorno Efectivo Apalancado','Retorno para SHARP AUMS','Retorno para SHARP NAV','Retorno Efectivo No Apalancado sin Equities','Retorno Efectivo Apalancado sin Equities']



nfd = nfd[['DATE','Retorno Efectivo No Apalancado','Retorno Efectivo Apalancado','Retorno Efectivo No Apalancado sin Equities','Retorno Efectivo Apalancado sin Equities']]

nfdr = pd.merge(nfd, rends[['DATE', 'SPX', 'ACWI', 'SHV', 'AGG', 'TLT']], on = 'DATE', how = 'left')

# =============================================================================
#  GRAFICAS NFD
# =============================================================================
nfdr.loc[len(nfdr)] = [pd.Timestamp(2024, 12, 31, 0)] + [0] * (len(nfdr.columns) - 1)
nfdr = nfdr.sort_values(by = 'DATE', ignore_index = True)

x = nfdr.drop('DATE', axis = 1).iloc[-1].sort_values(ascending = True)
# print(x)
ret_cols_eqts = ['Retorno Efectivo No Apalancado sin Equities', 'Retorno Efectivo Apalancado sin Equities']
ret_cols = ['Retorno Efectivo No Apalancado', 'Retorno Efectivo Apalancado']

#  GRAFICA NFD
plot_returns(path_pc, 'REND_NFD_YTD', nfdr, x, ret_cols_eqts, ret_cols, colors)

#  GRAFICA NFD sin equities
plot_returns(path_pc, 'REND_NFD_SIN_EQTS_YTD', nfdr, x, ret_cols, ret_cols_eqts, colors)