import pandas as pd
from functionsV1 import get_df, flows, rends_benchs, ret_data, plot_returns

path_pc = 'C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/DATOS'
path_to_db = f'{path_pc}/Bases de Datos'

colors = pd.read_excel(f'{path_pc}/INDICES.xlsx', sheet_name = 'COLORES')

### ============================================================================= ###
###                                  TABLA CON FX                                 ###
### ============================================================================= ###
#### FX ####
fx = get_df(path_to_db, 'SigCap.db', 'FX', where = "WHERE DATE >= '2023-12-31'")
mxnusd = [0] + fx['MXNUSD'].values.tolist()
fx = fx['EURUSD']

#### FX YTD ####
fx_ytd = get_df(path_to_db, 'SigCap.db', 'FX', where = "WHERE DATE >= '2024-12-31'")
mxnusd_ytd = fx_ytd.MXNUSD.values.tolist()
fx_ytd = fx_ytd.EURUSD

#### HCITY ####
hcity = get_df(path_to_db, 'SigCap.db', 'HCITY', where = 'WHERE DATE >= "2023-12-31"')
hcity['TRANSACTIONS'] = hcity.TRANSACTIONS * mxnusd

#### HCITY YTD ####
hcity_ytd = get_df(path_to_db, 'SigCap.db', 'HCITY', where = 'WHERE DATE >= "2024-12-31"')
hcity_ytd['TRANSACTIONS'] = hcity_ytd.TRANSACTIONS * mxnusd_ytd


### ============================================================================= ###
###                          Graficas 2019-2024 y 12M SCH                         ###
### ============================================================================= ###
# cols = ['CASH', 'LIQUIDITY', 'EQUITIES', 'FIXED_INCOME', 'ALTERNATIVES', 'STRUCTURED_NOTES', 'DERIVATIVES']
cols = ['CASH', 'LQDTY', 'EQTS', 'FIX_INC', 'ALTS', 'STR_NTS', 'DRVTS', 'LOAN_TOTAL']
loans = ['LOAN_USD', 'LOAN_EUR', 'LOAN_MXN', 'LOAN_UNSEC']

def port_flows(banks:list[str], pathDB:str, port:str, conds:str, condsl:str, hc:pd.DataFrame, fx_:pd.Series, names:list[str], dt:str = '"2023-12-31"', googl:bool = False):
    pfolio = pd.DataFrame({})    
    l = []
    for idx, bank in enumerate(banks):
        condsll = f'WHERE OP_LOAN_USD IS NOT NULL AND DATE >= {dt}' if port == 'SCH' and bank == 'SSZ' else condsl
        bnk = get_df(pathDB, f'SigCap_{port}.db', bank, where = conds)
        bnkl = get_df(pathDB, f'SigCap_{port}.db', f'{bank}LOAN', where = condsll)
        if idx == 0:
            pfolio['DATE'] = bnk.DATE
            pfolio[['CASH', 'LQDTY', 'EQTS', 'FIX_INC', 'ALTS', 'DRVTS', 'STR_NTS', 'LOAN_TOTAL', 'CASH_FLOW', 'LOAN_FLOW']] = 0
            if port == 'NFD' and googl:
                bnk.loc[bnk.DATE == '2025-09-19', 'DEPS'] = int(241.195 * 1398)
        pfolio = flows(bnkl, bnk, pfolio)
        l.append(bnkl)
    
    if port == 'NFD':
        nts_str = get_df(pathDB, 'SigCap.db', 'NOTAS_NOMINAL_NFD', conds)
        pfolio['STR_NTS'] = nts_str[['STRNTS_CTI_USD', 'STRNTS_JBR_USD', 'STRNTS_SSZ_USD', 'STRNTS_UBS_USD']].sum(axis = 1) + nts_str['STRNTS_JBR_EUR'] * fx_
        pfolio['LOAN_FLOW'] = l[0].OP_LOAN_EUR*fx_ + l[0].ADV_PAY_LOAN_EUR*fx_ + l[0].OP_LOAN_USD + l[0].ADV_PAY_LOAN_USD +\
                              l[1].OP_LOAN + l[1].ADV_PAY_LOAN + l[2].OP_LOAN + l[2].ADV_PAY_LOAN + l[3].OP_LOAN + l[3].ADV_PAY_LOAN
    
    pfolio['AUMS'] = pfolio[['CASH', 'LQDTY', 'EQTS', 'FIX_INC', 'ALTS', 'DRVTS', 'STR_NTS']].sum(axis = 1)
    pfolio['NAV'] = pfolio.AUMS + pfolio.LOAN_TOTAL
    
    pfolio['HOLD_CASH_FLOW'] = pfolio.CASH_FLOW.cumsum()
    pfolio['HOLD_LOAN_FLOW'] = pfolio.LOAN_FLOW.cumsum()
    
    hc['TRANCUM'] = hc.TRANSACTIONS.cumsum()
    
    pfolio['AUM_BS_ZR'] = pfolio.AUMS - pfolio.HOLD_CASH_FLOW - pfolio.HOLD_LOAN_FLOW
    pfolio['NAV_BS_ZR'] = pfolio.NAV - pfolio.HOLD_CASH_FLOW
    
    pfolio['AUM_SH'] = pfolio.AUMS - pfolio.CASH_FLOW - pfolio.LOAN_FLOW
    pfolio['NAV_SH'] = pfolio.NAV - pfolio.CASH_FLOW
    
    if port == 'SCH':
        pfolio['AUM_BS_ZR_WO_EQTS'] = pfolio.AUM_BS_ZR - hc.HCITY_SCH - hc.TRANCUM
        pfolio['NAV_BS_ZR_WO_EQTS'] = pfolio.NAV_BS_ZR - hc.HCITY_SCH - hc.TRANCUM
    else:
        pfolio['AUM_BS_ZR_WO_EQTS'] = pfolio.AUM_BS_ZR - pfolio.EQTS
        pfolio['NAV_BS_ZR_WO_EQTS'] = pfolio.NAV_BS_ZR - pfolio.EQTS
    
    cols_rend, cols_sharp = ['REND_AUM', 'REND_NAV', 'REND_AUM_WO_EQTS', 'REND_NAV_WO_EQTS'], ['REND_AUM_SHARP', 'REND_NAV_SHARP']
    act_rend, act_sharp = ['AUM_BS_ZR', 'NAV_BS_ZR', 'AUM_BS_ZR_WO_EQTS', 'NAV_BS_ZR_WO_EQTS'], ['AUM_SH', 'NAV_SH']
    pfolio[cols_rend] = pfolio[act_rend] / pfolio[act_rend].iloc[0] - 1
    pfolio[cols_sharp] = (pfolio[act_sharp] / pfolio[act_sharp].shift(+1) - 1).fillna(0)
    
    # rassets = [f'REND_{col}' for col in cols]
    # pfolio[rassets] = pfolio[cols] / pfolio[cols].iloc[0] - 1
    
    return pfolio

rends, benchs = rends_benchs(path_pc, 'INDICES')
rends_ytd, benchs_ytd = rends_benchs(path_pc, 'INDICES_2025')

csch = 'WHERE DEPS IS NOT NULL AND DATE >= "2023-12-31"'
cschl = 'WHERE OP_LOAN IS NOT NULL AND DATE >= "2023-12-31"'
sch = port_flows(['SSZ', 'DBK', 'GSS', 'UBS', 'JPM', 'MSY'], path_to_db, 'SCH', csch, cschl, hcity, fx, cols)

rasset_sch = [f'REND_{col}' for col in cols]
sch[rasset_sch] = sch[cols] / sch[cols].iloc[0] - 1
schr = pd.merge(sch, rends[['DATE', 'AGG', 'ACWI', 'Rusell 2000']], on = 'DATE', how = 'left')

### ================== EQTS ================== ###
sch_ = schr[['DATE', 'REND_EQTS', 'ACWI']].copy()
sch_['RE_ACWI'] = sch_.REND_EQTS - sch_.ACWI
sch_['DIF_POS'] = sch_.RE_ACWI.clip(lower = 0)
sch_['DIF_NEG'] = sch_.RE_ACWI.clip(upper = 0)
sch_ = sch_.drop('RE_ACWI', axis = 1)
x = sch_[sch_.columns[1:]].iloc[-1].sort_values(ascending = False)
plot_returns(path_pc, 'Graficas_PPT/REND_SCH_EQTS', sch_, x, [], [], colors, yTicks = 10, ncols = 7, diffs = True)

### ================== FIX_INC ================== ###
sch_ = schr[['DATE', 'REND_FIX_INC', 'AGG']].copy()
sch_['RE_AGG'] = sch_.REND_FIX_INC - sch_.AGG
sch_['DIF_POS'] = sch_.RE_AGG.clip(lower = 0)
sch_['DIF_NEG'] = sch_.RE_AGG.clip(upper = 0)
sch_ = sch_.drop('RE_AGG', axis = 1)
x = sch_[sch_.columns[1:]].iloc[-1].sort_values(ascending = False)
plot_returns(path_pc, 'Graficas_PPT/REND_SCH_FIXINC', sch_, x, [], [], colors, yTicks = 20, ncols = 7, diffs = True)

### ================== ALTS ================== ###
sch_ = schr[['DATE', 'REND_ALTS', 'Rusell 2000']].copy()
sch_['Rusell 2000'] = sch_['Rusell 2000'] + 0.025
sch_['RE_RTY'] = sch_.REND_ALTS - sch_['Rusell 2000']
sch_['DIF_POS'] = sch_.RE_RTY.clip(lower = 0)
sch_['DIF_NEG'] = sch_.RE_RTY.clip(upper = 0)
sch_ = sch_.drop('RE_RTY', axis = 1)
x = sch_[sch_.columns[1:]].iloc[-1].sort_values(ascending = False)
plot_returns(path_pc, 'Graficas_PPT/REND_SCH_ALTS', sch_, x, [], [], colors, yTicks = 5, ncols = 7, diffs = True)


### ================== EQTS wo HCITY ================== ###
schr['EQTS_wo_HC'] = schr.EQTS - hcity.HCITY_SCH - hcity.TRANSACTIONS.cumsum()
schr['REND_EQTS_wo_HC'] = schr.EQTS_wo_HC / schr.EQTS_wo_HC.iloc[0] - 1

sch_ = schr[['DATE', 'REND_EQTS_wo_HC', 'ACWI']].copy()
sch_['RE_ACWI'] = sch_.REND_EQTS_wo_HC - sch_.ACWI
sch_['DIF_POS'] = sch_.RE_ACWI.clip(lower = 0)
sch_['DIF_NEG'] = sch_.RE_ACWI.clip(upper = 0)
sch_ = sch_.drop('RE_ACWI', axis = 1)
x = sch_[sch_.columns[1:]].iloc[-1].sort_values(ascending = False)
plot_returns(path_pc, 'Graficas_PPT/REND_SCH_EQTS', sch_, x, [], [], colors, yTicks = 10, ncols = 7, diffs = True)