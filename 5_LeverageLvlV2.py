import numpy as np
import pandas as pd
from config import PARENT_PATH
from debugg.functionsV1 import get_df, levaumloannet, un_leveraged

path_pc = PARENT_PATH / 'DATOS'
path_to_db = path_pc / 'Bases de datos'
# path_to_db = f'{path_pc}/Bases de datos/'

port_banks = {
    'SCH': [{'GSS':'GS'}, {'UBS':'UBS'}, {'MSY':'MS'}, {'SSZ':'SS'}, {'DBK':'DB'}, {'JPM':'JPM'}],
    'NFD': [{'CTI':'CT'}, {'JBR':'JB'}, {'SSZ':'SS'}, {'UBS':'UBS'}],
    'WK': [{'MSY':'MS'}]
}

### ============================================================================= ###
###                                      PORTS                                    ###
### ============================================================================= ###
ports = []
for port, banks in port_banks.items():
    data = pd.DataFrame({})
    for idx, bank in enumerate(banks):
        bnk = get_df(path_to_db, f'SigCap_{port}.db', list(bank.keys())[0])
        if idx == 0:
            bnk = bnk[bnk.DATE > '2020-01-01']
            data['DATE'] = bnk.DATE
        
        data = levaumloannet(data, bnk, list(bank.values())[0])
    data = data.replace([np.nan, np.inf], 0)
    ports.append(data)
sch, nfd, wk = ports


### ============================================================================= ###
###                                      TOTAL                                    ###
### ============================================================================= ###
total = pd.DataFrame({})
total['DATE'] = nfd.DATE

total['NET_SCH'] = sch[[col for col in sch.columns if 'NET' in col]].sum(axis = 1)
total['NET_NFD'] = nfd[[col for col in nfd.columns if 'NET' in col]].sum(axis = 1)
total['NET_WK'] = wk['NET_MS']

total['LOAN_SCH'] = sch[[col for col in sch.columns if 'LOAN' in col]].sum(axis = 1) * -1
total['LOAN_NFD'] = nfd[[col for col in nfd.columns if 'LOAN' in col]].sum(axis = 1) * -1
total['LOAN_WK'] = wk['LOAN_MS'] * -1

total['LVRLV_SCH'] = total['LOAN_SCH'] / total['NET_SCH']
total['LVRLV_NFD'] = total['LOAN_NFD'] / total['NET_NFD']
total['LVRLV_WK'] = total['LOAN_WK'] / total['NET_WK']

total['NET'] = total.NET_SCH + total.NET_NFD + total.NET_WK
total['LOAN'] = total.LOAN_SCH + total.LOAN_NFD + total.LOAN_WK
total['LVRLV'] = total.LOAN / total.NET

total['PCT_SCH'] = total.NET_SCH / total.NET
total['PCT_NFD'] = total.NET_NFD / total.NET
total['PCT_WK'] = total.NET_WK / total.NET

total['PCT_LVR_SCH'] = total.LVRLV_SCH * total.PCT_SCH
total['PCT_LVR_NFD'] = total.LVRLV_NFD * total.PCT_NFD
total['PCT_LVR_WK'] = total.LVRLV_WK * total.PCT_WK

total['LVR_TOTAL_WEIGHTED'] = total.PCT_LVR_SCH + total.PCT_LVR_NFD + total.PCT_LVR_WK

total['PCT_LVR_SCH_W'] = total.PCT_LVR_SCH / total.LVRLV
total['PCT_LVR_NFD_W'] = total.PCT_LVR_NFD / total.LVRLV
total['PCT_LVR_WK_W'] = total.PCT_LVR_WK / total.LVRLV

total['LVR_SCH_W'] = total.PCT_LVR_SCH_W * total.LVRLV
total['LVR_NFD_W'] = total.PCT_LVR_NFD_W * total.LVRLV
total['LVR_WK_W'] = total.PCT_LVR_WK_W * total.LVRLV

# total.to_excel('C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Contributtion/leverage.xlsx', index = False)
# exit()


### ============================================================================= ###
###                                      PLOTS                                    ###
### ============================================================================= ###
un_leveraged(total, path_pc, 'Graficas_PPT/LEVERAGE_LV', [33, 7], pct = 'PCT_')
un_leveraged(total, path_pc, 'Graficas_PPT/LEVERAGE_LV_SCH', [16.5, 7], ['SCH'], 5.6, True, 2, 1, right = 0.92)
un_leveraged(total, path_pc, 'Graficas_PPT/LEVERAGE_LV_NFD', [16.5, 7], ['NFD'], 25, True, 2, 5, right = 0.92)