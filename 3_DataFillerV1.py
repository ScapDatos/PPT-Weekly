### ============================================= ###
### Notas:
###     - Actualizar la fecha al viernes inmediato anterior.
### ============================================= ###

import sqlite3
import pandas as pd
from config import DATE, MXNUSD, EURUSD, PARENT_PATH
from debugg.functions import del_record, insert_vals, update_vals, filling_blanks, filling_banks

dt = (pd.to_datetime(DATE, format = '%Y-%m-%d') - pd.Timedelta(days = 6)).strftime('%Y-%m-%d')

path_to_db = PARENT_PATH / 'DATOS' / 'Bases de Datos'
drop = True

atypes = {
    'CASH': 'CASH', 'LIQUIDITY': 'LQDTY', 'EQUITIES': 'EQTS', 'FIXED INCOME': 'FIX_INC',
    'ALTERNATIVES': 'ALTS', 'STRUCTURED NOTES': 'STR_NTS', 'DERIVATIVES': 'DRVTS',
    'UNI': 'UNI', 'HCITY': 'HCITY', 'DEPS': 'DEPS', 'WDRS': 'WDRS', 'COLATERAL': 'COL',
    'CREDITO': 'CREDITO'
}
asset = pd.read_excel(PARENT_PATH / 'Contributtion' / f'{DATE}-ASSET_ALLOC.xlsx')
asset_aux = asset[(asset.PORTFOLIO != '-') & (~asset.BANK.isin(['-', 'BVA']))].reset_index(drop = True)
asset_aux['ASSET_CLASS'] = asset_aux.ASSET_CLASS.map(atypes)
dw = pd.read_excel(PARENT_PATH / 'Contributtion' / 'HIST_DEPS_WDRS.xlsx')

loan_acct = asset[asset.ASSET_CLASS == 'CREDITO'].ACCT.unique()
loan_acct2 = [ac if ac != '1352' else '-' for ac in loan_acct]

dw['DATE']= pd.to_datetime(dw['DATE'], format='%d/%m/%Y').dt.strftime('%Y-%m-%d')
dw['ACCT'] = dw.ACCT.astype(str)

# Hacemos un filtro sobre la fecha
dw = dw[dw['DATE'] >= dt]
dw['ACCT'] = dw.ACCT.replace([9057, '9057'], '09057')

portfs = asset_aux.groupby(['PORTFOLIO', 'BANK', 'ACCT'])

### ============================================================================ ###
###                               LLENADO DE TABLAS                              ###
### ============================================================================ ###
if drop:
    del_record(path_to_db, DATE)
    exit()

dbs = []
bnk_port = {'NFD': [], 'SCH': [], 'WK': []}
for (port, bank, acct), group in portfs:
    with sqlite3.connect(f'{path_to_db}/SigCap_{port}.db') as conn:
        if acct in loan_acct:
            acct = '3762' if acct == '3762L' else acct
            loans = group[group.SYMBOL != '-'][['SYMBOL', 'MKT_VALUE_USD']].groupby('SYMBOL').sum()
            
            data = pd.read_sql_query(f'SELECT * FROM {bank}LOAN ORDER BY DATE DESC LIMIT 5;', conn)
            values = [DATE] + [0] * len(data.columns[1:])
            insert_vals(conn, data, f'{bank}LOAN', DATE, values)
            for idx, row in loans.iterrows():
                update_vals(conn, f'{bank}LOAN', idx, row.MKT_VALUE_USD, DATE)
            
            dw_acct = dw[dw.ACCT == acct].reset_index(drop = True)
            if not dw_acct.empty and acct != '3762':
                update_vals(conn, f'{bank}LOAN', dw_acct.TYPE[0], dw_acct.AMMOUNT[0], DATE)
        
        elif acct not in loan_acct2:
            asset_types = group[['ASSET_CLASS', 'MKT_VALUE_USD']].groupby('ASSET_CLASS').sum().round(2).to_dict()['MKT_VALUE_USD']
            dw_acct = dw[(dw.BANK == bank) & (dw.ACCT == acct)].reset_index(drop = True)
            dw_types = {}
            if not dw_acct.empty:
                multipliers = {'EUR': asset_aux['EURUSD'].iloc[0], 'MXN': asset_aux['MXNUSD'].iloc[0]}
                dw_acct['AMMOUNT'] = dw_acct.apply(lambda x: x.AMMOUNT * multipliers.get(x.CURRENCY, 1), axis = 1)
                dw_acct['CURRENCY'] = 'USD'
                dw_types = dw_acct.groupby("TYPE")["AMMOUNT"].sum().round(2).to_dict()
            
            data = pd.read_sql_query(f'SELECT * FROM {bank}{acct} ORDER BY DATE DESC LIMIT 5;', conn)
            values = [DATE] + [asset_types.get(col, 0.0) for col in data.columns[1:-3]]
            values += [sum(values[1:]), dw_types.get('DEPS', 0.0), dw_types.get('WDRS', 0.0)]
            insert_vals(conn, data, f'{bank}{acct}', DATE, values)

for (port, bank, acct), group in portfs:
    with sqlite3.connect(f'{path_to_db}/SigCap_{port}.db') as conn:    
        # Filling blanks
        if f'SigCap_{port}.db' not in dbs:
            filling_blanks(conn, dw, DATE)
            dbs.append(f'SigCap_{port}.db')
        
        if bank not in bnk_port['SCH'] and port == 'SCH':
            bnk = filling_banks(conn, bank, DATE)
            if bnk is not None:
                bnk_port[port].append(bnk)
        if bank not in bnk_port['NFD'] and port == 'NFD':
            bnk = filling_banks(conn, bank, DATE)
            bnk_port[port].append(bnk)
        if bank not in bnk_port['WK'] and port == 'WK':
            bnk = filling_banks(conn, bank, DATE)
            bnk_port[port].append(bnk)

bnk_port['SCH'].append('BNK1135')

### =========================================== ###
###            LLENADO DE TABLAS NAV            ###
### =========================================== ###
for db, port in zip(dbs, bnk_port):
    aux = {}
    with sqlite3.connect(f'{path_to_db}/{db}') as conn:
        cursor = conn.cursor()
        
        for table in bnk_port[port]:
            df = pd.read_sql_query(f'SELECT * FROM {table} ORDER BY DATE DESC LIMIT 5;', conn)
            aux[table] = float(df[df.DATE == DATE].NAV.iloc[0]) if table != 'BNK1135' else float(df[df.DATE == DATE].AUMS.iloc[0])
    
    with sqlite3.connect(f'{path_to_db}/SigCap.db') as conn:
        cursor = conn.cursor()
        data = pd.read_sql_query(f'SELECT * FROM PORT_{port}_NAV ORDER BY DATE DESC LIMIT 5;', conn)
        # values = [DATE] + list(aux.values()) + [sum(aux.values()), 'NO']
        values = [DATE] + [aux.get(bnk, 0.0) if bnk != 'BNK' else aux.get('BNK1135', 0.0) for bnk in data.columns[1:-2]] + [sum(aux.values()), 'NO']
        
        insert_vals(conn, data, f'PORT_{port}_NAV', DATE, values)

with sqlite3.connect(f'{path_to_db}/SigCap.db') as conn:
    cursor = conn.cursor()
    
    ## TABLA FX ##
    fx = pd.read_sql_query('SELECT * FROM FX ORDER BY DATE DESC LIMIT 5;', conn)
    values = [DATE, MXNUSD, EURUSD]
    insert_vals(conn, fx, 'FX', DATE, values)
    
    ## TABLA HCITY ##
    hcity = asset[asset.SHORT_NAME == 'HCITY'][['PORTFOLIO', 'MKT_VALUE_USD']].groupby('PORTFOLIO').sum()
    loan = asset[asset.SYMBOL == 'LOAN_MXN'][['PORTFOLIO', 'SYMBOL', 'MKT_VALUE_USD']].groupby('SYMBOL').sum()
    
    nav_city = pd.read_sql_query('SELECT * FROM HCITY ORDER BY DATE DESC LIMIT 5;', conn)
    values = [DATE, hcity.loc['SCH', 'MKT_VALUE_USD'], hcity.loc['NFD', 'MKT_VALUE_USD'], loan.loc['LOAN_MXN', 'MKT_VALUE_USD'], 0.0, 0.0, 'NO']
    insert_vals(conn, nav_city, 'HCITY', DATE, values)
    
    ## TABLA UNI ##
    try:
        cursor.execute(f'INSERT INTO UNICAJA (DATE, UNI) VALUES ("{DATE}", 0);')
        conn.commit()
    except:
        print(f"{cursor.execute('SELECT DATE FROM UNICAJA ORDER BY DATE DESC LIMIT 1;').fetchone()[0]} already added!")
    
    ## TABLA NOTAS_NOMINAL_NFD ##
    nts = asset[(asset.ASSET_CLASS == 'STRUCTURED NOTES') & (asset.PORTFOLIO == 'NFD')][['BANK', 'EXPOSITION_CCY', 'QUANTITY']]
    nts = nts.groupby(['BANK', 'EXPOSITION_CCY']).sum()
    
    notas = pd.read_sql_query('SELECT * FROM NOTAS_NOMINAL_NFD ORDER BY DATE DESC LIMIT 5;', conn)
    
    values = [DATE, nts.loc[('CTI', 'USD'), 'QUANTITY'], nts.loc[('JBR', 'USD  '), 'QUANTITY'], ### CORREGIR QUE ESTO NO SEA MANUAL
              nts.loc[('JBR', 'EUR  '), 'QUANTITY'], nts.loc[('SSZ', 'USD'), 'QUANTITY'], 0.0]
    insert_vals(conn, notas, 'NOTAS_NOMINAL_NFD', DATE, values)