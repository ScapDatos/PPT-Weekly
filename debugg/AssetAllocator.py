# =============================================================================
# Para el Asset Allocator Diario
# Notas:
#    - Verificar que las bases de datos estén actualizadas a la fecha reciente
# =============================================================================

import os
import re
import pymysql
import warnings
import pandas as pd

## IMPORTANTE: REVISAR ESTO, QUITA TODAS LAS ADVERTENCIAS. COMENTAR EN CASO DE MODIFICAR EL CODIGO
## A NIVEL PROFUNDO PARA REVISAR ERRORES Y/O ADVERTENCIAS
warnings.filterwarnings('ignore')

# date = (pd.Timestamp.today() - pd.Timedelta(days = 1)).strftime('%Y-%m-%d')
date = '2025-09-12'
date_files = '12092025'

mxnusd = 0.05420
eurusd = 1.1734
# ======================= Dictionary file of all accounts =======================
path_pc = "C:/Users/DATOS-INVERSIONES"
catPortf = pd.read_excel(f'{path_pc}/OneDrive/0. Nube Asset Mgmt/Contributtion/Diccionarios_MySQLMix.xlsx',
                         sheet_name = "Cuentas", converters = {'Cuenta':str}, engine = "openpyxl")
catPortf = catPortf[catPortf.Status.notnull()][['Banco', 'Portafolio', 'Cuenta', 'Empresa']].reset_index(drop = True)
assetClass = pd.read_excel(f'{path_pc}/OneDrive/0. Nube Asset Mgmt/Contributtion/Diccionarios_MySQLMix.xlsx',
                           sheet_name = "Asset_class", converters = {'Cuenta': str}, engine = "openpyxl")

def MktValueUSD(value, ccy): 
    """ Ajusta al monto de acuerdo a la moneda que se elija en ccy usando el 
    tipo de cambio de la primer sección"""
    if ccy == 'MXN':
        return value * 0.05327
    elif ccy == 'EUR':
        return value * 1.1669
    else:
        return value

print('\n', date)
data = pd.read_excel(f'{path_pc}/OneDrive/0. Nube Asset Mgmt/Contributtion/Template_MySQL.xlsx')
data['MKT_VALUE_USD'] = [MktValueUSD(data['MKT_VALUE'][i], data['CCY'][i]) for i in range(len(data))]
main_cols = data.columns

banks = ['DB', 'GS', 'JPM', 'MS']

bank_dirs = {}
for bank in banks:
    files = os.listdir(os.path.join(f'{path_pc}/OneDrive/0. Nube Asset Mgmt/Contributtion', bank))
    assets = []
    for file in files:
        if date_files in file:
            assets.append(os.path.join(f'{path_pc}/OneDrive/0. Nube Asset Mgmt/Contributtion', bank, file))
    bank_dirs[bank] = assets

def portafolio(data, col):
    for idx, row in catPortf.iterrows():
        if str(row.Cuenta) in data:
            return row[col]

# ========================= Connection to all databases =========================
ct_conn = pymysql.connect(host = host, user = user, password = pswd, database = 'CT')
jb_conn = pymysql.connect(host = host, user = user, password = pswd, database = 'JB')
jpm_conn = pymysql.connect(host = host, user = user, password = pswd, database = 'JPM')
ssz_conn = pymysql.connect(host = host, user = user, password = pswd, database = 'SSZ')
ubs_conn = pymysql.connect(host = host, user = user, password = pswd, database = 'UBS')


### =========================================================== ###
###                              UBS                            ###
### =========================================================== ###
ubs_query = f"""
            SELECT
                P.ACCT_NBR,
                (
                    SELECT S.DESCRIPTION
                    FROM UBS.SEC S
                    WHERE P.SEC_SYMBOL = S.SEC_SYMBOL
                    LIMIT 1
                ) AS DESCRIPTION,
                (
                    SELECT S.CUSIP
                    FROM UBS.SEC S
                    WHERE P.SEC_SYMBOL = S.SEC_SYMBOL
                    LIMIT 1
                ) AS CUSIP,
                P.QUANTITY, P.TRADE_AMT, P.BALANCE_DT,
                (
                    SELECT S.MATURITY_DT
                    FROM UBS.SEC S
                    WHERE P.SEC_SYMBOL = S.SEC_SYMBOL
                    LIMIT 1
                ) AS MATURITY_DT, P.SEC_SYMBOL
            FROM POSITION P
            WHERE P.BALANCE_DT = '{date};'
            """
ubs_data = pd.read_sql(ubs_query, con = ubs_conn)
ubs_data['PORTFOLIO'] = ubs_data.ACCT_NBR.apply(portafolio, col = 'Portafolio')
ubs_data['BANK'] = 'UBS'
ubs_data['ACCT'] = ubs_data.ACCT_NBR.str[-4:]
ubs_data['COMPANY'] = ubs_data.ACCT_NBR.apply(portafolio, col = 'Empresa')
ubs_data['ASSET_CLASS'] = ''
ubs_data['SUB_ASSET_CLASS'] = ''
ubs_data['SHORT_NAME'] = ''
ubs_data['INSTR_ID'] = ubs_data.SEC_SYMBOL
ubs_data['DESCRIPTION'] = ubs_data.DESCRIPTION.fillna('')
ubs_data['DESCRIPTION'] = [desc + ' ' + dt.strftime('%m/%d/%y') if 'TREASURY BILL' in desc else desc for desc, dt in zip(ubs_data.DESCRIPTION, ubs_data.MATURITY_DT)]
ubs_data['ISIN'] = ''
ubs_data['SYMBOL'] = ''
ubs_data['UNIT_COST'] = ''
ubs_data['UNIT_COST_CCY'] = 'USD'
ubs_data['CCY'] = 'USD'
ubs_data['MKT_VALUE'] = ubs_data.TRADE_AMT
ubs_data['MKT_VALUE_USD'] = ubs_data.TRADE_AMT
ubs_data['TOTAL_COST'] = ''
ubs_data['UNRELAIZED_GAIN/LOSS'] = ''
ubs_data['FXRATE'] = 1.0
ubs_data['EXPOSITION_CCY'] = 'USD'
ubs_data['Actualizado'] = ''

if not ubs_data.empty:
    data = pd.concat([data, ubs_data[main_cols]], axis = 0)
else:
    print(f'No hay datos de UBS para {date}. Revisar la info.')

# ### =========================================================== ###
# ###                              JPM                            ###
# ### =========================================================== ###
# jpm_query = f"""
#             SELECT
#                 ACCT_NUM, SEC_DESC_1, SEC_DESC_2, SEC_DESC_3, SEC_DESC_4,
#                 SEC_DESC_5, CUSIP, SETTLED_UNITS, SETTLED_COST_LOCAL,
#                 ACCT_BASE_CCY, SEC_CCY, SETTLED_MKT_VAL_LOCAL, ISIN,
#                 SETTLED_MKT_VAL_BASE, SETTLED_UNR_GL_BASE, FX_RATE,
#                 POS_DATE, SEC_NUM
#             FROM POSITION
#             WHERE POS_DATE = '{date}';
#             """
# jpm_cash = f"""
#             SELECT
#                 ACCT_NUM, CCY_LOCAL, EXCHANGE_RATE, SETT_CASH_BAL_BASE,
#                 SETT_CASH_BAL_LOCAL, ACCT_BASE_CCY
#             FROM CASH_BALANCE
#             WHERE DATE_CASH_BALANCE_FILE = '{date}';
#             """
# fvals = {'ERICSSON (LM) TEL-SP ADR    ': 'ERICSSON0XX1', 'JPM LI-LIQ LVNAV FD - USD - W - ACC  ISIN LU1873131988 SEDOL BFM4703  ': 'JOMLILIQ0XX1'}
# jpm_data = pd.read_sql(jpm_query, con = jpm_conn)
# jpm_csh = pd.read_sql(jpm_cash, con = jpm_conn)
# jpm_data['PORTFOLIO'], jpm_csh['PORTFOLIO'] = jpm_data.ACCT_NUM.apply(portafolio, col = 'Portafolio'), jpm_csh.ACCT_NUM.apply(portafolio, col = 'Portafolio')
# jpm_data['BANK'], jpm_csh['BANK'] = 'JPM', 'JPM'
# jpm_data['ACCT'], jpm_csh['ACCT'] = jpm_data.ACCT_NUM.str[-4:], jpm_csh.ACCT_NUM.str[-4:]
# jpm_data['COMPANY'], jpm_csh['COMPANY'] = jpm_data.ACCT_NUM.apply(portafolio, col = 'Empresa'), jpm_csh.ACCT_NUM.apply(portafolio, col = 'Empresa')
# jpm_data['ASSET_CLASS'], jpm_csh['ASSET_CLASS'] = '', ''
# jpm_data['SUB_ASSET_CLASS'], jpm_csh['SUB_ASSET_CLASS'] = '', ''
# jpm_data['SHORT_NAME'], jpm_csh['SHORT_NAME'] = '', ''
# jpm_data['DESCRIPTION'], jpm_csh['DESCRIPTION'] = jpm_data.SEC_DESC_1 + ' ' + jpm_data.SEC_DESC_2 + ' ' + jpm_data.SEC_DESC_3 + ' ' + jpm_data.SEC_DESC_4 + ' ' + jpm_data.SEC_DESC_5, 'CASH'
# jpm_data['INSTR_ID'], jpm_csh['INSTR_ID'] = jpm_data.SEC_NUM.fillna(jpm_data.DESCRIPTION.map(fvals)), jpm_csh.ACCT_NUM.str[-4:] + '-' + jpm_csh.CCY_LOCAL
# jpm_data['SYMBOL'], jpm_csh['SYMBOL'] = '', ''
# jpm_data['QUANTITY'], jpm_csh['QUANTITY'] = jpm_data.SETTLED_UNITS, jpm_csh.SETT_CASH_BAL_LOCAL
# jpm_data['UNIT_COST'], jpm_csh['UNIT_COST'] = jpm_data.SETTLED_COST_LOCAL, jpm_csh.EXCHANGE_RATE #/ jpm_data.SETTLED_UNITS
# jpm_data['UNIT_COST_CCY'], jpm_csh['UNIT_COST_CCY'] = jpm_data.ACCT_BASE_CCY, jpm_csh.ACCT_BASE_CCY
# jpm_data['CCY'], jpm_csh['CCY'] = jpm_data.SEC_CCY, jpm_csh.CCY_LOCAL
# jpm_data['MKT_VALUE'], jpm_csh['MKT_VALUE'] = jpm_data.SETTLED_MKT_VAL_BASE, jpm_csh.SETT_CASH_BAL_LOCAL
# jpm_data['MKT_VALUE_USD'], jpm_csh['MKT_VALUE_USD'] = jpm_data.SETTLED_MKT_VAL_BASE, jpm_csh.SETT_CASH_BAL_BASE
# jpm_data['TOTAL_COST'], jpm_csh['TOTAL_COST'] = jpm_data.SETTLED_COST_LOCAL, jpm_csh.SETT_CASH_BAL_LOCAL
# jpm_data['UNRELAIZED_GAIN/LOSS'], jpm_csh['UNRELAIZED_GAIN/LOSS'] = jpm_data.SETTLED_UNR_GL_BASE, 0
# jpm_data['FXRATE'], jpm_csh['FXRATE'] = jpm_data.FX_RATE, jpm_csh.EXCHANGE_RATE
# jpm_data['EXPOSITION_CCY'], jpm_csh['EXPOSITION_CCY'] = jpm_data.ACCT_BASE_CCY, jpm_csh.CCY_LOCAL
# jpm_data['Actualizado'], jpm_csh['Actualizado'] = '', ''
# jpm_csh['CUSIP'] = '0' + jpm_csh.CCY_LOCAL + 'PRAA7'
# jpm_csh['ISIN'] = ''

# if not jpm_data.empty:
#     data = pd.concat([data, jpm_data[main_cols], jpm_csh[main_cols]], axis = 0)
# else:
#     print(f'No hay datos de JP Morgan para {date}. Revisar la info.')

def insert_val(data):
    if 'CUENTA CORRIENTE' in data.DESCRIPTION:
        # print(data)
        if data.UNIT_COST_CCY.strip() == 'USD':
            data['INSTR_ID'] = '0'
        elif data.UNIT_COST_CCY.strip() == 'MXN':
            data['INSTR_ID'] = '1'
        elif data.UNIT_COST_CCY.strip() == 'EUR':
            data['INSTR_ID'] = '2'
        else:
            data['INSTR_ID'] = data.INSTR_ID
    elif 'ANTICIPO A' in data.DESCRIPTION:
        if data.UNIT_COST_CCY.strip() == 'USD':
            data['INSTR_ID'] = str(data['INSTR_ID']) + '-1'
        elif data.UNIT_COST_CCY.strip() == 'MXN':
            data['INSTR_ID'] = str(data['INSTR_ID']) + '-2'
        elif data.UNIT_COST_CCY.strip() == 'EUR':
            data['INSTR_ID'] = str(data['INSTR_ID']) + '-3'
        else:
            data['INSTR_ID'] = data.INSTR_ID
    
    return data

### =========================================================== ###
###                              JB                             ###
### =========================================================== ###
jb_query = f"""
            SELECT
                ZRNR, TITEL_WRG_TEXT, KTWRG_TEXT, BESTAND, LETZT_BUSALDO,
                KURZ_TEXT, AUSGABE_ZINSSATZ, BEZ, BEZ_GEGENPARTEI, BEK,
                TITELKURS2, UMRECH_KURS, KURSWERT, MARCHZINS, PER_DATUM,
                ISIN, TITEL
            FROM POSITION
            WHERE PER_DATUM = '{date}';
            """
jb_data = pd.read_sql(jb_query, con = jb_conn)
notes = pd.read_excel(f'{path_pc}/OneDrive/0. Nube Asset Mgmt/Notas Estructuradas/Graficas_Barreras.xlsx', sheet_name = 'NOTAS', engine = 'openpyxl')
jb_data = pd.merge(jb_data, notes[['ISIN', 'SHORT_NAME']], on = 'ISIN', how = 'left')
jb_data['KURZ_TEXT'] = jb_data.SHORT_NAME.combine_first(jb_data.KURZ_TEXT)

jb_data['PORTFOLIO'] = jb_data.ZRNR.apply(portafolio, col = 'Portafolio')
jb_data['BANK'] = 'JB'
jb_data['ACCT'] = jb_data.ZRNR.str[-4:]
jb_data['COMPANY'] = jb_data.ZRNR.apply(portafolio, col = 'Empresa')
jb_data['ASSET_CLASS'] = ''
jb_data['SUB_ASSET_CLASS'] = ''
jb_data['SHORT_NAME'] = ''
jb_data['INSTR_ID'] = [str(isin) if isin.strip() != '' else str(int(titel)) for isin, titel in zip(jb_data.ISIN, jb_data.TITEL)]
jb_data['DESCRIPTION'] = jb_data.KURZ_TEXT + ' ' + jb_data.AUSGABE_ZINSSATZ.astype(str) + '% ' + jb_data.BEZ + ' ' + jb_data.BEZ_GEGENPARTEI
jb_data['CUSIP'] = ''
jb_data['SYMBOL'] = ''
jb_data['QUANTITY'] = [let if 'CUENTA CORRIENTE' in bez else bes for bes, bez, let in zip(jb_data.BESTAND, jb_data.BEZ, jb_data.LETZT_BUSALDO)]
jb_data['UNIT_COST'] = ''
jb_data['UNIT_COST_CCY'] = [let if 'CUENTA CORRIENTE' in bez else bes for bes, bez, let in zip(jb_data.TITEL_WRG_TEXT, jb_data.BEZ, jb_data.KTWRG_TEXT)]
jb_data['CCY'] = 'USD'
jb_data['MKT_VALUE'] = jb_data.KURSWERT + jb_data.MARCHZINS
jb_data['MKT_VALUE_USD'] = jb_data.KURSWERT + jb_data.MARCHZINS
jb_data['TOTAL_COST'] = jb_data.BEK
jb_data['UNRELAIZED_GAIN/LOSS'] = ''
jb_data['FXRATE'] = jb_data.UMRECH_KURS
jb_data['EXPOSITION_CCY'] = jb_data.UNIT_COST_CCY
jb_data['Actualizado'] = ''
jb_data = jb_data.apply(insert_val, axis = 1)
# jb_data = jb_data[jb_data.ZRNR != '3176542'].reset_index(drop = True)

if not jb_data.empty:
    data = pd.concat([data, jb_data[main_cols]], axis = 0)
else:
    print(f'No hay datos de Julius Baer para {date}. Revisar la info.')

### =========================================================== ###
###                              CT                             ###
### =========================================================== ###
ct_query = f"""
            SELECT
                ACCT_NBR, FIN_INSTR_LONG_LINE1_DESC, FIN_INSTR_LONG_LINE2_DESC,
                FIN_INSTR_LONG_LINE3_DESC, DEF_TICKER_SYMB_CD, REMN_UNT, 
                REMN_COST_NOM_AMT, ACCT_REF_CCY_CD, NOM_CCY_CD, NOM_AMT,
                MKT_VAL_AMT, FX_RT, POSN_AS_OF_DT, FIN_INSTR_ID,
                ASSET_SUB_SUB_CLAS_CD, HIDG_COST_UNY_PRC_AMT, NOM_PRIN_AMT,
                UREL_GNLS_AMT
            FROM POSITION
            WHERE POSN_AS_OF_DT = '{date}';
            """
ct_data = pd.read_sql(ct_query, con = ct_conn)

ct_data['PORTFOLIO'] = ct_data.ACCT_NBR.apply(portafolio, col = 'Portafolio')
ct_data['BANK'] = 'CTI'
ct_data['ACCT'] = ct_data.ACCT_NBR.str[-5:]
ct_data['COMPANY'] = ct_data.ACCT_NBR.apply(portafolio, col = 'Empresa')
ct_data['ASSET_CLASS'] = ''
ct_data['SUB_ASSET_CLASS'] = ''
ct_data['SHORT_NAME'] = ''
ct_data['INSTR_ID'] = [sub_clas if instr_id is None else instr_id for instr_id, sub_clas in zip(ct_data.FIN_INSTR_ID, ct_data.ASSET_SUB_SUB_CLAS_CD)]
ct_data['DESCRIPTION'] = ct_data.FIN_INSTR_LONG_LINE1_DESC.fillna('') + ' ' + ct_data.FIN_INSTR_LONG_LINE2_DESC.fillna('') + ' ' + ct_data.FIN_INSTR_LONG_LINE3_DESC.fillna('')
ct_data['ISIN'] = ''
ct_data['CUSIP'] = ''
ct_data['SYMBOL'] = ct_data.DEF_TICKER_SYMB_CD
ct_data['QUANTITY'] = ct_data.REMN_UNT.combine_first(ct_data.NOM_PRIN_AMT).fillna(0)
ct_data['UNIT_COST'] = ct_data.HIDG_COST_UNY_PRC_AMT
ct_data['UNIT_COST_CCY'] = ct_data.ACCT_REF_CCY_CD
ct_data['CCY'] = ct_data.NOM_CCY_CD
ct_data['MKT_VALUE'] = [-amt if sclas == 'CLS FL18' else amt for amt, sclas in zip(ct_data.NOM_AMT, ct_data.ASSET_SUB_SUB_CLAS_CD)]
ct_data['MKT_VALUE_USD'] = [-amt if sclas == 'CLS FL18' else amt for amt, sclas in zip(ct_data.MKT_VAL_AMT, ct_data.ASSET_SUB_SUB_CLAS_CD)]
ct_data['TOTAL_COST'] = ct_data.MKT_VALUE_USD.where(ct_data.INSTR_ID == 'G9554E103', ct_data.REMN_COST_NOM_AMT / ct_data.FX_RT)
ct_data['UNRELAIZED_GAIN/LOSS'] = ct_data.UREL_GNLS_AMT / ct_data.FX_RT
ct_data['FXRATE'] = 1 / ct_data.FX_RT
ct_data['EXPOSITION_CCY'] = ct_data.NOM_CCY_CD
ct_data['Actualizado'] = ''

if not ct_data.empty:
    data = pd.concat([data, ct_data[main_cols]], axis = 0)
else:
    print(f'No hay datos de CITI para {date}. Revisar la info.')

### =========================================================== ###
###                              SSZ                            ###
### =========================================================== ###
ssz_query = f"""
            SELECT
                PORT_ID, INSTRUMENT_NAME, ISIN, NOMINAL_AMNT, 
                COST_PRICE, COST_PRICE_CCY, VALUE_PRICE_CCY,
                FX_RATE, PROC_DATE, IBAN
            FROM POSITION
            WHERE PROC_DATE = '{date}';
            """
ssz_data = pd.read_sql(ssz_query, con = ssz_conn)
ssz_data['ISIN'] = ssz_data.IBAN.combine_first(ssz_data.ISIN)

ssz_data['PORTFOLIO'] = ssz_data.PORT_ID.apply(portafolio, col = 'Portafolio')
ssz_data['BANK'] = 'SSZ'
ssz_data['ACCT'] = ssz_data.PORT_ID.str[:-2]
ssz_data['COMPANY'] = ssz_data.PORT_ID.apply(portafolio, col = 'Empresa')
ssz_data['ASSET_CLASS'] = ''
ssz_data['SUB_ASSET_CLASS'] = ''
ssz_data['SHORT_NAME'] = ''
ssz_data['INSTR_ID'] = ssz_data.ISIN.fillna('SSZLOAN0')
ssz_data['DESCRIPTION'] = ssz_data.INSTRUMENT_NAME
ssz_data['CUSIP'] = ''
ssz_data['SYMBOL'] = ''
ssz_data['QUANTITY'] = [0 if amnt < 0 else amnt for amnt in ssz_data.NOMINAL_AMNT]
ssz_data['UNIT_COST'] = ssz_data.COST_PRICE
ssz_data['UNIT_COST_CCY'] = ssz_data.COST_PRICE_CCY
ssz_data['CCY'] = 'USD'
ssz_data['MKT_VALUE'] = ssz_data.VALUE_PRICE_CCY * ssz_data.FX_RATE
ssz_data['MKT_VALUE_USD'] = ssz_data.VALUE_PRICE_CCY * ssz_data.FX_RATE
ssz_data['TOTAL_COST'] = 0
ssz_data['UNRELAIZED_GAIN/LOSS'] = (ssz_data.VALUE_PRICE_CCY - (ssz_data.COST_PRICE * ssz_data.NOMINAL_AMNT)) * ssz_data.FX_RATE
ssz_data['FXRATE'] = ssz_data.FX_RATE
ssz_data['EXPOSITION_CCY'] = ssz_data.COST_PRICE_CCY
ssz_data['Actualizado'] = ''

if not ssz_data.empty:
    data = pd.concat([data, ssz_data[main_cols]], axis = 0)
else:
    print(f'No hay datos de Santander Suiza para {date}. Revisar la info.\n')

for bank, dirs in bank_dirs.items():
    if bank == 'DB':
        print("Using DB holdings...")
        # Busca en los archivo de la carpeta DB
        for file in dirs:
            # lee el excel
            df_aux = pd.read_excel(file, engine = 'openpyxl', header = 12).drop('Unnamed: 0', axis = 1)
            
            df_aux['PORTFOLIO'] = df_aux['Account Number'].astype(str).str[-4:]#.apply(portafolio, col = 'Portafolio')
            df_aux['BANK'] = 'DBK'
            df_aux['ACCT'] = df_aux['Account Number'].astype(str).str[-4:]
            df_aux['COMPANY'] = df_aux['Account Number'].astype(str).str[-4:].apply(portafolio, col = 'Empresa')
            df_aux['ASSET_CLASS'] = ''
            df_aux['SUB_ASSET_CLASS'] = ''
            df_aux['SHORT_NAME'] = ''
            df_aux['DESCRIPTION'] = df_aux.Description
            df_aux['ISIN'] = ''
            df_aux['SYMBOL'] = ''
            df_aux['QUANTITY'] = df_aux.Quantity
            df_aux['UNIT_COST'] = ''
            df_aux['UNIT_COST_CCY'] = 'USD'
            df_aux['CCY'] = 'USD'
            df_aux['MKT_VALUE'] = df_aux['Market Value']
            df_aux['MKT_VALUE_USD'] = df_aux['Market Value']
            df_aux['TOTAL_COST'] = df_aux['Gain/Loss $'] - df_aux['Market Value']
            df_aux['UNRELAIZED_GAIN/LOSS'] = df_aux['Gain/Loss $']
            df_aux['MXNUSD'] = mxnusd
            df_aux['EURUSD'] = eurusd
            df_aux['EXPOSITION_CCY'] = ''
            df_aux['Actualizado'] = ''
            df_aux['FXRATE'] = df_aux.EXPOSITION_CCY.map({'USD':1, 'MXN':mxnusd, 'EUR':eurusd})
            df_aux['INSTR_ID'] = df_aux.CUSIP
            
            data = pd.concat([data, df_aux[main_cols]], axis = 0)
            
    elif bank == 'GS':
        print("Using GS holdings...")
        # Para cada archivo de GS
        for file in dirs:
            # Lee el archivo y Limpia las filas que no tienen Market Value
            # df_aux = pd.read_excel(file, engine = 'openpyxl', header=7).dropna(subset=["Market Value"]).reset_index(drop = True)
            df_aux = pd.read_csv(file, header = 9).dropna(subset = ['Market Value']).reset_index(drop = True)
            
            # Para cada activo en la tabla
            df_aux['PORTFOLIO'] = portafolio(file[-7:-4], col = 'Portafolio')
            df_aux['BANK'] = 'GSS'
            df_aux['ACCT'] = file[-7:-4]
            df_aux['COMPANY'] = portafolio(file[-7:-4], col = 'Empresa')
            df_aux['ASSET_CLASS'] = ''
            df_aux['SUB_ASSET_CLASS'] = ''
            df_aux['SHORT_NAME'] = ''
            df_aux['DESCRIPTION'] = df_aux.Description
            df_aux['ISIN'] = ''
            df_aux['CUSIP'] = ''
            df_aux['SYMBOL'] = df_aux.Symbol
            df_aux['QUANTITY'] = df_aux.Quantity.astype(str).str.replace(',', '').astype(float)
            df_aux['UNIT_COST'] = df_aux['Unit Cost'].astype(str).str.replace(',', '').astype(float)
            df_aux['UNIT_COST_CCY'] = 'USD'
            df_aux['CCY'] = 'USD'
            df_aux['MKT_VALUE'] = df_aux['Market Value'].astype(str).str.replace(',', '').astype(float)
            df_aux['MKT_VALUE_USD'] = df_aux['Market Value'].astype(str).str.replace(',', '').astype(float)
            df_aux['TOTAL_COST'] = df_aux['Current Cost'].astype(str).str.replace(',', '').astype(float)
            df_aux['UNRELAIZED_GAIN/LOSS'] = df_aux['Unrealized Gain/Loss'].astype(str).str.replace(',', '').astype(float)
            df_aux['MXNUSD'] = mxnusd
            df_aux['EURUSD'] = eurusd
            df_aux['EXPOSITION_CCY'] = ''
            df_aux['Actualizado'] = ''
            df_aux['FXRATE'] = df_aux.EXPOSITION_CCY.map({'USD':1, 'MXN':mxnusd, 'EUR':eurusd})
            df_aux['INSTR_ID'] = df_aux.SYMBOL
            
            data = pd.concat([data, df_aux[main_cols]], axis = 0)
            
    elif bank == 'JPM':
        print("Using JPM holdings...")
        # Busca en los archivo de la carpeta DB
        for file in dirs:
            # lee el excel
            df_aux = pd.read_csv(file).dropna(subset=["Description"]).reset_index(drop = True)
            # df_aux = pd.read_excel(file, engine = 'openpyxl').dropna(subset=["Description"]).reset_index(drop = True)
            df_aux['Local CCY'] = df_aux['Local CCY'].fillna(df_aux['Base CCY'])
            
            df_aux['PORTFOLIO'] = portafolio(file[-8:-4], col = 'Portafolio')
            df_aux['BANK'] = 'JPM'
            df_aux['ACCT'] = file[-8:-4]
            df_aux['COMPANY'] = portafolio(file[-8:-4], col = 'Empresa')
            df_aux['ASSET_CLASS'] = ''
            df_aux['SUB_ASSET_CLASS'] = ''
            df_aux['SHORT_NAME'] = ''
            df_aux['DESCRIPTION'] = df_aux.Description
            df_aux['ISIN'] = df_aux.DESCRIPTION.apply(lambda x: re.search(r'XS\w+', x).group() if re.search(r'XS\w+', x) else '')
            df_aux['SYMBOL'] = ''
            df_aux['QUANTITY'] = df_aux.Quantity.astype(str).str.replace(',', '').astype(float)
            df_aux['UNIT_COST'] = df_aux['Unit Cost'].astype(str).str.replace(',', '').astype(float)
            df_aux['UNIT_COST_CCY'] = df_aux['Base CCY']
            df_aux['CCY'] = df_aux['Local CCY']
            df_aux['MKT_VALUE'] = df_aux['Local Value'].astype(str).str.replace(',', '').astype(float)
            # df_aux['MKT_VALUE_USD'] = df_aux['Value'].astype(str).str.replace(',', '').astype(float)
            df_aux['MKT_VALUE_USD'] = df_aux.apply(lambda x: MktValueUSD(x['MKT_VALUE'], x['CCY']), axis = 1)
            df_aux['TOTAL_COST'] = df_aux['Local Cost'].astype(str).str.replace(',', '').astype(float)
            df_aux['UNRELAIZED_GAIN/LOSS'] = df_aux['Unrealized G/L Amt.'].astype(str).str.replace(',', '').astype(float)
            df_aux['MXNUSD'] = mxnusd
            df_aux['EURUSD'] = eurusd
            df_aux['EXPOSITION_CCY'] = ''
            df_aux['Actualizado'] = ''
            df_aux['FXRATE'] = df_aux.EXPOSITION_CCY.map({'USD':1, 'MXN':mxnusd, 'EUR':eurusd})
            df_aux['INSTR_ID'] = df_aux.CUSIP
            
            data = pd.concat([data, df_aux[main_cols]], axis = 0)
            
    elif bank == 'MS':
        print("Using MS holdings...")
        # Busca en los archivo de la carpeta DB
        for file in dirs:
            df_aux = pd.read_excel(file, header = 10, engine = 'openpyxl').dropna(subset=["Open Order"]).reset_index(drop = True)
            df_aux = df_aux[df_aux['Account Number'] != 'Total'].reset_index(drop = True)
            
            df_aux['PORTFOLIO'] = df_aux['Account Number'].astype(str).str[-4:].apply(portafolio, col = 'Portafolio')
            df_aux['BANK'] = 'MSY'
            df_aux['ACCT'] = df_aux['Account Number'].astype(str).str[-4:]
            df_aux['COMPANY'] = df_aux['Account Number'].astype(str).str[-4:].apply(portafolio, col = 'Empresa')
            df_aux['ASSET_CLASS'] = ''
            df_aux['SUB_ASSET_CLASS'] = ''
            df_aux['SHORT_NAME'] = ''
            df_aux['DESCRIPTION'] = df_aux.Name
            df_aux['ISIN'] = ''
            df_aux['SYMBOL'] = df_aux.Symbol
            df_aux['QUANTITY'] = df_aux.Quantity.apply(lambda x: 1 if x == '-' else str(x).replace(',', '')).astype(float)
            df_aux['UNIT_COST'] = ''
            df_aux['UNIT_COST_CCY'] = ''
            df_aux['CCY'] = 'USD'
            df_aux['MKT_VALUE'] = df_aux['Market Value ($)'].apply(lambda x: 0 if x == '-' else str(x).replace(',', '')).astype(float)
            df_aux['MKT_VALUE_USD'] = df_aux['Market Value ($)'].apply(lambda x: 0 if x == '-' else str(x).replace(',', '')).astype(float)
            df_aux['TOTAL_COST'] = df_aux['Total Cost ($)'].apply(lambda x: 0 if x == '-' else str(x).replace(',', '')).astype(float)
            df_aux['UNRELAIZED_GAIN/LOSS'] = df_aux['Unrealized Gain/Loss ($)'].apply(lambda x: 0 if x == '-' else str(x).replace(',', '')).astype(float)
            df_aux['MXNUSD'] = mxnusd
            df_aux['EURUSD'] = eurusd
            df_aux['EXPOSITION_CCY'] = ''
            df_aux['Actualizado'] = ''
            df_aux['FXRATE'] = df_aux.EXPOSITION_CCY.map({'USD':1, 'MXN':mxnusd, 'EUR':eurusd})
            df_aux['INSTR_ID'] = df_aux.CUSIP
            
            data = pd.concat([data, df_aux[main_cols]], axis = 0)


### =========================================================== ###
###              Completar y limpiar los datos                  ###
### =========================================================== ###
data = data.dropna(subset = 'PORTFOLIO').reset_index(drop = True)
data['SPONSOR'] = ''
data['QUANTITY'] = data.QUANTITY.fillna(0)
data['ACCT'] = data.ACCT.astype(str)
data['LTV %'] = 0.0

for i in data.index:
    df_aux = assetClass[assetClass.INSTR_ID == data.INSTR_ID[i]].reset_index(drop = True)
    if df_aux.shape[0] == 1:
        data.loc[i, 'ASSET_CLASS'] = df_aux.loc[0, 'NEW_ASSET_CLASS']
        data.loc[i, 'SUB_ASSET_CLASS'] = df_aux.loc[0, 'NEW_SUB_ASSET_CLASS']
        data.loc[i, 'SHORT_NAME'] = df_aux.loc[0, 'SHORT_NAME']
        data.loc[i, 'ISIN'] = df_aux.loc[0, 'ISIN']
        data.loc[i, 'CUSIP'] = df_aux.loc[0, 'CUSIP']
        data.loc[i, 'LTV %'] = df_aux.loc[0, 'LTV %']
        data.loc[i, 'SPONSOR'] = df_aux.loc[0, 'SPONSOR']
        # Coloca el valor del SYMBOL que corresponde al nombre
        if data.loc[i, 'SYMBOL'] == '' or data.loc[i, 'SYMBOL'] == None:
            data.loc[i, 'SYMBOL'] = df_aux.loc[0, 'SYMBOL']
    elif df_aux.shape[0] < 1:
        print(f"No está el valor: {data.loc[i, 'INSTR_ID']} | {data.loc[i, 'DESCRIPTION']} | {data.loc[i, 'CUSIP']}")
    elif df_aux.shape[0] > 1:
        print(f"El valor: {data.loc[i, 'INSTR_ID']} | {data.loc[i, 'DESCRIPTION']} | está duplicado en el diccionario, hace falta revisar")

# LTV %
data["LTV $"] = data["MKT_VALUE_USD"] * data["LTV %"]

# Calculamos el costo total de SSZ
if not ssz_data.empty:
    cond1 = (data.BANK == 'SSZ') & (data.ASSET_CLASS.isin(['COLATERAL', 'EQUITIES']))
    cond2 = (data.BANK == 'SSZ') & (data.ASSET_CLASS.isin(['STRUCTURED NOTES', 'LIQUIDITY']))
    cond3 = (data.BANK == 'SSZ') & (data.ASSET_CLASS == 'FIXED INCOME') & (data.SUB_ASSET_CLASS != 'CORPORATE BOND')
    cond4 = (data.BANK == 'SSZ') & (data.ASSET_CLASS == 'FIXED INCOME') & (data.SUB_ASSET_CLASS == 'CORPORATE BOND')
    cond5 = (data.BANK == 'SSZ') & (data.ASSET_CLASS == 'STRUCTURED NOTES')
    cond6 = (data.BANK == 'SSZ') & (data.ASSET_CLASS.isin(['LIQUIDITY', 'FIXED INCOME'])) & (data.SUB_ASSET_CLASS.isin(['CORPORATE BOND', 'DISTRESSED BOND', 'GOV SECURITIES', 'PREF STOCK']))
    cond7 = (data.BANK == 'SSZ') & (data.ASSET_CLASS == 'LIQUIDITY') & (data.SUB_ASSET_CLASS == 'GOV SECURITIES')

    data.loc[cond1, 'TOTAL_COST'] = data.loc[cond1, 'UNIT_COST'] * data.loc[cond1, 'QUANTITY']
    data.loc[cond2, 'TOTAL_COST'] = data.loc[cond2, 'MKT_VALUE_USD'] - data.loc[cond2, 'UNRELAIZED_GAIN/LOSS']
    data.loc[cond3, 'TOTAL_COST'] = data.loc[cond3, 'UNIT_COST'] * data.loc[cond3, 'QUANTITY']
    data.loc[cond4, 'TOTAL_COST'] = (data.loc[cond4, 'UNIT_COST'] / 100) * data.loc[cond4, 'QUANTITY']
    data.loc[cond5, 'TOTAL_COST'] = data.loc[cond5, 'TOTAL_COST'] / 100
    data.loc[cond5, 'UNRELAIZED_GAIN/LOSS'] = data.loc[cond5, 'MKT_VALUE_USD'] - (data.loc[cond5, 'QUANTITY'] * data.loc[cond5, 'FXRATE'])
    data.loc[cond6, 'UNRELAIZED_GAIN/LOSS'] = data.loc[cond6, 'MKT_VALUE_USD'] - (data.loc[cond6, 'QUANTITY'] * data.loc[cond6, 'UNIT_COST'] / 100 * data.loc[cond6, 'FXRATE'])
    data.loc[cond7, 'TOTAL_COST'] = (data.loc[cond7, 'UNIT_COST'] / 100) * data.loc[cond7, 'QUANTITY']

data['DESCRIPTION'] = data.DESCRIPTION.fillna('-').replace(['', '  '], '-')

# =============================================================================
#               Verificar si estan todas las cuentas (Corregir esto)
# =============================================================================
# print(f'No está(n) la(s) cuenta(s): {set(catPortf.Cuenta.unique()) - set(data.)}')

# =============================================================================
#                          Creación de la hoja de LTV
# =============================================================================
df_ltv = pd.DataFrame(columns=['BANK','MKT_VALUE_USD','LOAN','LTV','LTV-LOAN','PCT'])

for bnk in data['BANK'].unique():
    if bnk == 'SSZ' or bnk=='JBR':
        try:
            df_aux = data[(data['BANK']==bnk) & (data['ASSET_CLASS']!='CREDITO')][['BANK','MKT_VALUE_USD','LTV $']].groupby(['BANK']).sum()
        except:
            print('Agrega los valores del credito de JBR o SSZ al diccionario.')
        df_ltv.loc[bnk,'BANK'] = bnk
        df_ltv.loc[bnk,'MKT_VALUE_USD'] = df_aux['MKT_VALUE_USD'].loc[bnk]
        df_ltv.loc[bnk,'LTV'] = df_aux['LTV $'].loc[bnk]
        df_aux = data[(data['BANK']==bnk) & (data['ASSET_CLASS']=='CREDITO')][['BANK','MKT_VALUE_USD','LTV $']].groupby(['BANK']).sum()
        df_ltv.loc[bnk,'LOAN'] = df_aux['MKT_VALUE_USD'].loc[bnk]
        df_ltv.loc[bnk,'LTV-LOAN'] = df_ltv.loc[bnk,'LTV']+df_ltv.loc[bnk,'LOAN']
        df_ltv.loc[bnk,'PCT'] = str(round((df_ltv.loc[bnk,'LTV-LOAN']/df_ltv.loc[bnk,'LTV'])*100,2))+'%'
    elif bnk=='DBK':
        df_aux = data[((data['BANK']==bnk) & (data['ACCT']=='2269')) & (data['ASSET_CLASS']!='CREDITO') ][['BANK','MKT_VALUE_USD','LTV $']].groupby(['BANK']).sum()
        df_ltv.loc[bnk,'BANK'] = bnk
        df_ltv.loc[bnk,'MKT_VALUE_USD'] = df_aux['MKT_VALUE_USD'].loc[bnk]
        df_ltv.loc[bnk,'LTV'] = df_aux['LTV $'].loc[bnk]
        df_aux = data[(data['BANK']==bnk) & (data['ASSET_CLASS']=='CREDITO')][['BANK','MKT_VALUE_USD','LTV $']].groupby(['BANK']).sum()
        df_ltv.loc[bnk,'LOAN'] = df_aux['MKT_VALUE_USD'].loc[bnk]
        df_ltv.loc[bnk,'LTV-LOAN'] = df_ltv.loc[bnk,'LTV']+df_ltv.loc[bnk,'LOAN']
        df_ltv.loc[bnk,'PCT'] = str(round((df_ltv.loc[bnk,'LTV-LOAN']/df_ltv.loc[bnk,'LTV'])*100,2))+'%'    
    elif bnk=='CTI':
        df_aux = data[(data['BANK']==bnk) & (data['ASSET_CLASS']!='CREDITO')&(data['ACCT']!='11768')][['BANK','MKT_VALUE_USD','LTV $']].groupby(['BANK']).sum()
        df_ltv.loc[bnk,'BANK'] = bnk
        df_ltv.loc[bnk,'MKT_VALUE_USD'] = df_aux['MKT_VALUE_USD'].loc[bnk]        
        df_ltv.loc[bnk,'LTV'] = df_aux['MKT_VALUE_USD'].loc[bnk]
        df_aux = data[(data['BANK']==bnk) & (data['ASSET_CLASS']=='CREDITO')][['BANK','MKT_VALUE_USD','LTV $']].groupby(['BANK']).sum()
        df_ltv.loc[bnk,'LOAN'] = df_aux['MKT_VALUE_USD'].loc[bnk]
        df_ltv.loc[bnk,'LTV-LOAN'] = df_ltv.loc[bnk,'LTV']+df_ltv.loc[bnk,'LOAN']
        df_ltv.loc[bnk,'PCT'] = str(round((df_ltv.loc[bnk,'LTV-LOAN']/df_ltv.loc[bnk,'LTV'])*100,2))+'%'

df_ltv['MKT_VALUE_USD'] = df_ltv['MKT_VALUE_USD'].astype('float64')
df_ltv['LOAN'] = df_ltv['LOAN'].astype('float64')
df_ltv['LTV'] = df_ltv['LTV'].astype('float64')
df_ltv['LTV-LOAN'] = df_ltv['LTV-LOAN'].astype('float64')

with pd.ExcelWriter(f'{path_pc}/Documents/{date}-ASSET_ALLOC.xlsx', engine = "openpyxl") as writer:
    data.to_excel(writer, sheet_name = "CONSOLIDADO", index = False)
    data[(data["ASSET_CLASS"] != 'COLATERAL') & (data["ASSET_CLASS"] != 'CREDITO') ].to_excel(writer, sheet_name = "CONSOLIDADO SIN COLATERALES", index = False)
    df_ltv.to_excel(writer, sheet_name = "LTV", index = False)