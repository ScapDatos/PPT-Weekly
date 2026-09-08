# =============================================================================
# Para el Asset Allocator Diario
# Notas:
#    - Verificar que las bases de datos estén actualizadas a la fecha reciente
# Implementaciones
# =============================================================================

import os
import pymysql
import warnings
import pandas as pd
from dotenv import load_dotenv
from config import DATE, MXNUSD, EURUSD, PARENT_PATH

load_dotenv()

## IMPORTANTE: REVISAR ESTO, QUITA TODAS LAS ADVERTENCIAS. COMENTAR EN CASO DE MODIFICAR EL CODIGO
## A NIVEL PROFUNDO PARA REVISAR ERRORES Y/O ADVERTENCIAS
warnings.filterwarnings('ignore')

# =========================== Connection Info =================================
user = os.getenv('USER')       # User
pswd = os.getenv('PSWD')       # Password
host = os.getenv('HOST')       # Host

# ======================= Dictionary file of all accounts =======================
catPortf = pd.read_excel(PARENT_PATH / 'Contributtion' / 'Diccionarios_MySQL.xlsx',
                         sheet_name = "Cuentas", converters = {'Cuenta':str}, engine = "openpyxl")
catPortf = catPortf[catPortf.Status.notnull()][['Banco', 'Portafolio', 'Cuenta', 'Empresa']].reset_index(drop = True)
assetClass = pd.read_excel(PARENT_PATH / 'Contributtion' / 'Diccionarios_MySQL.xlsx',
                           sheet_name = "Asset_class", converters = {'Cuenta': str}, engine = "openpyxl")

def FXrate(data, mxnusd, eurusd):
    if data.CCY == 'MXN':
        return mxnusd
    elif data.CCY == 'EUR':
        return eurusd
    else:
        return 1.0

def accounts(data):
    accts = {'DDA00080-11614': '1614', 'DDA00080-94280': '4280',
             'GCR384-55557': '5557', 'CLS FL18-70002': '0002',
             'CH8408235006903701096-69037': '1096',
             'CH9308235006903702001-69037': '600001',
             'CH1508235006849002001-68490': '02001'}
    
    return accts.get(data.INSTR_ID, data.ACCT)

print('\n', DATE)
data = pd.read_excel(PARENT_PATH / 'Contributtion' / 'Template_MySQL.xlsx')

data['FXRATE'] = data.apply(FXrate, args = (MXNUSD, EURUSD), axis = 1)
data['MKT_VALUE_USD'] = data.MKT_VALUE * data.FXRATE
main_cols = data.columns

# Dictionaries where all historyc data is saved
ubs_dict = {}
jpm_dict = {}
jb_dic = {}
ct_dic = {}
ssz_dic = {}

def portafolio(data, col:str):
    for idx, acct in enumerate(catPortf.Cuenta):
        if str(acct) in data:
            return catPortf[col][idx]

# ========================= Connection to all databases =========================
ct_conn = pymysql.connect(host = host, user = user, password = pswd, database = 'CT')
jb_conn = pymysql.connect(host = host, user = user, password = pswd, database = 'JB')
jpm_conn = pymysql.connect(host = host, user = user, password = pswd, database = 'JPM')
ssz_conn = pymysql.connect(host = host, user = user, password = pswd, database = 'SSZ')
ubs_conn = pymysql.connect(host = host, user = user, password = pswd, database = 'UBS')
dbk_conn = pymysql.connect(host = host, user = user, password = pswd, database = 'DBK')
msy_conn = pymysql.connect(host = host, user = user, password = pswd, database = 'MSY')


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
                ) AS MATURITY_DT, P.SEC_SYMBOL,
                P.TRX_CD
            FROM POSITION P
            WHERE P.BALANCE_DT = "{DATE}";
            """
ubs_data = pd.read_sql(ubs_query, con = ubs_conn)
ubs_data['PORTFOLIO'] = ubs_data.ACCT_NBR.apply(portafolio, col = 'Portafolio')
ubs_data['BANK'] = 'UBS'
ubs_data['ACCT'] = ubs_data.ACCT_NBR.str[-5:].str.capitalize()
ubs_data['COMPANY'] = ubs_data.ACCT_NBR.apply(portafolio, col = 'Empresa')
ubs_data['ASSET_CLASS'] = ''
ubs_data['SUB_ASSET_CLASS'] = ''
ubs_data['SHORT_NAME'] = ''
ubs_data['INSTR_ID'] = ubs_data.SEC_SYMBOL + '-' + ubs_data.ACCT
ubs_data['DESCRIPTION'] = ubs_data.DESCRIPTION
ubs_data['ISIN'] = ''
ubs_data['SYMBOL'] = ubs_data.SEC_SYMBOL
ubs_data['UNIT_COST'] = ''
ubs_data['UNIT_COST_CCY'] = 'USD'
ubs_data['CCY'] = 'USD'
ubs_data['MKT_VALUE'] = ubs_data.TRADE_AMT
ubs_data['MKT_VALUE_USD'] = ubs_data.TRADE_AMT
ubs_data['TOTAL_COST'] = ''
ubs_data['UNRELAIZED_GAIN/LOSS'] = ''
ubs_data['FXRATE'] = 1.0
ubs_data['EXPOSITION_CCY'] = 'USD'
ubs_data['Actualizado'] = ubs_data.BALANCE_DT

ubs_data.loc[ubs_data['TRX_CD'] == 'wd', ['MKT_VALUE', 'MKT_VALUE_USD']] = -ubs_data.TRADE_AMT

if ubs_data.shape[0] != 0:
    data = pd.concat([data, ubs_data[main_cols]], axis = 0)
else:
    print(f'No hay datos de UBS para {DATE}. Revisar la info.')

### =========================================================== ###
###                              JPM                            ###
### =========================================================== ###
jpm_query = f"""
            SELECT
                ACCT_NUM, SEC_DESC_1, SEC_DESC_2, SEC_DESC_3, SEC_DESC_4,
                SEC_DESC_5, CUSIP, TRADED_UNITS, TRADED_COST_LOCAL,
                ACCT_BASE_CCY, SEC_CCY, TRADED_MKT_VAL_LOCAL, ISIN,
                TRADED_MKT_VAL_BASE, TRADED_UNR_GL_BASE, FX_RATE,
                POS_DATE, SEC_NUM, TICKER_SYMBOL, PICE_LOCAL
            FROM POSITION
            WHERE POS_DATE = '{DATE}';
            """
jpm_csh = f"""
            SELECT
                ACCT_NUM, CCY_LOCAL, EXCHANGE_RATE, TRADED_CASH_BAL_BASE,
                TRADED_CASH_BAL_LOCAL, ACCT_BASE_CCY, DATE_CASH_BALANCE_FILE
            FROM CASH_BALANCE
            WHERE DATE_CASH_BALANCE_FILE = '{DATE}';
            """
fvals = {'ERICSSON (LM) TEL-SP ADR    ': 'ERICSSON0XX1', 'JPM LI-LIQ LVNAV FD - USD - W - ACC  ISIN LU1873131988 SEDOL BFM4703  ': 'JOMLILIQ0XX1'}
jpm_data = pd.read_sql(jpm_query, con = jpm_conn)
jpm_cash = pd.read_sql(jpm_csh, con = jpm_conn)

jpm_data['PORTFOLIO'], jpm_cash['PORTFOLIO'] = jpm_data.ACCT_NUM.apply(portafolio, col = 'Portafolio'), jpm_cash.ACCT_NUM.apply(portafolio, col = 'Portafolio')
jpm_data['BANK'], jpm_cash['BANK'] = 'JPM', 'JPM'
jpm_data['ACCT'], jpm_cash['ACCT'] = jpm_data.ACCT_NUM.str[-4:], jpm_cash.ACCT_NUM.str[-4:]
jpm_data['COMPANY'], jpm_cash['COMPANY'] = jpm_data.ACCT_NUM.apply(portafolio, col = 'Empresa'), jpm_cash.ACCT_NUM.apply(portafolio, col = 'Empresa')
jpm_data['ASSET_CLASS'], jpm_cash['ASSET_CLASS'] = '', ''
jpm_data['SUB_ASSET_CLASS'], jpm_cash['SUB_ASSET_CLASS'] = '', ''
jpm_data['SHORT_NAME'], jpm_cash['SHORT_NAME'] = '', ''
jpm_data['DESCRIPTION'], jpm_cash['DESCRIPTION'] = jpm_data.SEC_DESC_1 + ' ' + jpm_data.SEC_DESC_2 + ' ' + jpm_data.SEC_DESC_3 + ' ' + jpm_data.SEC_DESC_4 + ' ' + jpm_data.SEC_DESC_5, 'CASH'
jpm_data['INSTR_ID'], jpm_cash['INSTR_ID'] = jpm_data.SEC_NUM.fillna(jpm_data.DESCRIPTION.map(fvals)) + '-' + jpm_data.ACCT, jpm_cash.ACCT_NUM.str[-4:] + '-' + jpm_cash.CCY_LOCAL
jpm_data['SYMBOL'], jpm_cash['SYMBOL'] = jpm_data.TICKER_SYMBOL, ''
jpm_data['QUANTITY'], jpm_cash['QUANTITY'] = jpm_data.TRADED_UNITS, jpm_cash.TRADED_CASH_BAL_LOCAL
jpm_data['UNIT_COST'], jpm_cash['UNIT_COST'] = jpm_data.PICE_LOCAL, jpm_cash.EXCHANGE_RATE
jpm_data['UNIT_COST_CCY'], jpm_cash['UNIT_COST_CCY'] = jpm_data.SEC_CCY, jpm_cash.CCY_LOCAL
jpm_data['CCY'], jpm_cash['CCY'] = jpm_data.ACCT_BASE_CCY, jpm_cash.ACCT_BASE_CCY
jpm_data['MKT_VALUE'], jpm_cash['MKT_VALUE'] = jpm_data.TRADED_MKT_VAL_LOCAL, jpm_cash.TRADED_CASH_BAL_LOCAL
jpm_data['MKT_VALUE_USD'], jpm_cash['MKT_VALUE_USD'] = jpm_data.TRADED_MKT_VAL_BASE, jpm_cash.TRADED_CASH_BAL_BASE
jpm_data['TOTAL_COST'], jpm_cash['TOTAL_COST'] = jpm_data.TRADED_COST_LOCAL, jpm_cash.TRADED_CASH_BAL_BASE
jpm_data['UNRELAIZED_GAIN/LOSS'], jpm_cash['UNRELAIZED_GAIN/LOSS'] = jpm_data.TRADED_UNR_GL_BASE, 0
jpm_data['FXRATE'], jpm_cash['FXRATE'] = jpm_data.FX_RATE, jpm_cash.EXCHANGE_RATE
jpm_data['EXPOSITION_CCY'], jpm_cash['EXPOSITION_CCY'] = jpm_data.SEC_CCY, jpm_cash.CCY_LOCAL
jpm_data['Actualizado'], jpm_cash['Actualizado'] = jpm_data.POS_DATE, jpm_cash.DATE_CASH_BALANCE_FILE
jpm_cash['CUSIP'] = '0' + jpm_cash.CCY_LOCAL + 'PRAA7'
jpm_cash['ISIN'] = ''
 
if jpm_data.shape[0] != 0:
    data = pd.concat([data, jpm_data[main_cols], jpm_cash[main_cols]], axis = 0)
else:
    print(f'No hay datos de JP Morgan para {DATE}. Revisar la info.')

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
                ISIN, TITEL, KURSWERT_TITWRG
            FROM POSITION
            WHERE PER_DATUM = '{DATE}';
            """
jb_data = pd.read_sql(jb_query, con = jb_conn)

jb_data['PORTFOLIO'] = jb_data.ZRNR.apply(portafolio, col = 'Portafolio')
jb_data['BANK'] = 'JBR'
jb_data['ACCT'] = jb_data.ZRNR.str[-4:]
jb_data['COMPANY'] = jb_data.ZRNR.apply(portafolio, col = 'Empresa')
jb_data['ASSET_CLASS'] = ''
jb_data['SUB_ASSET_CLASS'] = ''
jb_data['SHORT_NAME'] = ''
jb_data['INSTR_ID'] = [str(isin) + '-' + str(act) if isin.strip() != '' else str(int(titel)) + '-' + str(act) for isin, titel, act in zip(jb_data.ISIN, jb_data.TITEL, jb_data.ACCT)]
jb_data['DESCRIPTION'] = jb_data.KURZ_TEXT + ' ' + jb_data.AUSGABE_ZINSSATZ.astype(str) + '% ' + jb_data.BEZ + ' ' + jb_data.BEZ_GEGENPARTEI
jb_data['ISIN'] = jb_data.ISIN
jb_data['CUSIP'] = ''
jb_data['SYMBOL'] = ''
jb_data['QUANTITY'] = [let if 'CUENTA CORRIENTE' in bez else bes for bes, bez, let in zip(jb_data.BESTAND, jb_data.BEZ, jb_data.LETZT_BUSALDO)]
jb_data['UNIT_COST'] = ''
jb_data['UNIT_COST_CCY'] = [let if 'CUENTA CORRIENTE' in bez else bes for bes, bez, let in zip(jb_data.TITEL_WRG_TEXT, jb_data.BEZ, jb_data.KTWRG_TEXT)]
jb_data['CCY'] = 'USD'
jb_data['MKT_VALUE'] = [qty if mkt == 0 and 'CHROME' not in ac else mkt for ac, qty, mkt in zip(jb_data.KURZ_TEXT, jb_data.QUANTITY, jb_data.KURSWERT_TITWRG)]
jb_data['MKT_VALUE_USD'] = jb_data.KURSWERT + jb_data.MARCHZINS
jb_data['TOTAL_COST'] = jb_data.BEK
jb_data['UNRELAIZED_GAIN/LOSS'] = ''
jb_data['FXRATE'] = jb_data.UMRECH_KURS
jb_data['EXPOSITION_CCY'] = jb_data.UNIT_COST_CCY
jb_data['Actualizado'] = jb_data.PER_DATUM
jb_data = jb_data.apply(insert_val, axis = 1)
# jb_data = jb_data[jb_data.ZRNR != 3176542].reset_index(drop = True)

if jb_data.shape[0] != 0:
    data = pd.concat([data, jb_data[main_cols]], axis = 0)
else:
    print(f'No hay datos de Julius Baer para {DATE}. Revisar la info.')

### =========================================================== ###
###                              CT                             ### #CHECK
### =========================================================== ###
ct_query = f"""
            SELECT
                ACCT_NBR, FIN_INSTR_LONG_LINE1_DESC, FIN_INSTR_LONG_LINE2_DESC,
                FIN_INSTR_LONG_LINE3_DESC, DEF_TICKER_SYMB_CD, REMN_UNT, 
                REMN_COST_NOM_AMT, ACCT_REF_CCY_CD, NOM_CCY_CD, NOM_AMT,
                MKT_VAL_AMT, FX_RT, POSN_AS_OF_DT, FIN_INSTR_ID,
                ASSET_SUB_SUB_CLAS_CD, UREL_GNLS_AMT, HIDG_COST_UNY_PRC_AMT
            FROM POSITION
            WHERE POSN_AS_OF_DT = '{DATE}';
            """

ct_data = pd.read_sql(ct_query, con = ct_conn)
ct_data['PORTFOLIO'] = ct_data.ACCT_NBR.apply(portafolio, col = 'Portafolio')
ct_data['BANK'] = 'CTI'
ct_data['ACCT'] = ct_data.ACCT_NBR.str[-5:]
ct_data['COMPANY'] = ct_data.ACCT_NBR.apply(portafolio, col = 'Empresa')
ct_data['ASSET_CLASS'] = ''
ct_data['SUB_ASSET_CLASS'] = ''
ct_data['SHORT_NAME'] = ''
ct_data['INSTR_ID'] = ct_data.FIN_INSTR_ID.combine_first(ct_data.ASSET_SUB_SUB_CLAS_CD) + '-' + ct_data.ACCT
ct_data['DESCRIPTION'] = ct_data.FIN_INSTR_LONG_LINE1_DESC.fillna('') + ' ' + ct_data.FIN_INSTR_LONG_LINE2_DESC.fillna('') + ' ' + ct_data.FIN_INSTR_LONG_LINE3_DESC.fillna('')
ct_data['ISIN'] = ''
ct_data['CUSIP'] = ''
ct_data['SYMBOL'] = ct_data.DEF_TICKER_SYMB_CD
ct_data['QUANTITY'] = [amt if 'Western' in des else unt for unt, amt, des in zip(ct_data.REMN_UNT, ct_data.NOM_AMT, ct_data.FIN_INSTR_LONG_LINE1_DESC.fillna(''))]
ct_data['UNIT_COST'] = ct_data.HIDG_COST_UNY_PRC_AMT
ct_data['UNIT_COST_CCY'] = ct_data.NOM_CCY_CD
ct_data['CCY'] = ct_data.ACCT_REF_CCY_CD
ct_data['MKT_VALUE'] = [-amt if sclas == 'CLS FL18' else amt for amt, sclas in zip(ct_data.NOM_AMT, ct_data.ASSET_SUB_SUB_CLAS_CD)]
ct_data['MKT_VALUE_USD'] = [-amt if sclas == 'CLS FL18' else amt for amt, sclas in zip(ct_data.MKT_VAL_AMT, ct_data.ASSET_SUB_SUB_CLAS_CD)]
ct_data['TOTAL_COST'] = [amt if 'Western' in des else unt for unt, amt, des in zip(ct_data.REMN_COST_NOM_AMT, ct_data.NOM_AMT, ct_data.FIN_INSTR_LONG_LINE1_DESC.fillna(''))]
ct_data['UNRELAIZED_GAIN/LOSS'] = ct_data.UREL_GNLS_AMT / ct_data.FX_RT
ct_data['FXRATE'] = ct_data.FX_RT
ct_data['EXPOSITION_CCY'] = ct_data.NOM_CCY_CD
ct_data['Actualizado'] = ct_data.POSN_AS_OF_DT
ct_data['ACCT'] = ct_data.apply(accounts, axis = 1)

if ct_data.shape[0] != 0:
    data = pd.concat([data, ct_data[main_cols]], axis = 0)
else:
    print(f'No hay datos de CITI para {DATE}. Revisar la info.')

### =========================================================== ###
###                              DBK                            ###
### =========================================================== ###
dbk_query = f"""
            SELECT
                P.ACCT_NBR, P.SEC_DESC_1, P.SEC_DESC_2, P.SEC_DESC_3,
                P.SEC_DESC_4, P.CUSIP, P.ISIN, P.SEC_DESC_5, P.SEC_DESC_6,
                (
                    SELECT TX.CURR_TOTAL_COST
                    FROM DBK.TAXLOT TX
                    WHERE TX.CUSIP = P.CUSIP
                    LIMIT 1    
                ) AS TOTAL_COST,
                P.SEC_SYMBOL, P.TRADE_DT_QUANT, P.TRADE_DT_LIQ, P.POSITION_CCY,
                P.FX_RATE, P.ETL_DT
            FROM POSITION P
            WHERE P.ETL_DT = '{DATE}';
            """
dbk_ecbm = f"""
            SELECT
                ACCT_NBR, FUND_DESC, CUSIP, FUND_MAN, PRIN_BAL,
                ACC_DIV_BAL, CCY_CD, LAST_UP_DT
            FROM ECMB
            WHERE LAST_UP_DT = '{DATE}';
            """

dbk_data = pd.read_sql(dbk_query, con = dbk_conn)
dbk_ecbm = pd.read_sql(dbk_ecbm, con = dbk_conn)

dbk_data['PORTFOLIO'], dbk_ecbm['PORTFOLIO'] = dbk_data.ACCT_NBR.apply(portafolio, col = 'Portafolio'), dbk_ecbm.ACCT_NBR.apply(portafolio, col = 'Portafolio')
dbk_data['BANK'], dbk_ecbm['BANK'] = 'DBK', 'DBK'
dbk_data['ACCT'], dbk_ecbm['ACCT'] = dbk_data.ACCT_NBR.str[-5:-1], dbk_ecbm.ACCT_NBR.str[-4:]
dbk_data['COMPANY'], dbk_ecbm['COMPANY'] = dbk_data.ACCT_NBR.apply(portafolio, col = 'Empresa'), dbk_ecbm.ACCT_NBR.apply(portafolio, col = 'Empresa')
dbk_data['ASSET_CLASS'], dbk_ecbm['ASSET_CLASS'] = '', ''
dbk_data['SUB_ASSET_CLASS'], dbk_ecbm['SUB_ASSET_CLASS'] = '', ''
dbk_data['SHORT_NAME'], dbk_ecbm['SHORT_NAME'] = '', ''
dbk_data['INSTR_ID'], dbk_ecbm['INSTR_ID'] = dbk_data.CUSIP + '-' + dbk_data.ACCT, dbk_ecbm.CUSIP + '-' + dbk_ecbm.ACCT
dbk_data['DESCRIPTION'], dbk_ecbm['DESCRIPTION'] = dbk_data.SEC_DESC_1 + ' ' + dbk_data.SEC_DESC_2.fillna('') + ' ' + dbk_data.SEC_DESC_3.fillna('') + ' ' + dbk_data.SEC_DESC_4.fillna(''), dbk_ecbm.FUND_DESC
dbk_data['ISIN'], dbk_ecbm['ISIN'] = dbk_data.ISIN, ''
dbk_data['SYMBOL'], dbk_ecbm['SYMBOL'] = dbk_data.SEC_SYMBOL, dbk_ecbm.FUND_MAN
dbk_data['QUANTITY'], dbk_ecbm['QUANTITY'] = dbk_data.TRADE_DT_QUANT, dbk_ecbm.PRIN_BAL + dbk_ecbm.ACC_DIV_BAL
dbk_data['UNIT_COST'], dbk_ecbm['UNIT_COST'] = dbk_data.TOTAL_COST / dbk_data.QUANTITY, 1.0
dbk_data['UNIT_COST_CCY'], dbk_ecbm['UNIT_COST_CCY'] = dbk_data.POSITION_CCY, dbk_ecbm.CCY_CD
dbk_data['CCY'], dbk_ecbm['CCY'] = dbk_data.POSITION_CCY, dbk_ecbm.CCY_CD
dbk_data['MKT_VALUE'], dbk_ecbm['MKT_VALUE'] = dbk_data.TRADE_DT_LIQ, dbk_ecbm.PRIN_BAL + dbk_ecbm.ACC_DIV_BAL
dbk_data['MKT_VALUE_USD'], dbk_ecbm['MKT_VALUE_USD'] = dbk_data.TRADE_DT_LIQ, dbk_ecbm.PRIN_BAL + dbk_ecbm.ACC_DIV_BAL
dbk_ecbm['TOTAL_COST'] = dbk_ecbm.QUANTITY
dbk_data['UNRELAIZED_GAIN/LOSS'], dbk_ecbm['UNRELAIZED_GAIN/LOSS'] = dbk_data.MKT_VALUE - dbk_data.TOTAL_COST, 0.0
dbk_data['FXRATE'], dbk_ecbm['FXRATE'] = dbk_data.FX_RATE, 1.0
dbk_data['EXPOSITION_CCY'], dbk_ecbm['EXPOSITION_CCY'] = dbk_data.POSITION_CCY, dbk_ecbm.CCY_CD
dbk_data['Actualizado'], dbk_ecbm['Actualizado'] = dbk_data.ETL_DT, dbk_ecbm.LAST_UP_DT

if dbk_data.shape[0] != 0:
    data = pd.concat([data, dbk_data[main_cols], dbk_ecbm[main_cols]], axis = 0)
else:
    print(f'No hay datos de Deutsche Bank para {DATE}. Revisar la info.')

### =========================================================== ###
###                              MSY                            ###
### =========================================================== ###
msy_query = f"""
            SELECT
                ACCT, SEC_DESC, ISIN, CUSIP, SYMBOL, QUANTITY,
                CCY, MKT_LOCAL, MKT_BASE, TOTAL_COST, POS_DT
            FROM POSITION
            WHERE POS_DT = '{DATE}';
            """
msy_data = pd.read_sql(msy_query, con = msy_conn)

msy_data['PORTFOLIO'] = msy_data.ACCT.apply(portafolio, col = 'Portafolio')
msy_data['BANK'] = 'MSY'
msy_data['ACCT'] = msy_data.ACCT.str[-4:]
msy_data['COMPANY'] = msy_data.ACCT.apply(portafolio, col = 'Empresa')
msy_data['ASSET_CLASS'] = ''
msy_data['SUB_ASSET_CLASS'] = ''
msy_data['SHORT_NAME'] = ''
msy_data['INSTR_ID'] = msy_data.CUSIP.fillna('MSYCASHXX0') + '-' + msy_data.ACCT
msy_data['DESCRIPTION'] = msy_data.SEC_DESC
msy_data['UNIT_COST'] = msy_data.TOTAL_COST / msy_data.QUANTITY
msy_data['UNIT_COST_CCY'] = msy_data.CCY
msy_data['CCY'] = msy_data.CCY
msy_data['MKT_VALUE'] = msy_data.MKT_LOCAL
msy_data['MKT_VALUE_USD'] = msy_data.MKT_BASE
msy_data['UNRELAIZED_GAIN/LOSS'] = msy_data.MKT_VALUE_USD - msy_data.TOTAL_COST
msy_data['FXRATE'] = 1.0
msy_data['EXPOSITION_CCY'] = msy_data.CCY
msy_data['Actualizado'] = msy_data.POS_DT

if msy_data.shape[0] != 0:
    data = pd.concat([data, msy_data[main_cols]], axis = 0)
else:
    print(f'No hay datos de Morgan Stanley para {DATE}. Revisar la info.\n')

### =========================================================== ###
###                              SSZ                            ###
### =========================================================== ###
ssz_query = f"""
            SELECT
                PORT_ID, INSTRUMENT_NAME, ISIN, NOMINAL_AMNT, 
                COST_PRICE, COST_PRICE_CCY, VALUE_PRICE_CCY,
                FX_RATE, PROC_DATE, IBAN
            FROM POSITION
            WHERE PROC_DATE = '{DATE}'
            AND PORT_ID IN ('69037-1', '69037-2', '68490-1', '68490-2');
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
ssz_data['INSTR_ID'] = ssz_data.ISIN.fillna('SSZLOAN0') + '-' + ssz_data.ACCT
ssz_data['DESCRIPTION'] = ssz_data.INSTRUMENT_NAME
ssz_data['CUSIP'] = ''
ssz_data['SYMBOL'] = ''
ssz_data['QUANTITY'] = [0 if amnt < 0 else amnt for amnt in ssz_data.NOMINAL_AMNT]
ssz_data['UNIT_COST'] = ssz_data.COST_PRICE
ssz_data['UNIT_COST_CCY'] = ssz_data.COST_PRICE_CCY
ssz_data['CCY'] = 'USD'
ssz_data['MKT_VALUE'] = ssz_data.VALUE_PRICE_CCY
ssz_data['MKT_VALUE_USD'] = ssz_data.VALUE_PRICE_CCY * ssz_data.FX_RATE
ssz_data['TOTAL_COST'] = ssz_data.NOMINAL_AMNT * ssz_data.VALUE_PRICE_CCY
ssz_data['UNRELAIZED_GAIN/LOSS'] = (ssz_data.VALUE_PRICE_CCY - (ssz_data.COST_PRICE * ssz_data.NOMINAL_AMNT)) * ssz_data.FX_RATE
ssz_data['FXRATE'] = ssz_data.FX_RATE
ssz_data['EXPOSITION_CCY'] = ssz_data.COST_PRICE_CCY
ssz_data['Actualizado'] = ssz_data.PROC_DATE
ssz_data['ACCT'] = ssz_data.apply(accounts, axis = 1)

if ssz_data.shape[0] != 0:
    data = pd.concat([data, ssz_data[main_cols]], axis = 0)
else:
    print(f'No hay datos de Santander Suiza para {DATE}. Revisar la info.\n')


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
        data.loc[i,'SUB_ASSET_CLASS'] = df_aux.loc[0,'NEW_SUB_ASSET_CLASS']
        data.loc[i,'SHORT_NAME'] = df_aux.loc[0,'SHORT_NAME']
        data.loc[i,'LTV %'] = df_aux.loc[0,'LTV %']
        data.loc[i,'SPONSOR'] = df_aux.loc[0,'SPONSOR']
        data.loc[i,'ISIN'] = df_aux.loc[0,'ISIN']
        data.loc[i,'CUSIP'] = df_aux.loc[0,'CUSIP']
        # Coloca el valor del SYMBOL que corresponde al nombre
        # if data.loc[i,'SYMBOL'] == '' or data.loc[i,'SYMBOL'] == None:
        data.loc[i,'SYMBOL'] = df_aux.loc[0,'SYMBOL']
    elif df_aux.shape[0] < 1:
        print('No está el valor:', data.loc[i,'INSTR_ID'], '-', data.loc[i,'DESCRIPTION'], '-', data.loc[i,'CUSIP'] )
    elif df_aux.shape[0] > 1:
        print('El valor:', data.loc[i,'INSTR_ID'], data.loc[i,'DESCRIPTION'],' está duplicado en el catálogo, hace falta revisar')

# LTV %
data["LTV $"] = data["MKT_VALUE_USD"] * data["LTV %"]

comp = {
    'Davinci II': 'DVCII',
    'Cronus CV': 'CRO',
    'Cronus Private Inv': 'CPI',
    'Davinci': 'DVC',
    'GT': 'GT',
    'Kalispera CV': 'KLP',
    'NFD': 'NFD',
    'SC Holdings': 'SCH',
    'Signature Capital': 'SCAP',
    'WK Capital': 'WK',
    'Nemesis': 'NMS',
    '-': '-'
}
data['COMPANY'] = data.COMPANY.map(comp)

# Calculamos el costo total de SSZ
if ssz_data.shape[0] != 0:
    cond1 = (data.BANK == 'SSZ') & (data.ASSET_CLASS.isin(['COLATERAL', 'EQUITIES']))
    cond2 = (data.BANK == 'SSZ') & (data.ASSET_CLASS.isin(['STRUCTURED NOTES', 'LIQUIDITY']))
    cond3 = (data.BANK == 'SSZ') & (data.ASSET_CLASS == 'FIXED INCOME') & (data.SUB_ASSET_CLASS != 'CORPORATE BOND')
    cond4 = (data.BANK == 'SSZ') & (data.ASSET_CLASS == 'FIXED INCOME') & (data.SUB_ASSET_CLASS == 'CORPORATE BOND')
    cond5 = (data.BANK == 'SSZ') & (data.ASSET_CLASS == 'STRUCTURED NOTES') #& ()
    cond6 = (data.BANK == 'SSZ') & (data.ASSET_CLASS.isin(['LIQUIDITY', 'FIXED INCOME'])) & (data.SUB_ASSET_CLASS.isin(['CORPORATE BOND', 'DISTRESSED BOND', 'GOV SECURITIES', 'PREF STOCK']))
    cond7 = (data.BANK == 'SSZ') & (data.ASSET_CLASS == 'LIQUIDITY') & (data.SUB_ASSET_CLASS == 'GOV SECURITIES')
    cond8 = (data.BANK == 'SSZ') & (data.ASSET_CLASS.isin(['CASH', 'CREDITO']))

    data.loc[cond1, 'TOTAL_COST'] = data.loc[cond1, 'UNIT_COST'] * data.loc[cond1, 'QUANTITY']
    data.loc[cond2, 'TOTAL_COST'] = data.loc[cond2, 'MKT_VALUE_USD'] - data.loc[cond2, 'UNRELAIZED_GAIN/LOSS']
    data.loc[cond3, 'TOTAL_COST'] = data.loc[cond3, 'UNIT_COST'] * data.loc[cond3, 'QUANTITY']
    data.loc[cond4, 'TOTAL_COST'] = (data.loc[cond4, 'UNIT_COST'] / 100) * data.loc[cond4, 'QUANTITY']
    data.loc[cond5, 'TOTAL_COST'] = data.loc[cond5, 'TOTAL_COST'] / 100
    data.loc[cond8, 'TOTAL_COST'] = 0
    data.loc[cond8, 'UNRELAIZED_GAIN/LOSS'] = 0
    data.loc[cond5, 'UNRELAIZED_GAIN/LOSS'] = data.loc[cond5, 'MKT_VALUE_USD'] - (data.loc[cond5, 'QUANTITY'] * data.loc[cond5, 'FXRATE'])
    data.loc[cond6, 'UNRELAIZED_GAIN/LOSS'] = data.loc[cond6, 'MKT_VALUE_USD'] - (data.loc[cond6, 'QUANTITY'] * data.loc[cond6, 'UNIT_COST'] / 100 * data.loc[cond6, 'FXRATE'])
    data.loc[cond7, 'TOTAL_COST'] = (data.loc[cond7, 'UNIT_COST'] / 100) * data.loc[cond7, 'QUANTITY']

# =============================================================================
#               Verificar si estan todas las cuentas (Corregir esto)
# =============================================================================
# print(f'No está(n) la(s) cuenta(s): {set(catPortf.Cuenta.unique()) - set(data.)}')

# =============================================================================
#                          Creación de la hoja de LTV
# =============================================================================
df_ltv = pd.DataFrame(columns=['BANK','MKT_VALUE_USD','LOAN','LTV','LTV-LOAN','PCT'])

for bnk in data['BANK'].unique():
    # if bnk == 'SSZ' or bnk=='JBR':
    if bnk == 'JBR':
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

# with pd.ExcelWriter(f'{path_pc}/Contributtion/temp/Assets_Full/{DATE}-ASSET_ALLOC_PRUEBA.xlsx', engine = "openpyxl") as writer:
with pd.ExcelWriter(PARENT_PATH / 'Contributtion' / 'temp' / 'Assets_Full' / f'{DATE}-ASSET_ALLOC_PRUEBA.xlsx', engine = "openpyxl") as writer:
    # Crea la primer hoja de datos
    data.to_excel(writer, sheet_name = "CONSOLIDADO", index = False)
    # Crea la segunda hoja de datos
    data[(data["ASSET_CLASS"] != 'COLATERAL') & (data["ASSET_CLASS"] != 'CREDITO') ].to_excel(writer, sheet_name = "CONSOLIDADO SIN COLATERALES", index = False)
    # Crea la tercera hoja de datos
    df_ltv.to_excel(writer, sheet_name = "LTV", index = False)