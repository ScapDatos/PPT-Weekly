### ================================================== ###
### Verificar que cada posicion este en el diccionario
### de lo contrario, agregarlo correctamente.
### Notas:
###     - Actualizar la fecha al viernes inmediato anterior.
###     - Actualizar los tipos de cambio de MXNUSD, EURUSD.
### ================================================== ###

import os
import re
import pandas as pd

# ======================= Actualizaciones manuales ============================
date = "24072026"     # Fecha de los holdings (Debe coincidir con los holdings guardados)
mxnusd = 0.05728      # Tipo de cambio MXN->USD
eurusd = 1.1721       # Tipo de cambio EUR->USD

# Dirección de tu PC a la nube:
# contPath = f'C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Contributtion'
contPath = 'C:/Users/arnol/OneDrive/0. Nube Asset Mgmt/Contributtion'

# Lista que contiene las cuentas que no pueden faltar
catPortf = pd.read_excel(f'{contPath}/Diccionarios.xlsx', sheet_name = "Cuentas",
                         converters = {'Cuenta':str}, engine = "openpyxl")

ctas = catPortf[catPortf.Status == 'ACT'].Cuenta.values.tolist()

assetClass = pd.read_excel(f'{contPath}/Diccionarios.xlsx', sheet_name = "Asset_class",
                           converters = {'Cuenta':str}, engine = "openpyxl")

# =============================================================================
def portfolio(acct, col:str):
    """ Identifica a que portafolio pertenece la cuenta """
    dic = catPortf[catPortf['Cuenta'] == acct][['Portafolio', 'Empresa']]
    r = dic[col].values[0]
    return r

def MktValueUSD(value, ccy): 
    """ Ajusta al monto de acuerdo a la moneda que se elija en ccy usando el 
    tipo de cambio de la primer sección"""
    if ccy == 'MXN':
        return value * mxnusd
    elif ccy == 'EUR':
        return value * eurusd
    else:
        return value
    
# ========================= Obtiene todos los archivos de holdings ============
banks = ['CITI', 'DB', 'GS', 'JB', 'JPM', 'MS', 'SS', 'UBS']

bank_dirs = {}
for bank in banks:
    files = os.listdir(os.path.join(contPath, bank))
    assets = []
    for file in files:
        if date in file:
            assets.append(os.path.join(contPath, bank, file))
    bank_dirs[bank] = assets
# =============================================================================
# Crea el nuevo Dataframe con las opciones que requiere
df = pd.read_excel(f'{contPath}/Template.xlsx', engine = "openpyxl")
df['MXNUSD'] = mxnusd
df['EURUSD'] = eurusd

df['MKT_VALUE_USD'] = [MktValueUSD(df['MKT_VALUE'][i], df['CCY'][i] ) for i in range(len(df))]
main_cols = df.columns

# ==================== Obtención de los datos por archivo =====================
for bank, dirs in bank_dirs.items():
    if bank == "CITI":
        print("Using Citi holdings...")
        # Busca en los archivo de la carpeta SS         
        for file in dirs:
            # lee el excel
            df_aux = pd.read_csv(file)
            df_aux.loc[(df_aux.Description != 'LOAN') & (df_aux.Description != 'COMMERCIAL LOAN'), 'DESCRIPTION'] = df_aux.loc[(df_aux.Description != 'LOAN') & (df_aux.Description != 'COMMERCIAL LOAN'), 'Description']
            df_aux.loc[(df_aux.Description == 'LOAN') | (df_aux.Description == 'COMMERCIAL LOAN'), 'DESCRIPTION'] = 'CREDITO Citibank'
            
            df_aux['Account Number'] = df_aux['Account Number'].astype(str).str.replace(',', '').str.replace('(', '-').str.replace(')', '').str[-5:].str.replace('X', '')
            df_aux['PORTFOLIO'] = df_aux['Account Number'].apply(portfolio, col = 'Portafolio')
            df_aux['BANK'] = 'CTI'
            df_aux['ACCT'] = df_aux['Account Number']
            df_aux['COMPANY'] = df_aux['Account Number'].apply(portfolio, col = 'Empresa')
            df_aux['ASSET_CLASS'] = ''
            df_aux['SUB_ASSET_CLASS'] = ''
            df_aux['SHORT_NAME'] = ''
            df_aux['CUSIP'] = ''
            df_aux['SYMBOL'] = df_aux['Symbol']
            df_aux['QUANTITY'] = df_aux['Nominal Units'].astype(str).str.replace('(', '-').str.replace(')', '').str.replace('.000', '').astype(float)
            df_aux['UNIT_COST'] = df_aux['Average or Unit Cost']
            df_aux['UNIT_COST_CCY'] = df_aux['Reporting CCY']
            df_aux['CCY'] = df_aux['Nominal CCY']
            df_aux['MKT_VALUE'] = df_aux['Market Value (Nominal CCY)'].astype(str).str.replace(',', '').str.replace('(', '-').str.replace(')', '').astype(float)
            df_aux['MKT_VALUE_USD'] = df_aux['Current Value (Reporting CCY)'].astype(str).str.replace(',', '').str.replace('(', '-').str.replace(')', '').astype(float)
            df_aux['TOTAL_COST'] = df_aux['Total Cost Basis (Reporting CCY)'].astype(str).str.replace(',', '').astype(float)
            df_aux['UNRELAIZED_GAIN/LOSS'] = df_aux['Unrealized Gain/(Loss) (Reporting CCY)'].astype(str).str.replace(',', '').str.replace('(', '-').str.replace(')', '').astype(float)
            df_aux['MXNUSD'] = mxnusd
            df_aux['EURUSD'] = eurusd
            df_aux['EXPOSITION_CCY'] = ''
            df_aux['Actualizado'] = ''
            df_aux.loc[df_aux.DESCRIPTION == 'CREDITO Citibank', 'SYMBOL'] = 'LOAN'
            
            cond = (df_aux.Description == 'LOAN') | (df_aux.Description == 'COMMERCIAL LOAN')
            df_aux.loc[cond, 'MKT_VALUE'] = -1 * df_aux.loc[cond, 'MKT_VALUE']
            df_aux.loc[cond, 'MKT_VALUE_USD'] = -1 * df_aux.loc[cond, 'MKT_VALUE_USD']
            
            df = pd.concat([df, df_aux[main_cols]], axis = 0)
            
    elif bank == 'DB':
        print("Using DB holdings...")
        # Busca en los archivo de la carpeta DB
        for file in dirs:
            # lee el excel
            df_aux = pd.read_excel(file, engine = 'openpyxl', header = 12).drop('Unnamed: 0', axis = 1)
            
            df_aux['PORTFOLIO'] = df_aux['Account Number'].astype(str).str[-4:].apply(portfolio, col = 'Portafolio')
            df_aux['BANK'] = 'DBK'
            df_aux['ACCT'] = df_aux['Account Number'].astype(str).str[-4:]
            df_aux['COMPANY'] = df_aux['Account Number'].astype(str).str[-4:].apply(portfolio, col = 'Empresa')
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
            
            df = pd.concat([df, df_aux[main_cols]], axis = 0)
            
    elif bank == 'GS':
        print("Using GS holdings...")
        # Para cada archivo de GS
        for file in dirs:
            # Lee el archivo y Limpia las filas que no tienen Market Value
            # df_aux = pd.read_excel(file, engine = 'openpyxl', header=7).dropna(subset=["Market Value"]).reset_index(drop = True)
            df_aux = pd.read_csv(file, header = 9).dropna(subset = ['Market Value']).reset_index(drop = True)
            
            # Para cada activo en la tabla
            df_aux['PORTFOLIO'] = portfolio(file[-7:-4], col = 'Portafolio')
            df_aux['BANK'] = 'GSS'
            df_aux['ACCT'] = file[-7:-4]
            df_aux['COMPANY'] = portfolio(file[-7:-4], col = 'Empresa')
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
            
            df = pd.concat([df, df_aux[main_cols]], axis = 0)
            
    elif bank == 'JB':
        print("Using JB holdings...")
        # Busca en los archivo de la carpeta DB
        for file in dirs:
            # lee el excel
            df_aux = pd.read_csv(file).dropna(subset=["Instrument"]).reset_index(drop = True)
            
            df_aux['PORTFOLIO'] = portfolio(file[-8:-4], col = 'Portafolio')
            df_aux['BANK'] = 'JBR'
            df_aux['ACCT'] = file[-8:-4]
            df_aux['COMPANY'] = portfolio(file[-8:-4], col = 'Empresa')
            df_aux['ASSET_CLASS'] = ''
            df_aux['SUB_ASSET_CLASS'] = ''
            df_aux['SHORT_NAME'] = ''
            df_aux['DESCRIPTION'] = df_aux.Instrument
            df_aux['ISIN'] = ''
            df_aux['CUSIP'] = ''
            df_aux['SYMBOL'] = ''
            df_aux['QUANTITY'] = df_aux['Quantity / Nominal'].astype(str).str.replace("'", '').astype(float)
            df_aux['UNIT_COST'] = ''
            df_aux['UNIT_COST_CCY'] = df_aux.Curr
            df_aux['CCY'] = 'USD'
            df_aux['MKT_VALUE'] = df_aux['Total value USD'].astype(str).str.replace("'", '').astype(float)
            df_aux['MKT_VALUE_USD'] = df_aux['Total value USD'].astype(str).str.replace("'", '').astype(float)
            df_aux['TOTAL_COST'] = df_aux['Purch. price']
            df_aux['UNRELAIZED_GAIN/LOSS'] = ''
            df_aux['MXNUSD'] = mxnusd
            df_aux['EURUSD'] = eurusd
            df_aux['EXPOSITION_CCY'] = ''
            df_aux['Actualizado'] = ''
            
            df = pd.concat([df, df_aux[main_cols]], axis = 0)
            
    elif bank == 'JPM':
        print("Using JPM holdings...")
        # Busca en los archivo de la carpeta DB
        for file in dirs:
            # lee el excel
            df_aux = pd.read_csv(file).dropna(subset=["Description"]).reset_index(drop = True)
            # df_aux = pd.read_excel(file, engine = 'openpyxl').dropna(subset=["Description"]).reset_index(drop = True)
            df_aux['Local CCY'] = df_aux['Local CCY'].fillna(df_aux['Base CCY'])
            
            df_aux['PORTFOLIO'] = portfolio(file[-8:-4], col = 'Portafolio')
            df_aux['BANK'] = 'JPM'
            df_aux['ACCT'] = file[-8:-4]
            df_aux['COMPANY'] = portfolio(file[-8:-4], col = 'Empresa')
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
            
            df = pd.concat([df, df_aux[main_cols]], axis = 0)
            
    elif bank == 'MS':
        print("Using MS holdings...")
        # Busca en los archivo de la carpeta DB
        for file in dirs:
            df_aux = pd.read_excel(file, header = 10, engine = 'openpyxl').dropna(subset=["Open Order"]).reset_index(drop = True)
            df_aux = df_aux[df_aux['Account Number'] != 'Total'].reset_index(drop = True)
            
            df_aux['PORTFOLIO'] = df_aux['Account Number'].astype(str).str[-4:].apply(portfolio, col = 'Portafolio')
            df_aux['BANK'] = 'MSY'
            df_aux['ACCT'] = df_aux['Account Number'].astype(str).str[-4:]
            df_aux['COMPANY'] = df_aux['Account Number'].astype(str).str[-4:].apply(portfolio, col = 'Empresa')
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
            
            df = pd.concat([df, df_aux[main_cols]], axis = 0)
            
    elif bank == 'UBS':
        print("Using UBS holdings...")
        # Busca en los archivo de la carpeta UBS
        for file in dirs:
            # lee el excel
            df_aux = pd.read_csv(file, skiprows=1).fillna(0)
            
            # df_aux['PORTFOLIO'] = portfolio(df_aux['ACCOUNT NUMBER'][0].replace('(', '').replace(')', '').split()[1], col = 'Portafolio')
            # print(df_aux['ACCOUNT NUMBER'].str.replace('(', '').str.replace(')', '').str.split()[1])
            df_aux['ACCOUNT NUMBER'] = df_aux['ACCOUNT NUMBER'].str.replace('(', '').str.replace(')', '').str.split().str[1]
            df_aux['PORTFOLIO'] = df_aux['ACCOUNT NUMBER'].apply(portfolio, col = 'Portafolio')
            df_aux['BANK'] = 'UBS'
            df_aux['ACCT'] = df_aux['ACCOUNT NUMBER']#[0].replace('(', '').replace(')', '').split()[1]
            # df_aux['COMPANY'] = portfolio(df_aux['ACCOUNT NUMBER'][0].replace('(', '').replace(')', '').split()[1], col = 'Empresa')
            df_aux['COMPANY'] = df_aux['ACCOUNT NUMBER'].apply(portfolio, col = 'Empresa')
            df_aux['ASSET_CLASS'] = ''
            df_aux['SUB_ASSET_CLASS'] = ''
            df_aux['SHORT_NAME'] = ''
            df_aux['ISIN'] = ''
            df_aux['SYMBOL'] = ''
            df_aux['QUANTITY'] = df_aux.QUANTITY.astype(str).str.replace(',', '').astype(float)
            df_aux['UNIT_COST'] = ''
            df_aux['UNIT_COST_CCY'] = 'USD'
            df_aux['CCY'] = 'USD'
            df_aux['MKT_VALUE'] = df_aux.VALUE.astype(str).str.replace(',', '').str.replace('$', '').astype(float)
            df_aux['MKT_VALUE_USD'] = df_aux.VALUE.astype(str).str.replace(',', '').str.replace('$', '').astype(float)
            df_aux['TOTAL_COST'] = ''
            df_aux['UNRELAIZED_GAIN/LOSS'] = ''
            df_aux['MXNUSD'] = mxnusd
            df_aux['EURUSD'] = eurusd
            df_aux['EXPOSITION_CCY'] = ''
            df_aux['Actualizado'] = ''
            
            df = pd.concat([df, df_aux[main_cols]], axis = 0)
            
    elif bank == 'SS':
        print("Using SS holdings...")
        # Busca en los archivo de la carpeta SS
        for file in dirs:
            # lee el excel
            df_aux = pd.read_excel(file, header = 14)
            
            empty_rows = df_aux[df_aux.isnull().all(axis = 1)].index.tolist()
            splits = [0] + empty_rows + [df_aux.shape[0]]
            
            tablas = []
            for i in range(len(splits) - 1):
                start = splits[i] + (1 if i != 0 else 0)
                end = splits[i + 1]
                
                bloque = df_aux.iloc[start:end].dropna(how = 'all').reset_index(drop = True)
                
                if not bloque.empty:
                    nulls_row_0 = bloque.iloc[0].isnull().sum()
                    nulls_row_1 = bloque.iloc[1].isnull().sum()
                    if nulls_row_0 > nulls_row_1:
                        header = 1
                    else:
                        header = 0
                    
                    bloque.columns = bloque.iloc[header]
                    bloque = bloque[(header + 1):].reset_index(drop = True)
                    tablas.append(bloque)
            
            df_aux = []
            for idx, tabla in enumerate(tablas):
                c = ''
                cols = []
                for idx, col in enumerate(tabla.columns):
                    if col == 'CURRENCY':
                        col = f'CURRENCY{c}'
                        c = f'{idx}'
                    cols.append(col)
                tabla.columns = cols
                
                tabla['PORTFOLIO'] = portfolio(file[-10:-5], col = 'Portafolio')
                tabla['BANK'] = 'SSZ'
                tabla['ACCT'] = file[-10:-5]
                tabla['COMPANY'] = portfolio(file[-10:-5], col = 'Empresa')
                tabla['ASSET_CLASS'] = ''
                tabla['SUB_ASSET_CLASS'] = ''
                tabla['SHORT_NAME'] = ''
                
                if tabla.columns[0] == 'LIQUIDITY':
                    tabla['DESCRIPTION'] = tabla.LIQUIDITY + '-' + tabla.CURRENCY
                    tabla['ISIN'] = ''
                    tabla['QUANTITY'] = tabla['MARKET VALUE']
                    tabla['UNIT_COST'] = ''
                    tabla['UNIT_COST_CCY'] = ''
                    tabla['TOTAL_COST'] = 0
                    tabla["UNRELAIZED_GAIN/LOSS"] = 0
                else:
                    tabla['DESCRIPTION'] = tabla[tabla.columns[0]]
                    tabla['UNRELAIZED_GAIN/LOSS'] = tabla["UNREALIZED GAINS / LOSSES"]
                    tabla['UNIT_COST_CCY'] = tabla['CURRENCY']
                
                tabla['CUSIP'] = ''
                tabla['SYMBOL'] = ''
                tabla['CCY'] = 'USD'
                tabla['MKT_VALUE'] = tabla['BALANCE IN REF. CURRENCY']
                tabla['MKT_VALUE_USD'] = tabla['BALANCE IN REF. CURRENCY']
                tabla['MXNUSD'] = mxnusd
                tabla['EURUSD'] = eurusd
                tabla['EXPOSITION_CCY'] = ''
                tabla['Actualizado'] = ''
                
                if tabla.columns[0] in ['EMERGING EQUITIES STOCKS', 'US EQUITIES STOCKS', 'US EQUITIES FUNDS', 'EUROPEAN EQUITIES STOCKS', 'INVESTMENT GRADE FIXED INCOME FUNDS', 'EMERGING EQUITIES FUNDS']:
                    tabla['UNIT_COST'] = tabla['UNIT COST']
                    tabla['TOTAL_COST'] = tabla['UNIT COST'] * tabla['QUANTITY']
                elif tabla.columns[0] == 'STRUCTURED NOTES':
                    tabla['QUANTITY'] = tabla.NOMINAL
                    tabla['UNIT_COST'] = tabla['UNIT COST']
                    tabla['TOTAL_COST'] = tabla['BALANCE IN REF. CURRENCY'] - tabla['UNREALIZED GAINS / LOSSES']
                elif tabla.columns[0] == 'EMERGING FIXED INCOME BONDS':
                    tabla['QUANTITY'] = tabla.NOMINAL
                    tabla['UNIT_COST'] = ''
                    tabla['TOTAL_COST'] = tabla['BALANCE IN REF. CURRENCY'] - tabla['UNREALIZED GAINS / LOSSES']
                elif tabla.columns[0] == 'COMMERCIAL PAPER':
                    tabla['QUANTITY'] = tabla.NOMINAL
                    tabla['UNIT_COST'] = tabla['COST']
                    tabla['TOTAL_COST'] = tabla['BALANCE IN REF. CURRENCY'] - tabla['UNREALIZED GAINS / LOSSES']
                elif tabla.columns[0] == 'LIQUIDITY':
                    pass
                else:
                    tabla['QUANTITY'] = tabla.NOMINAL #''
                    tabla['UNIT_COST'] = tabla['UNIT COST'] if 'UNIT COST' in tabla.columns else tabla['COST']
                    tabla['TOTAL_COST'] = tabla['BALANCE IN REF. CURRENCY'] - tabla['UNREALIZED GAINS / LOSSES']
                
                tabla = tabla[main_cols].reset_index(drop = True)
                df_aux.append(tabla)
            
            df_aux = pd.concat(df_aux, axis = 0)
            df = pd.concat([df, df_aux[main_cols]], axis = 0)

# =============================================================================
#               Limpieza de la tabla y creacion de columnas faltantes
# =============================================================================
# Quitamos los que no tienen valores en "MKT_VALUE"
df = df.dropna(subset=["MKT_VALUE"]).reset_index(drop = True)
# Los que no tienen cantidad le ponemos 0
# df["QUANTITY"] = df["QUANTITY"].fillna(0)
# Todos los que traen guión en "Quantity" lo cambia por 0
df["QUANTITY"] = df["QUANTITY"].apply(lambda x: 0 if x=='-' else x).fillna(0).astype(float)
# Todos los que traen guión en "MARKET VALUE" lo cambia por 0
df["MKT_VALUE"] = df["MKT_VALUE"].apply(lambda x: 0 if x=='-' else x)
# Convierte la columna CUENTA a STR para su correcto manejo
df["ACCT"] = df["ACCT"].astype(str)
# Agrega la columna de sponsor
df["SPONSOR"] = ''
# LTV %
df["LTV %"] = 0.0

# =============================================================================
#                Sección para llenar las columnas
# =============================================================================
# for idx, desc in enumerate(df.DESCRIPTION):
for idx, row in df.iterrows():
    # df_aclass = assetClass[assetClass['NOMBRE'] == desc].reset_index(drop = True)
    df_aclass = assetClass[assetClass['NOMBRE'] == row.DESCRIPTION].reset_index(drop = True)
    
    if df_aclass.empty:
        # print(f"No está el valor: | {desc} | {df.loc[idx, 'CUSIP']} | {df.loc[idx, 'ISIN']}")
        print(f"No está el valor: | {row.DESCRIPTION} | {df.loc[idx, 'CUSIP']} | {df.loc[idx, 'ISIN']}")
    elif df_aclass.shape[0] == 1:
        df.loc[idx, 'ASSET_CLASS'] = df_aclass.loc[0, 'NEW_ASSET_CLASS']
        df.loc[idx, 'SUB_ASSET_CLASS'] = df_aclass.loc[0, 'NEW_SUB_ASSET_CLASS']
        df.loc[idx, 'SHORT_NAME'] = df_aclass.loc[0, 'SHORT_NAME']
        df.loc[idx, 'EXPOSITION_CCY'] = df_aclass.loc[0, 'EXPOSITION_CCY']
        df.loc[idx, 'LTV %'] = df_aclass.loc[0, 'LTV %']
        df.loc[idx, 'SPONSOR'] = df_aclass.loc[0, 'SPONSOR']
        if df.loc[idx, 'SYMBOL'] == '' or df.loc[idx, 'SYMBOL'] == None:
            df.loc[idx, 'SYMBOL'] = df_aclass.loc[0, 'SYMBOL']
    elif df_aclass.shape[0] > 1:
        for jdx, cus in enumerate(df_aclass.CUSIP):
            df_aclass2 = assetClass[assetClass['CUSIP'] == cus].reset_index(drop = True)
            if df_aclass2.empty:
                print(f"No está el valor: | {df_aclass.loc[jdx, 'NOMBRE']} | {cus} | {df_aclass.loc[jdx, 'ISIN']}")
            elif df_aclass2.shape[0] == 1:
                df.loc[idx, 'ASSET_CLASS'] = df_aclass2.loc[0, 'NEW_ASSET_CLASS']
                df.loc[idx, 'SUB_ASSET_CLASS'] = df_aclass2.loc[0, 'NEW_SUB_ASSET_CLASS']
                df.loc[idx, 'SHORT_NAME'] = df_aclass2.loc[0, 'SHORT_NAME']
                df.loc[idx, 'EXPOSITION_CCY'] = df_aclass2.loc[0, 'EXPOSITION_CCY']
                if df.loc[idx, 'SYMBOL'] == '' or df.loc[idx, 'SYMBOL'] == None:
                    df.loc[idx, 'SYMBOL'] = df_aclass2.loc[0, 'SYMBOL']
            elif df_aclass2.shape[0] > 1:
                print(f"El valor: | {df_aclass2.iloc[jdx, 'NOMBRE']} | {cus} | {df_aclass2.loc[jdx, 'ISIN']} esta duplicado en el catalogo. Revisar.")
    
    if row.BANK == 'JBR':
        df.loc[idx, 'ISIN'] = df_aclass.loc[0, 'ISIN']

# LTV %
df["LTV $"] = df["MKT_VALUE_USD"].astype(float) * df["LTV %"].astype(float)

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
df['COMPANY'] = df.COMPANY.map(comp)

# =============================================================================
#               Verificar si estan todas las cuentas
# =============================================================================
# for cta in df.ACCT.unique():
#     if cta not in ctas:
#         print(f'No esta la cuenta: {cta}')

for cta in ctas:
    if cta not in df['ACCT'].unique():
        print('No está la cuenta: ',cta)

# =============================================================================
#                          Creación de la hoja de LTV
# =============================================================================
df_ltv = pd.DataFrame(columns=['BANK','MKT_VALUE_USD','LOAN','LTV','LTV-LOAN','PCT'])

for bnk in df['BANK'].unique():
    if bnk == 'JBR': #bnk == 'SSZ' or bnk == 'JBR':
        try:
            df_aux = df[(df['BANK'] == bnk) & (df['ASSET_CLASS'] != 'CREDITO')][['BANK', 'MKT_VALUE_USD', 'LTV $']].groupby(['BANK']).sum()
        except:
            print('Agrega los valores del credito de JBR o SSZ al diccionario.')
        df_ltv.loc[bnk, 'BANK'] = bnk
        df_ltv.loc[bnk, 'MKT_VALUE_USD'] = df_aux['MKT_VALUE_USD'].loc[bnk]
        df_ltv.loc[bnk, 'LTV'] = df_aux['LTV $'].loc[bnk]
        df_aux = df[(df['BANK'] == bnk) & (df['ASSET_CLASS'] == 'CREDITO')][['BANK', 'MKT_VALUE_USD', 'LTV $']].groupby(['BANK']).sum()
        df_ltv.loc[bnk, 'LOAN'] = df_aux['MKT_VALUE_USD'].loc[bnk]
        df_ltv.loc[bnk, 'LTV-LOAN'] = df_ltv.loc[bnk, 'LTV']+df_ltv.loc[bnk, 'LOAN']
        df_ltv.loc[bnk, 'PCT'] = str(round((df_ltv.loc[bnk, 'LTV-LOAN'] / df_ltv.loc[bnk, 'LTV']) * 100, 2)) + '%'
    elif bnk == 'DBK':
        df_aux = df[((df['BANK'] == bnk) & (df['ACCT'] == '2269')) & (df['ASSET_CLASS'] != 'CREDITO') ][['BANK', 'MKT_VALUE_USD', 'LTV $']].groupby(['BANK']).sum()
        df_ltv.loc[bnk, 'BANK'] = bnk
        df_ltv.loc[bnk, 'MKT_VALUE_USD'] = df_aux['MKT_VALUE_USD'].loc[bnk]
        df_ltv.loc[bnk, 'LTV'] = df_aux['LTV $'].loc[bnk]
        df_aux = df[(df['BANK'] == bnk) & (df['ASSET_CLASS'] == 'CREDITO')][['BANK', 'MKT_VALUE_USD', 'LTV $']].groupby(['BANK']).sum()
        df_ltv.loc[bnk, 'LOAN'] = df_aux['MKT_VALUE_USD'].loc[bnk]
        df_ltv.loc[bnk, 'LTV-LOAN'] = df_ltv.loc[bnk, 'LTV'] + df_ltv.loc[bnk, 'LOAN']
        df_ltv.loc[bnk, 'PCT'] = str(round((df_ltv.loc[bnk, 'LTV-LOAN'] / df_ltv.loc[bnk, 'LTV']) * 100, 2)) + '%'
    elif bnk=='CTI':
        df_aux = df[(df['BANK']==bnk) & (df['ASSET_CLASS']!='CREDITO')&(df['ACCT']!='11768')][['BANK','MKT_VALUE_USD','LTV $']].groupby(['BANK']).sum()
        df_ltv.loc[bnk,'BANK'] = bnk
        df_ltv.loc[bnk,'MKT_VALUE_USD'] = df_aux['MKT_VALUE_USD'].loc[bnk]
        df_ltv.loc[bnk,'LTV'] = df_aux['MKT_VALUE_USD'].loc[bnk]
        df_aux = df[(df['BANK']==bnk) & (df['ASSET_CLASS']=='CREDITO')][['BANK','MKT_VALUE_USD','LTV $']].groupby(['BANK']).sum()
        df_ltv.loc[bnk,'LOAN'] = df_aux['MKT_VALUE_USD'].loc[bnk]
        df_ltv.loc[bnk,'LTV-LOAN'] = df_ltv.loc[bnk,'LTV']+df_ltv.loc[bnk,'LOAN']
        df_ltv.loc[bnk,'PCT'] = str(round((df_ltv.loc[bnk,'LTV-LOAN']/df_ltv.loc[bnk,'LTV'])*100,2))+'%'

df_ltv['MKT_VALUE_USD'] = df_ltv['MKT_VALUE_USD'].astype('float64')
df_ltv['LOAN'] = df_ltv['LOAN'].astype('float64')
df_ltv['LTV'] = df_ltv['LTV'].astype('float64')
df_ltv['LTV-LOAN'] = df_ltv['LTV-LOAN'].astype('float64')


# =============================================================================
#                         Creacion del archivo excel
# =============================================================================
with pd.ExcelWriter(f"{contPath}/{pd.to_datetime(date, format = '%d%m%Y').strftime('%Y-%m-%d')}-ASSET_ALLOC.xlsx", engine = 'openpyxl') as writer:
    # Crea la primer hoja de datos
    df.to_excel(writer, sheet_name = 'CONSOLIDADO', index=False)
    # Crea la segunda hoja de datos
    df[(df['ASSET_CLASS'] != 'COLATERAL') & (df['ASSET_CLASS'] != 'CREDITO')].to_excel(writer, sheet_name = 'CONSOLIDADO SIN COLATERALES', index = False)
    # Crea la tercera hoja de datos
    df_ltv.to_excel(writer, sheet_name = 'LTV', index = False)