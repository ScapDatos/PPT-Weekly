### ============================================= ###
### Notas:
###     - Actualizar la fecha al viernes inmediato anterior.
### ============================================= ###

import pandas as pd
from config import DATE, PARENT_PATH
from debugg.functions import treemap_aclass
pd.options.mode.chained_assignment = None

folderPath = PARENT_PATH / 'Contributtion'
folderPathWrite = PARENT_PATH / 'Asset Allocation' / 'Vistas'

df = pd.read_excel(f'{folderPath}/{DATE}-ASSET_ALLOC.xlsx', sheet_name = 'CONSOLIDADO SIN COLATERALES')

dfPlot = df[["PORTFOLIO", 'COMPANY', "ASSET_CLASS", "SUB_ASSET_CLASS", "BANK", "DESCRIPTION", "SHORT_NAME",
             "ACCT", "CCY", "MKT_VALUE", "MKT_VALUE_USD", "EXPOSITION_CCY" ]]

# dfPlot['BNK_ACCT'] = dfPlot.BANK.loc[:] + '-' + dfPlot.ACCT.loc[:].astype(str) # ORIGINAL
dfPlot['BNK_ACCT'] = dfPlot.COMPANY + '/' + dfPlot.BANK + '-' + dfPlot.ACCT.astype(str)
# dfPlot['BANK'] = dfPlot.COMPANY + '/' + dfPlot.BANK # AGREGADO POR MI
dfPlot['DESCRIPTION'] = dfPlot.DESCRIPTION.str.replace('/', '-')
dfPlot['DESCRIPTION_KEY'] = dfPlot['BNK_ACCT'] + '-' + dfPlot.DESCRIPTION

sHCITY = ['HOTELES CITY EXPRESS S.A. DE C*.V. CMN', 'HOTELES CITY EXPRE COM NPV', 'HOTELES CITY EXPRESS SAB DE',
          'HOTELES CITY EXPRE COM NPV, Ticker HCITY *', 'PROMOTORA DE HOTELES NORTE ISIN MX01HC000001 SEDOL BBL5196']

# Crea la tabla sin HCITY
dfPlotHcity = dfPlot[dfPlot.SHORT_NAME != 'HCITY']

# Filtra por valores que SOLO son positivos
dfPlot = dfPlot[dfPlot.MKT_VALUE_USD >= 0]
dfPlot['MKT_VALUE_USD'] = dfPlot.MKT_VALUE_USD.round(2)

# Quita los valores vacios
dfPlot = dfPlot.dropna().reset_index(drop = True)

# Filtra por valores que SOLO son positivos para HCITY
dfPlotHcity = dfPlotHcity[dfPlotHcity.MKT_VALUE_USD >= 0]
dfPlotHcity['MKT_VALUE_USD'] = dfPlotHcity.MKT_VALUE_USD.round(2)

# Quita los valores vacios para HCITY
dfPlotHcity = dfPlotHcity.dropna().reset_index(drop = True)

### ======================================================= ###
###                   Grafica CONSOLIDADO                   ###
### ======================================================= ###
# Calculo del NAV completo
nav_full = df.MKT_VALUE_USD.sum()

# TreeMap todo junto
treemap_aclass(dfPlot, title = '1 - ASSET MIX Completo',
               fields = ['ASSET_CLASS', 'SUB_ASSET_CLASS', 'PORTFOLIO', 'SHORT_NAME', 'BNK_ACCT'],
               path = folderPathWrite, nav = f'CONSOLIDADO | ${nav_full:,.0f} | 100%', html = True)

### ======================================================= ###
###          Grafica CONSOLIDADO sHCITY + VCFunds           ###
### ======================================================= ###
# Calculo del NAV completo
dfAux = dfPlotHcity[dfPlotHcity.PORTFOLIO != 'NFD']
navSHCITY = dfAux.MKT_VALUE_USD.sum()

# TreeMap todo junto
treemap_aclass(dfAux, title = '1 - ASSET MIX Consolidado sHcity + VCFunds',
               fields = ['ASSET_CLASS', 'SUB_ASSET_CLASS', 'PORTFOLIO', 'SHORT_NAME', 'BNK_ACCT'],
               path = folderPathWrite, nav = f'CONSOLIDADO | ${navSHCITY:,.0f} | 100%', html = True)

### ======================================================= ###
###                   Grafica por monedas                   ###
### ======================================================= ###
ccy_dict = {'USD': '#C3FFEB', 'EUR': '#AEE4FF', 'MXN': '#FAFFD0'}
treemap_aclass(dfPlot, title = '1 - ASSET MIX BY CURRENCY',
               fields = ["EXPOSITION_CCY", "ASSET_CLASS", "SUB_ASSET_CLASS", "PORTFOLIO", "BNK_ACCT", "SHORT_NAME"],
               path = folderPathWrite, nav = f'CONSOLIDADO | ${nav_full:,.0f} | 100%', html = True,
               ccy = True, dic = ccy_dict)

dfAux = dfPlot#.copy()
dfAux['BANK'] = dfAux.BANK.apply(lambda x: x if x != '-' else 'VCL')
treemap_aclass(dfAux, title = '1 - ASSET MIX BY BROKER',
              fields = ["BANK","ASSET_CLASS", "SUB_ASSET_CLASS", "BNK_ACCT", "SHORT_NAME"],
              path = folderPathWrite, nav = f'CONSOLIDADO | ${nav_full:,.2f} | 100%')

### ======================================================= ###
###             Grafica por Portafolio y Banco              ###
### ======================================================= ###
for portfolio in dfPlot.PORTFOLIO.unique():
    ### ============== GRAFICA POR PORTAFOLIO ============== ###
    if portfolio != '-':
        dfAux = dfPlot[dfPlot.PORTFOLIO == portfolio]
        nav = dfAux.MKT_VALUE_USD.sum()
        treemap_aclass(dfAux, title = f'2 - ASSET MIX {portfolio}',
                       fields = ['ASSET_CLASS', 'SUB_ASSET_CLASS', 'BNK_ACCT', 'SHORT_NAME'],
                       path = folderPathWrite, nav = f'{portfolio} | ${nav:,.2f} | 100%')
        
        treemap_aclass(dfAux, title = f'2 - ASSET MIX DIV-{portfolio}',
                       fields = ['ASSET_CLASS', 'SUB_ASSET_CLASS', 'BNK_ACCT', 'SHORT_NAME'],
                       path = folderPathWrite, nav = f'{portfolio} | ${nav:,.2f} | {(nav / nav_full) * 100:,.2f}%',
                       w = 16, h = 15)
        
        ### ========= GRAFICA POR PORTAFOLIO SIN HCITY ========= ###
        dfAuxCity = dfPlotHcity[dfPlotHcity.PORTFOLIO == portfolio]
        navhcity = dfAuxCity.MKT_VALUE_USD.sum()
        treemap_aclass(dfAuxCity, title = f'3 - ASSET MIX {portfolio} SIN HCITY',
                       fields = ['ASSET_CLASS', 'SUB_ASSET_CLASS', 'BNK_ACCT', 'SHORT_NAME'],
                       path = folderPathWrite, nav = f'{portfolio} | ${navhcity:,.2f} | 100%',
                       html = True)
        
        ### ================ GRAFICA POR BANCOS ================ ###
        for bank in dfAux.BANK.unique():
            if bank not in ['-', 'BNK', 'BVA']:
                nav = dfAux[dfAux.BANK == bank].MKT_VALUE_USD.sum()
                treemap_aclass(dfAux[dfAux.BANK == bank], title = f'4 - ASSET MIX {portfolio} {bank}',
                               fields = ['ASSET_CLASS', 'SUB_ASSET_CLASS', 'SHORT_NAME'],
                               path = folderPathWrite, nav = f'{portfolio} | {bank} | ${nav:,.2f} | 100%')

### ======================================================= ###
###             Grafica por NFD Notas | Lo demas            ###
### ======================================================= ###
dfAux = dfPlot[dfPlot.PORTFOLIO == 'NFD']
dfAux.loc[:, 'SUB'] = [dfAux.SHORT_NAME[snid].split(' ')[1] if dfAux.ASSET_CLASS[snid] == 'STRUCTURED NOTES' else '-'  for snid in dfAux.index]
dfAux.loc[:, 'COLOR'] = [dfAux.ASSET_CLASS[i] if dfAux.SUB[i] == '-' else dfAux.SUB[i] for i in dfAux.index]
treemap_aclass(dfAux, title = "5 - ASSET MIX NFD STRUCTURED NOTES",
              fields = ["ASSET_CLASS", "SUB_ASSET_CLASS","SUB", "BNK_ACCT", "SHORT_NAME"],
              path = folderPathWrite, nav = f'NFD | ${dfAux.MKT_VALUE_USD.sum():,.2f} | 100%',
              colorCol = 'COLOR', html = True)

### ======================================================= ###
###   Grafica por NFD Notas | Portafolio/Broker/Subyacente  ###
### ======================================================= ###
dfAux = dfPlot[dfPlot.ASSET_CLASS == 'STRUCTURED NOTES']
dfAux['SUB'] = [dfAux.SHORT_NAME[snid].split(' ')[1] if dfAux.ASSET_CLASS[snid] == 'STRUCTURED NOTES' else '-'  for snid in dfAux.index]
dfAux['COLOR'] = [dfAux.ASSET_CLASS[i] if dfAux.SUB[i] == '-' else dfAux.SUB[i] for i in dfAux.index ]
treemap_aclass(dfAux, title = "6 - ASSET MIX STRUCTURED NOTES BY BROKER",
              fields = ['PORTFOLIO','BANK',"SUB", "SHORT_NAME"],
              path = folderPathWrite, nav = f'NFD | ${dfAux.MKT_VALUE_USD.sum():,.2f} | 100%',
              colorCol='COLOR', html=True)

### ======================================================= ###
###               Grafica por BROKER SIN NOTAS              ###
### ======================================================= ###
dfAux = dfPlot[dfPlot.ASSET_CLASS != 'STRUCTURED NOTES']
dfAux['BANK'] = dfAux.BANK.apply(lambda x: x if x != '-' else 'VCL')
treemap_aclass(dfAux, title = '6 - ASSET MIX STRUCTURED NOTES BY BROKER',
              fields = ["BANK","ASSET_CLASS", "SUB_ASSET_CLASS", "PORTFOLIO", "SHORT_NAME", "BNK_ACCT"],
              path = folderPathWrite, nav = f'CONSOLIDADO | ${dfAux.MKT_VALUE_USD.sum():,.2f} | 100%')

### ======================================================= ###
###                  Grafica ETF BY SPONSOR                 ###
### ======================================================= ###
dfAux = df[df.SUB_ASSET_CLASS == "ETF"]
dfAux["BNK_ACCT"] = dfAux["BANK"] + '-' + dfAux["ACCT"].astype(str)
dfAux = dfAux[dfAux.MKT_VALUE_USD >= 0]
dfAux["MKT_VALUE_USD"] = dfAux.MKT_VALUE_USD.round(2)

dic_sponsor = {
    'Vanguard Group Inc': '#C2B994',
    'BlackRock Fund Advisors': '#89687E',
    'Invesco Capital Management LLC': '#97B7B3',
    'DBX Advisors LLC': '#867B74'
}

treemap_aclass(dfAux, title = '7 - ETF MIX BY SPONSOR',
               fields = ['SPONSOR', 'PORTFOLIO', 'ASSET_CLASS', 'SHORT_NAME', 'BNK_ACCT'],
               path = folderPathWrite, nav = f'CONSOLIDADO ETFs | ${dfAux.MKT_VALUE_USD.sum():,.2f} | 100%',
               html = True, ccy = True, dic = dic_sponsor)

### ======================================================= ###
###                     Grafica POR DEUDA                   ###
### ======================================================= ###
data = pd.read_excel(f'{folderPath}/{DATE}-ASSET_ALLOC.xlsx', engine = 'openpyxl')
data = data[data.ASSET_CLASS == 'CREDITO']
data['MKT_VALUE_USD'] = data.MKT_VALUE_USD * -1

treemap_aclass(data, title = "8 - ASSET MIX BY DEBT",
              fields = ["PORTFOLIO","BANK","DESCRIPTION","MKT_VALUE_USD"],
              path = folderPathWrite, nav = f'DEBT | ${data.MKT_VALUE_USD.sum():,.2f} | 100%')