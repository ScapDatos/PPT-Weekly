import pandas as pd
from config import DATE, PARENT_PATH
from debugg.functionsV1 import get_df, stacked_areas, portfolioswoHCITY, waterfall

date12m = (pd.to_datetime(DATE, format = '%Y-%m-%d') - pd.DateOffset(years = 1)).strftime('%Y-%m-%d')
path_pc = PARENT_PATH / 'DATOS'
path_to_db = path_pc / 'Bases de datos'
# path_to_db = f'{path_pc}/Bases de datos'

colors = {
    'CASH':	           '#97B7B3',
    'LIQUIDITY':       '#637976',
    'EQUITIES':        '#475161',
    'FIXED_INCOME':    '#8F8F8F',
    'ALTERNATIVES':    '#867B74',
    'STRUCTURED_NOTES':'#C2B994',
    'DERIVATIVES':     '#89687E',
    'LOAN':            '#FF7C80',
    'LOAN_USD':        '#FFD1B2',
    'LOAN 2':          '#FFD1B2',
    'LOAN_EUR':        '#FF7C80',
    'BANKINTER':       '#FF9BD7',
    'LOAN_MXN':        '#EC9092',
    'LOAN 3':          '#EC9092',
    'LOAN 4':          '#D9B2B2',
    'UNI':             '#5C7D96',
    'LOAN_UNSEC':      '#FFD1B2',
    'Morgan Stanley':  '#637976',
    'Goldman Sachs':   '#475161',
    'Deutsche Bank':   '#8F8F8F',
    'Santander Suiza': '#867B74',
    'TD Ameritrade':   '#C2B994',
    'JP Morgan':       '#89687E',
    'Julius Baer':     '#475161',
    # 'UBS':             'rgba( 99,121,118,  1 )', #'#637976'
    'UBS':             '#89A184',
    'Citibank':        '#8F8F8F',
    ' Santander Suiza':'#89687E',
    'Bankinter':       '#5C7D96',
    'NAV':             '#0000FF'
}

### ============================================================================= ###
###                              Graficas SCH COMPLETO                            ###
### ============================================================================= ###
cols = ['Goldman Sachs', 'UBS','Morgan Stanley','Deutsche Bank','Santander Suiza','JP Morgan','Bankinter','NAV']
sch = get_df(path_to_db, 'SigCap.db', 'PORT_SCH_NAV')
sch.columns = ['DATE','Deutsche Bank','Goldman Sachs','Santander Suiza','JP Morgan','Morgan Stanley', 'UBS','Bankinter','NAV','LABEL']
stacked_areas(path_pc, 'Graficas_PPT/SCHOLDINGS', sch, cols, 10000000, 3, colors = colors, labels = True, bnk = True, bbox = (0.001, 1.06))
stacked_areas(path_pc, 'Graficas_PPT/SCHOLDINGS12M', sch, cols, 10000000, 1, colors = colors, startDate = date12m, bnk = True, bbox = (0.001, 1.06))

### ============================================================================= ###
###                          Graficas 2019-2024 y 12M SCH                         ###
### ============================================================================= ###
ncols = ['DATE', 'CASH', 'LIQUIDITY', 'EQUITIES', 'FIXED_INCOME', 'ALTERNATIVES', 'STRUCTURED_NOTES', 'DERIVATIVES', 'LOAN', 'AUMS', 'NAV', 'DEPS', 'WDRS', 'LABEL', 'VLINE']
cols = ['CASH','LIQUIDITY','EQUITIES','FIXED_INCOME','ALTERNATIVES','STRUCTURED_NOTES','DERIVATIVES']

# GOLDMAN SACHS
gss = get_df(path_to_db, 'SigCap_SCH.db', 'GSS')
gss.columns = ncols
stacked_areas(path_pc, 'Graficas_PPT/GS', gss, cols + ['NAV'], 2000000, 1, colors = colors, startDate = '2020-01-01', labels = True, vline = True, bbox = (0.001, 1.06), left = 0.04)
stacked_areas(path_pc, 'Graficas_PPT/GS12M', gss, cols + ['NAV'], 2000000, 1, colors = colors, startDate = date12m, vline = True, bbox = (0.001, 1.07), left = 0.04)

# UBS
ubs = get_df(path_to_db, 'SigCap_SCH.db', 'UBS')
ubs.columns = ncols
stacked_areas(path_pc, 'Graficas_PPT/UBSSCH', ubs, cols + ['NAV'], 2000000, 4, colors = colors, startDate = '2020-01-01', labels = False, vline = True, bbox = (0.001, 1.06), left = 0.033)
stacked_areas(path_pc, 'Graficas_PPT/UBSSCH12M', ubs, cols + ['NAV'], 2000000, 1, colors = colors, startDate = date12m, vline = True, bbox = (0.001, 1.07), left = 0.033)

# MORGAN STANLEY
msy = get_df(path_to_db, 'SigCap_SCH.db', 'MSY')
msy.columns = ncols
stacked_areas(path_pc, 'Graficas_PPT/MS', msy, cols + ['LOAN', 'NAV'], 2000000, 1, colors = colors, startDate = '2020-01-01', labels = True)
stacked_areas(path_pc, 'Graficas_PPT/MS12M', msy, cols + ['LOAN', 'NAV'], 2000000, 1, colors = colors, startDate = date12m, left = 0.046)

# DEUTSCHE BANK
dbk = get_df(path_to_db, 'SigCap_SCH.db', 'DBK')
dbk.columns = ncols
stacked_areas(path_pc, 'Graficas_PPT/DB', dbk, cols + ['LOAN', 'NAV'], 2000000, 1, colors = colors, startDate = '2020-01-01', labels = True)
stacked_areas(path_pc, 'Graficas_PPT/DB12M', dbk, cols + ['LOAN', 'NAV'], 2000000, 1, colors = colors, startDate = date12m, left = 0.033)

# SANTANDER SUIZA
ssz = get_df(path_to_db, 'SigCap_SCH.db', 'SSZ')
ncolss = ['DATE','CASH','LIQUIDITY','EQUITIES','FIXED_INCOME','ALTERNATIVES','STRUCTURED_NOTES','DERIVATIVES','LOAN_USD','LOAN_EUR','LOAN_MXN','AUMS','NAV','DEPS','WDRS','LABEL','VLINE']
ssz.columns = ncolss

uni = get_df(path_to_db, 'SigCap.db', 'UNICAJA')
ssz = pd.merge(ssz, uni, how = 'left', on = 'DATE')
ssz['EQUITIES'] = ssz.EQUITIES - ssz.UNI

stacked_areas(path_pc, 'Graficas_PPT/SS', ssz, cols + ['LOAN_USD', 'LOAN_EUR', 'LOAN_MXN', 'NAV', 'UNI'], 4000000, 3, colors = colors, 
              startDate = '2019-12-31', labels = True)
stacked_areas(path_pc, 'Graficas_PPT/SS12M', ssz, cols + ['LOAN_USD', 'LOAN_EUR', 'LOAN_MXN', 'NAV', 'UNI'], 4000000, 3, colors = colors,
              startDate = '2014-10-31', bbox = (0.001, 1.11), ncols = 6)

# JP MORGAN
ncolsj = ['DATE', 'CASH', 'LIQUIDITY', 'EQUITIES', 'FIXED_INCOME', 'ALTERNATIVES', 'STRUCTURED_NOTES', 'DERIVATIVES', 'LOAN', 'LOAN_UNSEC', 'AUMS', 'NAV', 'DEPS', 'WDRS', 'LABEL', 'VLINE']
jpm = get_df(path_to_db, 'SigCap_SCH.db', 'JPM')
jpm.columns = ncolsj

stacked_areas(path_pc, 'Graficas_PPT/JPM', jpm, cols + ['LOAN', 'LOAN_UNSEC', 'NAV'], 3000000, 1, colors = colors, startDate = '2020-01-01', labels = True, bbox = (0.001, 1.11), ncols = 6)
stacked_areas(path_pc, 'Graficas_PPT/JPM12M', jpm, cols + ['LOAN', 'LOAN_UNSEC', 'NAV'], 3000000, 1, colors = colors, startDate = date12m, ncols = 7)

### ============================================================================= ###
###                              Graficas NFD COMPLETO                            ###
### ============================================================================= ###
nfd = get_df(path_to_db, 'SigCap.db', 'PORT_NFD_NAV')
nfd.columns = ['DATE', 'Citibank','Julius Baer', ' Santander Suiza', 'UBS', 'NAV', 'LABEL']
cols = ['Julius Baer', 'Citibank', 'UBS', ' Santander Suiza', 'NAV']
stacked_areas(path_pc, 'Graficas_PPT/NFD', nfd, cols, 1000000, 2, colors = colors, startDate = '2019-06-30', bnk = True, bbox = (0.001, 1.05), left = 0.046)
stacked_areas(path_pc, 'Graficas_PPT/NFD12M', nfd, cols, 1000000, 1, colors = colors, startDate = date12m, bnk = True, bbox = (0.001, 1.05), left = 0.046)

### ============================================================================= ###
###                          Graficas 2019-2024 y 12M NFD                         ###
### ============================================================================= ###
ncols = ['DATE','CASH','LIQUIDITY','EQUITIES','FIXED_INCOME','ALTERNATIVES','STRUCTURED_NOTES','DERIVATIVES','LOAN','AUMS','NAV','DEPS','WDRS','LABEL','VLINE']
cols = ['CASH','LIQUIDITY','EQUITIES','FIXED_INCOME','ALTERNATIVES','STRUCTURED_NOTES','DERIVATIVES','LOAN','NAV']

# CITIBANK
ct = get_df(path_to_db, 'SigCap_NFD.db', 'CTI')
ct.columns = ncols

stacked_areas(path_pc, 'Graficas_PPT/CTI', ct, cols, 2000000, 1, colors = colors, startDate = '2022-10-31', labels = True, left = 0.046)
stacked_areas(path_pc, 'Graficas_PPT/CTI12M', ct, cols, 2000000, 1, colors = colors, startDate = date12m, left = 0.046)

# JULIUS BAER
jbr = get_df(path_to_db, 'SigCap_NFD.db', 'JBR')
jbr.columns = ncols

stacked_areas(path_pc, 'Graficas_PPT/JB', jbr, cols, 2000000, 1, colors = colors, startDate = '2022-06-01', labels = True, left = 0.046)
stacked_areas(path_pc, 'Graficas_PPT/JB12M', jbr, cols, 2000000, 1, colors = colors, startDate = date12m, left = 0.046)

# SANTANDER SUIZA
ssz = get_df(path_to_db, 'SigCap_NFD.db', 'SSZ')
ssz.columns = ncols

stacked_areas(path_pc, 'Graficas_PPT/SSNFD', ssz, cols, 2000000, 1, colors = colors, startDate = '2019-12-31', labels = False, left = 0.046)
stacked_areas(path_pc, 'Graficas_PPT/SSNFD12M', ssz, cols, 1000000, 1, colors = colors, startDate = date12m, left = 0.046)

# UBS
ubs = get_df(path_to_db, 'SigCap_NFD.db', 'UBS')
ubs.columns = ncols

stacked_areas(path_pc, 'Graficas_PPT/UBS', ubs, cols, 2000000, 1, colors = colors, startDate = '2019-05-01', labels = True, vline = True, left = 0.046)
stacked_areas(path_pc, 'Graficas_PPT/UBS12M', ubs, cols, 100000, 1, colors = colors, startDate = date12m, vline = True, left = 0.033)

### ============================================================================= ###
###                               Graficas WK COMPLETO                            ###
### ============================================================================= ###
wk = get_df(path_to_db, 'SigCap_WK.db', 'MSY')
wk.columns = ncols

stacked_areas(path_pc, 'Graficas_PPT/MSWK', wk, cols, 250000, 1, colors = colors, startDate = '2021-04-30', left = 0.046)
stacked_areas(path_pc, 'Graficas_PPT/MSWK12M', wk, cols, 250000, 1, colors = colors, startDate = date12m, left = 0.046)

wknav = get_df(path_to_db, 'SigCap.db', 'PORT_WK_NAV')
wknav.columns = ['DATE', 'Morgan Stanley', 'NAV', 'LABEL']

stacked_areas(path_pc, 'Graficas_PPT/WK', wknav, ['Morgan Stanley', 'NAV'], 250000, 1, colors = colors, startDate = '2021-04-30', bnk = True, bbox = (0.001, 1.05), left = 0.033)
stacked_areas(path_pc, 'Graficas_PPT/WK12M', wknav, ['Morgan Stanley', 'NAV'], 250000, 1, colors = colors, startDate = date12m, bnk = True, bbox = (0.001, 1.05), left = 0.033)


### ============================================================================= ###
###                              Graficas SCH Y HCITY                             ###
### ============================================================================= ###
hcity = get_df(path_to_db, 'SigCap.db', 'HCITY')
hcity_ = hcity[['DATE', 'HCITY_SCH', 'LOANS_SCH', 'LABEL']]

axy = {'N': (55, -15), 'NwoH': (55, 0), 'H': (55, -15)}
portfolioswoHCITY(path_pc, 'Graficas_PPT/SCHOLDINGS_W_HCITY', sch.drop('LABEL', axis = 1), hcity_, xTicks = 3, yTicks = 10000000)

### ============================================================================= ###
###                              Graficas NFD Y HCITY                             ###
### ============================================================================= ###
hcity_ = hcity[['DATE', 'HCITY_NFD', 'LOANS_NFD', 'LABEL']]

axy = {'N': (55, 0), 'NwoH': (55, 0), 'H': (55, 0)}
portfolioswoHCITY(path_pc, 'Graficas_PPT/NFD_W_HCITY', nfd.drop('LABEL', axis = 1), hcity_, xTicks = 2, yTicks = 1000000)


### ============================================================================= ###
###                          Grafica de cascada SCH Y NFD                         ###
### ============================================================================= ###
# cols_list = ['NAV', 'Goldman Sachs', 'UBS', 'Morgan Stanley', 'Deutsche Bank', 'Santander Suiza', 'JP Morgan', 'Bankinter']
cols_list = ['NAV', 'UBS', 'Morgan Stanley', 'Deutsche Bank', 'Santander Suiza', 'JP Morgan', 'Bankinter']
difs = sch[cols_list].iloc[-1] - sch[cols_list].iloc[-2]
wfSCH = [sch.NAV.iloc[-2], difs['UBS'], difs['Morgan Stanley'], difs['Deutsche Bank'], difs['Santander Suiza'],
         difs['JP Morgan'], difs['Bankinter'], sch['NAV'].iloc[-1]]
# xticks = ['LAST NAV', 'GS', 'UBS', 'MS', 'DB', 'SS', 'JPM', 'BNK', 'ACTUAL NAV']
xticks = ['LAST NAV', 'UBS', 'MS', 'DB', 'SS', 'JPM', 'BNK', 'ACTUAL NAV']

waterfall(path_pc, wfSCH, sch, xticks, name = 'Graficas_PPT/SCHOLDINGS_WATERFALL', minmax = (500000, 500000)) #(,500000)


cols_list = ['NAV', 'Citibank', 'Julius Baer', ' Santander Suiza', 'UBS']
difs = nfd[cols_list].iloc[-1] - nfd[cols_list].iloc[-2]
wfNFD = [nfd.NAV.iloc[-2], difs['Citibank'], difs['Julius Baer'], difs[' Santander Suiza'], difs['UBS'], nfd.NAV.iloc[-1]]
xticks = ['LAST NAV', 'CT', 'JB', 'SS', 'UBS', 'ACTUAL NAV']

waterfall(path_pc, wfNFD, nfd, xticks, name = 'Graficas_PPT/NFD_WATERFALL', minmax = (150000, 100000))