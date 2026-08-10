import pandas as pd

path = 'C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Contributtion'
names = ['Bank', 'Acct_Num', 'Description', 'Identifier', 'Trade_Date', 'Settlement_Date', 'Local_Amount', 'Local_CCY', 'Type', 'Reporting_Amount']

### ======================================================================== ###
###                                    SSZ                                   ###
### ======================================================================== ###

def interests2(data):
    if 'TRANSFER' in data.DETAIL and '15' in data['VALUE DATE']: #data['CREDITS/DEBITS'] < 0:
        data['LOAN'] = data.DETAIL[-3:]
    elif 'INTERESES LOAN' in data.DETAIL:
        data['LOAN'] = data.DETAIL[-3:]
    elif ('DEBIT INTEREST' in data.DETAIL) and ('31' in data['BOOKING DATE'] or '30' in data['BOOKING DATE']):
        data['LOAN'] = data.DETAIL[-3:]
    else:
        data['LOAN'] = 'None'
    
    return data

def interests(data):
    if ('TRANSFER' in data.DETAIL) and (data['VALUE DATE'].day == 15): #data['CREDITS/DEBITS'] < 0:
        data['LOAN'] = f'Loan_{data.DETAIL[-3:]}'
    elif 'INTERESES LOAN' in data.DETAIL:
        data['LOAN'] = f'Loan_{data.DETAIL[-3:]}'
    elif ('DEBIT INTEREST' in data.DETAIL) and (data['VALUE DATE'].day == 1):
        data['LOAN'] = f'Loan_{data.DETAIL[-3:]}'
    elif 'CORPORATE ACTIVITY /' in data.DETAIL:
        data['LOAN'] = 'Coupons_SSZ'
    elif data.DETAIL.startswith('COUPONS /'):
        data['LOAN'] = 'Dividends_SSZ'
    elif 'WITHHOLDING TAX' in data.DETAIL:
        data['LOAN'] = 'Tax_SSZ'
    elif 'FEE' in data.DETAIL:
        data['LOAN'] = 'Fee_SSZ'
    elif 'WITHHOLDING TAX' in data.DETAIL:
        data['LOAN'] = 'Tax_SSZ'
    else:
        data['LOAN'] = 'None'
        
    if ('TRANSFER' in data.DETAIL) and (data['CREDITS/DEBITS'] > 0):
        data['LOAN'] = 'None'
    else:
        data['LOAN'] = data['LOAN']
    
    if ('CORRECTION OF DEBIT' in data.DETAIL):
        data['LOAN'] = None
    
    return data

# ssz = pd.read_excel(f'{path}/SS/Transactions/01Jan2023_02Jul2025_SSZ1_69037.xlsx', header = 10, engine = 'openpyxl').drop(['BALANCE', 'CURRENCY.1'], axis = 1)
ssz = pd.read_excel(f'{path}/SS/Transactions/TRNS_Jan2024-26Jun2026.xlsx', header = 10, engine = 'openpyxl').drop(['BALANCE', 'CURRENCY.1'], axis = 1)
# ssz = pd.read_excel(f'{path}/SS/Transactions/TRNS_2025.xlsx', header = 10, engine = 'openpyxl').drop(['BALANCE', 'CURRENCY.1'], axis = 1)
fx = pd.read_excel(f'{path}/SS/Transactions/TipoCambio.xlsx', engine = 'openpyxl')
mxnusd = fx[['Date', 'MXNUSD']]
eurusd = fx[['Date', 'EURUSD']]

def frate(data):
    if data.CURRENCY == 'MXN':
        fxr = mxnusd[mxnusd.Date == data['BOOKING DATE']]['MXNUSD'].values[0]
        data['VALUE'] = fxr * data['CREDITS/DEBITS']
    elif data.CURRENCY == 'EUR':
        fxr = eurusd[eurusd.Date == data['BOOKING DATE']]['EURUSD'].values[0]
        data['VALUE'] = fxr * data['CREDITS/DEBITS']
    else:
        data['VALUE'] = data['CREDITS/DEBITS']
        
    return data


ssz[['VALUE DATE', 'BOOKING DATE']] = ssz[['VALUE DATE', 'BOOKING DATE']].apply(pd.to_datetime, format = '%d %b %Y')

ssz = ssz.apply(interests, axis = 1)
ssz = ssz.apply(frate, axis = 1)

ssz = ssz[ssz.LOAN != 'None']
ssz['PORTFOLIO'] = ssz.PORTFOLIO.str.split('/').str[1]
ssz.insert(0, 'Bank', 'SSZ')
ssz.columns = names

# ssz.to_excel(f'{path}/SS/Transactions/Transactions_SSZ3.xlsx', index = False)

### ======================================================================== ###
###                                    JPM                                   ###
### ======================================================================== ###

types = ['Interest', 'Misc. Disbursement', 'Taxes', 'Corporate Interest', 'Fees', 'Foreign Dividend', 'Swap Cashflow', 'Cash Distribution']
cols = ['Account Name', 'Description', 'Cusip', 'Trade Date', 'Settlement Date', 'Amount Local', 'Local Currency', 'Type', 'Amount USD']

dtyp = {
    'Interest': 'Interest_JPM',
    'Misc. Disbursement': 'Loan_JPM',
    'Taxes': 'Taxes_JPM',
    'Fees': 'Fees_JPM',
    'Foreign Dividend': 'Dividends_JPM',
    'Corporate Interest': 'Interest_JPM',
    'Swap Cashflow': 'Dividends_JPM',
    'Cash Distribution': 'Dividends_JPM'
}

acc0004 = pd.read_csv(f'{path}/JPM/Transactions/01Jan2024_26Jun2026_JPM_0004.csv')
acc2001 = pd.read_csv(f'{path}/JPM/Transactions/01Jan2024_26Jun2026_JPM_2001.csv')
acc4005 = pd.read_csv(f'{path}/JPM/Transactions/01Jan2024_26Jun2026_JPM_4005.csv')
acc7005 = pd.read_csv(f'{path}/JPM/Transactions/01Jan2024_26Jun2026_JPM_7005.csv')
acc7008 = pd.read_csv(f'{path}/JPM/Transactions/01Jan2024_26Jun2026_JPM_7008.csv')
# acc0004 = pd.read_csv(f'{path}/JPM/Transactions/TRNS_JPM_0004_2025.csv')
# acc2001 = pd.read_csv(f'{path}/JPM/Transactions/TRNS_JPM_2001_2025.csv')
# acc4005 = pd.read_csv(f'{path}/JPM/Transactions/TRNS_JPM_4005_2025.csv')
# acc7005 = pd.read_csv(f'{path}/JPM/Transactions/TRNS_JPM_7005_2025.csv')
# acc7008 = pd.read_csv(f'{path}/JPM/Transactions/TRNS_JPM_7008_2025.csv')

# loan6512 = pd.read_csv(f'{path}/JPM/Transactions/Loan12months_6512.csv')
# loan6512 = pd.read_csv(f'{path}/JPM/Transactions/Loan25Jul2024-27Aug2025_6512.csv')
loan6512 = pd.read_csv(f'{path}/JPM/Transactions/Loan12months_6512_upto26Jun2026.csv')
loan7536 = pd.read_csv(f'{path}/JPM/Transactions/Loan12months_7536_upto26Jun2026.csv')

loan7536 = loan7536[loan7536.Description != 'DRAW'].reset_index(drop = True)
loan6512 = loan6512[loan6512.Description != 'DRAW'].reset_index(drop = True)

acc0004 = acc0004[acc0004.Type.isin(types)][cols].reset_index(drop = True)
acc2001 = acc2001[acc2001.Type.isin(types)][cols].reset_index(drop = True)
acc4005 = acc4005[acc4005.Type.isin(types)][cols].reset_index(drop = True)
acc7005 = acc7005[acc7005.Type.isin(types)][cols].reset_index(drop = True)
acc7008 = acc7008[acc7008.Type.isin(types)][cols].reset_index(drop = True)

acc0004['Type'] = acc0004.Type.map(dtyp)
acc2001['Type'] = acc2001.Type.map(dtyp)
acc4005['Type'] = acc4005.Type.map(dtyp)
acc7005['Type'] = acc7005.Type.map(dtyp)
acc7008['Type'] = acc7008.Type.map(dtyp)

loan7536['Account Name'] = '7536'
loan7536['Description'] = loan7536.Description
loan7536['Cusip'] = ''
loan7536['Trade Date'] = loan7536.Date
loan7536['Settlement Date'] = loan7536.Date
loan7536['Amount Local'] = loan7536.Amount
loan7536['Local Currency'] = 'USD'
loan7536['Type'] = 'Loan_JPM'
loan7536['Amount USD'] = loan7536.Amount

loan6512['Account Name'] = '6512'
loan6512['Description'] = loan6512.Description
loan6512['Cusip'] = ''
loan6512['Trade Date'] = loan6512.Date
loan6512['Settlement Date'] = loan6512.Date
loan6512['Amount Local'] = loan6512.Amount
loan6512['Local Currency'] = 'USD'
loan6512['Type'] = 'Loan_JPM'
loan6512['Amount USD'] = loan6512.Amount

loan6512 = loan6512[cols]
loan7536 = loan7536[cols]

jpm = pd.concat([acc0004, acc2001, acc4005, acc7005, acc7008, loan6512, loan7536], axis = 0)
jpm[['Trade Date', 'Settlement Date']] = jpm[['Trade Date', 'Settlement Date']].apply(pd.to_datetime, format = '%m/%d/%Y', errors = 'coerce')
jpm.insert(0, 'Bank', 'JPM')
jpm.columns = names

# print(jpm.head())
# jpm.to_excel(f'{path}/JPM/Transactions/Transactions_JPM1.xlsx', index = False)


### ======================================================================== ###
###                                    DBK                                   ###
### ======================================================================== ###

# dbk = pd.read_excel(f'{path}/DB/Transactions/01Jan2024_22Ago2025_ALL.xlsx', engine = 'openpyxl')
dbk = pd.read_excel(f'{path}/DB/Transactions/TRNS_Jan2024-26Jun2026.xlsx', engine = 'openpyxl')
loan = pd.read_excel(f'{path}/DB/Transactions/LOANS_Jan2024-26Jun2026.xlsx', engine = 'openpyxl')

loan['Bank'] = 'DBK'
loan['Acct_Num'] = loan['Account Number']
loan['Description'] = loan['To/From']
loan['Identifier'] = 'DBKLoan'
loan['Trade_Date'] = loan['Transaction Date']
loan['Settlement_Date'] = loan['Transaction Date']
loan['Local_Amount'] = loan['Funds Subtracted']
loan['Local_CCY'] = loan['Base Currency']
# loan['Type'] = 'Loan_DBK'
loan['Reporting_Amount'] = loan['Funds Subtracted']
loan[['Trade_Date', 'Settlement_Date']] = loan[['Trade_Date', 'Settlement_Date']].apply(pd.to_datetime, format = '%m/%d/%Y')
loan = loan[(loan.Trade_Date.dt.year >= 2024)].reset_index(drop = True)

def loans(data):
    # if 'INTEREST PAYME' in data.Description:
    if 'Loan Payment' in data.Description:
        data['Type'] = 'Loan_DBK'
    else:
        data['Type'] = 'Other'
    
    return data
loan = loan.apply(loans, axis = 1)
loan = loan[loan.Type != 'Other'].reset_index(drop = True)

cols = ['Account Number', 'Transaction Description', 'Security Identifier', 'Trade Date', 'Settlement Date', 'Net Transaction Amount – Local Currency',
        'Local Currency', 'Transaction Type ', 'Net Transaction Amount – Reporting  Currency']
types = ['Interest', 'Dividend', 'Taxes Withheld']

dtyp = {
    'Interest': 'Interest_DB',
    'Dividend': 'Dividends_DB',
    'Taxes Withheld': 'Taxes_DB',
}

dbk = dbk[dbk['Transaction Type '].isin(types)][cols].reset_index(drop = True)
dbk['Transaction Type '] = dbk['Transaction Type '].map(dtyp)

dbk[['Trade Date', 'Settlement Date']] = dbk[['Trade Date', 'Settlement Date']].apply(pd.to_datetime, format = '%m/%d/%Y', errors = 'coerce')
dbk.insert(0, 'Bank', 'DBK')
dbk.columns = names

dbk = pd.concat([dbk, loan[dbk.columns]], axis = 0)
# print(dbk.head())
# dbk.to_excel(f'{path}/DB/Transactions/Transactions_DB.xlsx', index = False)

# ### ======================================================================== ###
# ###                                    GSS                                   ###
# ### ======================================================================== ###

# # gss = pd.read_excel(f'{path}/GS/Transactions/01Jan2024_22Ago2025_GS_657.xlsx', header = 7, engine = 'openpyxl')
# gss = pd.read_excel(f'{path}/GS/Transactions/TRNS_2025.xlsx', header = 7, engine = 'openpyxl').dropna()
# gss['Activity Date'] = gss['Activity Date'].apply(pd.to_datetime, format = '%m/%d/%Y')

# gss['Acct_Num'] = 'XXXX657'
# gss['Identifier'] = '-----'
# gss['Trade_Date'] = gss['Activity Date']
# gss['Settlement_Date'] = gss['Activity Date']
# gss['Local_Amount'] = gss['Settlement Amount']
# gss['Local_CCY'] = gss['Display Currency']
# gss['Type'] = gss['Activity Type']
# gss['Reporting_Amount'] = gss['Settlement Amount']
# gss['Bank'] = 'GSS'

# types = ['Withdrawal', 'Deposit']
# gss = gss[gss.Type.isin(types)][names].reset_index(drop = True)

# dtyp = {
#     'Withdrawal': 'Taxes_GS',
#     'Deposit': 'Interest_GS'
# }

# gss['Type'] = gss.Type.map(dtyp)

# # print(gss.head())
# # gss.to_excel(f'{path}/GS/Transactions/Transactions_GS.xlsx', index = False)

### ======================================================================== ###
###                                    MSY                                   ###
### ======================================================================== ###

msy = pd.read_excel(f'{path}/MS/Transactions/TRNS_Jan2024-26Jun2026.xlsx', header = 6, engine = 'openpyxl')
msy[['Activity Date', 'Transaction Date']] = msy[['Activity Date', 'Transaction Date']].apply(pd.to_datetime, format = '%m/%d/%Y')

msy['Acct_Num'] = msy.Account
msy['Identifier'] = msy.Cusip
msy['Trade_Date'] = msy['Activity Date']
msy['Settlement_Date'] = msy['Transaction Date']
msy['Local_Amount'] = msy['Amount($)']
msy['Local_CCY'] = 'USD'
msy['Type'] = msy.Activity
msy['Reporting_Amount'] = msy['Amount($)']
msy['Bank'] = 'MSY'

types = ['Dividend', 'Dividend - Adjustment', 'Dividend Reinvestment', 'Dividend Stock', 'Interest Income', 'Miscellaneous Income',
         'Other Debits', 'Qualified Dividend', 'Service Fee', 'Service Fee Adj', 'Tax Withholding', 'Tax Withholding Adj',
         'Tax Withholding-Adj', 'Cashless Dividend', 'Transfer In']
msy = msy[msy.Type.isin(types)][names].reset_index(drop = True)

dtyp = {
    'Dividend': 'Dividends_MS',
    'Dividend - Adjustment': 'Dividends_MS',
    'Dividend Reinvestment': 'Dividends_MS',
    'Dividend Stock': 'Dividends_MS',
    'Qualified Dividend': 'Dividends_MS',
    'Interest Income': 'Interest_MS',
    'Miscellaneous Income': 'Interest_MS',
    'Other Debits': 'Taxes_MS',
    'Service Fee': 'Taxes_MS',
    'Service Fee Adj': 'Taxes_MS',
    'Tax Withholding': 'Taxes_MS',
    'Tax Withholding Adj': 'Taxes_MS',
    'Tax Withholding-Adj': 'Taxes_MS',
    'Cashless Dividend': 'Dividends_MS',
    'Transfer In': 'Dividends_MS'
}

msy['Type'] = msy.Type.map(dtyp)

# print(msy.head())
# msy.to_excel(f'{path}/MS/Transactions/Transactions_MS.xlsx', index = False)

### ======================================================================== ###
###                                Consolidado                               ###
### ======================================================================== ###

# data = pd.concat([ssz, jpm, dbk, gss, msy], axis = 0)
data = pd.concat([ssz, jpm, dbk, msy], axis = 0)
data['Year'] = data.Trade_Date.dt.year

grouped = data.groupby(['Year', 'Bank', 'Type']).Reporting_Amount.sum().reset_index()
print(grouped)

grouped.to_excel(f'{path}/TRNS_CLEAN_2026.xlsx', index = False)