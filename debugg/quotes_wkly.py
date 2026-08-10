import re
# import numpy as np
import pandas as pd
from pathlib import Path

PATH = Path.home() / 'OneDrive' / '0. Nube Asset Mgmt'
drange = '20260804-10'

data = pd.read_excel(PATH / 'Cotizaciones' / f'Raw_{drange}.xlsx')
# data = pd.read_excel(f'{path}/Cotizaciones/Raw_{drange}.xlsx')
# data = data.dropna(ignore_index = True)
data['fecDate'] = data.fecDate.apply(pd.to_datetime, errors = 'coerce')

def quotes(path:str, quotes:pd.DataFrame, ccy:str):
    notas = pd.read_excel(f'{path}/Notas Estructuradas/Notas_Template.xlsx', sheet_name = f'CUPONES {ccy}', engine = 'openpyxl').T.reset_index()
    cols = [f'Date_{i+1}' for i in range(len(notas.columns))]
    notas.columns = cols
    
    notas['Date_1'] = notas.Date_1.apply(lambda x: re.sub(r'\.\d+$', '', x))
    notas['Date_1'] = notas.Date_1.str.replace('B-', '')
    
    single = quotes[quotes['Basket Type'] == 'Single'][['Barrier', 'Basket', 'Coupon PA']].reset_index(drop = True)
    single['Basket'] = single.Basket.str.split().str[0]
    single['Barrier'] = 100 - single.Barrier.str.replace(',00%', '').astype(int)
    single['Barrier'] = single['Barrier'].astype(str) + '%'
    
    wof = quotes[quotes['Basket Type'] == 'WoF'][['Barrier', 'Basket', 'Coupon PA']].reset_index(drop = True)
    wof['Basket'] = wof.Basket.str.replace('; ', '|')
    wof['Barrier'] = 100 - wof.Barrier.str.replace(',00%', '').astype(int)
    wof['Barrier'] = wof['Barrier'].astype(str) + '%'
    
    data = pd.DataFrame(columns = notas.loc[1:, 'Date_2'].astype(str) + '-' + notas.loc[1:, 'Date_1'].astype(str))
    
    for idxs, sing in single.iterrows():
        for idxn, nota in notas.iterrows():        
            if (sing['Basket'] == nota['Date_2']) and (sing['Barrier'] == nota['Date_1']):
                col = f"{sing['Basket']}-{sing['Barrier']}"
                data.loc[0, col] = sing['Coupon PA']
    
    data.loc[0, 'MXEF-25%'] = single.loc[(single.Basket == 'EEM') & (single.Barrier == '25%'), 'Coupon PA'].values[0]
    data.loc[0, 'MXEF-35%'] = single.loc[(single.Basket == 'EEM') & (single.Barrier == '35%'), 'Coupon PA'].values[0]
    data.loc[0, 'MXEF-30%'] = single.loc[(single.Basket == 'EEM') & (single.Barrier == '30%'), 'Coupon PA'].values[0]
    data.loc[0, 'MXEF-40%'] = single.loc[(single.Basket == 'EEM') & (single.Barrier == '40%'), 'Coupon PA'].values[0]
    data.loc[0, 'IBOV-25%'] = single.loc[(single.Basket == 'EWZ') & (single.Barrier == '25%'), 'Coupon PA'].values[0]
    data.loc[0, 'IBOV-35%'] = single.loc[(single.Basket == 'EWZ') & (single.Barrier == '35%'), 'Coupon PA'].values[0]
    data.loc[0, 'IBOV-30%'] = single.loc[(single.Basket == 'EWZ') & (single.Barrier == '30%'), 'Coupon PA'].values[0]
    data.loc[0, 'IBOV-40%'] = single.loc[(single.Basket == 'EWZ') & (single.Barrier == '40%'), 'Coupon PA'].values[0]
    if ccy == 'USD':
        data.loc[0, 'SPY-20%'] = single.loc[(single.Basket == 'SPX') & (single.Barrier == '20%'), 'Coupon PA'].values[0]
        data.loc[0, 'QQQ-20%'] = single.loc[(single.Basket == 'NDX') & (single.Barrier == '20%'), 'Coupon PA'].values[0]
        data.loc[0, 'IWM-20%'] = single.loc[(single.Basket == 'IWN') & (single.Barrier == '20%'), 'Coupon PA'].values[0]
    if ccy == 'EUR':
        data.loc[0, 'SX5E|SPX-20%'] = wof.loc[(wof.Basket == 'SPX|SX5E') & (wof.Barrier == '20%'), 'Coupon PA'].values[0]
        data.loc[0, 'SX5E|SPX|RTY-20%'] = wof.loc[(wof.Basket == 'SPX|SX5E|RTY') & (wof.Barrier == '20%'), 'Coupon PA'].values[0]

        data.loc[0, 'SX5E|SPX-30%'] = wof.loc[(wof.Basket == 'SPX|SX5E') & (wof.Barrier == '30%'), 'Coupon PA'].values[0]
        data.loc[0, 'SX5E|SPX|RTY-30%'] = wof.loc[(wof.Basket == 'SPX|SX5E|RTY') & (wof.Barrier == '30%'), 'Coupon PA'].values[0]

        data.loc[0, 'SX5E|SPX-25%'] = wof.loc[(wof.Basket == 'SPX|SX5E') & (wof.Barrier == '25%'), 'Coupon PA'].values[0]
        data.loc[0, 'SX5E|SPX|RTY-25%'] = wof.loc[(wof.Basket == 'SPX|SX5E|RTY') & (wof.Barrier == '25%'), 'Coupon PA'].values[0]

    for idxw, wofs in wof.iterrows():
        for idxn, nota in notas.iterrows():
            if (wofs['Basket'] == nota['Date_2']) and (wofs['Barrier'] == nota['Date_1']):
                col = f"{wofs['Basket']}-{wofs['Barrier']}"
                data.loc[0, col] = wofs['Coupon PA']
    
    data.insert(0, 'Date', quotes.fecDate.iloc[0])
    return data

groups = data.groupby(['fecDate', 'Currency'])

usd = []
eur = []
for (date, ccy), group in groups:
    data_filtered = quotes(PATH, group, ccy)
    
    if ccy == 'USD':
        usd.append(data_filtered)
    elif ccy == 'EUR':
        eur.append(data_filtered)
    else:
        print(f'Revisar la moneda: {ccy}')

usd = pd.concat(usd, axis = 0, ignore_index = True)
eur = pd.concat(eur, axis = 0, ignore_index = True)

with pd.ExcelWriter(PATH / 'Cotizaciones' / 'Clean' / f'Clean_{drange}.xlsx', engine = 'openpyxl') as writer:
    usd.to_excel(writer, sheet_name = 'USD', index = False)
    eur.to_excel(writer, sheet_name = 'EUR', index = False)