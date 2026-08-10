import pandas as pd
import re

date = pd.Timestamp.today().strftime('%Y%m%d')
# date = '20250807'

path = 'C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt'
data = pd.read_excel(f'{path}/Cotizaciones/Raw_{date}.xlsx', engine = 'openpyxl')

def quotes(path:str, quotes:pd.DataFrame, ccy:str):
    notas = pd.read_excel(f'{path}/Notas Estructuradas/Notas_Template.xlsx', sheet_name = f'CUPONES {ccy}', engine = 'openpyxl').T.reset_index()
    cols = [f'Date_{i+1}' for i in range(len(notas.columns))]
    notas.columns = cols
    
    notas['Date_1'] = notas.Date_1.apply(lambda x: re.sub(r'\.\d+$', '', x))
    notas['Date_1'] = notas.Date_1.str.replace('B-', '')
    
    quotes = quotes[quotes.Currency == ccy].reset_index(drop = True)
    
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
    
    return data

usd = quotes(path = path, quotes = data, ccy = 'USD')
eur = quotes(path = path, quotes = data, ccy = 'EUR')

with pd.ExcelWriter(f'{path}/Cotizaciones/Clean/Quotes_{date}.xlsx', engine = 'openpyxl') as writer:
    usd.to_excel(writer, sheet_name = 'USD', index = False)
    eur.to_excel(writer, sheet_name = 'EUR', index = False)