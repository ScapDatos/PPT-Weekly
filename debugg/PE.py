import re
import pandas as pd
import matplotlib.pyplot as plt
from scipy.stats.mstats import winsorize

start = '2005-01-01'
end = '2025-12-11'
sector = 'S&P 500 INFO TECH INDEX'

hoja = 'PE'
img_path = 'C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/DATOS/Graficas_PPT/Multiplos'
sheets = {'PE':'PE', 'PE ACWI':'PEint', 'PE SX5E':'PESX5E', 'PE SX5Efwd':'PESX5Efwd'}
benchs = {'PE':'S&P 500 INDEX', 'PE fwd':'S&P 500 INDEX', 'PE ACWI':'MSCI ACWI Index', 'PE SX5E':'EURO STOXX 50 Price EUR', 'PE SX5Efwd':'EURO STOXX 50 Price EUR'}

if hoja in sheets.keys():
    titulo = ' P/E Trailing 12M'
elif hoja == 'PE fwd':
    titulo = ' P/E Forward'
else:
    titulo = 'Revisar'
    
path_pptx = f'C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Equity/{sheets[hoja]}.pptx'

pe_ratio = pd.read_excel('C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Equity/Base valuación relativa.xlsx', sheet_name = hoja)
pe_ratio['Fecha'] = pd.to_datetime(pe_ratio.Fecha)
pe_ratio.index = pe_ratio.Fecha
pe_ratio = pe_ratio.drop('Fecha', axis = 1)

DICC = pd.read_excel('C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Equity/Base valuación relativa.xlsx', sheet_name = 'DICC')
DICC_ = DICC.set_index('Ticker')['Name'].to_dict()

pe_ratio.rename(columns = DICC_, inplace = True)

WIND = DICC.loc[DICC['Wndorización'] == 'Si', 'Name'].tolist()

def windor(df:pd.DataFrame, cols:list[str], limits:tuple = (0.05, 0.05)):
    df_cpy = df.copy()
    for col in cols:
        if col in df_cpy.columns:
            df_cpy[col] = winsorize(df_cpy[col], limits = limits)
    
    return df_cpy

def mult_plots(df:pd.DataFrame, data:pd.DataFrame, col:str, idx:int, img_name:str, ls_sig2:str = '-.', perc:bool = True, add:float = 0.0,
               relative:bool = False):
    fig, ax = plt.subplots(figsize = (12, 6))
    y = df[df[col] != 0.0][col]
    x = y.index
    # ax.plot(df.index, df[col], ls = '-', label = col, c = '#0E8388')
    ax.plot(x, y, ls = '-', label = col, c = '#0E8388')
    
    mean = data[col].mean()
    std = data[col].std()
    
    ax.axhline(y = mean, ls = '--', label = f'$\mu$: {mean:.2f}', c = 'k')
    ax.axhline(y = mean + std, ls = '-', label = rf'+1$\sigma$: {mean + std:.2f}', c = 'k')
    ax.axhline(y = mean - std, ls = '-', label = f'-1$\sigma$: {mean - std:.2f}', c = 'k')
    ax.axhline(y = mean + 2*std, ls = ls_sig2, label = f'+2$\sigma$: {mean + 2*std:.2f}', c = 'k')
    ax.axhline(y = mean - 2*std, ls = ls_sig2, label = f'-2$\sigma$: {mean - 2*std:.2f}', c = 'k')
    
    last_val = df[col].iloc[-1]
    last_date = df.index[-1]
    
    if perc:
        ax.axhline(y = last_val * 0.8, ls = ':', c = 'b', label = f'80% último valor: {last_val * 0.8:.2f}')
        ax.axhline(y = last_val * 0.75, ls = '-.', c = 'b', label = f'75% último valor: {last_val * 0.75:.2f}')
        ax.axhline(y = last_val * 0.7, ls = '--', c = 'b', label = f'70% último valor: {last_val * 0.7:.2f}')
        ax.axhline(y = last_val * 0.6, ls = '-', c = 'b', label = f'60% último valor: {last_val * 0.6:.2f}')
    
    plt.annotate(f'{last_val:.2f}', xy = (last_date, last_val), xytext = (last_date, last_val + add),
                 arrowprops = dict(facecolor = 'b', arrowstyle = '->'), fontsize = 12, color = '#0E8388',
                 fontname = 'Arial', fontweight = 'bold')
    ax.set_xlabel('Date', fontdict = dict(family = 'Arial', size = 12, weight = 'bold'))
    ylabeln = 'Relative Price to Earnings' if relative else 'Price to Earnings'
    ax.set_ylabel(ylabeln, fontdict = dict(family = 'Arial', size = 12, weight = 'bold'))
    
    ax.spines['top'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    title = f'{col} Relative {titulo}' if relative else f'{col} {titulo}'
    ax.set_title(title, fontdict = dict(family = 'Arial', size = 16, weight = 'bold'))
    ax.legend(prop = dict(family = 'Arial', size = 12, weight = 'bold'))
    
    path = f'{img_path}/{img_name}_{idx}.png'
    plt.savefig(path)
    ax.grid()
    plt.close(fig)

pe_ratio = windor(pe_ratio, WIND, (0.05, 0.1))
pe = pe_ratio.loc[start:end]

for idx, col in enumerate(pe.columns):
    mult_plots(pe, pe_ratio, col, idx, 'PE', add = 2.0)

bench = benchs[hoja]
VRPE = pe_ratio.copy()
for col in VRPE.columns:
    VRPE[col] = VRPE[col] / VRPE[bench]

VRPE = VRPE.drop(bench, axis = 1)
VRPE_ = VRPE.loc[start:end]
for idx, col in enumerate(VRPE_.columns):
    mult_plots(VRPE_, VRPE, col, idx, 'RPE', '-.', False, 0.04, True)




def extract_num(filename):
    num = re.findall(r'\d+', filename)
    return int(num[0]) if num else -1

