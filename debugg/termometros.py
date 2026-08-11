import numpy as np
import pandas as pd
from pathlib import Path
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from scipy.stats.mstats import winsorize

def windor(df:pd.DataFrame, cols:list[str], limits:tuple = (0.05, 0.05)):
    df_cpy = df.copy()
    for col in cols:
        if col in df_cpy.columns:
            df_cpy[col] = winsorize(df_cpy[col], limits = limits)
    
    return df_cpy


def term_plots(data:pd.DataFrame, file_name:str, path):
    data['Date'] = pd.to_datetime(data.Date)
    data[data.columns[1:]] = data[data.columns[1:]].astype(float, errors = 'ignore') # REVISAR LOS ERRORES CUANDO SON #NA #NA

    indexes = ['SPX', 'NDX', 'SX5E', 'RTY']

    for index in indexes:
        plt.rcParams['axes.xmargin'] = 0
        fig, ax = plt.subplots(figsize = (16.3/2.54, 16.3/2.54))
        # fig, ax = plt.subplots(figsize = (15.5/2.54, 12.5/2.54))
        y = data[f'{index} Index']
        y = y.replace('#N/A Requesting Data...', np.nan)

        last_val = y.iloc[-1]
        y = winsorize(y.astype(float), limits = (0.05, 0.1)) if index == 'RTY' else y
        x = data.Date

        ax.plot(x, y, 'k-')

        mean = y.mean()
        std = y.std()

        ax.axhline(y = mean, ls = '--', c = '#333E50')
        ax.axhline(y = mean + std, ls = '-.', c = '#0E8388')
        ax.axhline(y = mean - std, ls = '-.', c = '#0E8388')
        ax.axhline(y = mean + 2*std, ls = '-.', c = '#6DAAAA')
        ax.axhline(y = mean - 2*std, ls = '-.', c = '#6DAAAA')

        last_date = x.iloc[-1]

        ax.axhline(y = last_val * 0.8, ls = '-', c = '#FF0000')
        ax.axhline(y = last_val * 0.75, ls = '-', c = '#DC143C')
        ax.axhline(y = last_val * 0.7, ls = '-', c = '#990000')
        ax.axhline(y = last_val * 0.6, ls = '-', c = '#800000')

        leg = {r'$\mu$':mean, r'-1$\sigma$':mean-std, r'+1$\sigma$':mean+std, r'-2$\sigma$':mean-2*std, r'+2$\sigma$':mean+2*std,
            '20%':last_val*0.8, '25%':last_val*0.75, '30%':last_val*0.7, '40%':last_val*0.6, index:last_val}
        leg = dict(sorted(leg.items(), key = lambda item: item[1], reverse = True))
        colors = {r'$\mu$':'#333E50', r'-1$\sigma$':'#0E8388', r'+1$\sigma$':'#0E8388', r'-2$\sigma$':'#6DAAAA', r'+2$\sigma$':'#6DAAAA',
                '20%':'#FF0000', '25%':'#DC143C', '30%':'#990000', '40%':'#800000', index:'white'}
        
        ax.xaxis.set_major_locator(mdates.MonthLocator(interval = 12))
        ax.xaxis.set_major_formatter(mdates.DateFormatter(r'$\bf{%b}$-$\bf{%y}$'))

        ax.yaxis.set_major_locator(plt.MultipleLocator(5))
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: rf'$\bf{x:.0f}.0x$'))

        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        ax.set_axisbelow(True)
        ax.tick_params(axis = 'x', rotation = 90)

        y_step = 0.05
        y_total = (len(leg) - 1) * y_step
        y_start = 0.5 + y_total / 2
        txts = []

        # plt.subplots_adjust(left=0.088, right=0.856, top=1, bottom=0.126) # ANUAL
        plt.subplots_adjust(left=0.088, right=0.855, top=1, bottom=0.11)
        for i, (key, val) in enumerate(leg.items()):
            y_pos = y_start - i * y_step
            t = ax.text(
                x = 1.02,
                y = y_pos,
                s = f'{key}: {val:.1f}x',
                transform = ax.transAxes,
                color = colors[key],
                bbox = dict(
                    facecolor = 'none' if key != index else 'black',
                    edgecolor = 'none',
                    alpha = 0 if key != index else 1
                ),
                va = 'center', ha = 'left'
            )
            txts.append(t)


        fig.canvas.draw()
        renderer = fig.canvas.get_renderer()
        max_w = max(t.get_window_extent(renderer = renderer).width for t in txts)
        wfp = fig.get_size_inches()[0] * fig.dpi
        margin_right = max_w / wfp
        
        # plt.show()

        fig.savefig(path / 'DATOS' / 'Graficas_PPT'/ 'Termometros' / f'{index}_{file_name}.png', dpi = 300)


PATH = Path.home() / 'OneDrive' / '0. Nube Asset Mgmt'
FILES = ['Métricas Índices con Bancos Europa.xlsm', 'Métricas Índices Fwd con Bancos Europa.xlsm']

for idx, file in enumerate(FILES):
    path = PATH  / 'Notas Estructuradas' / file
    if idx == 0:
        fname = 'PE'
    else:
        fname = 'FwdPE'
    
    data = pd.read_excel(path, sheet_name = 'PE', skiprows = 2)
    data = data.rename(columns = {'Unnamed: 1': 'Date'}).drop('Unnamed: 0', axis = 1)[['Date', 'SPX Index', 'NDX Index', 'RTY Index', 'SX5E Index']][:5220]

    term_plots(data, fname, PATH)