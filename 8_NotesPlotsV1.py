### ============================================= ###
### Notas:
###     - Archivos a actualizar antes de ejecutar
###       * .../Nube/DATOS/Notas Presentacion (New).xlsx
###         -- Hoja de VIX
###       * .../Nube/DATOS/Graficas Barreras.xlsx
###         -- Hoja de SERIES
###         -- Hoja de NOTAS (en caso de haber notas llamadas o nuevas notas)
### ============================================= ###

import os
import pandas as pd
import plotly.graph_objects as go
from config import DATE, PARENT_PATH
from debugg.functions import vix_plot

# path_pc = 'C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt'

path_notes = PARENT_PATH / 'Notas Estructuradas'
today = pd.to_datetime(DATE).date()
vix = pd.read_excel(path_notes / 'Notas Presentación (New).xlsx', sheet_name = 'VIX', skiprows = 1)
# vix['FECHA'] = vix.FECHA.dt.to_pydatetime()

# print(type(vix.FECHA.iloc[0]))
# print(pd.__version__)
# exit()
# =============================================================================
#    GRAFICA VIXXO 
# =============================================================================
cols = ['RTY', 'NDX', 'SPX', 'VIXXO', 'VIXXO3M', 'VIXXO6M', 'Max. Funding Rate']
xys = [(0, 0), (0, 0), (0, 0), (0, 0), (0, -30), (0, 30), (0, 0)]
vix_plot(vix, cols, PARENT_PATH, 'VIXXO', xys, 110)

# =============================================================================
# GRAFICA VIX
# =============================================================================
cols = ['RTY', 'NDX', 'SPX', 'VIX', 'Max. Funding Rate']
xys = [(0, 0), (0, 0), (0, 20), (0, 0), (0,0)]
vix_plot(vix, cols, PARENT_PATH, 'VIX', xys, 90)

# =============================================================================
#  Graficas barreras
# =============================================================================
notes = pd.read_excel(path_notes / 'Graficas_Barreras - copia.xlsx', engine = 'openpyxl', sheet_name = 'NOTAS')
series = pd.read_excel(path_notes / 'Graficas_Barreras - copia.xlsx', engine = 'openpyxl', sheet_name = 'SERIES')
colors = pd.read_excel(path_notes / 'Graficas_Barreras - copia.xlsx', engine = 'openpyxl', sheet_name = 'COLORS')

series['Date'] = series.Date.dt.date

# notes = notes[notes.Estatus == 'ON GOING'].reset_index(drop = True)
notes = notes[(notes.Estatus == 'ON GOING') & (notes.IN_AA == 'YES') & (notes.SHORT_NAME != 'Nota DXY 08/27/2027')].reset_index(drop = True)
notes['Strike'] = 1
notes = notes[['ISIN', 'SHORT_NAME', 'Barrera', 'CCY', 'Nominal', 'Trade Date', 'Call_Period', 'CVD 1','CVD 2','CVD 3','CVD 4','CVD 5','CVD 6','CVD 7','CVD 8','Strike','BANK']]
notes['Barrera'] = 1 - notes.Barrera

# Elimina todas las imagenes de la direccion de Notas 
# para escribir las nuevas
# imgs_dir = f'{path_pc}/DATOS/Graficas_PPT/NOTAS/'
imgs_dir = PARENT_PATH / 'DATOS' / 'Graficas_PPT' / 'Notas'

c = 0
for file in os.listdir(imgs_dir):
    file_path = os.path.join(imgs_dir, file)
    if os.path.isfile(file_path):
        os.remove(file_path)
        c += 1
print(f'{c} images removed')

# Escribe las nuevas imagenes de notas
for idx, row in notes.iterrows():
    names = row.SHORT_NAME.split(' ')[1].split('|')
    underls = pd.DataFrame()
    underls['Date'] = series[series.Date >= row['Trade Date'].date()].Date
    
    last_vals = []
    plot = go.Figure()
    
    for name in names:
        name_serie = f'{name} Index'
        underls[name_serie] = series[series.Date >= row['Trade Date'].date()][name_serie]
        underls[name] = underls[name_serie] / underls[name_serie].iloc[0]
        
        color_idx = names.index(name)
        clr = colors.COLOR.iloc[color_idx]
        plot.add_trace(go.Scatter(
            x = underls.Date, y = underls[name], name = name,
            mode = 'lines+markers', line = {'width': 2},
            marker = dict(color = clr, size = 3),
            fillcolor = clr, line_color = clr
        ))
        last_vals.append(underls[name].iloc[-1])
    
    min_val = min(last_vals)
    
    plot.add_annotation(
        x = underls.Date.iloc[-1], y = row.Barrera,
        ax = underls.Date.iloc[-1], ay = min_val,
        xref = 'x', yref = 'y', axref = 'x', ayref = 'y',
        arrowhead = 1, arrowsize = 1, arrowwidth = 2,
        arrowcolor = 'blue'
    )
    plot.add_hline(
        y = 1, name = 'Strike', line = {'width': 2, 'dash': 'dash'},
        fillcolor = '#C7BBAC', line_color = '#C7BBAC'
    )
    plot.add_hline(
        y = row.Barrera, name = 'Barrier', line = {'width': 2, 'dash': 'dash'},
        fillcolor = '#C00000', line_color = '#C00000'
    )
    obs_list = row[[f'CVD {i}' for i in range(1, 9)]].dropna().to_list()
    
    for obs in obs_list:
        plot.add_vline(x = obs, line = dict(width = 1, dash = 'dash'))
        if obs.date() > today:
            max_x = obs
            break
    
    plot.add_annotation(
        x = underls.Date.iloc[0], y = row.Barrera + 0.04, # y = row.Barrera + 0.04,
        ax = underls.Date.iloc[0], ay = row.Barrera + 0.04, # ay = row.Barrera + 0.04,
        xref = 'x', yref = 'y', axref = 'x', ayref = 'y',
        xanchor = 'left',
        text = f'Distance {round((min_val - row.Barrera) * 100, 2)}%',
        font = dict(family = 'Arial Black', color = '#000000'),
        # showarrow = False, arrowhead = 1, arrowsize = 1,
        # arrowwidth = 2, arrowcolor = 'blue'
    )
    plot.add_annotation(
        x = underls.Date.iloc[0], y = row.Barrera + 0.09, # y = row.Barrera + 0.09,
        ax = underls.Date.iloc[0], ay = row.Barrera + 0.09, # ay = row.Barrera + 0.09
        xref = 'x', yref = 'y', axref = 'x', ayref = 'y',
        xanchor = 'left',
        text = f'CVD: {row[f"CVD {int(row.Call_Period)}"].strftime("%d/%m/%Y")}',
        font = dict(family = 'Arial Black', color = '#000000')
    )
    plot.update_layout(
        plot_bgcolor = '#FFFFFF', showlegend = True, font = dict(family = 'Arial', size = 10, color = '#000000'),
        xaxis = {'dtick': 'M1'}, xaxis_tickformat = '%b-%y', yaxis_tickformat = ',.0%', margin = dict(l=0,r=0,t=50,b=0),
        title = f'{row.BANK}-{row.SHORT_NAME.replace("Nota", "")}|{row.Barrera:,.0%}|${row.Nominal/1000000:,.2f} Mn {row.CCY}',
        legend = dict(orientation = 'h', yanchor = 'top', xanchor = 'center', y = 1.1, x = 0.5,
                       font = dict(family = 'Arial', size = 12, color = '#000000')),
        xaxis_range = [underls.Date.iloc[0], max_x + pd.Timedelta(days = 3)],
        yaxis_range = [row.Barrera - 0.02, underls[names].to_numpy().max() + 0.2] # Sin convertir a Numpy la escala cambia un poco (se ve un poco mejor a mi parecer)
    )
    plot.update_xaxes(
        ticks = 'outside', showline = True, linewidth = 2, linecolor = 'black', tickcolor = 'black', tickangle = -90
    )
    plot.update_yaxes(
        ticks = 'outside', showline = True, linewidth = 2, linecolor = 'black', gridwidth = 1, gridcolor = 'LightGray',
        showgrid = True
    )
    plot.write_image(
        # f'{path_pc}/DATOS/Graficas_PPT/Notas/{idx}.png', width = 11 * 37.795276, height = 7 * 37.795276, scale = 1
        imgs_dir / f'{idx}.png', width = 11 * 37.795276, height = 7 * 37.795276, scale = 1
    )
    # plot.show()


### =========================================================== ###
###                        GRAFICA DXY                          ###
### =========================================================== ###

import numpy as np
# import pandas as pd
# import matplotlib.pyplot as plt
# import plotly.graph_objects as go

# path_notes = 'C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Notas Estructuradas'
# path_pc = 'C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt'
notes = pd.read_excel(f'{path_notes}/Graficas_Barreras - copia.xlsx', engine = 'openpyxl', sheet_name = 'NOTAS')
series = pd.read_excel(f'{path_notes}/Graficas_Barreras - copia.xlsx', engine = 'openpyxl', sheet_name = 'SERIES')
colors = pd.read_excel(f'{path_notes}/Graficas_Barreras - copia.xlsx', engine = 'openpyxl', sheet_name = 'COLORS')

notes = notes[notes.Estatus == 'ON GOING'].reset_index(drop = True)
notes['Strike'] = 1

bs = 97.84

def rend_dxy(x:pd.Series, base:int):
    return np.maximum(100, (x / base - 1) * 182 + 100)
# def rend_dxy(x:pd.Series, base:int):
#     return np.maximum(0, (x / base - 1) * 1.82)

dxy = series[series.Date >= pd.to_datetime('13/08/2025', format = '%d/%m/%Y')][['Date', 'DXY Index']]
dxy['DXY Index'] = dxy['DXY Index'].apply(rend_dxy, base = bs) / 100
# dxy['Date'] = dxy.Date.dt.strftime('%Y-%m-%d') #'%b-%y'
plot = go.Figure()
clr = colors.COLOR.iloc[0]
plot.add_trace(go.Scatter(
    x = dxy.Date, y = dxy['DXY Index'], name = 'DXY',
    mode = 'lines+markers', line = {'width': 2},
    marker = dict(color = clr, size = 3),
    fillcolor = clr, line_color = clr
))
plot.add_annotation(
    x = dxy.Date.iloc[-1], y = 1,
    ax = dxy.Date.iloc[-1], ay = 1,
    xref = 'x', yref = 'y', axref = 'x', ayref = 'y',
    arrowhead = 1, arrowsize = 1, arrowwidth = 2,
    arrowcolor = 'blue'
)
plot.add_hline(
    y = 1, name = 'Strike', line = {'width': 2, 'dash': 'dash'},
    fillcolor = '#C7BBAC', line_color = '#C7BBAC'
)
plot.add_hline(
    y = 1, name = 'Barrier', line = {'width': 2, 'dash': 'dash'},
    fillcolor = '#C00000', line_color = '#C00000'
)

plot.update_layout(
    plot_bgcolor = '#FFFFFF', showlegend = True, font = dict(family = 'Arial', size = 10, color = '#000000'),
    xaxis = {'dtick': 'M1'}, xaxis_tickformat = '%b-%y', yaxis_tickformat = ',.2%', margin = dict(l=0,r=0,t=50,b=0),
    title = f'JBR-DXY 08/27/2027|100%|$0.5 Mn USD',
    legend = dict(orientation = 'h', yanchor = 'top', xanchor = 'center', y = 1.1, x = 0.5,
                    font = dict(family = 'Arial', size = 12, color = '#000000')),
    xaxis_range = [dxy.Date.iloc[0], pd.to_datetime('27122025', format = '%d%m%Y') + pd.Timedelta(days = 3)],
    yaxis_range = [1 - 0.001, dxy['DXY Index'].to_numpy().max() + 0.01] # Sin convertir a Numpy la escala cambia un poco (se ve un poco mejor a mi parecer)
)
plot.update_xaxes(
    ticks = 'outside', showline = True, linewidth = 2, linecolor = 'black', tickcolor = 'black', tickangle = -90
)
plot.update_yaxes(
    ticks = 'outside', showline = True, linewidth = 2, linecolor = 'black', gridwidth = 1, gridcolor = 'LightGray',
    showgrid = True
)
plot.write_image(
    # f'{path_pc}/DATOS/Graficas_PPT/Notas/{idx+1}.png', width = 11 * 37.795276, height = 7 * 37.795276, scale = 1
    imgs_dir / f'{idx+1}.png', width = 11 * 37.795276, height = 7 * 37.795276, scale = 1
)
# plot.show()