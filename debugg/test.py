import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from functionsV1 import plot_confs


def vix_plot(path:str, title:str, data:pd.DataFrame, columns:list[str], xTicks:int = 1, yTicks:int = 50,
             leg_size:int = 9, ncols:int = 8, legLoc:str = 'left', annot_size:int = 9, sec_y:bool = True,
             bbox:tuple = (0.001, 0.95)):
    """
    
    """
    
    data_ = data.copy()
    data_[['RTY', 'NDX', 'SPX', 'Max. Funding Rate']] = data_[['RTY', 'NDX', 'SPX', 'Max. Funding Rate']] * 100
    
    plt.rcParams['axes.xmargin'] = 0
    fig, ax = plt.subplots(figsize=(33/2.54, 16.18/2.54))
    ax2 = ax.twinx() if sec_y else None
    
    info = {
        'RTY': {'c':'#867B74', 'val':data_.RTY.iloc[-1], 'suf':'RTY'},
        'NDX': {'c':'#7CAFDD', 'val':data_.NDX.iloc[-1], 'suf':'NDX'},
        'SPX': {'c':'#475161', 'val':data_.SPX.iloc[-1], 'suf':'SPX'},
        'VIX': {'c':'#0000FF', 'val':data_.VIX.iloc[-1], 'suf':'VIX'},
        'VIXXO': {'c':'#0000FF', 'val':data_.VIXXO.iloc[-1], 'suf':'VIXXO'},
        'VIXXO3M': {'c':'#FFC000', 'val':data_.VIXXO3M.iloc[-1], 'suf':'VIXXO3M'},
        'VIXXO6M': {'c':'k', 'val':data_.VIXXO6M.iloc[-1], 'suf':'VIXXO6M'},
        'Max. Funding Rate': {'c':'#8F8F8F', 'val':data_['Max. Funding Rate'].iloc[-1], 'suf':'Fund Rate'}
    }
    info = dict(sorted(((k, v) for k, v in info.items() if k in columns), key = lambda item: item[1]['val'], reverse = True))
    
    labels_right = [f"{v['val']:.2f} {k}" if 'VIX' in k else f"{v['val']/100:.2%} {v['suf']}" for k, v in info.items()]
    color_list, fontc_list = [v['c'] for v in info.values()], ['w'] * len(info)
    
    for idx, col in enumerate(columns):
        if 'VIX' in col:
            ax.plot(
                data_.FECHA,
                data_[col],
                c = info[col]['c'],
                ls = '-.',
                lw = 2,
                label = col
            )
        else:
            ax2.plot(
                data_.FECHA,
                data_[col],
                c = info[col]['c'],
                ls = ':' if col == 'Max. Funding Rate' else '-',
                lw = 2,
                label = col if col == 'Max. Funding Rate' else f'{col} Coupons'
            )
            ax2.legend(
                loc = f'upper left',
                bbox_to_anchor = (0.001, 1.0),
                fontsize = 9,
                frameon = False,
                ncol = 8,
                prop = {'family':'Arial'}
            )
    
    margin_right = plot_confs(fig, ax, labels_right, 0.05, 0.0525, 0, annot_size, color_list, fontc_list, xTicks,
                              yTicks, leg_size, ncols, legLoc, bbox, yformat = lambda x, _: rf'$\bf{x:.0f}$',
                              ax2 = ax2, x2Ticks = 1)
    
    plt.subplots_adjust(left = 0.03, right = 1 - margin_right - 0.063, top = 1, bottom = 0.103)
    fig.savefig(f'{path}/{title}.png', dpi = 300)
    plt.close()
    # plt.show()


path_pc = 'C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt'
path_notes = f'{path_pc}/Notas Estructuradas'
vix = pd.read_excel(f'{path_notes}/Notas Presentación (New).xlsx', sheet_name = 'VIX', skiprows = 1)
today = pd.to_datetime('2026-01-23')#.date()

# =============================================================================
#    GRAFICA VIXXO 
# =============================================================================
cols = ['RTY', 'NDX', 'SPX', 'VIXXO', 'VIXXO3M', 'VIXXO6M', 'Max. Funding Rate']
# vix_plot(path_pc, '/DATOS/Graficas_PPT/VIXXO', vix, cols)

# =============================================================================
# GRAFICA VIX
# =============================================================================
cols = ['RTY', 'NDX', 'SPX', 'VIX', 'Max. Funding Rate']
# vix_plot(path_pc, '/DATOS/Graficas_PPT/VIX', vix, cols, yTicks = 10)



# =============================================================================
#  Graficas barreras
# =============================================================================
notes = pd.read_excel(f'{path_notes}/Graficas_Barreras.xlsx', engine = 'openpyxl', sheet_name = 'NOTAS')
series = pd.read_excel(f'{path_notes}/Graficas_Barreras.xlsx', engine = 'openpyxl', sheet_name = 'SERIES')
colors = pd.read_excel(f'{path_notes}/Graficas_Barreras.xlsx', engine = 'openpyxl', sheet_name = 'COLORS')

notes = notes[(notes.Estatus == 'ON GOING') & (notes.IN_AA == 'YES') & (notes.SHORT_NAME != 'Nota DXY 08/27/2027')].reset_index(drop = True)
notes['Strike'] = 1
notes = notes[['ISIN', 'SHORT_NAME', 'Barrera', 'CCY', 'Nominal', 'Trade Date', 'Call_Period', 'CVD 1','CVD 2','CVD 3','CVD 4','CVD 5','CVD 6','CVD 7','CVD 8','Strike','BANK']]
notes['Barrera'] = (1 - notes.Barrera) * 100

imgs_dir = f'{path_pc}/DATOS/Graficas_PPT/NOTAS'
c = 0

for file in os.listdir(imgs_dir):
    file_path = os.path.join(imgs_dir, file)
    if os.path.isfile(file_path):
        os.remove(file_path)
        c += 1
print(f'{c} images removed')

def str_notes(path:str, title:str, nts:pd.DataFrame, ser:pd.DataFrame, colors:pd.DataFrame, today, xTicks:int = 1,
              yTicks:int = 20, leg_size:int = 9, ncols:int = 8, legLoc:str = 'center', annot_size:int = 9, bbox:tuple = (0.5, 1.05)):
    """
    
    """
    
    for idx, row in nts.iterrows():
        names = row.SHORT_NAME.split()[1].split('|')
        underls = pd.DataFrame()
        underls['Date'] = ser[ser.Date >= row['Trade Date']].Date
        
        last_vals = []
        
        # plt.rcParams['axes.xmargin'] = 0
        fig, ax = plt.subplots(figsize=(11/2.54, 7/2.54))
        
        for name in names:
            name_serie = f'{name} Index'
            vals = ser[ser.Date >= row['Trade Date']][name_serie]
            underls[name] = vals / vals.iloc[0] * 100
            
            color_idx = names.index(name)
            clr = colors.COLOR.iloc[color_idx]
            
            ax.plot(
                underls.Date,
                underls[name],
                c = clr,
                ls = '-',
                marker = 'o',
                lw = 1,
                ms = 1.5,
                label = name
            )
            last_vals.append(underls[name].iloc[-1])
        
        min_val = min(last_vals)
        
        ax.annotate(
            '',
            xy = (underls.Date.iloc[-1], row.Barrera),
            xytext = (underls.Date.iloc[-1], min_val),
            arrowprops = dict(
                arrowstyle = '-|>',
                lw = 1.5,
                color = 'b'
            )
        )
        
        ax.axhline(
            y = 100,
            # xmax = 0.95,
            c = '#C7BBAC',
            ls = '--',
            lw = 1
        )
        ax.axhline(
            y = row.Barrera,
            # xmax = 0.9,
            c = '#C00000',
            ls = '--',
            lw = 1
        )
        
        obs_list = row[[f'CVD {i}' for i in range(1, 9)]].dropna().to_list()
        for obs in obs_list:
            ax.axvline(
                x = obs,
                ymax = 0.9,
                c = 'gray',
                ls = '--',
                lw = 1
            )
            if obs > today:
                max_x = obs
                break
        
        distance = f'Distance: {min_val - row.Barrera:.2f}%'
        cvd = f'CVD: {row[f"CVD {int(row.Call_Period)}"].strftime("%d/%m/%Y")}'
        
        
        margin_right = plot_confs(fig, ax, [''], 0, 0, 0, annot_size, ['w'], ['w'], xTicks, yTicks, leg_size, ncols,
                                  legLoc, bbox, yformat = lambda x, _: rf'$\bf{x:.0f}\%$', glw = 0.5)
       
        ax.annotate(
            f'{distance}\n{cvd}',
            xy = (row['Trade Date'], row.Barrera),
            xytext = (-5, 10),
            textcoords = 'offset points',
            fontsize = 7,
            color = 'k',
            fontweight = 'bold',
            bbox = dict(
                facecolor = 'none',
                edgecolor = 'none'
            ),
            va = 'center', ha = 'left'
        )
        
        ax.set_ylim(row.Barrera - 2, underls[names].max().max() + 100 - row.Barrera)
        # ax.set_ylim(row.Barrera - 2, 210 - row.Barrera)
        ax.tick_params(axis = 'both', labelsize = 7, labelfontfamily = 'Arial')
        
        plt.title(f'{row.BANK}-{row.SHORT_NAME.replace("Nota", "")}|{row.Barrera:,.0f}%|${row.Nominal/1000000:,.2f} Mn {row.CCY}', fontsize = 9)
        plt.subplots_adjust(left = 0.093, right = 0.99, top = 0.935, bottom = 0.183)
        fig.savefig(f'{path}/{idx}.png', dpi = 300)
        plt.close()
        # plt.show()

str_notes(imgs_dir, '', notes, series, colors, today)