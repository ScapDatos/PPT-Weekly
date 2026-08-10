import pandas as pd
from pathlib import Path
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from functionsV1 import get_df, port_flows

path_pc = Path.home() / 'OneDrive' / '0. Nube Asset Mgmt' / 'DATOS'
# path_pc = 'C:/Users/diego/OneDrive/0. Nube Asset Mgmt/DATOS'
path_to_db = f'{path_pc}/Bases de Datos'

colors = pd.read_excel(f'{path_pc}/INDICES.xlsx', sheet_name = 'COLORES')

### ============================================================================= ###
###                              TABLA CON FX Y HCITY                             ###
### ============================================================================= ###
#### FX ####
fx = get_df(path_to_db, 'SigCap.db', 'FX', where = "WHERE DATE >= '2023-12-31'")
mxnusd = [0] + fx['MXNUSD'].values.tolist()
fx = fx['EURUSD']

#### FX YTD ####
fx_ytd = get_df(path_to_db, 'SigCap.db', 'FX', where = "WHERE DATE >= '2025-12-31'")
mxnusd_ytd = fx_ytd.MXNUSD.values.tolist()
fx_ytd = fx_ytd.EURUSD

#### HCITY ####
hcity = get_df(path_to_db, 'SigCap.db', 'HCITY', where = 'WHERE DATE >= "2023-12-31"')
hcity['TRANSACTIONS'] = hcity.TRANSACTIONS * mxnusd

#### HCITY YTD ####
hcity_ytd = get_df(path_to_db, 'SigCap.db', 'HCITY', where = 'WHERE DATE >= "2025-12-31"')
hcity_ytd['TRANSACTIONS'] = hcity_ytd.TRANSACTIONS * mxnusd_ytd

### ============================================================================= ###
###                              TABLAS POR PORTAFOLIO                            ###
### ============================================================================= ###
#### SCH ####
csch = 'WHERE DEPS IS NOT NULL AND DATE >= "2023-12-31"'
cschl = 'WHERE OP_LOAN IS NOT NULL AND DATE >= "2023-12-31"'
cols = ['DATE','Retorno Efectivo No Apalancado','Retorno Efectivo Apalancado','Retorno para Sharp AUMS', 'Retorno para Sharp NAV','Retorno Efectivo No Apalancado sin HCITY','Retorno Efectivo Apalancado sin HCITY']
sch = port_flows(['SSZ', 'DBK', 'GSS', 'UBS', 'JPM', 'MSY'], path_to_db, 'SCH', csch, cschl, hcity, fx, cols)

#### SCH YTD ####
csch_ytd = 'WHERE DEPS IS NOT NULL AND DATE >= "2025-12-31"'
cschl_ytd = 'WHERE OP_LOAN IS NOT NULL AND DATE >= "2025-12-31"'
sch_ytd = port_flows(['SSZ', 'DBK', 'GSS', 'UBS', 'JPM', 'MSY'], path_to_db, 'SCH', csch_ytd, cschl_ytd, hcity_ytd, fx_ytd, cols, dt = '"2025-12-31"')


hcity['TRANCUM'] = hcity.TRANSACTIONS.cumsum()
hcity_ytd['TRANCUM'] = hcity_ytd.TRANSACTIONS.cumsum()

sch.rename(columns = {
    'EQTS': 'EQUITIES',
    'LQDTY': 'LIQUIDITY',
    'FIX_INC': 'FIXED_INCOME',
    'ALTS': 'ALTERNATIVES',
    'DRVTS': 'DERIVATIVES',
    'STR_NTS': 'STRUCTURED_NOTES',
    'LOAN_TOTAL': 'LOAN'
}, inplace = True)
sch_ytd.rename(columns = {
    'EQTS': 'EQUITIES',
    'LQDTY': 'LIQUIDITY',
    'FIX_INC': 'FIXED_INCOME',
    'ALTS': 'ALTERNATIVES',
    'DRVTS': 'DERIVATIVES',
    'STR_NTS': 'STRUCTURED_NOTES',
    'LOAN_TOTAL': 'LOAN'
}, inplace = True)

sch_w_hcity = pd.merge(sch, hcity[['DATE', 'HCITY_SCH', 'LOANS_SCH', 'TRANSACTIONS', 'TRANCUM']], how='left', on = 'DATE')
sch_w_hcity_ytd = pd.merge(sch_ytd, hcity_ytd[['DATE', 'HCITY_SCH', 'LOANS_SCH', 'TRANSACTIONS', 'TRANCUM']], how='left', on = 'DATE')

loans = pd.read_excel('C:/Users/arnol/OneDrive/0. Nube Asset Mgmt/Notas Estructuradas/Graficas_Barreras - copia.xlsx', sheet_name = 'INTPAYS_SCH')
loans = loans[loans.columns[:19]]

loans['Date'] = loans.Date + pd.to_timedelta((4 - loans.Date.dt.dayofweek) % 7, unit = 'D')
values_col = [c for c in loans.columns if c != 'Date']
weekly = loans.groupby('Date', as_index = False).sum()
weekly['LOAN_USD'] = weekly[values_col[:-4]].sum(axis = 1) * -1
weekly['OPERATIONS'] = weekly[values_col[-4:]].sum(axis = 1) * -1
weekly.loc[weekly.Date == '2025-02-14', 'LOAN_USD'] = -3e5 #16e4

sch_w_hcity = pd.merge(sch_w_hcity, weekly[['Date', 'LOAN_USD', 'OPERATIONS']], how = 'left', left_on = 'DATE', right_on = 'Date').fillna(0)
sch_w_hcity_ytd = pd.merge(sch_w_hcity_ytd, weekly[['Date', 'LOAN_USD', 'OPERATIONS']], how = 'left', left_on = 'DATE', right_on = 'Date').fillna(0)

sch_w_hcity.loc[sch_w_hcity.index[-1], ['OPERATIONS', 'LOAN_USD']] = -10
sch_w_hcity_ytd.loc[sch_w_hcity_ytd.index[-1], ['OPERATIONS', 'LOAN_USD']] = -10

with pd.ExcelWriter('C:/Users/arnol/OneDrive/0. Nube Asset Mgmt/Contributtion/exp_rends.xlsx', engine = 'openpyxl') as writer:
    sch_w_hcity.to_excel(writer, sheet_name = 'SCH', index = False)
    sch_w_hcity_ytd.to_excel(writer, sheet_name = 'SCH_YTD', index = False)
exit()
# print(sch_w_hcity.LOAN_USD.min())

# rends = pd.read_excel('C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Contributtion/exp_rends.xlsx', sheet_name = 'SCH')
# rends_ytd = pd.read_excel('C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Contributtion/exp_rends.xlsx', sheet_name = 'SCH_YTD')

# print(rends.LOAN_USD.min())

colors = {
    'CASH':	           '#97B7B3',
    'LIQUIDITY':       '#637976',
    'EQUITIES':        '#475161',
    'FIXED_INCOME':    '#8F8F8F',
    'ALTERNATIVES':    '#867B74',
    'STRUCTURED_NOTES':'#C2B994',
    'DERIVATIVES':     '#89687E',
    'OPERATIONS':      "#9AEC79",
    'LOAN':            '#FF7C80',
    'LOAN_USD':        '#FFD1B2',
    'INTS_SCH':        '#FFD1B2',
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


def plot_confs(figure, ax, labels:list[str], y_step:int, y_step2:int, counter:int, annot_size:int, color_list:list[str], fontc_list:list[str],
               xTicks:int, yTicks:int, leg_size:int, ncols:int, legLoc:str, bbox:tuple, yformat:str, ax2 = None, flag:bool = False,
               right:float = 1.5, x2Ticks:int = 25, glw:float = 0.5):
    
    ax.grid(axis = 'y', color = 'lightgray', lw = glw, alpha = 0.5)
    ax.tick_params(axis = 'x', rotation = 90)
    ax.tick_params(axis = 'both', labelsize = 10, labelfontfamily = 'Arial')
    
    ax.legend(
        loc = f'upper {legLoc}',
        bbox_to_anchor = bbox,
        fontsize = leg_size,
        frameon = False,
        ncol = ncols,
        prop = {'family':'Arial'}
    )
    
    nlines = [label.count('\n') + 1 for label in labels]
    steps = [n * (y_step - 0.03) if n > 1 else y_step for n in nlines]
    if counter != 0:
        steps[-counter-1:] = [y_step] + [y_step2] * counter
    y_start = 0.5 + (sum(steps) - y_step) / 2
    txts = []
    
    for label, step, fc, bc in zip(labels, steps, fontc_list, color_list):
        txt = ax.text(
            x = (1.07 if ax2 is not None else 1.01) if not flag else right,
            y = y_start,
            s = label,
            transform = ax.transAxes,
            fontsize = annot_size,
            color = fc,
            bbox = dict(
                facecolor = bc,
                edgecolor = 'none'
            ),
            va = 'center', ha = 'left'
        )
        y_start -= step
        txts.append(txt)
    
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval = xTicks))
    ax.xaxis.set_major_formatter(mdates.DateFormatter(r'$\bf{%b}$-$\bf{%y}$'))
    
    ax.yaxis.set_major_locator(plt.MultipleLocator(yTicks))
    ax.yaxis.set_major_formatter(plt.FuncFormatter(yformat))
    
    ymin1, ymax1 = ax.get_ylim()
    frac = (0 - ymin1) / (ymax1 - ymin1)
    
    if ax2 is not None:
        
        ymin2, ymax2 = ax2.get_ylim()
        newmax = -ymin2 / frac + ymin2
        ax2.set_ylim(ymin2, newmax)
        
        ax2.tick_params(axis = 'both', labelsize = 9, labelfontfamily = 'Arial')
        ax2.yaxis.set_major_locator(plt.MultipleLocator(x2Ticks))
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: rf'$\bf\${x/1e3:.0f}k$'))
        ax2.spines['right'].set_visible(False)
        ax2.spines['top'].set_visible(False)

        ax2.legend(
            loc = f'upper center',
            bbox_to_anchor = (0.24, 1.06),
            fontsize = leg_size,
            frameon = False,
            ncol = ncols,
            prop = {'family':'Arial'}
        )
    
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    # ax.set_axisbelow(True)
    
    figure.canvas.draw()
    renderer = figure.canvas.get_renderer()
    max_w = max(t.get_window_extent(renderer = renderer).width for t in txts)
    wfp = figure.get_size_inches()[0] * figure.dpi
    margin_right = max_w / wfp
    
    return margin_right

sorter     =['NAV',' Santander Suiza','Julius Baer','Citibank','UBS','JP Morgan',
             'Goldman Sachs','Morgan Stanley','Deutsche Bank','Santander Suiza',
             'Bankinter','TD Ameritrade','UNI','CASH','LIQUIDITY','EQUITIES',
             'FIXED_INCOME','ALTERNATIVES','STRUCTURED_NOTES','DERIVATIVES',
             'BANKINTER', 'OPERATIONS','LOAN_EUR','LOAN_USD','LOAN_MXN','LOAN','LOAN_UNSEC']

def stacked_areas(path:str, title:str, data:pd.DataFrame, namesList:list[str], yTicks:str|int, xTicks:str|int,
                  colors:dict, annot_size:int = 10, startDate:str = '', labels:bool = False, bnk:bool = False,
                  vline:bool = False, leg_size:float = 10, ncols:int = 8, legLoc:str = 'left', 
                  bbox:tuple = (0.001, 1.1), left:float = 0.053, xt2:int = 1e5):
    """
    Descript:
        Genera una imagen de areas para los movimientos dentro del portafolio y la guarda dentro de la ruta especificada. 
    
    Args:
        path (str): Ruta hacia donde se debe guardar el archivo generado.
        pname (str): Nombre del archivo generado (sin extension).
        data (DataFrame): Tabla de datos del cual se obtiene la informacion para generar la grafica.
        namesList (list[str]): Lista de columnas a usar de la tabla dada.
        yTicks (str or int): Formato de etiqueta en el eje 'y' de la grafica generada (debe ser compatible con Plotly).
        xTicks (str or int): Formato de etiqueta en el eje 'x' de la grafica generada (debe ser compatible con Plotly).
        colors (dict): Diccionario que contiene los colores para cada grafica y banco.
        startDate (str)(default = ''): Fecha desde donde se tomaran los datos para las graficas.
        fdays (int)(default = 430): Numero de dias agregados al rango del eje 'x' para agregar las anotaciones.
        delayPixels (int)(default = 0): Numero de pixeles para ajustar correctamente las anotaciones.
        labels (bool)(default = False): Valor para decidir si poner anotaciones dentro de la grafica (generadas automaticamente).
    Returns:
        (None) Guarda la imagen generada en la ruta especificada.
    """
    if startDate != '':
        data = data[data.DATE.dt.strftime('%Y-%m-%d').between(startDate, str(data.DATE.iloc[-1]))].reset_index(drop = True)
    
    df_aux = data[namesList].iloc[-1].dropna().to_frame().reset_index()
    df_aux.columns = ['Name', 'Value']
    df_aux['Color'] = df_aux.Name.map(colors)
    
    notSumVals = ['NAV', 'LOAN', 'LOAN_USD', 'LOAN_EUR', 'LOAN_MXN', 'LOAN_UNSEC', 'OPERATIONS']
    blackFontVals = ['LOAN', 'LOAN_USD', 'LOAN_EUR', 'LOAN_MXN', 'LOAN_UNSEC', 'BANKINTER', 'STRUCTURED_NOTES', 'CASH', 'TD Ameritrade']
    primAxisCols = ['Julius Baer', 'Citibank', 'UBS', 'Goldman Sachs', 'Morgan Stanley', 'Deutsche Bank', 'Santander Suiza',
                    'TD Ameritrade', 'HCITY', 'BANKINTER', 'CASH', 'LIQUIDITY', 'EQUITIES', 'FIXED_INCOME', 'ALTERNATIVES',
                    'STRUCTURED_NOTES', 'DERIVATES', 'UNI']
    secAixsCols = ['OPERATIONS', 'LOAN', 'LOAN_USD', 'LOAN_EUR', 'LOAN_MXN', 'LOAN_UNSEC', 'JP Morgan', ' Santander Suiza']
    mapping = {'Julius Baer':'JBR', 'Citibank':'CTI', 'UBS':'UBS', 'Goldman Sachs':'GSS', 'Morgan Stanley':'MSY', 'Deutsche Bank':'DBK', 
               'Santander Suiza':'SSZ', 'JP Morgan':'JPM', ' Santander Suiza':'SSZ', 'Bankinter':'BKR'}
    
    df_aux['Pct'] = round(df_aux.Value / df_aux[~df_aux.Name.isin(notSumVals)].Value.sum() * 100, 1)
    
    nav = df_aux[df_aux.Name == 'NAV'].Value.values[0]
    
    # df_aux = df_aux.sort_values('Name', key = lambda col: col.map(lambda e: sorter_draw.index(e))).reset_index(drop = True)
    df_aux = df_aux.sort_values('Name', key = lambda col: col.map(lambda e: sorter.index(e))).reset_index(drop = True)
    
    cols = [col for col in data.columns if col in primAxisCols]
    cols2 = [col for col in data.columns if col in secAixsCols]
    if title not in ['SCHOLDINGS', 'NFD']:
        data['SUM_PAXIS'] = data[cols].sum(axis = 1)
    else:
        data['SUM_PAXIS'] = data[cols + ['NAV']].sum(axis = 1)
    data['SUM_SAXIS'] = data[cols2].sum(axis = 1)
    second_axis_names = set(df_aux[df_aux.Name.isin(secAixsCols)].Name)
    
    plt.rcParams['axes.xmargin'] = 0
    fig, ax = plt.subplots(figsize=(33/2.54, 15/2.54))
    ax2 = ax.twinx()
    
    labels_right, color_list, fontc_list = [f'${nav:,.1f}'], ['b'], ['w']
    loan_names = {'OPERATIONS', 'LOAN', 'LOAN_USD', 'LOAN_EUR', 'LOAN_MXN', 'LOAN_UNSEC'}
    
    ax.plot(
        data.DATE,
        data.NAV,
        'o-b',
        lw = 2,
        ms = 3,
        label = 'NAV'
    )
    
    low_asset = pd.Series([0]*len(data.DATE))
    low_loan = pd.Series([0]*len(data.DATE))
    c = 0
    
    for idx, row in df_aux.iterrows():
        if row.Name != 'NAV':
            target_ax = ax2 if (row.Name == 'LOAN_USD') or (row.Name == 'OPERATIONS') else ax
            target_ax.fill_between(
                data.DATE,
                low_loan if row.Name in second_axis_names else low_asset, # 0
                low_loan + data[row.Name] if row.Name in second_axis_names else low_asset + data[row.Name], # data[row.Name]
                color = row.Color,
                label = row.Name,
                lw = 0
            )
            
            if row.Value != 0:
                txt = f'${row.Value:,.1f}\n{row.Pct}% - {mapping[row.Name]}' if bnk else f'${row.Value:,.1f}\n{row.Pct}%'
                txt_loan = f'${row.Value:,.1f}'
                labels_right.append(txt_loan if row.Name in loan_names else txt)
                color_list.append(row.Color)
                fontc_list.append('k' if row.Name in blackFontVals else 'w')
                c += 1 if row.Name in loan_names else 0
            low_loan += data[row.Name] if row.Name in second_axis_names else 0
            low_asset += data[row.Name] if row.Name not in second_axis_names else 0
    
    if labels:
        data['updn'] = data.LABEL.apply(lambda x: 20 if x == 'UP' else -20)
        dl = data[data.LABEL != 'NO']
        for idx, row in dl.iterrows():
            ax.annotate(
                f'${row.NAV:,.1f}',
                xy = (row.DATE, row.NAV),
                xytext = (1, row.updn),
                textcoords = 'offset points',
                fontsize = 9,
                color = 'k',
                bbox = dict(
                    facecolor = 'w',
                    edgecolor = 'none'
                ),
                va = 'center', ha = 'left',
                arrowprops = dict(
                    arrowstyle = '-|>',
                    lw = 1.5,
                    color = 'k'
                )
            )
    if vline:
        lines = data[data.VLINE != 'NO']
        for idx, row in lines.iterrows():
            ax.axvline(
                x = row.DATE,
                ymax = 0.98,
                ls = '-.',
                c = '#187F3D' if row.VLINE == 'YES_G' else '#C00000',
                lw = 1
            )
    
    ymin = min(low_loan.min().min(), data.NAV.min().min())
    ymax = max(low_asset.max().max(), data.NAV.max().max())
    ax.set_ylim(ymin - 500000, ymax + 500000)
    
    margin_right = plot_confs(fig, ax, labels_right, 0.078, 0.057, c, annot_size, color_list, fontc_list, xTicks, yTicks,
                              leg_size, ncols, legLoc, bbox, lambda x, _: rf'$\bf\${x/1000000:.0f}M$', ax2, x2Ticks = xt2)
    
    plt.subplots_adjust(left = 0.052, right = 0.856, top = 0.928, bottom = 0.109)
    fig.savefig(f'{path}/{title}.png', dpi = 300)
    plt.close()
    # plt.show()




cols = ['CASH', 'LIQUIDITY', 'EQUITIES', 'FIXED_INCOME', 'ALTERNATIVES', 'STRUCTURED_NOTES', 'DERIVATIVES', 'OPERATIONS', 'LOAN_USD', 'LOAN', 'NAV']
stacked_areas(path_pc, 'Graficas_PPT/SCH_EXP', sch_w_hcity, cols, 5e6, 1, colors = colors, startDate = '2023-12-31', bbox = (0.001, 1.1), left = 0.04, xt2 = 5e4)
stacked_areas(path_pc, 'Graficas_PPT/SCH_EXP_YTD', sch_w_hcity_ytd, cols, 5e6, 1, colors = colors, startDate = '2026-01-02', bbox = (0.001, 1.1), left = 0.04, xt2 = 1e5)