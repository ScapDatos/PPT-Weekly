import pptx
import sqlite3
import numpy as np
import pandas as pd
import plotly.express as px
from pptx.util import Pt
from pptx.dml.color import RGBColor
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

### ======================================================= ###
###                     General Functions                   ###
### ======================================================= ###
def get_df(pathToDB:str, dbName:str, tableName:str, where:str = None):
    """
    Descrit:
        Manda a llamar una tabla desde la base de datos seleccionada.
    
    Args:
        pathToDB (str): Ruta hacia la base de datos, ingresata como cadena de texto.
        dbName (str): Nombre de la base de datos a ser leida.
        tableName (str): Nombre de la tabla dentro de la base de datos dada a ser leida.
        where (str): Clausula 'where' dentro del query por si se quiere filtrar aun mas la tabla (opcional).
    Returns:
        pandas.DataFrame
    """
    
    conn = sqlite3.connect(f'{pathToDB}/{dbName}')
    
    if where is not None:
        query = f"""
            SELECT * FROM {tableName}
            {where}
            ORDER BY DATE;
            """
    else:
        query = f"""
                SELECT * FROM {tableName}
                ORDER BY DATE;
                """
    df = pd.read_sql_query(query, con = conn)
    df['DATE'] = pd.to_datetime(df.DATE, format = '%Y-%m-%d')
    conn.close()
    
    return df

def plot_confs(figure, ax, labels:list[str], y_step:int, y_step2:int, counter:int, annot_size:int, color_list:list[str], fontc_list:list[str],
               xTicks:int, yTicks:int, leg_size:int, ncols:int, legLoc:str, bbox:tuple, yformat:str, ax2 = None, flag:bool = False,
               right:float = 0.958, x2Ticks:int = 25, glw:float = 2):
    
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
    
    if ax2 is not None:
        ax2.tick_params(axis = 'both', labelsize = 10, labelfontfamily = 'Arial')
        ax2.yaxis.set_major_locator(plt.MultipleLocator(x2Ticks))
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: rf'$\bf{x:.0f}\%$'))
        ax2.spines['right'].set_visible(False)
        ax2.spines['top'].set_visible(False)
    
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.set_axisbelow(True)
    
    figure.canvas.draw()
    renderer = figure.canvas.get_renderer()
    max_w = max(t.get_window_extent(renderer = renderer).width for t in txts)
    wfp = figure.get_size_inches()[0] * figure.dpi
    margin_right = max_w / wfp
    
    return margin_right


### ======================================================= ###
###                         Codigo 3                        ###
### ======================================================= ###
def del_record(path_db:str, date:str):
    """
    Descript:
        Elimina los registros de todas las tablas en cada base de datos siempre que la fecha coincida
        con la que se ingreso (date).
    
    Args:
        path_db (str): Ruta hacia las bases de datos.
        date (str): Fecha que actua como identificador unico (ID) para la eliminacion de datos.
    Returns:
        (None) Elimina los datos en cuestion. 
    """
    # dbss = ['SCAP_S', 'SCAP_N', 'SCAP_W', 'SCAP']
    dbss = ['SigCap_SCH', 'SigCap_NFD', 'SigCap_WK', 'SigCap']
    for dbs in dbss:
        print(f'\n==================== {dbs} ====================')
        conn = sqlite3.connect(f'{path_db}/{dbs}.db')
        cursor = conn.cursor()

        cursor.execute(
            """
            SELECT name FROM sqlite_master
            WHERE type = 'table' AND name NOT LIKE 'sqlite_%';
            """
        )

        tablas = cursor.fetchall()

        for (tabla,) in tablas:
            try:
                sql = f'DELETE FROM {tabla} WHERE DATE = ?;'
                cursor.execute(sql, (date,))
                print(f'Datos eliminados para {date} de la tabla: {tabla}. Filas afectadas: {cursor.rowcount}')
            except sqlite3.OperationalError as e:
                print(f'No se pudo eliminar de la tabla {tabla}: {e}')

        conn.commit()
        conn.close()

def insert_vals(conn:sqlite3.connect, data:pd.DataFrame, table:str, date:str, values:list|tuple):
    """
    Descript:
        Ingresa los valores a las tablas especificadas.
    
    Args:
        conn (sqlite.Connection): Objeto de conexion a SQLite desde Python.
        data (DataFrame): Tabla de datos obtenida.
        table (str): Nombre de la tabla a la que se le insertaran los datos.
        date (str): Fecha que actua como identificador unico (ID) para la insercion de datos.
        values (list or tuple): Valores a insertar en la tabla.
    Returns:
        (None) Valores ingresados en la tabla.
    """
    cursor = conn.cursor()
    
    if date != data.DATE[0]:
        cols = ', '.join(data.columns)
        placeholders = ', '.join(['?'] * len(data.columns))
        sql = f'INSERT INTO {table} ({cols}) VALUES ({placeholders});'
        cursor.execute(sql, values)
        conn.commit()

def update_vals(conn:sqlite3.connect, table:str, column:str, value:int|float, date:str):
    """
    Descript:
        Actualiza valores de la tabla anteriormente llenada.
    
    Args:
        conn (sqlite.Connection): Objeto de conexion a SQLite desde Python.
        table (str): Nombre de la tabla a la que se le insertaran los datos.
        column (str): Columna en la que insertaran los datos.
        value (int, float): Valore a insertar en la tabla.
        date (str): Fecha que actua como identificador unico (ID) para la insercion de datos.
    Returns:
        (None) Valores actualizados en la tabla.
    """
    cursor = conn.cursor()
    sql = f'UPDATE {table} SET {column} = ? WHERE DATE = ?;'
    cursor.execute(sql, [value, date])
    conn.commit()

def filling_blanks(conn:sqlite3.connect, dw:pd.DataFrame, date:str):
    """
    Descript:
        LLena espacios en blanco en las tablas.
    
    Args:
        conn (sqlite.Connection): Objeto de conexion a SQLite desde Python.
        dw (DataFrame): Tabla con la informacion de depositos y retiros.
        date (str): Fecha que actua como identificador unico (ID) para la insercion de datos.
    Returns:
        (None) Valores en blanco correctamente llenados.
    """
    tables = pd.read_sql_query('SELECT * FROM sqlite_master WHERE type = "table" ORDER BY name;', conn)
    
    for table in tables.name:
        if len(table) > 3:
            data = pd.read_sql_query(f'SELECT * FROM {table} ORDER BY DATE DESC LIMIT 5;', conn)
            dw_acct = dw[dw.ACCT == table[3:]].reset_index(drop = True)
            values = [date] + [0] * (len(data.columns[1:]))
            insert_vals(conn, data, table, date, values)
            if not dw_acct.empty:
                update_vals(conn, table, dw_acct.TYPE[0], dw_acct.AMMOUNT[0], date)

def filling_banks(conn:sqlite3.connect, bnk:str, date:str):
    """
    Descript:
        Llena la tabla al banco correspondiente con los datos de cada cuenta.
    
    Args:
        conn (sqlite.Connection): Objecto de conexion a SQLite desde Python.
        bnk (str): Nombre del banco al cual se debe llenar su tabla.
        date (str): Fecha que actua como identificador unico (ID) para la insercion de datos.
    Returns:
        bnk (str)
    """
    if bnk != 'BNK':
        summary_accts = pd.DataFrame(columns = ['DATE', 'CASH', 'LQDTY', 'EQTS', 'FIX_INC', 'ALTS', 'STR_NTS', 'DRVTS', 'AUMS', 'DEPS', 'WDRS'])
        accts_bank = pd.read_sql_query(f'SELECT * FROM sqlite_master WHERE type = "table" AND NAME LIKE "%{bnk}%" ORDER BY name;', conn)
        loan_names = ['LOAN', 'LOAN_USD', 'LOAN_EUR', 'LOAN_MXN', 'LOAN_UNSEC'] + [f'LOAN_{i+1}' for i in range(5)]
        
        for table in accts_bank.name:
            if len(table) > 3 and table != f'{bnk}LOAN':
                data = pd.read_sql_query(f'SELECT * FROM {table} WHERE DATE = "{date}";', conn)
                summary_accts.loc[len(summary_accts)] = data.loc[0]
        summary_accts = summary_accts.sum()
        
        bank_df = pd.read_sql_query(f'SELECT * FROM {bnk} ORDER BY DATE DESC LIMIT 5;', conn)
        loan = pd.read_sql_query(f'SELECT * FROM {bnk}LOAN ORDER BY DATE DESC LIMIT 5;', conn)
        loan = loan[loan.DATE == date][[l for l in loan.columns if l in loan_names]].sum()
        values = [date] + [float(summary_accts.loc[c]) for c in bank_df.columns[1:8]] + \
            [float(loan[c]) if bnk not in ['DBK', 'JBR'] else float(loan.sum()) for c in bank_df.columns[8:-6]] + \
            [float(summary_accts.loc['AUMS']), float(round(summary_accts.loc['AUMS'] + loan.sum(), 2)), 
             float(summary_accts.loc['DEPS']), float(summary_accts.loc['WDRS']), 'NO', 'NO']
        
        insert_vals(conn, bank_df, bnk, date, values)
        
        return bnk


### ======================================================= ###
###                         Codigo 4                        ###
### ======================================================= ###
sorter     =['NAV',' Santander Suiza','Julius Baer','Citibank','UBS','JP Morgan',
             'Goldman Sachs','Morgan Stanley','Deutsche Bank','Santander Suiza',
             'Bankinter','TD Ameritrade','UNI','CASH','LIQUIDITY','EQUITIES',
             'FIXED_INCOME','ALTERNATIVES','STRUCTURED_NOTES','DERIVATIVES',
             'BANKINTER','LOAN_EUR','LOAN_USD','LOAN_MXN','LOAN','LOAN_UNSEC']

def stacked_areas(path:str, title:str, data:pd.DataFrame, namesList:list[str], yTicks:str|int, xTicks:str|int,
                  colors:dict, annot_size:int = 10, startDate:str = '', labels:bool = False, bnk:bool = False,
                  vline:bool = False, leg_size:float = 10, ncols:int = 8, legLoc:str = 'left', 
                  bbox:tuple = (0.001, 1.1), left:float = 0.053):
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
    
    notSumVals = ['NAV', 'LOAN', 'LOAN_USD', 'LOAN_EUR', 'LOAN_MXN', 'LOAN_UNSEC']
    blackFontVals = ['LOAN', 'LOAN_USD', 'LOAN_EUR', 'LOAN_MXN', 'LOAN_UNSEC', 'BANKINTER', 'STRUCTURED_NOTES', 'CASH', 'TD Ameritrade']
    primAxisCols = ['Julius Baer', 'Citibank', 'UBS', 'Goldman Sachs', 'Morgan Stanley', 'Deutsche Bank', 'Santander Suiza',
                    'TD Ameritrade', 'HCITY', 'BANKINTER', 'CASH', 'LIQUIDITY', 'EQUITIES', 'FIXED_INCOME', 'ALTERNATIVES',
                    'STRUCTURED_NOTES', 'DERIVATES', 'UNI']
    secAixsCols = ['LOAN', 'LOAN_USD', 'LOAN_EUR', 'LOAN_MXN', 'LOAN_UNSEC', 'JP Morgan', ' Santander Suiza']
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
    
    labels_right, color_list, fontc_list = [f'${nav:,.1f}'], ['b'], ['w']
    loan_names = {'LOAN', 'LOAN_USD', 'LOAN_EUR', 'LOAN_MXN', 'LOAN_UNSEC'}
    
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
            ax.fill_between(
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
                              leg_size, ncols, legLoc, bbox, lambda x, _: rf'$\bf\${x/1000000:.0f}M$')
    
    plt.subplots_adjust(left=left, right=1 - margin_right - 0.014, top=0.891, bottom=0.109)
    fig.savefig(f'{path}/{title}.png', dpi = 300)
    plt.close()
    # plt.show()

def portfolioswoHCITY(path:str, title:str, data_port:pd.DataFrame, data_city:pd.DataFrame, xTicks:int, yTicks:int,
                      labels:bool = True):
    """
    Descript:
        Muestra la imagen generada de los portafolios sin considerar HCITY.
    
    Args:
        path (str): Ruta hacia donde se debe guardar el archivo generado.
        data_port (DataFrame): Tabla de datos del portafolio para generar la grafica.
        data_city (DataFrame): Tabla de datos de HCITY para generar la grafica.
        name (str): Nombre del archivo generado (sin extension).
        fdays (int): Numero de dias agregados al rango del eje 'x' para agregar las anotaciones.
        axy (dict): Diccionario que contenga las coordenadas de las etiquetas.
        ydtick (int): Formato de etiqueta en el eje 'y' de la grafica generada (debe ser compatible con Plotly).
    Returns:
        (None) Guarda la imagen generada en la ruta especificada.
    """
    dataWHcity = pd.merge(data_port, data_city, on = 'DATE')
    dataWHcity.columns = data_port.columns.tolist() + ['HCITY', 'HCITY_LOAN', 'LABEL']
    dataWHcity['NAV WITHOUT HCITY'] = dataWHcity.NAV - dataWHcity.HCITY - dataWHcity.HCITY_LOAN
    
    plt.rcParams['axes.xmargin'] = 0
    fig, ax = plt.subplots(figsize=(33/2.54, 15/2.54))
    
    labels_right, colors_list, fontc_list = [], [], []
    confs = {'NAV':{'color':'k', 'val':dataWHcity.NAV.iloc[-1]},
             'HCITY':{'color':'#89C6EE', 'val':dataWHcity.HCITY.iloc[-1]},
             'NAV WITHOUT HCITY':{'color':'b', 'val':dataWHcity['NAV WITHOUT HCITY'].iloc[-1]}}
    confs = dict(sorted(confs.items(), key = lambda item: item[1]['val']))
    
    for key, val in confs.items():
        if key == 'HCITY':
            ax.fill_between(
                dataWHcity.DATE,
                0, dataWHcity[key],
                # color = val,
                color = val['color'],
                label = key,
                lw = 0
            )
        else:
            ax.plot(
                dataWHcity.DATE,
                dataWHcity[key],
                f'o-{val["color"]}',
                lw = 2,
                ms = 3,
                label = key,
            )
        labels_right.append(f'${val["val"]:,.1f}')
        colors_list.append(val['color'])
        fontc_list.append('w' if key != 'HCITY' else 'k')
    
    if labels:
        dataWHcity['updn'] = dataWHcity.LABEL.apply(lambda x: 20 if x == 'UP' else -20)
        dl = dataWHcity[dataWHcity.LABEL != 'NO']
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
    
    ymin = dataWHcity[['HCITY', 'NAV', 'NAV WITHOUT HCITY']].min().min()
    ymax = dataWHcity[['HCITY', 'NAV', 'NAV WITHOUT HCITY']].max().max()
    ax.set_ylim(ymin - 500000, ymax + 500000)
    labels_right, colors_list, fontc_list = labels_right[::-1], colors_list[::-1], fontc_list[::-1]
    
    margin_right = plot_confs(fig, ax, labels_right, 0.05, 0.0525, 0, 9, colors_list, fontc_list, xTicks, yTicks,
                              9, 8, 'left', (0.001, 1.05), yformat = lambda x, _: rf'$\bf\${x/1000000:.0f}M$')
    
    plt.subplots_adjust(left=0.048, right=1 - margin_right - 0.014, top=0.971, bottom=0.109)
    fig.savefig(f'{path}/{title}.png', dpi = 300)
    plt.close()
    # plt.show()

def waterfall(path: str, wfall: pd.DataFrame | list, data_port: pd.DataFrame, x: list[str], name: str, minmax: tuple):
    """
    Muestra la grafica de cascada del portafolio con las contribuciones de cada banco.
    """

    values = np.array(wfall, dtype=float)
    x_pos = np.arange(len(values))

    cumulative = np.zeros(len(values))
    cumulative[1:] = np.cumsum(values[:-1])
    cumulative[-1] = 0
    
    colors = ['#222A35'] + ['#548235' if v >= 0 else '#C00000' for v in values[1:-1]] + ['#222A35']

    plt.rcParams['axes.xmargin'] = 1e-2
    fig, ax = plt.subplots(figsize=(33/2.54, 16.18/2.54))

    ax.bar(
        x_pos,
        values,
        bottom=cumulative,
        color=colors
    )
    
    for i in range(1, len(values)):
        ax.plot([x_pos[i-1], x_pos[i]], [cumulative[i-1] + values[i-1], cumulative[i-1] + values[i-1]], 
                color='gray', linewidth=1)

    for i, v in enumerate(values):
        y = cumulative[i] + v
        ax.text(
            x_pos[i],
            y + 1e3 if v >= 0 else y - 4e3,
            f'${v:,.1f}',
            ha='center',
            va='bottom' if v >= 0 else 'top',
            fontsize=11,
            fontweight='bold'
        )

    ax.set_xticks(x_pos)
    ax.tick_params(axis = 'both', labelsize = 12)
    ax.set_xticklabels(x, fontweight='bold')

    y_min = min(data_port.NAV.iloc[-1], data_port.NAV.iloc[-2]) - minmax[0]
    y_max = max(data_port.NAV.iloc[-1], data_port.NAV.iloc[-2]) + minmax[1]
    ax.set_ylim(y_min, y_max)

    ax.yaxis.set_major_formatter(lambda x, _: rf'$\bf\${x/1000000:.2f}M$')

    ax.grid(axis='y', color='LightGray', linewidth=1)
    for spine in ax.spines.values():
        spine.set_linewidth(1)
        spine.set_color('black')
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.set_axisbelow(True)
    
    plt.subplots_adjust(left=0.069, right=1, top=1, bottom=0.038)
    plt.savefig(f'{path}/{name}.png', dpi = 300)
    plt.show()


### ======================================================= ###
###                         Codigo 5                        ###
### ======================================================= ###
def levaumloannet(data:pd.DataFrame, bdf:pd.DataFrame, bank:str):
    """
    Descript:
        Genera el desempeno del portafolio contando todos los bancos.
    
    Args:
        data (DataFrame): Datos del portafolio a considerar.
        bdf (DataFrame): Datos del banco para calular el desempeno.
        bank (str): Siglas del banco a considerar.
    Returns:
        pd.DataFrame
    """
    cols = ['CASH', 'LQDTY', 'EQTS', 'FIX_INC', 'ALTS', 'STR_NTS', 'DRVTS']
    loan = [col for col in bdf.columns if 'LOAN' in col]
    bdf['LOAN_'] = bdf[loan].sum(axis = 1)
    bdf = bdf[bdf.DATE > '2020-01-01'].reset_index(drop = True)
    
    data = data.reset_index(drop = True)
    data[f'AUM_{bank}'] = bdf[cols].sum(axis = 1)
    data[f'LOAN_{bank}'] = bdf['LOAN_']
    data[f'NET_{bank}'] = bdf.NAV
    data[f'LVRLV_{bank}'] = -data[f'LOAN_{bank}'] / data[f'NET_{bank}']
    
    return data

def un_leveraged(data:pd.DataFrame, path:str, title:str, msrs:tuple[float]|list[float], cols:list[str] = ['NFD', 'SCH', 'Total'], ymax:int=None,
                 unique:bool = False, xTicks:int = 1, yTicks:int = 2, leg_size:int = 10, annot_size:int = 10, flag:bool = True, right:float = 0.952,
                 pct:str = ''):
    plt.rcParams['axes.xmargin'] = 0
    fig, ax = plt.subplots(figsize = (msrs[0]/2.54, msrs[1]/2.54))
    
    low = pd.Series([0]*len(data.DATE))
    labels_right, color_list, fontc_list = [], [], []
    
    for port in cols:
        if port == 'Total' or unique:
            y = data[f'LVRLV_{port}'] if unique else data.LVR_TOTAL_WEIGHTED
            ax.plot(
                data.DATE,
                y,
                'o-b',
                lw = 1.5,
                ms = 2,
                label = f'Leverage {port}'
            )
            labels_right.append(f'{y.iloc[-1]:.2f} X')
            color_list.append('b')
            fontc_list.append('w')
        else:
            ax.fill_between(
                data.DATE,
                low,
                low + data[f'LVR_{port}_W'],
                color = '#8F8F8F' if port == 'NFD' else '#C2B994',
                label = f'Leverage {port}',
                lw = 0
            )
            low += data[f'LVR_{port}_W']
            labels_right.append(f"{data[f'{pct}LVR_{port}_W'].iloc[-1]*100:,.2f}%")
            color_list.append('#8F8F8F' if port == 'NFD' else '#C2B994')
            fontc_list.append('w' if port == 'NFD' else 'k')
    
    if ymax is not None:
        ax.set_ylim(0, ymax)
    
    margin_right = plot_confs(fig, ax, labels_right, 0.13, 0.0525, 0, annot_size, color_list, fontc_list, xTicks, yTicks,
                              leg_size, 8, 'right', (1, 1.07), lambda x, _: rf'$\bf{x:.1f}x$', flag = flag, right = right)
    
    plt.tight_layout(pad = 0.05)
    fig.savefig(f'{path}/{title}.png', dpi = 300)
    # plt.show()


### ======================================================= ###
###                         Codigo 6                        ###
### ======================================================= ###
def Bench60_40(df, serie1, serie2):
    """ Funcion que calcula el rendimiento con base en un DF con fecha y 
    los dos indices que se desea"""
    # Definicion de variables necesarias
    titulosSixty = [60000]
    titulosForty = [40000]
    titulos_tot = [100000]
    rend_sf = [0.0]
    for i in range(1,len(df)):
        # Calculo de titulos
        titulos_1 = (titulosSixty[i-1] / df[serie1][i-1]) * df[serie1][i]
        titulos_2 = (titulosForty[i-1] / df[serie2][i-1]) * df[serie2][i]
        # Suma de titulos 
        titulos_tot.append( titulos_1+titulos_2 )
        # Calculo de titulos que seran 60/40
        titulosSixty.append(titulos_tot[i] * 0.6)
        titulosForty.append(titulos_tot[i] * 0.4)
        # Calculo de rendimientos
        rend_sf.append((titulos_tot[i] / titulos_tot[0]) - 1)
    # Agrega una columna que se llama rendimiento
    return rend_sf

def flows(loan:pd.DataFrame, bank:pd.DataFrame, port:pd.DataFrame):
    """
    Descript:
        Calcula los flujos de cada portafolio.
    
    Args:
        loan (DataFrame): Datos de la deuda del banco.
        bank (DataFrame): Datos del banco a considerar.
        port (DataFrame): Datos del portafolio a llenar.
    Returns:
        pd.DataFrame
    """
    bank = bank.reset_index(drop = True)
    loan = loan.drop('LOAN_FLOW', axis = 1).reset_index(drop = True)
    port = port.reset_index(drop = True)
    
    ll = [col for col in loan.columns if 'OP_LOAN' in col or 'ADV_PAY_LOAN' in col]
    loan_ = loan.drop(ll, axis = 1)
    l = [col for col in loan_.columns if 'LOAN' in col]
    
    port['CASH'] += bank['CASH']
    port['LQDTY'] += bank['LQDTY']
    port['EQTS'] += bank['EQTS']
    port['FIX_INC'] += bank['FIX_INC']
    port['ALTS'] += bank['ALTS']
    port['DRVTS'] += bank['DRVTS']
    port['STR_NTS'] += bank['STR_NTS']
    port['LOAN_TOTAL'] += loan_[l].sum(axis = 1)
    port['CASH_FLOW'] += bank['DEPS'] + bank['WDRS']
    port['LOAN_FLOW'] += loan[ll].sum(axis = 1)
    
    return port

def port_flows(banks:list[str], pathDB:str, port:str, conds:str, condsl:str, hc:pd.DataFrame, fx_:pd.Series, names:list[str], dt:str = '"2023-12-31"', googl:bool = False):
    """
    Descript:
        Calcula los rendimientos de cada portafolio con y sin equities
    
    Args:
        banks (list[str]): Lista de bancos que hay en el portafolio.
        pathDB (str): Ruta hacia la base de datos, ingresata como cadena de texto.
        port (str): Nombre del portafolio.
        conds (str): Condicionales en la clausula WHERE en formato SQL.
        condsl (str): Condicionales en la clausula WHERE en formato SQL para la tabla de creditos.
        hc (DataFrame): Tabla de HCITY para los calculos necesarios.
        fx_ (pd.Series): Series de los tipos de cambio EURUSD.
        names (list[str]): Lista con los nuevos nombres de la tabla a devolver.
    Returns:
        (DataFrame) Tabla de rendiemientos del portafolio.
    """
    pfolio = pd.DataFrame({})    
    l = []
    for idx, bank in enumerate(banks):
        condsll = f'WHERE OP_LOAN_USD IS NOT NULL AND DATE >= {dt}' if port == 'SCH' and bank == 'SSZ' else condsl
        bnk = get_df(pathDB, f'SigCap_{port}.db', bank, where = conds)
        bnkl = get_df(pathDB, f'SigCap_{port}.db', f'{bank}LOAN', where = condsll)
        if idx == 0:
            pfolio['DATE'] = bnk.DATE
            pfolio[['CASH', 'LQDTY', 'EQTS', 'FIX_INC', 'ALTS', 'DRVTS', 'STR_NTS', 'LOAN_TOTAL', 'CASH_FLOW', 'LOAN_FLOW']] = 0
            if port == 'NFD' and googl:
                bnk.loc[bnk.DATE == '2025-09-19', 'DEPS'] = int(241.195 * 1398)
        pfolio = flows(bnkl, bnk, pfolio)
        l.append(bnkl)
    
    if port == 'NFD':
        nts_str = get_df(pathDB, 'SigCap.db', 'NOTAS_NOMINAL_NFD', conds)
        pfolio['STR_NTS'] = nts_str[['STRNTS_CTI_USD', 'STRNTS_JBR_USD', 'STRNTS_SSZ_USD', 'STRNTS_UBS_USD']].sum(axis = 1) + nts_str['STRNTS_JBR_EUR'] * fx_
        pfolio['LOAN_FLOW'] = l[0].OP_LOAN_EUR*fx_ + l[0].ADV_PAY_LOAN_EUR*fx_ + l[0].OP_LOAN_USD + l[0].ADV_PAY_LOAN_USD +\
                              l[1].OP_LOAN + l[1].ADV_PAY_LOAN + l[2].OP_LOAN + l[2].ADV_PAY_LOAN + l[3].OP_LOAN + l[3].ADV_PAY_LOAN
    
    pfolio['AUMS'] = pfolio[['CASH', 'LQDTY', 'EQTS', 'FIX_INC', 'ALTS', 'DRVTS', 'STR_NTS']].sum(axis = 1)
    pfolio['NAV'] = pfolio.AUMS + pfolio.LOAN_TOTAL
    
    pfolio['HOLD_CASH_FLOW'] = pfolio.CASH_FLOW.cumsum()
    pfolio['HOLD_LOAN_FLOW'] = pfolio.LOAN_FLOW.cumsum()
    
    hc['TRANCUM'] = hc.TRANSACTIONS.cumsum()
    
    pfolio['AUM_BS_ZR'] = pfolio.AUMS - pfolio.HOLD_CASH_FLOW - pfolio.HOLD_LOAN_FLOW
    pfolio['NAV_BS_ZR'] = pfolio.NAV - pfolio.HOLD_CASH_FLOW
    
    pfolio['AUM_SH'] = pfolio.AUMS - pfolio.CASH_FLOW - pfolio.LOAN_FLOW
    pfolio['NAV_SH'] = pfolio.NAV - pfolio.CASH_FLOW
    
    if port == 'SCH':
        pfolio['AUM_BS_ZR_WO_EQTS'] = pfolio.AUM_BS_ZR - hc.HCITY_SCH - hc.TRANCUM
        pfolio['NAV_BS_ZR_WO_EQTS'] = pfolio.NAV_BS_ZR - hc.HCITY_SCH - hc.TRANCUM
    else:
        pfolio['AUM_BS_ZR_WO_EQTS'] = pfolio.AUM_BS_ZR - pfolio.EQTS
        pfolio['NAV_BS_ZR_WO_EQTS'] = pfolio.NAV_BS_ZR - pfolio.EQTS
    
    cols_rend, cols_sharp = ['REND_AUM', 'REND_NAV', 'REND_AUM_WO_EQTS', 'REND_NAV_WO_EQTS'], ['REND_AUM_SHARP', 'REND_NAV_SHARP']
    act_rend, act_sharp = ['AUM_BS_ZR', 'NAV_BS_ZR', 'AUM_BS_ZR_WO_EQTS', 'NAV_BS_ZR_WO_EQTS'], ['AUM_SH', 'NAV_SH']
    pfolio[cols_rend] = pfolio[act_rend] / pfolio[act_rend].iloc[0] - 1
    pfolio[cols_sharp] = (pfolio[act_sharp] / pfolio[act_sharp].shift(+1) - 1).fillna(0)
    
    # pfolio = pfolio[['DATE','REND_AUM','REND_NAV','REND_AUM_SHARP','REND_NAV_SHARP','REND_AUM_WO_EQTS','REND_NAV_WO_EQTS']]
    # pfolio.columns = names
    
    return pfolio

def rends_benchs(pathPC:str, file:str):
    benchs = pd.read_excel(f'{pathPC}/{file}.xlsx', sheet_name = 'INDICES', skiprows = 8)
    benchs = benchs.loc[1:].reset_index(drop = True)
    benchs['DATE'] = benchs.DATE.apply(pd.to_datetime)
    benchs[benchs.columns[1:]] = benchs[benchs.columns[1:]].apply(pd.to_numeric)
    
    rends = pd.DataFrame({})
    rends['DATE'] = benchs.DATE

    # ncols = ['SPX', 'ACWI', 'BRK', 'AGG', 'TLT', 'SHV']
    # acols = ['SPX Index', 'MXWO Index', 'BRK/B US Equity', 'AGG US EQUITY', 'TLT US EQUITY', 'SHV US EQUITY']
    ncols = ['SPX', 'ACWI', 'BRK', 'AGG', 'TLT', 'SHV', 'Rusell 2000']
    acols = ['SPX Index', 'MXWO Index', 'BRK/B US Equity', 'AGG US EQUITY', 'TLT US EQUITY', 'SHV US EQUITY', 'RTY Index']
    rends[ncols] = benchs[acols] / benchs[acols].iloc[0] - 1

    rends['SPX/TLT(60/40)'] = Bench60_40(benchs[['DATE', 'SPX Index', 'TLT US EQUITY']], 'SPX Index', 'TLT US EQUITY')
    rends['SPX/AGG(60/40)'] = Bench60_40(benchs[['DATE', 'SPX Index', 'AGG US EQUITY']], 'SPX Index', 'AGG US EQUITY')
    rends['ACWI/TLT(60/40)'] = Bench60_40(benchs[['DATE', 'MXWO Index', 'TLT US EQUITY']], 'MXWO Index', 'TLT US EQUITY')
    rends['ACWI/AGG(60/40)'] = Bench60_40(benchs[['DATE', 'MXWO Index', 'AGG US EQUITY']], 'MXWO Index', 'AGG US EQUITY')
    rends['BRK/TLT(60/40)'] = Bench60_40(benchs[['DATE', 'BRK/B US Equity', 'TLT US EQUITY']], 'BRK/B US Equity', 'TLT US EQUITY')
    rends['BRK/AGG(60/40)'] = Bench60_40(benchs[['DATE', 'BRK/B US Equity', 'AGG US EQUITY']], 'BRK/B US Equity', 'AGG US EQUITY')
    
    return rends, benchs

def ret_data(data:pd.DataFrame, bench:pd.DataFrame, ncols:list[str], acols:list[str], diffs:bool = False, indexes:list[str] = None):
    datac = data.copy()
    
    datac[ncols] = bench[acols] / bench[acols].iloc[0] - 1
    datac = datac[['DATE'] + ncols]
    if diffs:
        datac['differ'] = datac[indexes[0]] - datac[indexes[1]]
        datac['DIF_POS'] = datac.differ.clip(lower = 0)
        datac['DIF_NEG'] = datac.differ.clip(upper = 0)
        datac = datac.drop('differ', axis = 1)
    x_ = datac[datac.columns[1:]].iloc[-1].sort_values(ascending = False)
    
    return datac, x_

def plot_returns(path:str, title:str, data:pd.DataFrame, x:pd.Series, cols_out:list[str], cols_in:list[str],
                  colors:pd.DataFrame, leg_size:int = 9, xTicks:int = 1, yTicks:int = 5, annots:bool = False, 
                  annot_size:int = 9, w:float = 3, sec_y:bool = False, ncols:int = 5, diffs:bool = False,
                  legLoc:str = 'center', bbox:tuple = (0.5, 1.05), top:float = 0.974):
    """
    Descript:
        Grafica los retornos del portafolio (Apalancado y No Apalancado) respecto a los benchmarks.
    
    Args:
        path (str): Ruta hacia donde se debe guardar el archivo generado.
        title (str): Nombre del archivo generado (sin extension).
        data (DataFrame): Datos del portafolio para graficar.
        x (Series): Serie de datos del portafolio con las etiquetas como indices.
        cols_out (list[str]): Lista de columnas que no se van a considerar para la grafica.
        cols_in (list[str]): Lista de columnas que se van a considerar para la grafica.
        colors (DataFrame): Tabla de colores que se le asignara a cada instrumento en la grafica.
        leg_size (int)(default = 12): Tamano de la leyenda dentro de la grafica.
        yTicks (int)(default = 5): Formato de intervalo de separacion en el eje 'y' de la grafica generada (debe ser compatible con Plotly).
        annots (bool)(default = False): Agrega texto extra a las anotaciones, solo se usa en casos especificos.
        annot_size (int)(default = 20): Tamano de fuente de cada anotacion.
        w (int)(default = 3): Ancho de la imagen generada.
        sec_y (bool)(default = False): Si se desea usar el eje 'y' secundario.
        ncols (int)(default = 5): Numero de columnas en la leyenda del grafico.
    
    Returns:
        (None) Grafica los rendimientos.
    """
    plt.rcParams['axes.xmargin'] = 0
    fig, ax = plt.subplots(figsize=(33/2.54, 16.18/2.54))
    ax2 = ax.twinx() if sec_y else None
    
    
    labels_right, color_list, fontc_list = [], [], []
    
    for idx, col in enumerate(x.index):
        if col not in cols_out:
            width = w if col not in cols_in else 6
            star = '★' if col in cols_in else ''
            # star = r'$\star$' if col in cols_in else ''
            
            color = colors.loc[colors.LABEL == col, 'COLOR'].values[0]
            dash = colors.loc[colors.LABEL == col, 'DASH_MLIB'].values[0]
            font_color = colors.loc[colors.LABEL == col, 'FONT_COLOR'].values[0]
            vix = True if col == 'VIX' else False
            note = ',(RHS)' if col == 'VIX' else ''
            
            target_ax = ax2 if (vix) else ax
            
            if diffs:
                if col in ['DIF_NEG', 'DIF_POS']:
                    ax.fill_between(
                        data.DATE,
                        0,
                        data[col] * 100,
                        color = color,
                        lw = 0 #3
                    )
                if data.DIF_POS.iloc[-1] != 0 and col == 'DIF_NEG':
                    continue
            
            target_ax.plot(
                data.DATE,
                data[col] * 100,
                color = color,
                lw = width if col not in ['DIF_NEG', 'DIF_POS'] else 0,
                ls = dash,
                label = col if col not in ['DIF_NEG', 'DIF_POS'] else None
            )
            txt = f'{x.index[idx]}{note}, {x.iloc[idx]*100:,.1f}%{star}' if annots else f'{x.iloc[idx]*100:,.1f}%{star}'
            labels_right.append(txt)
            color_list.append(color)
            fontc_list.append(font_color)
    
    drop_cols = cols_out + ['DATE', 'VIX'] if 'VIX' in list(x.index) else cols_out + ['DATE']
    ymin = data.drop(drop_cols, axis=1).min().min() * 100
    ymax = data.drop(drop_cols, axis=1).max().max() * 100
    ax.set_ylim(ymin - 5, ymax + 5)
    
    margin_right = plot_confs(fig, ax, labels_right, 0.05, 0.0525, 0, annot_size, color_list, fontc_list, xTicks, yTicks,
                              leg_size, ncols, legLoc, bbox, yformat = lambda x, _: rf'$\bf{x:.1f}\%$',
                              ax2 = ax2 if sec_y else None)
    
    plt.subplots_adjust(left=0.055, right=1 - margin_right - (0.059 if sec_y else 0.014), top=top, bottom=0.109)
    fig.savefig(f'{path}/{title}.png', dpi = 300)
    plt.close()
    # plt.show()