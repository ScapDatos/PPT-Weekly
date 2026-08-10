import pptx
import sqlite3
import numpy as np
import pandas as pd
import plotly.express as px
from pptx.util import Pt
from pptx.dml.color import RGBColor
import plotly.graph_objects as go
from plotly.subplots import make_subplots
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
               xTicks:int, yTicks:int, leg_size:int, ncols:int, legLoc:str, bbox:tuple, yformat:str, ax2 = None):
    
    ax.grid(axis = 'y', color = 'lightgray', lw = 2, alpha = 0.5)
    ax.set_facecolor('white')
    ax.tick_params(axis = 'x', rotation = 90)
    ax.tick_params(axis = 'both', labelsize = 10)
    
    ax.legend(
        loc = f'upper {legLoc}',
        bbox_to_anchor = bbox,
        fontsize = leg_size,
        frameon = False,
        ncol = ncols
    )
    
    nlines = [label.count('\n') + 1 for label in labels]
    steps = [n * (y_step - 0.03) if n > 1 else y_step for n in nlines]
    if counter != 0:
        steps[-counter-1:] = [y_step] + [y_step2] * counter
    y_start = 0.5 + (sum(steps) - y_step) / 2
    txts = []
    
    for label, step, fc, bc in zip(labels, steps, fontc_list, color_list):
        txt = ax.text(
            x = 1.07 if ax2 is not None else 1.01,
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
        ax2.tick_params(axis = 'both', labelsize = 10)
        ax2.yaxis.set_major_locator(plt.MultipleLocator(25))
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
###                         Codigo 2                        ###
### ======================================================= ###
def treemap_aclass(data:pd.DataFrame, title:str, fields:list[str], path:str, nav:str = '', colorCol:str = 'ASSET_CLASS',
                   mxd:int = 4, html:bool = True, w:float = 33, h:float = 16.18, ccy:bool = False, dic: dict = None):
    """
    Descript:
        Graficas de profuncidad por portafolio, moneda, notas, etc.
    
    Args:
        data (DataFrame): Pandas DataFrame de donde se sacaran los datos para la grafica.
        title (str): Nombre que se le dara a la grafica una vez guardada, el nombre es sin la extension del tipo de archivo.
        fielda (list[str]): Lista de los campos que se van a graficar, estos definiran la profundida de las grafica.
        path (str): Ruta en donde se guardara la imagen.
        nav (str): Valor del NAV en formato de cadena de texto.
        colorCol (str): Nombre de la columna a la que se le asignaran los colores.
        mxd (int): Nivel maximo de profunidad para el treemap.
        html (bool): Decide si la grafica se guarda en formato html o no.
        w (float): Ancho de la imagen.
        h (float): Alto de la imagen.
        ccy (bool): Especifica si se requieren cambios en otros parametros (Se usa en las graficas de CCY y Sponsors).
        dic (dict): Solo empleado si el valor ccy se fijo en True. Diccionario con los valores y colores a modificar.
    Returns:
        (None) Shows the images and saves them in the format specified.
    """
    
    fields.insert(0, px.Constant('CONSOLIDADO'))
    color_pallette={'(?)':'lightgrey', 'CASH':'#97B7B3', 'LIQUIDITY':'#49615E', 'EQUITIES':'#475161',
                    'FIXED INCOME':'#8F8F8F', 'ALTERNATIVES':'#867B74', 'STRUCTURED NOTES':'#C2B994',
                    'DERIVATIVES':'#89687E', 'LOAN':'#FF7C80', 'LOAN 2':'#FFD1B2', 'LOAN 3':'#EC9092',
                    'LOAN 4':'#D9B2B2', 'SX5E':'#579FDF', 'SPX':'#4DA2E3', 'SPX|RTY':'#B6D9FF',
                    'SX5E|RTY':'#6AB4F0', 'SPX|SX5E':'#1F81E7', 'NKY|RTY|SPX':'#005CBA',
                    'SPX|RTY|SX5E':'#003883', 'SPX|NKY':'#3277A8', 'SPX|NKY|SX5E':'#4E84AD',
                    'SPX|NDX|RTY':'#2C89E6', 'SX5E|NKY':'#2A5DBD', 'HSI':'#AB63FA', 'CREDITO': '#FF7C80'}
    
    fig = px.treemap(data,                                  # La tabla de los datos
                     path = fields,                         # Orden de la profundidad
                     values = 'MKT_VALUE_USD',              # Columna usada para los calculos
                     color = data[colorCol],                # Columna de colores
                     color_discrete_map = color_pallette,   # Colores de acuerdo al valor que tenga la columna
                     maxdepth = mxd,                        # Opcion de profundidad
                     title = '')                            # Titulo del grafico
    
    fig.update_layout(font = dict(size = 25),                          # Tamano de la fuente
                      hoverlabel = {'font_size': 20},                # Tamano de las etiquetas
                      margin = {'l': 0, 'r': 0, 't': 0, 'b': 0})    # Ancho de los margenes de la figura
    
    fig.update_traces(marker_depthfade = True,                                                                   # Efecto de colores con profundidad
                      textposition = 'middle center',                                                            # Posicion del texto
                      marker = {'cornerradius': 5},                                                              # Redondea las esquinas con ese radio
                      texttemplate = '%{label}<br>%{percentEntry:.2%}',                                          # Formato de los titulos: <LABEL> \n <Valor redondeado a 2 decimales \n <> <Porcentaje redondeado a 2 decimales>
                      hovertemplate = '%{label}<br>Market Value: %{value:$,.2f}<br>Weight: %{percentEntry:.2%}', # Lo que muestra al pasar el mouse por encima
                      selector = {'type': 'treemap'})                                                            # Tipo de grafico
    
    # Lista con los nombres y valores
    labels = fig.data[0].labels.tolist()
    values = fig.data[0].values.T.tolist()
    
    ix = 0
    for idx, lab in enumerate(labels):
        if lab == 'CONSOLIDADO':
            ix = idx
    
    # Coloca el texto de los titulos de cada nivel de profundidad
    titles = [f'{lab} | ${val:,.0f} | {(val / max(values)) * 100:,.2f}%' for lab, val in zip(labels, values)]
    if nav != '':
        titles[ix] = nav
    
    # Se actualizan las etiquetas de la figura
    fig.data[0].labels = titles
    
    # Esto es para las graficas de CCY y Sponsors
    if ccy:
        mktvalue_db = data[['SHORT_NAME', 'MKT_VALUE', 'CCY']]
        customdata = []
        for parent in fig.data[0].parents:
            data_aux = mktvalue_db[mktvalue_db.SHORT_NAME == parent.split('/')[-1]]
            if not data_aux.empty:
                if data_aux.CCY.iloc[-1] != 'EUR':
                    customdata.append('${:,.0f}'.format(data_aux.MKT_VALUE.iloc[-1]))
                else:
                    customdata.append('â¬{:,.0f}'.format(data_aux.MKT_VALUE.iloc[-1]))
            else:
                customdata.append(' ')
                
        fig.data[0].customdata = customdata
        fig.data[0].texttemplate = '%{label}<br>%{percentEntry:.2%}<br>%{customdata}'
        
        list_of_colors = []
        for i, id in enumerate(fig.data[0]['ids']):
            if (len(id.split('/')) == 2 )and (id.split('/')[1] in dic.keys()):
                list_of_colors.append(dic[id.split('/')[1]])
            else:
                list_of_colors.append(fig.data[0]['marker']['colors'][i])
        fig.data[0]['marker']['colors'] = list_of_colors
    
    # Se guarda en formato HTML en caso de haber sido especificado
    if html:
        fig.write_html(f'{path}/{title}.html')
    
    # Guarda la figura para la presentacion
    fig.write_image(f'{path}/{title}.png',
                    width = w * 37.795276,  # Ancho de la imagen (tanano_cm * cm_to_pixels)
                    height = h * 37.795276, # Altura de la imagen (tanano_cm * cm_to_pixels)
                    scale = 1,              # Zoom de la imagen
                    # engine = 'orca'         # Motor que genera la imagen
                    )
    
    # Muestra la imagen
    fig.show()


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
# Orden en el que se grafican las columnas
# Los primeros en la lista son los ultimos en dibujarse
sorter_draw=['NAV',' Santander Suiza','Julius Baer','Citibank','UBS','JP Morgan',
             'Goldman Sachs','Morgan Stanley','Deutsche Bank','Santander Suiza',
             'Bankinter','TD Ameritrade','LOAN_EUR','LOAN_USD','LOAN_MXN','LOAN',
             'LOAN_UNSEC','UNI','CASH','LIQUIDITY','EQUITIES','FIXED_INCOME',
             'ALTERNATIVES','STRUCTURED_NOTES','DERIVATIVES','BANKINTER']

#  Orden de las etiquetas
sorter_labels=['NAV',' Santander Suiza','Julius Baer','Citibank','UBS',
               'JP Morgan','Goldman Sachs','Morgan Stanley','Deutsche Bank',
               'Santander Suiza','TD Ameritrade','LOAN_MXN','LOAN_USD','LOAN_EUR',
               'LOAN_UNSEC','LOAN','UNI','CASH','LIQUIDITY','EQUITIES',
               'FIXED_INCOME','ALTERNATIVES','STRUCTURED_NOTES','DERIVATIVES',
               'BANKINTER','Bankinter']
def stacked_areas(path:str, pname:str, data:pd.DataFrame, namesList:list[str], yTicks:str|int, xTicks:str|int,
                  colors:dict, startDate:str = '', fdays:int = 430, delayPixels:int = 0, labels:bool = False,
                  bnk:bool = False, vline:bool = False):
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
    
    df_aux['Pct'] = round(df_aux.Value / df_aux[~df_aux.Name.isin(notSumVals)].Value.sum()*100, 1)
    
    nav = df_aux[df_aux.Name == 'NAV'].Value.values[0]
    
    df_aux = df_aux.sort_values('Name', key = lambda col: col.map(lambda e: sorter_draw.index(e))).reset_index(drop = True)
    
    cols = [col for col in data.columns if col in primAxisCols]
    cols2 = [col for col in data.columns if col in secAixsCols]
    if pname not in ['SCHOLDINGS', 'NFD']:
        data['SUM_PAXIS'] = data[cols].sum(axis = 1)
    else:
        data['SUM_PAXIS'] = data[cols + ['NAV']].sum(axis = 1)
    data['SUM_SAXIS'] = data[cols2].sum(axis = 1)
    
    plot = go.Figure()
    second_axis_names = set(df_aux[df_aux.Name.isin(secAixsCols)].Name)
    for idx, row in df_aux.iterrows():
        if row.Name != 'NAV':
            sg = 'two' if row.Name in second_axis_names else 'one'
            plot.add_trace(go.Scatter(
                name = row.Name,
                x = data.DATE,
                y = data[row.Name],
                stackgroup = sg,
                fillcolor = row.Color,
                mode = 'none',
                fill = 'tonexty'
            ))
    plot.add_trace(go.Scatter(
        name = 'NAV',
        x = data.DATE,
        y = data.NAV,
        mode = 'lines+markers',
        marker = dict(color = '#0000FF', size = 6),
        fillcolor = '#0000FF',
        line = {'width': 4}
    ))
    fix_date = data.DATE.iloc[-1] + pd.Timedelta(days = fdays)
    # limit = fix_date + pd.Timedelta(days = 60)
    
    plot.update_xaxes(
        ticks = 'outside',
        showline = True,
        linewidth = 2,
        linecolor = 'black',
        tickcolor = 'black',
        tickangle = -90
    )
    plot.update_yaxes(
        ticks = 'outside',
        showline = True,
        linewidth = 2,
        linecolor = 'black',
        gridwidth = 1,
        gridcolor = 'LightGray',
        showgrid = True
    )
    
    df_aux = df_aux[(df_aux.Value != 0.0) & (df_aux.Name != 'NAV')]
    df_aux = df_aux.sort_values('Name', key = lambda col: col.map(lambda e: sorter_labels.index(e))).reset_index(drop = True)
    
    loan_names = {'LOAN', 'LOAN_USD', 'LOAN_EUR', 'LOAN_MXN', 'LOAN_UNSEC'}
    count_neg = 0
    glob_val = 0
    ay_vals = []
    for idx, row in df_aux.iterrows():
        if row.Name in loan_names:
            y_coordinate = -(21 * idx) - delayPixels
            text = '${:,.1f}'.format(row.Value)
            count_neg += 1
        else:
            y_coordinate = -(43 * idx) - delayPixels + (21 * count_neg)
            text = f'${row.Value:,.1f}<br>{row.Pct}% - {mapping[row.Name]}' if bnk else f'${row.Value:,.1f}<br>{row.Pct}%' 
            glob_val += row.Value
        ay_vals.append(y_coordinate)
        
        plot.add_annotation(
            x = fix_date,
            y = min(df_aux.Value),
            ax = -15, #-30 en caso de querer la anotacion de glob_vals
            ay = y_coordinate,
            xanchor = 'right',
            yanchor = 'bottom',
            arrowcolor = 'white',
            bgcolor = row.Color,
            text = text,
            font = dict(family = 'Arial', size = 15, color = '#000000' if row.Name in blackFontVals else '#FFFFFF')
        )
    
    plot.add_annotation(x = fix_date, y = min(df_aux.Value), ax = -15, ay = y_coordinate-43, xanchor = "right", yanchor = "bottom", #ax = -30 en caso de querer la anotacion de glob_vals
                        arrowcolor = 'white', bgcolor = '#0000FF', text = f'${nav:,.1f}',
                        font = dict(family = "Arial", size = 15, color = "#ffffff"))
    
    # plot.add_annotation(x = fix_date, y = min(df_aux.Value), ax = -25, ay = sum(ay_vals) / len(ay_vals), xanchor = "right", yanchor = "bottom", arrowcolor = 'white',
    #                     bgcolor = '#000000', text = f'${glob_val:,.1f}', textangle = 90,
    #                     font = dict(family = "Arial", size = 15, color = "#ffffff"))
    
    if labels:
        data['updn'] = data.LABEL.apply(lambda x: -20 if x == 'UP' else 20)
        dl = data[data.LABEL != 'NO']
        for idx, row in dl.iterrows():
            plot.add_annotation(x = row.DATE, y = row.NAV, ax = 20, ay = row.updn, bgcolor = '#FFFFFF',
                                text = f'${row.NAV:,.1f}', font = dict(family = 'Arial', size = 15, color = '#000000'))
    
    if vline:
        lines = data[data.VLINE != 'NO']
        for idx, row in lines.iterrows():
            plot.add_vline(x = row.DATE, line = dict(width = 1, dash = 'dashdot', color = '#187F3D' if row.VLINE == 'YES_G' else '#C00000'))
    
    plot.update_layout(
        plot_bgcolor = '#FFFFFF',
        showlegend = True,
        font = dict(family = 'Arial Black', size = 16, color = '#000000'),
        xaxis_tickformat = '%b-%y',
        yaxis_tickformat = '$~s',
        margin = dict(l = 0, r = 0, t = 20, b = 20),
        # legend = dict(orientation = 'h', yanchor = 'bottom', xanchor = 'left', y = -0.28, 
        #                 font = dict(family = 'Arial', size = 13, color = '#000000')),
        legend = dict(orientation = 'h', yanchor = 'top', xanchor = 'left', y = 1.12, x = 0, # y = 1
                        font = dict(family = 'Arial', size = 13, color = '#000000')),#, entrywidth = 70, entrywidthmode="pixels"),
        xaxis_range = [data['DATE'][0], fix_date], 
        yaxis_range = [min(data['SUM_SAXIS']) - 100000, max(data['SUM_PAXIS'].to_list()+data['NAV'].to_list()) + 1000000],
        yaxis = dict(dtick = yTicks),
        xaxis = dict(dtick = xTicks)
    )
    
    plot.write_image(
        f'{path}/Graficas_PPT/{pname}.png',
        width = 33 * 37.795276,
        height = 15 * 37.795276,
        scale = 1
    )

def portfolioswoHCITY(path:str, data_port:pd.DataFrame, data_city:pd.DataFrame, name:str, fdays:int, axy:dict, ydtick:int):
    """
    Descript:
        Muestra la imagen generada de los portafolios sin considerar HCITY.
    
    Args:
        path (str): Ruta hacia donde se debe guardar el archivo generado.
        data_port (DataFrame): Tabla de datos del portafolio para generar la grafica.
        data_city (DataFrame): Tabla de datos de HCITY para generar la grafica.
        name (str): Nombre del archivo generado (sin extension).
        fdays (int): Numero de dias agregados al rango del eje 'x' para agregar las anotaciones.
        axy (dict): Dicctionario que contenga las coordenadas de las etiquetas.
        ydtick (int): Formato de etiqueta en el eje 'y' de la grafica generada (debe ser compatible con Plotly).
    Returns:
        (None) Guarda la imagen generada en la ruta especificada.
    """
    dataWHcity = pd.merge(data_port, data_city, on = 'DATE')
    dataWHcity.columns = data_port.columns.tolist() + ['HCITY', 'HCITY_LOAN']
    dataWHcity['NAVwoHCITY'] = dataWHcity.NAV - dataWHcity.HCITY - dataWHcity.HCITY_LOAN
    
    fix_date = dataWHcity.DATE.iloc[-1] + pd.Timedelta(days = fdays)
    
    plot = go.Figure()
    plot.add_trace(go.Scatter(
        name = 'NAV WITHOUT HCITY',
        x = dataWHcity.DATE,
        y = dataWHcity.NAVwoHCITY,
        mode = 'lines+markers',
        marker = dict(color = '#0000FF', size = 6),
        fillcolor = '#0000FF',
        line_color = '#0000FF',
        line = {'width': 4}
    ))
    plot.add_trace(go.Scatter(
        name = 'NAV',
        x = dataWHcity.DATE,
        y = dataWHcity.NAV,
        mode = 'lines+markers',
        marker = dict(color = '#000000', size = 6),
        fillcolor = '#FF8C19',
        line_color = '#000000',
        line = {'width': 4}
    ))
    plot.add_trace(go.Scatter(
        name = 'HCITY',
        x = dataWHcity.DATE,
        y = dataWHcity.HCITY,
        stackgroup = 'one',
        fillcolor = 'rgba(59, 160, 227,0.6)',
        line_color = 'rgba(59, 160, 227, 0.6)'
    ))
    
    plot.add_annotation(
        x = dataWHcity.DATE.iloc[-1],
        y = dataWHcity.NAV.iloc[-1],
        ax = axy['N'][0],
        ay = axy['N'][1],
        bgcolor = '#000000',
        text = f'${dataWHcity.NAV.iloc[-1]:,.1f}',
        font = dict(family = 'Arial', size = 15, color = '#FFFFFF')
    )
    plot.add_annotation(x = dataWHcity.DATE.iloc[-1], y = dataWHcity.NAVwoHCITY.iloc[-1], ax = axy['NwoH'][0], ay = axy['NwoH'][1], bgcolor = '#0000FF',
                        text = f'${dataWHcity.NAVwoHCITY.iloc[-1]:,.1f}', font = dict(family = 'Arial', size = 15, color = '#FFFFFF'))
    plot.add_annotation(x = dataWHcity.DATE.iloc[-1], y = dataWHcity.HCITY.iloc[-1], ax = axy['H'][0], ay = axy['H'][1], bgcolor = 'rgba(59, 160, 227, 0.6)',
                        text = f'${dataWHcity.HCITY.iloc[-1]:,.1f}', font = dict(family = 'Arial', size = 15, color = '#000000'))
    plot.update_layout(
        plot_bgcolor = '#FFFFFF',
        xaxis_tickformat = '%b-%y',
        yaxis_tickformat = '$~s',
        font = dict(family = 'Arial Black', size = 16, color = '#000000'),
        autosize = False,
        width = 150*32/2.5,
        height = 150*15/2.5,
        xaxis_range = [dataWHcity.DATE[0], fix_date],
        margin = dict(l = 0, r = 55, t = 0, b = 55),
        showlegend = True,
        # legend = dict(orientation = 'h', yanchor = 'bottom', xanchor = 'left', x = -0.04, y = -0.225,
        #               font = dict(family = 'Arial', size = 14, color = '#000000')),
        legend = dict(orientation = 'h', yanchor = 'top', xanchor = 'left', x = 0.0, y = 1.0,
                      font = dict(family = 'Arial', size = 14, color = '#000000')),
        yaxis = dict(dtick = ydtick),
        xaxis = dict(dtick = 'M2')
    )
    plot.update_xaxes(
        ticks = 'outside',
        showline = True,
        linewidth = 2,
        linecolor = 'black',
        tickcolor = 'black',
        tickangle = -90
    )
    plot.update_yaxes(
        ticks = 'outside',
        showline = True,
        linewidth = 2,
        linecolor = 'black',
        gridwidth = 1,
        gridcolor = 'LightGray',
        showgrid = True
    )
    plot.write_image(
        f'{path}/Graficas_PPT/{name}.png',
        width = 33*37.795276,
        height = 15.57*37.795276,
        scale = 1
    )
    plot.show()

def waterfall(path:str, wfall:pd.DataFrame|list, data_port:pd.DataFrame, x:list[str], name:str, minmax:tuple):
    """
    Descript:
        Muestra la grafica de cascada del portafolio con las contribuciones de cada banco.
    
    Args:
        path (str): Ruta hacia donde se debe guardar el archivo generado.
        wfall (DataFrame): Datos del portafolio con las contribuciones.
        data_port (DataFrame): Tabla de datos del portafolio para generar la grafica.
        x (list[str]): Etiquetas que se colocaran en la grafica (eje x).
        name (str): Nombre del archivo generado (sin extension).
        minmax (tuple): Valores minimo y maximo (respectivamente) que se suman a los extremos de la grafica (sobre el eje y).
    Returns:
        (None) Guarda la imagen generada en la ruta especificada.
    """
    plot = go.Figure()
    plot.add_trace(go.Waterfall(
        orientation = 'v',
        textposition = 'outside',
        measure = ['absolute'] + ['relative']*(len(wfall)-2) + ['total'],
        x = x,
        y = wfall,
        text = [f'${i:,.1f}' for i in wfall],
        decreasing = {'marker': {'color': '#C00000'}},
        increasing = {'marker': {'color': '#548235'}},
        totals = {'marker': {'color': '#222A35'}}
    ))
    plot.update_layout(
        yaxis_tickformat = '$~s',
        yaxis_range = [min(data_port.NAV.iloc[-1], data_port.NAV.iloc[-2]) - minmax[0], max(data_port.NAV.iloc[-1], data_port.NAV.iloc[-2]) + minmax[1]],
        plot_bgcolor = '#FFFFFF',
        font = dict(family = 'Arial Black', size = 16, color = '#000000'),
        margin = dict(l = 0, r = 0, t = 0, b = 0),
        autosize = False
    )
    plot.update_xaxes(
        showline = True,
        linewidth = 2,
        linecolor = 'black',
        tickcolor = 'Black'
    )
    plot.update_yaxes(
        showline = True,
        linewidth = 2,
        linecolor = 'black',
        showgrid = True,
        gridwidth = 1,
        gridcolor = 'LightGray'
    )
    plot.write_image(
        f'{path}/Graficas_PPT/{name}.png',
        width = 33 * 37.37795276,
        height = 16.18 * 37.795276,
        scale = 1
    )
    plot.show()


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
    
    pfolio = pfolio[['DATE','REND_AUM','REND_NAV','REND_AUM_SHARP','REND_NAV_SHARP','REND_AUM_WO_EQTS','REND_NAV_WO_EQTS']]
    pfolio.columns = names
    
    return pfolio

def rends_benchs(pathPC:str, file:str):
    benchs = pd.read_excel(f'{pathPC}/{file}.xlsx', sheet_name = 'INDICES', skiprows = 8)
    benchs = benchs.loc[1:].reset_index(drop = True)
    benchs['DATE'] = benchs.DATE.apply(pd.to_datetime)
    benchs[benchs.columns[1:]] = benchs[benchs.columns[1:]].apply(pd.to_numeric)
    
    rends = pd.DataFrame({})
    rends['DATE'] = benchs.DATE

    ncols = ['SPX', 'ACWI', 'BRK', 'AGG', 'TLT', 'SHV']
    acols = ['SPX Index', 'MXWO Index', 'BRK/B US Equity', 'AGG US EQUITY', 'TLT US EQUITY', 'SHV US EQUITY']
    rends[ncols] = benchs[acols] / benchs[acols].iloc[0] - 1

    rends['SPX/TLT(60/40)'] = Bench60_40(benchs[['DATE', 'SPX Index', 'TLT US EQUITY']], 'SPX Index', 'TLT US EQUITY')
    rends['SPX/AGG(60/40)'] = Bench60_40(benchs[['DATE', 'SPX Index', 'AGG US EQUITY']], 'SPX Index', 'AGG US EQUITY')
    rends['ACWI/TLT(60/40)'] = Bench60_40(benchs[['DATE', 'MXWO Index', 'TLT US EQUITY']], 'MXWO Index', 'TLT US EQUITY')
    rends['ACWI/AGG(60/40)'] = Bench60_40(benchs[['DATE', 'MXWO Index', 'AGG US EQUITY']], 'MXWO Index', 'AGG US EQUITY')
    rends['BRK/TLT(60/40)'] = Bench60_40(benchs[['DATE', 'BRK/B US Equity', 'TLT US EQUITY']], 'BRK/B US Equity', 'TLT US EQUITY')
    rends['BRK/AGG(60/40)'] = Bench60_40(benchs[['DATE', 'BRK/B US Equity', 'AGG US EQUITY']], 'BRK/B US Equity', 'AGG US EQUITY')
    
    return rends, benchs

def plot_returns(path:str, title:str, data:pd.DataFrame, x:pd.Series, cols_out:list[str], cols_in:list[str],
                 colors:pd.DataFrame, delay_pixels:int = 30, days_right:int = 65, leg_size:int = 12, 
                 ytick:int = 5, annots:bool = False, annot_size:int = 20, w = 3, sec_y = False):
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
        delay_pixels (int)(default = 30): Numero de pixeles para ajustar correctamente las anotaciones en vertical.
        days_right (int)(default = 65): Numero de dias agregados al rango del eje 'x' para agregar las anotaciones.
        leg_size (int)(default = 12): Tamano de la leyenda dentro de la grafica.
        ytick (int)(default = 5): Formato de intervalo de separacion en el eje 'y' de la grafica generada (debe ser compatible con Plotly).
        annots (bool)(default = False): Agrega texto extra a las anotaciones, solo se usa en casos especificos.
        annot_size (int)(default = 20): Tamano de fuente de cada anotacion.
        w (int)(default = 3): Ancho de la imagen generada.
        sec_y (bool)(default = False): Si se desea usar el eje 'y' secundario.  
    
    Returns:
        (None) Grafica los rendimientos.
    """
    plot = make_subplots(specs=[[{"secondary_y": sec_y}]])
    
    for idx, col in enumerate(x.index):
        if col not in cols_out:
            width = 6 if col in cols_in else w
            star = ' ★' if col in cols_in + cols_out else ''
            
            color = colors[colors.LABEL == col].COLOR.reset_index(drop = True).iloc[0]
            dash = colors[colors.LABEL == col].DASH.reset_index(drop = True).iloc[0]
            font_color = colors[colors.LABEL == col].FONT_COLOR.reset_index(drop = True).iloc[0]
            y_coordinate = -((annot_size + 8) * idx) + delay_pixels
            vix = True if col == 'VIX' else False
            note = ',(RSH)' if col == 'VIX' else ''
            
            plot.add_trace(go.Scatter(
                name = col,
                x = data.DATE,
                y = data[col] * 100,
                mode = 'lines',
                marker = {'color': color, 'size': 5},
                fillcolor = color,
                line_color = color,
                line = {'width': width, 'dash': dash}
            ), secondary_y = vix)
            txt = f'{x.index[idx]}{note}, {x.iloc[idx]*100:,.1f}%{star}' if annots else f'{x.iloc[idx]*100:,.1f}%{star}'
            plot.add_annotation(
                x = data.DATE.iloc[-1] + pd.Timedelta(days = days_right),
                y = x[1:].min(),
                ax = 0, ay = y_coordinate,
                xanchor = 'right',
                yanchor = 'bottom',
                arrowcolor = 'black',
                bgcolor = color,
                text = txt,
                font = {'family': 'Arial', 'size': annot_size, 'color': font_color}
            )
        else:
            delay_pixels += 28

    plot.update_xaxes(
        ticks = 'outside',
        linewidth = 2,
        linecolor = 'black',
        tickcolor = 'black',
        tickangle = -90
    )
    plot.update_yaxes(
        ticks = 'outside',
        linewidth = 2,
        linecolor = 'black',
        gridwidth = 1,
        gridcolor = 'LightGray',
        showgrid = True,
        # fixedrange = True, ## REMOVER
        secondary_y = False
    )
    if sec_y:
        # plot.update_layout(
        #     plot_bgcolor = '#FFFFFF', showlegend = True, font = {'family': 'Arial Black', 'size': 14, 'color': '#000000'},
        #     xaxis = {'dtick': 'M1'}, xaxis_tickformat = '%b-%y', yaxis = {'dtick': ytick, 'tickformat': '.1f', 'ticksuffix': '%'},
        #     yaxis2 = {'dtick': 25, 'tickformat': '.1f', 'ticksuffix': '%'}, margin = dict(l = 0, r = 0, t = 20, b = 20),
        #     legend = dict(orientation = 'h', yanchor = 'bottom', xanchor = 'left', y = -0.3, font = dict(family = 'Arial', size = leg_size, color = '#000000')),
        #     xaxis_range = [data.DATE.iloc[0], data.DATE.iloc[-1] + pd.Timedelta(days = days_right)],
        #     yaxis_range = [data[data.columns[1:]].min() * 100, data[data.columns[1:]].max() * 100]
        # )
        plot.update_layout(
            plot_bgcolor = '#FFFFFF', showlegend = True, font = {'family': 'Arial Black', 'size': 14, 'color': '#000000'},
            xaxis = {'dtick': 'M1'}, xaxis_tickformat = '%b-%y', yaxis = {'dtick': ytick, 'tickformat': '.1f', 'ticksuffix': '%'},
            yaxis2 = {'dtick': 25, 'tickformat': '.1f', 'ticksuffix': '%'}, margin = dict(l = 0, r = 0, t = 20, b = 20),
            legend = dict(orientation = 'h', yanchor = 'bottom', xanchor = 'center', y = 1.02, x = 0.5, font = dict(family = 'Arial', size = leg_size, color = '#000000')),
            xaxis_range = [data.DATE.iloc[0], data.DATE.iloc[-1] + pd.Timedelta(days = days_right)],
            # yaxis_range = [data[data.columns[1:]].min().min() * 100, data[data.columns[1:]].max().max() * 100]
            yaxis_range = [data[data.columns[1:]].min() * 100, data[data.columns[1:]].max() * 100]
        )
    else:
        # plot.update_layout(
        #     plot_bgcolor = '#FFFFFF', showlegend = True, font = {'family': 'Arial Black', 'size': 14, 'color': '#000000'},
        #     xaxis = {'dtick': 'M1'}, xaxis_tickformat = '%b-%y', yaxis = {'dtick': ytick}, yaxis_tickformat = '.1f', yaxis_ticksuffix = '%',
        #     margin = dict(l = 0, r = 0, t = 20, b = 20),
        #     legend = dict(orientation = 'h', yanchor = 'bottom', xanchor = 'left', y = -0.3, font = dict(family = 'Arial', size = leg_size, color = '#000000')),
        #     xaxis_range = [data.DATE.iloc[0], data.DATE.iloc[-1] + pd.Timedelta(days = days_right)],
        #     yaxis_range = [data[data.columns[1:]].min() * 100, data[data.columns[1:]].max() * 100]
        # )
        plot.update_layout(
            plot_bgcolor = '#FFFFFF', showlegend = True, font = {'family': 'Arial Black', 'size': 14, 'color': '#000000'},
            xaxis = {'dtick': 'M1'}, xaxis_tickformat = '%b-%y', yaxis = {'dtick': ytick}, yaxis_tickformat = '.1f', yaxis_ticksuffix = '%',
            margin = dict(l = 0, r = 0, t = 20, b = 20),
            legend = dict(orientation = 'h', yanchor = 'bottom', xanchor = 'center', y = 1.05, x = 0.5, font = dict(family = 'Arial', size = leg_size, color = '#000000')),
            xaxis_range = [data.DATE.iloc[0], data.DATE.iloc[-1] + pd.Timedelta(days = days_right)],
            yaxis_range = [data[data.columns[1:]].min() * 100, data[data.columns[1:]].max() * 100 + 10]
            # yaxis_range = [data[data.columns[1:]].min().min() * 100, data[data.columns[1:]].max().max() * 100 + 10]
        )
    plot.write_image(f'{path}/Graficas_PPT/{title}.png', width = 33*37.795276, height = 16.18*37.795276, scale = 1)
    plot.show()

def plot_diffs(path:str, title:str, data:pd.DataFrame, colors:pd.DataFrame, xsys:list[tuple], xy:list[int]):
    """
    Descript:
        Grafica las diferencias respecto a los benchmarks.
    
    Args:
        path (str): Ruta hacia donde se debe guardar el archivo generado.
        title (str): Nombre del archivo generado (sin extension).
        data (DataFrame): Datos del portafolio para graficar.
        colors (DataFrame): Tabla de colores que se le asignara a cada instrumento en la grafica.
        xsys (list[tuple]): Lista con las coordenadas de cada etiqueta.
        xy (list[int]): Coordenadas de la flecha, solo en caso de que esté habilitada dentro de las anotaciones.
    Returns:
        (None) Grafica las diferencias.
    """
    cols = ['DIF_POS', 'DIF_NEG']
    clrs = ['#85B998', '#DF7F7F']
    
    plot = go.Figure()
    for c, col in zip(clrs, cols):
        plot.add_trace(go.Scatter(
            name = '',
            x = data.DATE,
            y = data[col] * 100,
            marker = {'color': c, 'size': 5},
            showlegend = False,
            fillcolor = c,
            line_color = c,
            line = {'width': 3},
            fill = 'tozeroy'
        ))
        if data.DIF_POS.iloc[-1] != 0 and col == 'DIF_NEG':
            continue
        plot.add_annotation(
            x = data.DATE.iloc[-1],
            y = data[col].iloc[-1] * 100,
            ax = xy[0], ay = xy[1],
            xanchor = 'left',
            yanchor = 'bottom',
            arrowcolor = 'black',
            bgcolor = c,
            text = f'{data[col].iloc[-1]*100:,.1f}%',
            font = dict(family = 'Arial', size = 20, color = '#000000')
        )
    
    data = data.drop(cols, axis = 1)
    for idx, name in enumerate(data.columns[1:]):
        color = colors[colors['LABEL'] == name].COLOR.reset_index(drop = True).iloc[0]
        dash = colors[colors['LABEL'] == name].DASH.reset_index(drop = True).iloc[0]
        star = ' ★' if name == 'Retorno Efectivo Apalancado' else ''
        plot.add_trace(go.Scatter(
            name = name,
            x = data.DATE,
            y = data[name] * 100,
            mode = 'lines',
            marker = {'color': color, 'size': 5},
            fillcolor = color,
            line_color = color,
            line = {'width': 3, 'dash': dash}
        ))
        plot.add_annotation(
            x = data.DATE.iloc[-1],
            y = data[name].iloc[-1] * 100,
            ax = xsys[idx][0], ay = xsys[idx][1],
            xanchor = 'left',
            yanchor = 'bottom',
            arrowcolor = 'black',
            bgcolor = color,
            text = f'{data[name].iloc[-1]*100:,.1f}%{star}',
            font = dict(family = 'Arial', size = 20, color = '#000000')
        )
    plot.update_xaxes(ticks="outside",showline=True,linewidth=2,linecolor='black',tickcolor="black", tickangle=-90)
    plot.update_yaxes(ticks="outside",showline=True,linewidth=2,linecolor='black',gridwidth=1,gridcolor='LightGray',showgrid=True, )

    # plot.update_layout(plot_bgcolor='#FFFFFF',showlegend=True, font=dict(family='Arial Black',size=14,color='#000000'),
    #                 xaxis=dict(dtick='M1'),xaxis_tickformat='%b-%y',yaxis=dict(dtick=5),yaxis_tickformat='.1f',yaxis_ticksuffix='%',
    #                 margin=dict(l=0,r=0,t=20,b=20),legend=dict(orientation='h',yanchor='bottom',xanchor='left',y=-0.3,font=dict(family='Arial',size=12,color='#000000')),
    #                 xaxis_range=[data.DATE.iloc[0], data.DATE.iloc[-1] + pd.Timedelta(days=50)],
    #                 yaxis_range=[data[data.columns[1:]].min() * 100, data[data.columns[1:]].max() * 100])
    plot.update_layout(plot_bgcolor='#FFFFFF',showlegend=True, font=dict(family='Arial Black',size=14,color='#000000'),
                    xaxis=dict(dtick='M1'),xaxis_tickformat='%b-%y',yaxis=dict(dtick=5),yaxis_tickformat='.1f',yaxis_ticksuffix='%',
                    margin=dict(l=0,r=0,t=20,b=20),legend=dict(orientation='h',yanchor='bottom',xanchor='center',y=1.02,x=0.5,font=dict(family='Arial',size=12,color='#000000')),
                    xaxis_range=[data.DATE.iloc[0], data.DATE.iloc[-1] + pd.Timedelta(days=50)],
                    yaxis_range=[data[data.columns[1:]].min() * 100, data[data.columns[1:]].max() * 100])
    plot.write_image(f'{path}/Graficas_PPT/{title}.png', width = 33*37.795276, height = 16.18*37.795276, scale = 1)
    plot.show()


### ======================================================= ###
###                         Codigo 7                        ###
### ======================================================= ###
# def asset_metrics(data:pd.DataFrame, dataVS:pd.DataFrame, mkt_sum:float):
#     """
#     Descript:
#         Calcula todas las metricas para la creacion del archivo Contribution.
    
#     Args:
#         data (DataFrame): Datos del Asset actual.
#         dataVS (DataFrame): Datos del Asset de la semana inmediata anterior.
#         mkt_sum (float): Suma del Market Value de la semana inmediata anterior.
#     Returns:
#         pd.DataFrame
#     """
#     df_ = pd.DataFrame(columns = ['BANK', 'ACCT', 'ASSET_CLASS', 'TOTAL_COST', 'DESCRIPTION', 'SHORT_NAME',
#                                   'LTV %', 'PRICE_NEW', 'QUANTITY_NEW', 'PRICE_OLD', 'QUANTITY_OLD'])
#     acs = data.ACCT.unique().tolist() + dataVS.ACCT.unique().tolist()
#     res = []
#     # for acc in data.ACCT.unique():
#     for acc in acs:
#         df_aux = dataVS[dataVS.ACCT == acc]
#         result = pd.merge(data[data.ACCT == acc], df_aux[['DESCRIPTION', 'QUANTITY', 'MKT_VALUE_USD', 'PRICE']],
#                           on = 'DESCRIPTION', suffixes = ('_NEW', '_OLD'), how = 'outer') # 15/04/2026 se agrego el param."how"
#         res.append(result)
    
#     df_ = pd.concat(res, axis = 0, ignore_index = True)    
#     df_['LTV $'] = round(df_['LTV %'] * df_.MKT_VALUE_USD_NEW, 5)
#     df_['Return'] = round(df_.PRICE_NEW / df_.PRICE_OLD - 1, 5)
#     df_['Weight'] = round(df_.MKT_VALUE_USD_OLD / mkt_sum, 3)
#     df_['Contribution'] = round(df_.Weight * df_.Return, 5)
#     df_['PPP'] = abs(round(df_.TOTAL_COST / df_.QUANTITY_NEW, 1))
#     df_['Weekly Change $'] = round(df_.MKT_VALUE_USD_NEW - df_.MKT_VALUE_USD_OLD, 2)
#     df_['P&L $'] = round(df_.MKT_VALUE_USD_NEW - abs(df_.TOTAL_COST), 1)
#     df_['P&L %'] = round(df_['P&L $'] / df_.TOTAL_COST, 3)
    
#     df_['P&L %'] = df_['P&L %'].apply(lambda x: 0 if (x == np.inf) or (x == -np.inf) else x)
    
#     df_ = df_.drop('BANK', axis = 1).sort_values('Contribution', ascending = False)
    
#     return df_

def asset_metrics(data:pd.DataFrame, dataVS:pd.DataFrame, mkt_sum:float):
    """
    Descript:
        Calcula todas las metricas para la creacion del archivo Contribution.
    
    Args:
        data (DataFrame): Datos del Asset actual.
        dataVS (DataFrame): Datos del Asset de la semana inmediata anterior.
        mkt_sum (float): Suma del Market Value de la semana inmediata anterior.
    Returns:
        pd.DataFrame
    """
    dataVS_ = dataVS.groupby(['PORTFOLIO', 'BANK', 'ACCT', 'INSTR_ID'], as_index = False)[['QUANTITY', 'MKT_VALUE_USD']].sum()
    data_ = data.groupby(['PORTFOLIO', 'BANK', 'ACCT', 'INSTR_ID'], as_index = False)[['QUANTITY', 'MKT_VALUE_USD']].sum()
    
    df = pd.merge(dataVS_, data_, on = 'INSTR_ID', how = 'outer', suffixes = ('_OLD', '_NEW'))
    for col in ['PORTFOLIO', 'BANK', 'ACCT']:
        df[f'{col}_OLD'] = df[f'{col}_OLD'].combine_first(df[f'{col}_NEW'])
        df[f'{col}_NEW'] = df[f'{col}_NEW'].combine_first(df[f'{col}_OLD'])
    df['MKT_VALUE_USD_NEW'] = df['MKT_VALUE_USD_NEW'].fillna(0)
    df['MKT_VALUE_USD_OLD'] = df['MKT_VALUE_USD_OLD'].fillna(0)
    
    cols = ['INSTR_ID', 'ASSET_CLASS', 'DESCRIPTION', 'SHORT_NAME', 'LTV %']
    lookup = pd.concat([data[cols], dataVS[cols]]).drop_duplicates(subset = 'INSTR_ID')
    df_ = pd.merge(df, lookup, how = 'left', on = 'INSTR_ID').drop(['PORTFOLIO_NEW', 'BANK_NEW', 'ACCT_NEW'], axis = 1)
    df_ = df_.rename(columns = {'PORTFOLIO_OLD': 'PORTFOLIO', 'BANK_OLD':'BANK', 'ACCT_OLD':'ACCT'})
    # print(df_)
    # exit()
    
    df_['LTV $'] = df_['LTV %'] * df_.MKT_VALUE_USD_NEW
    df_['Return'] = df_.MKT_VALUE_USD_NEW / df_.MKT_VALUE_USD_OLD - 1
    df_['Weight'] = df_.MKT_VALUE_USD_OLD / mkt_sum
    df_['Contribution'] = (df_.MKT_VALUE_USD_NEW - df_.MKT_VALUE_USD_OLD) / mkt_sum
    # df_['PPP'] = abs(round(df_.TOTAL_COST / df_.QUANTITY_NEW, 1))
    df_['Weekly Change $'] = df_.MKT_VALUE_USD_NEW - df_.MKT_VALUE_USD_OLD
    # df_['P&L $'] = round(df_.MKT_VALUE_USD_NEW - abs(df_.TOTAL_COST), 1)
    # df_['P&L %'] = round(df_['P&L $'] / df_.TOTAL_COST, 3)
    
    # df_['P&L %'] = df_['P&L %'].apply(lambda x: 0 if (x == np.inf) or (x == -np.inf) else x)
    
    # df_ = df_.drop('BANK', axis = 1).sort_values('Contribution', ascending = False)
    
    return df_.sort_values('Contribution', ascending = False)

def bar_plot(data:pd.DataFrame, path:str, name:str, positive:bool):
    """
    Descript:
        Genera las graficas de contribuciones (positivas y negativas).s
    
    Args:
        data (DataFrame): Datos a tomar en cuenta para el calculo de las contribuciones.
        path (str): Ruta hacia donde se debe guardar el archivo generado.
        positive (bool): Indica si las contribuciones son positivas o negativas (True o False, respectivamente).
    Returns:
        (None) Genera la grafica del top de contribuciones por portafolio.
    """
    # data = data[(data.ASSET_CLASS != 'LIQUIDITY') & (data.ASSET_CLASS != 'COLATERAL') & (data.ASSET_CLASS != 'CASH')]
    data = data[~data.ASSET_CLASS.isin(['LIQUIDITY', 'CASH', 'COLATERAL', 'CREDITO'])].dropna(subset = ['QUANTITY_NEW', 'QUANTITY_OLD'])
    
    if positive:
        df_filtered = data[['SHORT_NAME','Contribution','Weekly Change $']][data['Weekly Change $'] > 0]
    else:
        df_filtered = data[['SHORT_NAME','Contribution','Weekly Change $']][data['Weekly Change $'] < 0]
    
    df_filtered = df_filtered.groupby('SHORT_NAME', as_index = False).sum().sort_values('Weekly Change $', ascending = not positive)
    names = df_filtered.SHORT_NAME.unique()[:10]
    df_filtered = df_filtered[df_filtered.SHORT_NAME.isin(names)]
    
    if len(names) < 10:
        for i in range(len(names), 10):
            df_filtered.loc[len(df_filtered)] = [' ' * i, 0, 0]
    
    df_filtered['Text'] = [f'${wch:,.0f}' if (positive and wch > 0) or (not positive and wch < 0) else ''
                           for wch in df_filtered['Weekly Change $']]
    
    plot = go.Figure()
    plot.add_trace(go.Bar(
        x = df_filtered['Weekly Change $'], y = df_filtered.SHORT_NAME, text = df_filtered.Text, orientation = 'h',
        marker = dict(color = 'rgb(24, 127, 61)' if positive else 'rgb(192, 0, 0)'), textangle = 0
    ))
    plot.update_layout(
        yaxis = dict(zeroline = False, side = 'right' if positive else 'left'), font = dict(size = 11),
        plot_bgcolor = 'white', barmode = 'stack', margin = dict(l = 0, r = 0, t = 0, b = 0)
    )
    plot.update_xaxes(showticklabels = False)
    plot.write_image(
        f'{path}/{name}.png', width = 16 * 37.795276, height = 15 * 37.795276
    )

def process_portfs(data:pd.DataFrame, dataVS:pd.DataFrame, port:str, path:str, name_top:str, name_bot:str):
    """
    Descript:
        Procesa el portafolio con las metricas y las graficas. Ademas de generar el diccionario con los datos para
        la creacion del archivo de Contribution.
    
    Args:
        data (DataFrame): Datos del Asset actual.
        dataVS (DataFrame): Datos del Asset de la semana inmediata anterior.
        port (str): Portafolio con el que se esta trabajando.
        path (str): Ruta hacia donde se debe guardar el archivo generado.
        name_top (str): Nombre de la grafica de los mayores contribuidores.
        name_bot (str): Nombre de la grafica de los menores contribuidores.
    Returns:
        dict
    """
    df = data[(data.PORTFOLIO == port) & (~data.BANK.isin(['BVA', 'BNK', '-']))]
    dfVS = dataVS[(dataVS.PORTFOLIO == port) & (~dataVS.BANK.isin(['BVA', 'BNK', '-']))]
    
    mkt_last = dfVS.MKT_VALUE_USD.sum()
    
    # df = df[(df.ASSET_CLASS != 'LIQUIDITY') & (df.ASSET_CLASS != 'CASH')]
    # dfVS = dfVS[(dfVS.ASSET_CLASS != 'LIQUIDITY') & (dfVS.ASSET_CLASS != 'CASH')]
    
    # cols = ['ACCT', 'SHORT_NAME', 'DESCRIPTION', 'ASSET_CLASS', 'LTV %', 'LTV $', 'QUANTITY_OLD', 'QUANTITY_NEW', 'PRICE_OLD', 'PRICE_NEW',
    #         'MKT_VALUE_USD_NEW', 'Weight', 'PPP', 'TOTAL_COST', 'Weekly Change $', 'Return', 'Contribution', 'P&L %', 'P&L $']
    cols = ['BANK', 'ACCT', 'INSTR_ID', 'SHORT_NAME', 'DESCRIPTION', 'ASSET_CLASS', 'LTV %', 'LTV $', 'QUANTITY_OLD', 'QUANTITY_NEW',
            'MKT_VALUE_USD_NEW', 'MKT_VALUE_USD_OLD', 'Weight', 'Weekly Change $', 'Return', 'Contribution']
    data_ = asset_metrics(df, dfVS, mkt_last)
    data_ = data_[cols]
    
    bar_plot(data_, path, name_top, True)
    bar_plot(data_, path, name_bot, False)
    
    banks = {f'{port}': data_}
    for bank in df.BANK.unique():
        if bank not in ['BVA', 'BNK', '-']:
            df_bank = df[df.BANK == bank].reset_index(drop = True)
            dfVS_bank = dfVS[dfVS.BANK == bank].reset_index(drop = True)
            
            mkt_bank = dfVS_bank.MKT_VALUE_USD.sum()
            data_bank = asset_metrics(df_bank, dfVS_bank, mkt_bank)
            data_bank = data_bank[cols]
            banks[f'{port}{bank}'] = data_bank
    
    return banks


### ======================================================= ###
###                         Codigo 8                        ###
### ======================================================= ###

def vix_plot(data:pd.DataFrame, columns:list[str], path:str, name:str, axys:list[tuple], fdays:int):
    """
    Descript:
        Grafica los niveles del VIX comparando con ciertos indices.
    
    Args:
        data (DataFrame): Datos historicos de VIX.
        columns (list[str]): Lista de columnas a considerar dentro de los datos.
        path (str): Ruta hacia donde se debe guardar el archivo generado.
        axys (list[tuple]): Lista de coordenadas de cada una de las anotaciones.
        fdays (int): Numero de dias agregados al rango del eje 'x' para agregar las anotaciones.
    Returns:
        (None) Grafica de VIX.
    """
    plot = make_subplots(specs = [[{'secondary_y': True}]])
    idxs = ['RTY', 'NDX', 'SPX', 'VIX']
    colors = {
        'RTY': '#867B74',
        'NDX': '#7CAFDD',
        'SPX': '#475161',
        'VIX': '#0000FF',
        'VIXXO': '#0000FF',
        'VIXXO3M': '#FFC000',
        'VIXXO6M': '#000000',
        'Max. Funding Rate': '#8F8F8F'
    }
    
    for idx, col in enumerate(columns):
        secy = True if col in idxs[:-1] else False
        ref = 'y2' if col in idxs[:-1] else None
        if col in idxs[:-1]:
            dash = None
        elif col == 'Max. Funding Rate':
            dash = 'dot'
        else:
            dash = 'dash'
        txt = 'Fund Rate' if col == 'Max. Funding Rate' else col
        val = data[col].iloc[-1] if 'VIX' in col else data[col].iloc[-1] * 100
        per = '%' if col in idxs[:-1] + ['Max. Funding Rate'] else ''
        plot.add_trace(go.Scatter(
            x = data.FECHA,
            y = data[col],
            name = f'{col} Coupons' if col in idxs[:-1] else col,
            mode = 'lines+markers',
            marker = dict(color = colors[col], size = 6),
            fillcolor = colors[col],
            line_color = colors[col],
            line = {'width': 4, 'dash': dash}
        ), secondary_y = secy)
        
        plot.add_annotation(
            x = data.FECHA.iloc[-1] + pd.Timedelta(days = 3),
            y = data[col].iloc[-1],
            yref = ref, ax = axys[idx][0], ay = axys[idx][1],
            xanchor = 'left', yanchor = 'bottom',
            arrowcolor = 'black', bgcolor = colors[col],
            text = f'{val:,.2f}{per} {txt}',
            font = dict(family = 'Arial', size = 13, color = '#FFFFFF')
        )
    plot.update_layout(
        plot_bgcolor = '#FFFFFF', showlegend = True, font = dict(family = 'Arial Black', size = 13, color = '#000000'),
        xaxis = dict(dtick = 'M1'), xaxis_tickformat = '%b-%y', yaxis_tickformat = '.1f%', yaxis2 = {'tickformat': ',.2%'},
        margin = dict(l = 0, r = 0, t = 0, b = 0),
        legend = dict(orientation = 'h', yanchor = 'top', xanchor = 'right', x = 1, y = 1,
                      font = dict(family = 'Arial', size = 13, color = '#000000')),
        xaxis_range = [data.FECHA.iloc[0], data.FECHA.iloc[-1] + pd.Timedelta(days = fdays)]
    )
    plot.update_xaxes(
        ticks = 'outside', linewidth = 2, linecolor = 'black', tickcolor = 'black'
    )
    plot.update_yaxes(
        ticks = 'outside', linewidth = 2, linecolor = 'black', showgrid = True,
        gridwidth = 1, gridcolor = 'LightGray', secondary_y = False
    )
    plot.write_image(
        f'{path}/DATOS/Graficas_PPT/{name}.png', width = 33 * 37.795276, height = 16.18 * 37.795276, scale = 1
    )
    plot.show()


### ======================================================= ###
###                         Codigo 9                        ###
### ======================================================= ###

def slide_look(ppt:pptx.Presentation, num_slide:int, img_path:str, txt:str):
    """
    Descript:
        Inserta las imagenes dentro de la diapositivamente, especificamente las imagenes de profundidad.
    
    Args:
        ppt (pptx.Presentation): Objeto de tipo pptx para escribir sobre la presentacion desde Python.
        num_slide (int): Numero de diapositiva en donde se quiere agregar informacion.
        img_path (str): Ruta a la imagen que se quiere insertar en la diapositiva.
        txt (str): Texto al pie de la diapositiva.
    Returns:
        (None) Imagenes insertadas en la diapositiva.
    """
    lamina_n = ppt.slides[num_slide]
    placeholder = lamina_n.placeholders[13]
    placeholder.insert_picture(img_path)
    placeholder = lamina_n.placeholders[15]
    placeholder.text = txt

def slide_area(path:str, ppt:pptx.Presentation, num_slide:int, df:pd.DataFrame, img_path:str, serie:str = 'NAV',
               sheetname:str = 'x', txt:str = None):
    """
    Descript:
        Inserta cuadros de texto, imagenes y pie de diapositiva dentro de la presentacion.
    
    Args:
        path (str): Ruta hacia donde se debe guardar el archivo generado.
        ppt (pptx.Presentation): Objeto de tipo pptx para escribir sobre la presentacion desde Python.
        num_slide (int): Numero de diapositiva en donde se quiere agregar informacion.
        df (DataFrame): Datos para agregar el aumento o disminucion del portafolio semana a semana.
        img_path (str): Ruta a la imagen que se quiere insertar en la diapositiva.
        serie (str)(default = 'NAV'): Columna a considerar para la comparacion de rendimiento.
        sheetname (str)(default = 'x'): Nombre de la hoja del archivo Contribution para usar los datos.
        txt(str)(Optional): Texto al pie de la diapositiva.
    Returns:
        (None) Informacion insertada dentro de la diapositiva.
    """
    actual = f'${df[serie].iloc[-1]:,.1f}'
    dif1w = f'${df[serie].iloc[-1] - df[serie].iloc[-2]:,.1f}'
    
    lamina_n = ppt.slides[num_slide]
    tabla = lamina_n.shapes[3].table
    
    tabla.cell(0, 1).text = actual
    tabla.cell(1, 1).text = dif1w
    
    tabla.cell(0, 1).text_frame.paragraphs[0].runs[0].font.size = Pt(16)
    tabla.cell(1, 1).text_frame.paragraphs[0].runs[0].font.size = Pt(16)
    
    trend = ''
    
    if df[serie].iloc[-1] - df[serie].iloc[-2] > 0:
        tabla.cell(1, 1).text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(24, 127, 61)
        trend = 'up'
    elif df[serie].iloc[-1] - df[serie].iloc[-2] < 0:
        tabla.cell(1, 1).text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(192, 0, 0)
        trend = 'down'
    else:
        tabla.cell(1, 1).text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(0, 0, 0)
        
    if sheetname != 'x':
        # cont = pd.read_excel(f'{path}/Contributtion/Contributtion.xlsx', sheet_name = sheetname)
        cont = pd.read_excel(f'{path}/Contributtion/Contribution_Weekly.xlsx', sheet_name = sheetname)
        # cont = cont[~cont.ASSET_CLASS.isin(['LIQUIDITY', 'CASH'])]
        cont = cont[~cont.ASSET_CLASS.isin(['LIQUIDITY', 'CASH', 'COLATERAL', 'CREDITO'])].dropna(subset = ['QUANTITY_NEW', 'QUANTITY_OLD'])

        # Coloca el texto del valor actual en las celda 0,1 y 1,1 de la tabla
        if trend == 'up':
            cont = cont.sort_values('Contribution', ascending = False, ignore_index = True)
            tabla2 = lamina_n.shapes[4].table
            if cont.ASSET_CLASS.iloc[0] == 'EQUITIES':
                peso='$'
            else:
                peso=''
            tabla2.cell(0, 0).text = f'▲ {peso}{cont.SHORT_NAME.iloc[0]}'
            tabla2.cell(0,0).text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(24, 127, 61)
            tabla2.cell(0,0).text_frame.paragraphs[0].runs[0].font.size = Pt(14)
        elif trend=='down':
            cont = cont.sort_values('Contribution', ascending = True, ignore_index = True)
            tabla2 = lamina_n.shapes[4].table
            if cont.ASSET_CLASS.iloc[0] == 'EQUITIES':
                peso='$'
            else:
                peso=''
            tabla2.cell(0,0).text = f'▼ {peso}{cont.SHORT_NAME.iloc[0]}'
            tabla2.cell(0,0).text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(192, 0, 0)
            tabla2.cell(0,0).text_frame.paragraphs[0].runs[0].font.size = Pt(14)
        else:
            tabla2 = lamina_n.shapes[4].table
            tabla2.cell(0,0).text = 'Sin cambios'
            tabla2.cell(0,0).text_frame.paragraphs[0].runs[0].font.size = Pt(14)

    placeholder = lamina_n.placeholders[13]
    placeholder.insert_picture(img_path)
    placeholder = lamina_n.placeholders[15]
    placeholder.text = txt