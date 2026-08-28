# trns_daily.py — ETL de transacciones de 7 custodios → Excel homologado
#
# Salida: <fecha>-TRANSACTIONS.xlsx con una hoja por banco, TODAS, y cuatro hojas de
# auditoria que dejan traza de todo lo que el ETL corrige:
#   Excepciones     conceptos sin mapear e instrumentos sin clasificar
#   Filtradas       filas descartadas, con el motivo (dedup de JBR y de JPM)
#   Reversas        cancelaciones marcadas por el banco, con valor antes y despues
#   Signos          filas re-firmadas a la convencion unica, con valor antes y despues
#
# Las correcciones no triviales estan documentadas junto al codigo que las aplica. El
# analisis que las respalda esta en conciliacion.ipynb y verificacion_jpm.ipynb.
import os
import pandas as pd
import pymysql
from warnings import filterwarnings

filterwarnings('ignore')

# ========================= Configuracion =========================
user = 'diego.hdz'
pswd = 'Equity.2024$'
host = '172.16.10.55'        # Solo responde con VPN

path_pc = "C:/Users/emili/OneDrive/0. Nube Asset Mgmt"
base_dir = os.path.dirname(os.path.abspath(__file__))
report_date = pd.Timestamp.today().strftime('%Y-%m-%d')

data = pd.read_excel(os.path.join(base_dir, 'Diccionarios_actualizados_v3.xlsx'),
                     sheet_name = "Asset_class", converters = {'Cuenta': str},
                     engine = "openpyxl")
data = data[['INSTR_ID', 'NEW_ASSET_CLASS', 'NEW_SUB_ASSET_CLASS', 'SHORT_NAME']]

main_cols = ['FECHA_TRADE', 'FECHA_VALOR', 'CONCEPTO', 'TITULOS', 'PRECIO',
             'MONTO_NETO', 'CORRETAJE', 'CCY', 'INSTR_ID', 'NEW_ASSET_CLASS',
             'NEW_SUB_ASSET_CLASS', 'SHORT_NAME']
info_cols = ['DESCRIPCION', 'NOMBRE_VALOR', 'TIPO_VALOR', 'REFERENCIA', 'ACCT']
cols_final = main_cols + info_cols

dict_ids = set(data['INSTR_ID'].astype(str))

def best_instr_id(*candidatos):
    """Primer candidato que exista en el diccionario; el ultimo es el fallback."""
    best = candidatos[-1].astype(str)
    for cand in reversed(candidatos[:-1]):
        cand = cand.astype(str)
        best = cand.where(cand.isin(dict_ids), best)
    return best

desc_orig = {}

def find_exceptions(df, banco):
    sin_concepto, sin_clase = df.CONCEPTO.isna(), df.NEW_ASSET_CLASS.isna()
    exc = df[sin_concepto | sin_clase].copy()
    motivos = ['; '.join(m for m, flag in [('Concepto sin mapear', c),
                                           ('Instrumento sin clasificar', s)] if flag)
               for c, s in zip(sin_concepto[exc.index], sin_clase[exc.index])]
    exc.insert(0, 'BANCO', banco)
    exc.insert(1, 'MOTIVO', motivos)
    exc.insert(2, 'DESCRIPCION_ORIGINAL', desc_orig[banco][exc.index])
    return exc

# ========================= Conexiones =========================
ct_conn  = pymysql.connect(host=host, user=user, password=pswd, database='CT')
jb_conn  = pymysql.connect(host=host, user=user, password=pswd, database='JB')
jpm_conn = pymysql.connect(host=host, user=user, password=pswd, database='JPM')
ssz_conn = pymysql.connect(host=host, user=user, password=pswd, database='SSZ')
ubs_conn = pymysql.connect(host=host, user=user, password=pswd, database='UBS')
msy_conn = pymysql.connect(host=host, user=user, password=pswd, database='MSY')
dbk_conn = pymysql.connect(host=host, user=user, password=pswd, database='DBK')

### ========================== JB ========================== ###

def insert_val(descript):
    """Sufijo de moneda para los prestamos ANTICIPO A."""
    if pd.isna(descript) or 'ANTICIPO A' not in str(descript):
        return ''
    for ccy, sufijo in (('USD', '-1'), ('MXN', '-2'), ('EUR', '-3')):
        if ccy in descript:
            return sufijo
    return ''

jb_query = """
    SELECT T.BUCH_DAT, T.VAL_DAT, C.DESCRIPTION,
           T.TIT_KURS, T.UMSATZ, T.KTO_WRG_ISO_CD, TIT_KURS_WRG_ISO_CD,
           T.COURTAGE, T.ISIN,
           (SELECT P.TITEL FROM JB.POSITION P
            WHERE P.ANTEIL_DEAL_NO = T.TR_REFNR LIMIT 1) AS TITEL,
           T.TRANS_TEXT_1, T.TRANS_TEXT_2, T.TRANS_TEXT_3,
           T.KTO_ART, T.ZRNR, T.VAL_KURZ_TEXT, T.TR_REFNR, T.STORNO_CD
    FROM JB.TRANSACTIONS T
    INNER JOIN JB.TRX_TYPES_CODES C ON C.GA_CODE = T.GA_CD
    ORDER BY BUCH_DAT DESC
"""

jb_map = {'Foreign option issue convertible': 'Asset Purchased',
          'Notes convertible': 'Asset Purchased',
          'Payment / distribution of capital profit': 'Interest',
          'Coupons from sage deposit items': 'Interest',
          'Call': 'Asset Sold',
          'Fixed-term loan (principal)': 'Loan Payment',
          'Interest on fixed term loan': 'Interest',
          'Account transfer': 'Cash Withdrawal',
          'Reimbursement money market': 'Asset Sold',
          'Fiduciary on sight': 'Asset Purchased',
          'Foreign exchange spot': 'FX Spot',
          'Swift entry': 'Cash Deposit',
          'Swift outgoing (telex, letter)': 'Cash Withdrawal',
          'Commissions': 'Fees Paid - Asset Account',
          'Payment of interest': 'Interest',
          'Other fees': 'Fees Paid - Asset Account',
          'Closing interest': 'Interest',
          'Transaction fees': 'Fees Paid - Asset Account'}

jb_data = pd.read_sql(jb_query, con = jb_conn)
jb_data['CCY'] = [str(isin) if isin.strip() != '' else str(titel) for isin, titel in zip(jb_data.KTO_WRG_ISO_CD, jb_data.TIT_KURS_WRG_ISO_CD)]
jb_data['ISIN'] = jb_data.ISIN.fillna('')
jb_data['TITEL'] = jb_data.TITEL.fillna(100000)
jb_data['INSTR_ID'] = [str(isin) if isin.strip() != '' else str(int(titel)) for isin, titel in zip(jb_data.ISIN, jb_data.TITEL)]
jb_data['DESCRIPT'] = jb_data.TRANS_TEXT_1.astype(str) + ' ' + jb_data.TRANS_TEXT_2.astype(str) + ' ' + jb_data.TRANS_TEXT_3.astype(str)

jb_ccy  = pd.Series([insert_val(d) for d in jb_data.DESCRIPT], index = jb_data.index)
jb_zrnr = '-' + jb_data.ZRNR.astype('Int64').astype(str).str[-4:]
jb_data['INSTR_ID'] = best_instr_id(jb_data.INSTR_ID + jb_zrnr + jb_ccy,
                                    jb_data.INSTR_ID + jb_ccy,
                                    jb_data.INSTR_ID + jb_zrnr,
                                    jb_data.INSTR_ID)

filtradas = []
reversas  = []
signos    = []

# Convencion unica de signo, para que sumar entre bancos signifique algo.
# MONTO_NETO se expresa como FLUJO DE EFECTIVO: negativo = sale dinero de la cuenta.
# TITULOS se expresa como MOVIMIENTO DE POSICION: negativo = salen titulos.
SIGNO_FLUJO = {
    'Asset Sold': +1, 'Cash Deposit': +1, 'Dividends': +1, 'Interest': +1,
    'Distribution': +1, 'Return of Capital': +1,
    'Asset Purchased': -1, 'Cash Withdrawal': -1, 'Fees Paid - Asset Account': -1,
    'Capital Call': -1, 'Fund Expense': -1, 'Loan Payment': -1,
}
SIGNO_TITULOS = {
    'Asset Purchased': +1, 'Transfer Shares In': +1,
    'Asset Sold': -1, 'Transfer Shares Out': -1,
}

def normalizar_signo(df, banco, monto = False, titulos = False, excluir = None):
    """Reexpresa MONTO_NETO y/o TITULOS en la convencion unica de arriba.

    SOLO se aplica a los bancos cuya columna de origen viene SIN signo, donde el signo
    actual no lleva informacion y forzarlo no destruye nada. Verificado contra la base:

      JBR  UMSATZ        0 negativos de 850     -> monto sin signo
      UBS  TRADE_AMT     0 negativos de 356     -> monto sin signo
      MSY  TOTAL_AMT     2 negativos de 1,365   -> monto sin signo en la practica
      MSY  QUANTITY      0 negativos            -> titulos sin signo
      DBK  QUANTITY      0 negativos            -> titulos sin signo (el monto ya lo
                                                   firma el ETL con SIGN3, no se toca)

    NO se aplica a CT (44% de netbase negativo), JPM (24%) ni SSZ (19%): ahi el signo si
    distingue entradas de salidas y sobreescribirlo convertiria reembolsos en cargos.

    `excluir` deja fuera las filas cuyo signo YA es intencional -- es decir, las reversas
    que se van a tratar con accion 'registrar' (JPM, DBK, JBR), donde el banco ya mando el
    signo contrario a proposito.

    OJO con el orden: las reversas de accion 'invertir' (MSY) NO se excluyen. Vienen con el
    signo del original, que es justamente lo que hay que normalizar primero; recien despues
    se invierten. Al reves da el resultado equivocado: en el T-bill 912797RV1, excluirla
    dejaba -150,000 titulos en vez de -50,000.
    """
    if excluir is None:
        excluir = pd.Series(False, index = df.index)
    excluir = pd.Series(excluir, index = df.index).fillna(False).astype(bool)

    for campo, mapa, activo in (('MONTO_NETO', SIGNO_FLUJO,   monto),
                                ('TITULOS',    SIGNO_TITULOS, titulos)):
        if not activo:
            continue
        esperado = df.CONCEPTO.map(mapa)
        valor    = pd.to_numeric(df[campo], errors = 'coerce')
        # Solo se tocan las filas con regla, con valor y cuyo signo actual difiere.
        cambia = (esperado.notna() & valor.notna() & (valor != 0) & ~excluir
                  & ((valor > 0) != (esperado > 0)))
        if cambia.any():
            traza = df.loc[cambia, ['FECHA_TRADE', 'CONCEPTO', 'INSTR_ID', 'ACCT']].copy()
            traza.insert(0, 'BANCO', banco)
            traza.insert(1, 'CAMPO', campo)
            traza['VALOR_ANTES']   = valor[cambia].values
            traza['VALOR_DESPUES'] = (valor[cambia].abs() * esperado[cambia]).values
            signos.append(traza)
            df.loc[cambia, campo] = valor[cambia].abs() * esperado[cambia]
    return df

def descartar(df, mascara, motivo, banco = 'JBR'):
    """Quita las filas marcadas y las guarda en la hoja Filtradas con su motivo.

    El banco va como parametro porque ya no solo lo usa JBR: el dedup del reporte diario
    de JPM manda aqui sus ~3,000 descartes."""
    mascara = pd.Series(mascara, index = df.index).fillna(False).astype(bool)
    fuera = df[mascara].copy()
    if len(fuera):
        fuera.insert(0, 'BANCO', banco)
        fuera.insert(1, 'MOTIVO_DESCARTE', motivo)
        filtradas.append(fuera)
    return df[~mascara]

def marcar_reversa(df, mascara, banco, columna, motivo, accion = 'invertir'):
    """Registra las cancelaciones en la hoja Reversas y aplica la correccion del banco.

    Los 7 custodios entregan las cancelaciones de forma distinta, asi que la accion NO puede
    ser la misma para todos. Se verifico fila por fila contra la tabla origen antes de elegir
    cada una (ver Documentacion_conciliacion.html):

      'invertir'  La fila anula a otra, pero el banco la manda con el signo del ORIGINAL, no
                  con el contrario. Hay que voltearlo para que la suma neta cuadre.
                  Caso MSY: 3 filas de +50,000 donde la tercera era una cancelacion.

      'registrar' La suma neta YA es correcta y tocarla la romperia; solo se deja traza.
                  Caso JPM: la reversa viene con el signo opuesto de fabrica (INTEREST -0.17
                  contra su original +0.17, emparejados por REVERSAL_REC_ENTRY_NUM).
                  Caso DBK: el banco manda el par completo, una compra '+' y una venta '-'
                  con el mismo REF_NBR, que ya se cancelan entre si.
                  Caso JBR: STORNO_CD con valores 'A'/'S'/'1' sin semantica confirmada; se
                  deja traza sin tocar los numeros hasta poder validarlo contra un extracto.

      'descartar' La columna marca la fila CANCELADA, no la que cancela, y el banco vuelve a
                  registrar la operacion aparte. Invertirla daria neto cero en vez de la
                  operacion real, asi que hay que quitarla.
                  Caso SSZ: compra marcada por 1,160,863.75 re-registrada como 1,160,873.42.
    """
    mascara = pd.Series(mascara, index = df.index).fillna(False).astype(bool)
    if not mascara.any():
        return df

    traza = df[mascara].copy()
    traza.insert(0, 'BANCO', banco)
    traza.insert(1, 'ACCION', accion)
    traza.insert(2, 'MARCADA_POR', columna)
    traza.insert(3, 'MOTIVO', motivo)
    # Se guardan los valores tal como los mando el banco, antes de corregir, para poder
    # comparar contra lo que quedo en la hoja del custodio.
    traza = traza.rename(columns = {'TITULOS': 'TITULOS_ORIGINAL',
                                    'MONTO_NETO': 'MONTO_ORIGINAL'})
    if accion == 'invertir':
        traza['TITULOS_CORREGIDO'] = -traza.TITULOS_ORIGINAL
        traza['MONTO_CORREGIDO']   = -traza.MONTO_ORIGINAL
    elif accion == 'descartar':
        traza['TITULOS_CORREGIDO'] = 0
        traza['MONTO_CORREGIDO']   = 0
    else:                                        # 'registrar': no se toca nada
        traza['TITULOS_CORREGIDO'] = traza.TITULOS_ORIGINAL
        traza['MONTO_CORREGIDO']   = traza.MONTO_ORIGINAL
    reversas.append(traza)

    if accion == 'invertir':
        df.loc[mascara, ['TITULOS', 'MONTO_NETO']] *= -1
        df.loc[mascara, 'DESCRIPCION'] = df.loc[mascara, 'DESCRIPCION'].astype(str) + ' (CANCELADA)'
    elif accion == 'descartar':
        df = df[~mascara]
    elif accion == 'registrar':
        df.loc[mascara, 'DESCRIPCION'] = df.loc[mascara, 'DESCRIPCION'].astype(str) + ' (REVERSA)'
    else:
        raise ValueError(f"accion desconocida: {accion!r}")
    return df

jb_data = descartar(jb_data,
                    (jb_data['KTO_ART'] == 2120) & jb_data['DESCRIPTION'].isin(
                        ['Reimbursement money market', 'Fiduciary on sight',
                         'Fixed-term loan (principal)']),
                    'Copia por partida doble (KTO_ART 2120); la gemela se conserva')
jb_data = descartar(jb_data,
                    (jb_data['ZRNR'] == 3176543) & (jb_data['DESCRIPTION'] == 'Account transfer'),
                    'Account transfer repetido en la cuenta 3176543; la gemela se conserva')

jb_full = pd.merge(jb_data, data, how = 'left', on = 'INSTR_ID').rename(columns = {
    'BUCH_DAT': 'FECHA_TRADE', 'VAL_DAT': 'FECHA_VALOR', 'DESCRIPTION': 'CONCEPTO',
    'TIT_KURS': 'PRECIO', 'UMSATZ': 'MONTO_NETO', 'COURTAGE': 'CORRETAJE',
    'VAL_KURZ_TEXT': 'NOMBRE_VALOR', 'TR_REFNR': 'REFERENCIA', 'ZRNR': 'ACCT'})
jb_full['TITULOS'] = 0
jb_full['TIPO_VALOR'] = None
jb_full['DESCRIPCION'] = jb_full.CONCEPTO
desc_orig['JBR'] = jb_full.CONCEPTO.copy()
jb_full['CONCEPTO'] = jb_full.CONCEPTO.map(jb_map)

# STORNO_CD trae 'A'(2), 'S'(2) y '1'(4) sobre 850 filas, todas de la cuenta 3176543 y del
# mismo instrumento (BSKT/BBVA 27). 'Storno' es reversa en aleman, pero la semantica de cada
# valor no esta confirmada y las de KTO_ART 2120 ya las quita el descarte de arriba. Se deja
# traza sin tocar los numeros hasta poder validarlo contra un extracto de JBR.
jb_rev = jb_full.STORNO_CD.astype(str).str.strip().isin(['1', 'A', 'S']) if 'STORNO_CD' in jb_full \
         else pd.Series(False, index = jb_full.index)
jb_full = jb_full[cols_final]
# Sin `excluir`: las filas de STORNO_CD se registran solo porque su semantica no esta
# confirmada, no porque su signo sea intencional -- UMSATZ viene sin signo tambien en ellas.
# Excluirlas las dejaba positivas contra el resto en negativo.
jb_full = normalizar_signo(jb_full, 'JBR', monto = True)
jb_full = marcar_reversa(jb_full, jb_rev, 'JBR', 'STORNO_CD',
                         "STORNO_CD en ('1','A','S') -- SIN CONFIRMAR, solo traza",
                         'registrar')

### ========================== CT ========================== ###

ct_query = """
    SELECT T.ACCOUNT, T.TRADE_DATE, T.SETTLEMENT_DATE,
           D.DESCRIPT, T.SHARES, T.PRICE, T.NETBASE,
           T.TRADE_CURRENCY, T.SECURITY,
           T.transaction_id AS REFERENCIA
    FROM TRANSACTIONS T
    LEFT JOIN TRX_CODES D ON D.CODES = T.TRANSACTION_CODE
    ORDER BY T.TRADE_DATE DESC
"""

ct_data = pd.read_sql(ct_query, con = ct_conn)
ct_full = pd.merge(ct_data, data, how = 'left', left_on = 'SECURITY', right_on = 'INSTR_ID').drop('SECURITY', axis = 1)
ct_full = ct_full[ct_full.DESCRIPT != 'Change in Interest Rate on Loans']

# Dedup Loan Payment / Cash Withdrawal (cada pago aparece dos veces)
lp_keys = ct_full.loc[ct_full.DESCRIPT == 'Loan Payment', 'NETBASE']
cw_mask = (ct_full.DESCRIPT == 'Cash Withdrawal') & ct_full.NETBASE.isin(lp_keys.values)
ct_full = ct_full[~cw_mask]

ct_full['CORRETAJE'] = 0
ct_full = ct_full.rename(columns = {
    'ACCOUNT': 'ACCT', 'TRADE_DATE': 'FECHA_TRADE', 'SETTLEMENT_DATE': 'FECHA_VALOR',
    'DESCRIPT': 'CONCEPTO', 'SHARES': 'TITULOS', 'PRICE': 'PRECIO',
    'NETBASE': 'MONTO_NETO', 'TRADE_CURRENCY': 'CCY'})

# Respaldo desde los CSVs del portal Citi por si CT.TRANSACTIONS vuelve a quedar vacia.
# Estuvo vacia mucho tiempo y esta rama era la unica que corria; desde 2026-08-27 la tabla
# ya trae datos (265 filas, 2025-01-13 a 2026-08-19) y el flujo normal es el de la base.
if ct_full.empty:
    import glob
    ct_csvs = sorted(glob.glob(os.path.join(base_dir, 'Transactions_*.csv')))
    if ct_csvs:
        ct_portal = pd.concat([pd.read_csv(f) for f in ct_csvs], ignore_index = True)
        ct_portal['Account Number'] = ct_portal['Account Number'].str.replace(
            r'=T\("(.*)"\)', r'\1', regex = True)

        def _parse_monto(s):
            if pd.isna(s): return 0.0
            s = str(s).replace(',', '')
            if s.startswith('(') and s.endswith(')'):
                return -float(s[1:-1])
            return float(s)

        ct_portal['Amount (Reporting CCY)'] = ct_portal['Amount (Reporting CCY)'].apply(_parse_monto)
        ct_portal['Quantity'] = ct_portal['Quantity'].apply(_parse_monto)
        ct_portal['Type'] = ct_portal['Type'].str.strip().replace({
            'Fees paid - Asset account': 'Fees Paid - Asset Account'})
        ct_portal = ct_portal[~ct_portal['Type'].isin(
            ['Change in Interest rate on Loans', 'Other'])]

        # Dedup Loan Payment / Cash Withdrawal en portal
        lp = ct_portal[ct_portal['Type'] == 'Loan Payment'][
            ['Date Range', 'Amount (Reporting CCY)']].copy()
        lp_key = lp['Date Range'].astype(str) + '|' + lp['Amount (Reporting CCY)'].astype(str)
        cw = ct_portal[ct_portal['Type'] == 'Cash Withdrawal'].copy()
        cw_key = cw['Date Range'].astype(str) + '|' + cw['Amount (Reporting CCY)'].astype(str)
        ct_portal = ct_portal.drop(cw[cw_key.isin(lp_key.values)].index)

        ct_portal['SECURITY'] = ct_portal['CUSIP'].fillna('CT-EFECTIVO')
        ct_full = pd.merge(ct_portal, data, how = 'left',
                           left_on = 'SECURITY', right_on = 'INSTR_ID')
        ct_full = ct_full.rename(columns = {
            'Account Number': 'ACCT', 'Date Range': 'FECHA_TRADE',
            'Value Date': 'FECHA_VALOR', 'Type': 'CONCEPTO',
            'Quantity': 'TITULOS', 'Amount (Reporting CCY)': 'MONTO_NETO',
            'Nominal CCY': 'CCY', 'Security Description': 'NOMBRE_VALOR',
            'Transaction Ref # or Deal ID': 'REFERENCIA'})
        ct_full['PRECIO'] = 0
        ct_full['CORRETAJE'] = 0
        ct_full['DESCRIPCION'] = ct_full['Description']
        ct_full['TIPO_VALOR'] = ct_full['Asset Class']
        print(f'  CT portal: {len(ct_csvs)} CSVs, {len(ct_full)} filas tras dedup')
    else:
        ct_full = pd.read_excel(os.path.join(base_dir, '2025-10-03-TRANSACTIONS.xlsx'),
                                sheet_name = 'CT')[['ACCT'] + main_cols]
        ct_full['DESCRIPCION'] = ct_full.CONCEPTO
        ct_full[['NOMBRE_VALOR', 'TIPO_VALOR', 'REFERENCIA']] = None

# --- CT: completar columnas que la rama de base de datos no produce ---
# ct_query no trae DESCRIPCION, NOMBRE_VALOR ni TIPO_VALOR (no existen en TRANSACTIONS).
# El fallo estuvo latente mientras CT.TRANSACTIONS estaba vacia y el flujo caia siempre al
# respaldo de CSVs del portal, que si las arma. Al empezar a llegar datos (265 filas,
# 2025-01-13 a 2026-08-19) el corte a cols_final tronaba con KeyError.
# Se rellenan solo si faltan, para no pisar lo que ya construye la rama del portal.
if 'DESCRIPCION' not in ct_full:
    ct_full['DESCRIPCION'] = ct_full.CONCEPTO
for _col in ('NOMBRE_VALOR', 'TIPO_VALOR'):
    if _col not in ct_full:
        ct_full[_col] = None

ct_full = ct_full[cols_final]
desc_orig['CT'] = ct_full.CONCEPTO.copy()

### ========================== JPM ========================== ###

jpm_query = """
    SELECT T.DATE_OF_TRANSACTION, T.SETTLEMENT_DATE,
           T.TYPE_OF_TRANSACTION, T.UNITS,
           T.NET_AMNT_BASE, T.TRANSACTION_CCY,
           T.TRANS_DESC_1, T.CUSIP, T.ACCT_NUM,
           T.MAJOR_ASSET_CLASS_CODE, T.RECORD_ENTRY_NUM,
           T.TRANS_DESC_2, T.TRANS_DESC_3, T.TRANS_DESC_4,
           T.TRANSACTION_DESC_5, T.TRANSACTION_DESC_6,
           T.PRICE_PER_SHARE, T.COMMISSION_AMNT_BASE,
           T.REVERSAL_INDICATOR
    FROM TRANSACTIONS T
    ORDER BY T.SETTLEMENT_DATE DESC
"""

jpm_map = {
    'FOREIGN DIVIDEND': 'Dividends',
    'PURCHASE': 'Asset Purchased',
    'CORPORATE INTEREST': 'Interest',
    'SALE': 'Asset Sold',
    'MISC. DISBURSEMENT': 'Fees Paid - Asset Account',
    'INTEREST': 'Interest',
    'ADJUST CRYVL': 'Change in Interest Rate on Loans',
    'MISC. INCOME': 'Interest',
    'ATM TRANSFER': 'Cash Deposit',
    'DERV CSHFLOW': 'Asset Purchased',
    'MISC. RECEIPT': 'Cash Deposit',
    'PURCHASE OPT': 'Asset Purchased',
    'SELL OPTION': 'Asset Sold',
    'SPOT FX': 'FX Spot',
    'COMMISSIONS': 'Fees Paid - Asset Account',
    'REDEMPTION': 'Asset Sold',
    'RECEIPT OF ASSETS': 'Transfer Shares In',
    'FOREIGN INTEREST': 'Interest',
    'FATCA TX WH': 'Fees Paid - Asset Account',
    'MEMO': 'Fees Paid - Asset Account',
    'FREE DELIVERY': 'Transfer Shares Out',
    'DIVIDEND': 'Dividends',
    'STOCK DIVIDEND': 'Dividends',
    'PRINCIPAL PAYMENT': 'Loan Payment',
    'WRITE OPTION': 'Asset Sold',
    'EXPIRED OPT': 'Asset Sold'
}

fvals = {'ERICSSON (LM) TEL-SP ADR': 'ERICSSON0XX1',
         'JPM LI-LIQ LVNAV FD - USD - W - ACC  ISIN LU1873131988 SEDOL BFM4703  ': 'JOMLILIQ0XX1'}

jpm_data = pd.read_sql(jpm_query, con = jpm_conn)
jpm_data['DESCRIPTION'] = jpm_data.TRANS_DESC_1
jpm_data['INSTR_ID'] = jpm_data.CUSIP.fillna(jpm_data.DESCRIPTION.map(fvals))
jpm_acct = jpm_data.ACCT_NUM.astype(str).str[-4:]
jpm_data['INSTR_ID'] = best_instr_id(jpm_data.INSTR_ID.astype(str) + '-' + jpm_acct,
                                     jpm_data.INSTR_ID,
                                     jpm_acct + '-' + jpm_data.TRANSACTION_CCY.astype(str))

jpm_full = pd.merge(jpm_data, data, how = 'left', on = 'INSTR_ID').rename(columns = {
    'DATE_OF_TRANSACTION': 'FECHA_TRADE', 'SETTLEMENT_DATE': 'FECHA_VALOR',
    'TYPE_OF_TRANSACTION': 'CONCEPTO', 'UNITS': 'TITULOS', 'NET_AMNT_BASE': 'MONTO_NETO',
    'TRANSACTION_CCY': 'CCY', 'TRANS_DESC_1': 'NOMBRE_VALOR',
    'MAJOR_ASSET_CLASS_CODE': 'TIPO_VALOR', 'RECORD_ENTRY_NUM': 'REFERENCIA',
    'ACCT_NUM': 'ACCT', 'PRICE_PER_SHARE': 'PRECIO', 'COMMISSION_AMNT_BASE': 'CORRETAJE'})
jpm_full['DESCRIPCION'] = jpm_full.CONCEPTO
desc_orig['JPM'] = jpm_full.CONCEPTO.copy()
jpm_full['CONCEPTO'] = jpm_full.CONCEPTO.map(jpm_map)

# --- Fondos privados: reclasificar por frase en la descripcion ---
jpm_desc = ['NOMBRE_VALOR', 'TRANS_DESC_2', 'TRANS_DESC_3', 'TRANS_DESC_4',
            'TRANSACTION_DESC_5', 'TRANSACTION_DESC_6']
jpm_texto = (jpm_full[jpm_desc].fillna('').agg(' '.join, axis = 1)
             .str.replace(r'\s+', ' ', regex = True).str.strip().str.upper())

jpm_alt = [(r'COMMITMENT TO',         'Commitment'),
           (r'-\s*RETURN OF CAPITAL', 'Return of Capital'),
           (r'-\s*DISTRIBUTIONS?\b',  'Distribution'),
           (r'-\s*MANAGEMENT FEES?',  'Fund Expense'),
           (r'-\s*FUND EXPENSES?',    'Fund Expense'),
           (r'-\s*CONDUIT',           'Fund Expense'),
           (r'-\s*TAX EXPENSE',       'Fund Expense'),
           (r'-\s*ORIGINATION FEE',   'Fund Expense'),
           (r'-\s*INTEREST TRUE-UP',  'Interest'),
           (r'-\s*INVESTMENT\b',      'Capital Call')]

jpm_es_alt = jpm_full.NEW_ASSET_CLASS.eq('ALTERNATIVES')
jpm_evento = pd.Series(pd.NA, index = jpm_full.index, dtype = object)
for patron, nuevo in jpm_alt:
    aun_sin = jpm_es_alt & jpm_evento.isna() & jpm_texto.str.contains(patron, regex = True, na = False)
    jpm_evento[aun_sin] = nuevo
jpm_full['CONCEPTO'] = jpm_evento.fillna(jpm_full.CONCEPTO)
jpm_full.loc[jpm_es_alt, 'DESCRIPCION'] = jpm_texto[jpm_es_alt]

# Se calcula antes del corte de columnas, que deja fuera REVERSAL_INDICATOR.
jpm_rev = jpm_full.REVERSAL_INDICATOR.astype(str).str.strip().eq('R')

jpm_full = jpm_full[cols_final]

# JPM manda la reversa YA con el signo contrario al original -- verificado emparejando por
# REVERSAL_REC_ENTRY_NUM: INTEREST -0.17 contra su original +0.17, MISC. INCOME -22,471.50
# contra FOREIGN INTEREST +22,471.50. La suma neta ya es correcta; invertirla la romperia.
jpm_full = marcar_reversa(jpm_full, jpm_rev, 'JPM', 'REVERSAL_INDICATOR',
                          "REVERSAL_INDICATOR = 'R'", 'registrar')

# --- JPM reemite cada movimiento en curso todos los dias habiles ---
# JPM.TRANSACTIONS no es un libro de operaciones, es un reporte diario de estado: un
# movimiento abierto se vuelve a emitir cada dia, con el mismo RECORD_ENTRY_NUM, hasta que
# el banco lo asienta. Leerlas todas inflaba la hoja 3.7x (4,335 filas contra 1,314 reales).
#
# Verificado contra los estados de cuenta del banco (jpm/*.csv, ver verificacion_jpm.ipynb):
# por instrumento, el numero de REFERENCIA unicas coincide exacto con las filas del extracto
# (59=59, 60=60, 34=34, 19=19, 16=16), y la ULTIMA fila de cada grupo es la que el banco
# asienta en el 97% de los casos -- 100% en los grupos cuyo monto varia.
#
# OJO: esto NO es un drop_duplicates(). Las filas de un grupo no siempre son iguales:
#   A) devengo   -- el monto crece cada dia (dividendo 16.57 -> 86.70 en 16 filas; el banco
#                   paga 86.70, sumarlas todas lo inflaria a mas de 800)
#   B) pendiente -- cambia concepto y signo al liquidar (7x 'Fees Paid' +9,312.80 y luego
#                   'Capital Call' -9,312.80; el banco solo asienta la ultima)
#   C) copia     -- solo cambia la fecha
# En los tres casos la buena es la ultima, de ahi el sort por fecha + tail(1).
jpm_ult = jpm_full.sort_values('FECHA_TRADE').groupby('REFERENCIA', dropna = True).tail(1).index
jpm_full = descartar(jpm_full,
                     jpm_full.REFERENCIA.notna() & ~jpm_full.index.isin(jpm_ult),
                     'JPM reemite el movimiento cada dia hasta asentarlo; se conserva la ultima',
                     'JPM')

### ========================== SSZ ========================== ###

ssz_query = """
    SELECT TRADE_DATE, SETTLEMENT_DATE, TRX_DES,
           NOMINAL_AMNT, TRX_PRICE, NET_AMNT,
           PRODUCT_CCY, PORT_ID, ISIN, IBAN,
           PRODUCT_DES, TRX_TYPE, TRX_REF, REVERSAL_DATE
    FROM TRANSACTIONS
    WHERE PORT_ID NOT IN ('64520-1', '64520-2')
    ORDER BY TRADE_DATE DESC
"""

ssz_map = {
    'Dividends': 'Dividends',
    'Coupons': 'Interest',
    'Outward Payment': 'Cash Withdrawal',
    'Currency Sale': 'Asset Sold',
    'Currency Purchase': 'Asset Purchased',
    'SEC EXCHANGE RECD': 'Transfer Shares In',
    'Transfer': 'Cash Withdrawal',
    'Management Fees': 'Fees Paid - Asset Account',
    'MainteNULLce Fees': 'Fees Paid - Asset Account',   # corrupto en la base: nan→NULL
    'Reimb.of Interest': 'Interest',
    'Payment of Principal': 'Loan Payment',
    'SECURITY SALE': 'Asset Sold',
    'NEW ISSUE PURCH': 'Asset Purchased',
    'SECURITY PURCHASE': 'Asset Purchased',
    'SEC TRANSFER OUT': 'Transfer Shares Out',
    'Debit Interest': 'Interest',
    'REDEMPTION': 'Asset Sold',
    'Redemption': 'Asset Sold',
    'Redemption Sale': 'Asset Sold',
    'Loan fees': 'Fees Paid - Asset Account',
    'FUND REDEMPTION': 'Asset Sold',
    'Third Party Fee': 'Fees Paid - Asset Account',
    'Inward Transfer Payment': 'Cash Deposit',
    'Correction of Deb. Int.': 'Interest',
    'STK SPLIT PAID': 'Transfer Shares Out',
    'STK SPLIT RECD': 'Transfer Shares In'
}

ssz_data = pd.read_sql(ssz_query, con = ssz_conn)
ssz_data['ISIN'] = ssz_data.ISIN.combine_first(ssz_data.IBAN)
ssz_plain_id = ssz_data.ISIN.fillna('SSZLOAN0')
ssz_port_familia = ssz_data.PORT_ID.astype(str).str.split('-').str[0]
ssz_data['INSTR_ID'] = best_instr_id(ssz_plain_id + '-' + ssz_port_familia, ssz_plain_id)

ssz_full = pd.merge(ssz_data, data, how = 'left', on = 'INSTR_ID').rename(columns = {
    'TRADE_DATE': 'FECHA_TRADE', 'SETTLEMENT_DATE': 'FECHA_VALOR', 'TRX_DES': 'CONCEPTO',
    'NOMINAL_AMNT': 'TITULOS', 'TRX_PRICE': 'PRECIO', 'NET_AMNT': 'MONTO_NETO',
    'PRODUCT_CCY': 'CCY', 'PRODUCT_DES': 'NOMBRE_VALOR', 'TRX_TYPE': 'TIPO_VALOR',
    'TRX_REF': 'REFERENCIA', 'PORT_ID': 'ACCT'})
ssz_full['CORRETAJE'] = 0
ssz_full['DESCRIPCION'] = ssz_full.CONCEPTO
desc_orig['SSZ'] = ssz_full.CONCEPTO.copy()
ssz_full['CONCEPTO'] = ssz_full.CONCEPTO.map(ssz_map)

# En SSZ, REVERSAL_DATE marca la fila CANCELADA, no la que cancela: el banco vuelve a
# registrar la operacion como una fila nueva con otra referencia. Verificado por producto:
# la compra marcada de 1,160,863.75 (ref ...30800138.1) reaparece como 1,160,873.42
# (ref ...30900129.1), y la venta de MICROSOFT por 157,290.56 tiene una gemela sin marcar.
# Invertirla dejaria neto cero en vez de la operacion real, asi que se descarta.
ssz_rev = ssz_full.REVERSAL_DATE.notna()
ssz_full = ssz_full[cols_final]
ssz_full = marcar_reversa(ssz_full, ssz_rev, 'SSZ', 'REVERSAL_DATE',
                          'REVERSAL_DATE no nulo; el banco la volvio a registrar aparte',
                          'descartar')

### ========================== UBS ========================== ###

ubs_query = """
    SELECT T.TRADE_DT, T.SETTLEMENT_DT,
           COALESCE(C.DESCRIPT, T.TRX_CD) AS DESCRIPT, T.QUANTITY,
           T.TRADE_AMT, T.SEC_SYMBOL, T.acct_nbr AS ACCT_NBR,
           T.comment AS COMENTARIO, T.sec_typ AS SEC_TYP
    FROM TRANSACTIONS T
    LEFT JOIN TRX_CODES C ON T.TRX_CD = C.CODES
    ORDER BY T.TRADE_DT DESC
"""

ubs_map = {
    'Deposit': 'Cash Deposit',
    'Withdrawal': 'Cash Withdrawal',
    'Sell': 'Asset Sold',
    'Buy': 'Asset Purchased',
    'Deliver Out - Long Position': 'Transfer Shares Out',
    'Deliver In - Long Position': 'Transfer Shares In',
    'Interest': 'Interest',
    'Return of Capital - Long Position': 'Asset Sold',
    'Dividend': 'Dividends',
    'Accrued Interest - Buy': 'Interest',
    'dv': 'Dividends'
}

ubs_data = pd.read_sql(ubs_query, con = ubs_conn)
ubs_acct_suffix = ubs_data.ACCT_NBR.astype(str).str[-5:].str.upper()
ubs_data['SEC_SYMBOL'] = best_instr_id(ubs_data.SEC_SYMBOL.astype(str) + '-' + ubs_acct_suffix, ubs_data.SEC_SYMBOL)

ubs_full = pd.merge(ubs_data, data, how = 'left', left_on = 'SEC_SYMBOL',
                    right_on = 'INSTR_ID').rename(columns = {
    'TRADE_DT': 'FECHA_TRADE', 'SETTLEMENT_DT': 'FECHA_VALOR', 'DESCRIPT': 'CONCEPTO',
    'QUANTITY': 'TITULOS', 'TRADE_AMT': 'MONTO_NETO', 'COMENTARIO': 'NOMBRE_VALOR',
    'SEC_TYP': 'TIPO_VALOR', 'ACCT_NBR': 'ACCT'})
ubs_full['PRECIO'] = 0
ubs_full['CORRETAJE'] = 0
ubs_full['CCY'] = 'USD'
ubs_full['REFERENCIA'] = None
ubs_full['DESCRIPCION'] = ubs_full.CONCEPTO
desc_orig['UBS'] = ubs_full.CONCEPTO.copy()
ubs_full['CONCEPTO'] = ubs_full.CONCEPTO.map(ubs_map)
ubs_full = ubs_full[cols_final]
# UBS no expone columna de reversa, pero TRADE_AMT viene sin signo (0 negativos de 356).
ubs_full = normalizar_signo(ubs_full, 'UBS', monto = True)

### ========================== MSY ========================== ###

msy_query = """
    SELECT T.TRADE_DT, T.SETTLE_DT,
           COALESCE(T.ALT_TRAN_CD, T.TRAN_CD) AS DESCRIPT,
           T.QUANTITY, T.PRICE, T.TOTAL_AMT,
           T.COMM_BASE, T.CUSIP, T.ACCT,
           T.SEC_DESC, T.SEC_CD, T.CURR_ORDER,
           T.CANCEL_IND
    FROM TRANSACTIONS T
    ORDER BY T.TRADE_DT DESC
"""

msy_map = {
    'BUY': 'Asset Purchased', 'SELL': 'Asset Sold',
    'DEP': 'Cash Deposit',    'WITH': 'Cash Withdrawal',
    'DIV': 'Dividends',       'INT': 'Interest',
    'OTH': 'Interest',
    'MFEE': 'Fees Paid - Asset Account',
    'EXP': 'Fees Paid - Asset Account',
    'FGTAX': 'Fees Paid - Asset Account',
    'NRTAX': 'Fees Paid - Asset Account',
    'REC': 'Transfer Shares In',
    'DEL': 'Transfer Shares Out',
    '44410IDL': 'Dividends',
    '44410ERL': 'Asset Purchased'
}

msy_data = pd.read_sql(msy_query, con = msy_conn)
msy_base = msy_data.CUSIP.fillna('MSYCASHXX0').astype(str)
msy_data['INSTR_ID'] = best_instr_id(msy_base + '-' + msy_data.ACCT.astype(str).str[-4:], msy_base)
msy_data['CCY'] = 'USD'

msy_full = pd.merge(msy_data, data, how = 'left', on = 'INSTR_ID').rename(columns = {
    'TRADE_DT': 'FECHA_TRADE', 'SETTLE_DT': 'FECHA_VALOR', 'DESCRIPT': 'CONCEPTO',
    'QUANTITY': 'TITULOS', 'PRICE': 'PRECIO', 'TOTAL_AMT': 'MONTO_NETO',
    'COMM_BASE': 'CORRETAJE', 'SEC_DESC': 'NOMBRE_VALOR', 'SEC_CD': 'TIPO_VALOR',
    'CURR_ORDER': 'REFERENCIA'})
msy_full['DESCRIPCION'] = msy_full.CONCEPTO
desc_orig['MSY'] = msy_full.CONCEPTO.copy()
msy_full['CONCEPTO'] = msy_full.CONCEPTO.map(msy_map)

# --- Cancelaciones de MSY ---
# TRANSACTIONS.CANCEL_IND = 'C' marca la fila que anula una operacion previa, pero el query
# no la traia y ALT_TRAN_CD sigue diciendo 'SELL', asi que msy_map la mandaba a 'Asset Sold'
# como si fuera una venta mas. Caso que lo destapo: el T-bill 912797RV1 del 07-ene-2026
# aparecia 3 veces por 50,000 -- parecia una venta de 150,000. ActivityMS.xlsx trae 2 'Sold'
# y 1 'Cancel Sell', y MSY.POSITION confirma que la posicion solo bajo 50,000 ese dia.
# No se puede emparejar la cancelacion con su original (PREV_ORDER apunta a la fila misma),
# asi que se invierte el signo igual que hace el extracto: 50000 + 50000 - 50000 = 50000.
msy_rev = msy_full.CANCEL_IND.eq('C')       # antes del corte, que deja fuera CANCEL_IND
msy_full = msy_full[cols_final]
# Sin `excluir`: la fila cancelada tambien viene sin signo, hay que normalizarla ANTES de
# invertirla. Excluirla dejaba -150,000 titulos en el T-bill en vez de -50,000.
msy_full = normalizar_signo(msy_full, 'MSY', monto = True, titulos = True)
msy_full = marcar_reversa(msy_full, msy_rev, 'MSY', 'CANCEL_IND',
                          "CANCEL_IND = 'C'", 'invertir')

### ========================== DBK ========================== ###

dbk_query = """
    SELECT T.PROC_DT, T.SETTLE_DT, T.BACT_CD, T.TRN_TYPE, T.BUY_SELL,
           T.QUANTITY, T.NET_AMT_USD, T.SIGN3, T.COM_USD, T.USD_PC,
           T.CCY_IND, T.CUSIP, T.ACCT_NBR,
           T.SYMBOL, T.SEC_TYPE_CD, T.REF_NBR, T.CANCEL_CD
    FROM TRANSACTIONS T
    ORDER BY T.SETTLE_DT DESC
"""

dbk_data = pd.read_sql(dbk_query, con = dbk_conn)
# SIGN3 perspectiva del titulo: '+' = entran titulos = sale dinero
dbk_data['MONTO'] = dbk_data.NET_AMT_USD * dbk_data.SIGN3.map({'+': -1}).fillna(1)
dbk_base = dbk_data.CUSIP.astype(str)
dbk_data['INSTR_ID'] = best_instr_id(dbk_base + '-' + dbk_data.ACCT_NBR.astype(str).str[:-1].str[-4:],
                                     dbk_base)

# Concepto derivado: las 6 reglas se aplican en orden, la ultima gana
dbk_es_trade = dbk_data.TRN_TYPE == 'T'
concepto = pd.Series('Cash Withdrawal', index = dbk_data.index)
concepto[dbk_data.MONTO >= 0] = 'Cash Deposit'
concepto[dbk_data.BACT_CD.isin(['*NRAD', '*NRAR'])] = 'Fees Paid - Asset Account'
concepto[dbk_data.BACT_CD == '*DVDS'] = 'Dividends'
concepto[dbk_es_trade & (dbk_data.BUY_SELL == 'B')] = 'Asset Purchased'
concepto[dbk_es_trade & (dbk_data.BUY_SELL == 'S')] = 'Asset Sold'

dbk_data = dbk_data.rename(columns = {'USD_PC': 'PRECIO'})
dbk_full = pd.merge(dbk_data, data, how = 'left', on = 'INSTR_ID').rename(columns = {
    'PROC_DT': 'FECHA_TRADE', 'SETTLE_DT': 'FECHA_VALOR', 'BACT_CD': 'DESCRIPCION',
    'QUANTITY': 'TITULOS', 'MONTO': 'MONTO_NETO', 'COM_USD': 'CORRETAJE',
    'CCY_IND': 'CCY', 'SYMBOL': 'NOMBRE_VALOR', 'SEC_TYPE_CD': 'TIPO_VALOR',
    'REF_NBR': 'REFERENCIA', 'ACCT_NBR': 'ACCT'})
desc_orig['DBK'] = dbk_full.DESCRIPCION.copy()
dbk_full['CONCEPTO'] = concepto

# DBK manda el par completo de la cancelacion: la fila marcada con CANCEL_CD es la compra
# ('+', 15,000 de 91282CBC4) y junto a ella viene la venta ('-') con el mismo REF_NBR
# P1H4UX. Ya se anulan entre si, asi que solo se deja traza.
dbk_rev = dbk_full.CANCEL_CD.notna() if 'CANCEL_CD' in dbk_full else pd.Series(False, index = dbk_full.index)
dbk_full = dbk_full[cols_final]
dbk_full = normalizar_signo(dbk_full, 'DBK', titulos = True, excluir = dbk_rev)
dbk_full = marcar_reversa(dbk_full, dbk_rev, 'DBK', 'CANCEL_CD',
                          'CANCEL_CD no nulo; el par compra/venta ya se anula solo',
                          'registrar')

# ========================= Reporte =========================
hojas = {'JBR': jb_full, 'CT': ct_full, 'JPM': jpm_full, 'SSZ': ssz_full, 'UBS': ubs_full,
         'MSY': msy_full, 'DBK': dbk_full}
excepciones = pd.concat([find_exceptions(df, banco) for banco, df in hojas.items()], ignore_index = True)

todas = pd.concat([df.assign(BANCO = banco) for banco, df in hojas.items() if not df.empty],
                  ignore_index = True)
todas = todas[['BANCO'] + [c for c in todas.columns if c != 'BANCO']]
for col in ('FECHA_TRADE', 'FECHA_VALOR'):
    todas[col] = pd.to_datetime(todas[col], errors = 'coerce')
todas = todas.sort_values('FECHA_TRADE', ascending = False).reset_index(drop = True)

descartadas = (pd.concat(filtradas, ignore_index = True) if filtradas
               else pd.DataFrame(columns = ['BANCO', 'MOTIVO_DESCARTE']))
revertidas = (pd.concat(reversas, ignore_index = True) if reversas
              else pd.DataFrame(columns = ['BANCO', 'ACCION', 'MARCADA_POR', 'MOTIVO']))
resignadas = (pd.concat(signos, ignore_index = True) if signos
              else pd.DataFrame(columns = ['BANCO', 'CAMPO', 'VALOR_ANTES', 'VALOR_DESPUES']))
futuras = todas[todas.FECHA_TRADE > pd.Timestamp(report_date)].copy()

with pd.ExcelWriter(os.path.join(base_dir, f'{report_date}-TRANSACTIONS.xlsx'), engine = 'openpyxl') as writer:
    todas.to_excel(writer, sheet_name = 'TODAS', index = False)
    for banco, df in hojas.items():
        df.to_excel(writer, sheet_name = banco, index = False)
    excepciones.to_excel(writer, sheet_name = 'Excepciones', index = False)
    descartadas.to_excel(writer, sheet_name = 'Filtradas', index = False)
    revertidas.to_excel(writer, sheet_name = 'Reversas', index = False)
    resignadas.to_excel(writer, sheet_name = 'Signos', index = False)
    futuras.to_excel(writer, sheet_name = 'Fechas_futuras', index = False)

print(f'{report_date}-TRANSACTIONS.xlsx')
print(f'  TODAS          {len(todas):>6} filas')
for banco, df in hojas.items():
    print(f'  {banco:<14} {len(df):>6}')
print(f'  Excepciones    {len(excepciones):>6}')
print(f'  Filtradas      {len(descartadas):>6}')
print(f'  Reversas       {len(revertidas):>6}')
print(f'  Signos         {len(resignadas):>6}')
print(f'  Fechas_futuras {len(futuras):>6}')
