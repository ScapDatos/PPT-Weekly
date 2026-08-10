### ============================================= ###
### Notas:
###     - Actualizar la fecha al viernes inmediato anterior.
### ============================================= ###

import pandas as pd
from os import listdir
from pptx import Presentation
from pptx.dml.color import RGBColor
from config import DATE, PARENT_PATH
from debugg.functions import get_df, slide_look, slide_area

# ppt_date = "12062026"
# ppt_datetime = pd.to_datetime(ppt_date, format = '%d%m%Y')
ppt_datetime = pd.to_datetime(DATE, format = '%Y-%m-%d')

# path_pc = 'C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt'
# path_to_db = f'{path_pc}/DATOS/Bases de datos'
path_pc = PARENT_PATH
path_to_db = PARENT_PATH / 'DATOS' / 'Bases de datos'

foot_note = f'Información al {ppt_datetime.strftime("%d/%m/%Y")}'
ppt = Presentation(f'{path_pc}/DATOS/Presentaciones/Template.pptx')

# ============================= 1 - Portada ===================================
lamina_n = ppt.slides[0]
for shape in lamina_n.shapes:
    if shape.name=='CuadroTexto 22':
        fecha_der = f'{ppt_datetime.day} {ppt_datetime.month_name(locale = "es_ES")} {ppt_datetime.year}'
        ppt_date_last = ppt_datetime - pd.Timedelta(days = 4)
        fecha_izq = f'{ppt_date_last.day} {ppt_date_last.month_name(locale = "es_ES")} {ppt_date_last.year}'
        shape.text = f'{fecha_izq} - {fecha_der}'
        shape.text_frame.paragraphs[0].runs[0].font.bold = True
        shape.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(137, 104, 126)

### ================================================ ###
###      Distribucion del portafolio en general      ###
### ================================================ ###
nslide = 1          # Diapositiva 2 - Consolidado por monedas
slide_look(ppt, nslide, f'{path_pc}/Asset Allocation/Vistas/1 - ASSET MIX BY CURRENCY.png', foot_note)

nslide += 1         # Diapositiva 3 - Consolidado Asset Mix
slide_look(ppt, nslide, f'{path_pc}/Asset Allocation/Vistas/1 - ASSET MIX Completo.png', foot_note)

nslide += 1         # Diapositiva 4 - Consolidado Asse Mix by BrokerS
slide_look(ppt, nslide, f'{path_pc}/Asset Allocation/Vistas/1 - ASSET MIX BY BROKER.png', foot_note)

nslide += 1         # Diapositiva 5 - Consolidados Asset Mix (wo/ WK-Actinver)
lamina_n = ppt.slides[nslide]
placeholder = lamina_n.placeholders[13]                             # Asset Mix NFD
image_path = f'{path_pc}/Asset Allocation/Vistas/2 - ASSET MIX DIV-NFD.png'
placeholder_picture = placeholder.insert_picture(image_path)
placeholder = lamina_n.placeholders[15]                             # Asset Mix SCH
image_path = f'{path_pc}/Asset Allocation/Vistas/2 - ASSET MIX DIV-SCH.png'
placeholder_picture = placeholder.insert_picture(image_path)
placeholder = lamina_n.placeholders[16]
placeholder.text = foot_note

nslide += 1         # Diapositiva 6 - Consolidado | Leverage Multiples
lamina_n = ppt.slides[nslide]
placeholder = lamina_n.placeholders[13]                             # Leverage NFD
image_path = f'{path_pc}/DATOS/Graficas_PPT/LEVERAGE_LV_NFD.png'
placeholder_picture = placeholder.insert_picture(image_path)
placeholder = lamina_n.placeholders[15]                             # Leverage SCH
image_path = f'{path_pc}/DATOS/Graficas_PPT/LEVERAGE_LV_SCH.png'
placeholder_picture = placeholder.insert_picture(image_path)
placeholder = lamina_n.placeholders[16]                             # Leverage
image_path = f'{path_pc}/DATOS/Graficas_PPT/LEVERAGE_LV.png'
placeholder_picture = placeholder.insert_picture(image_path)
placeholder = lamina_n.placeholders[17]
placeholder.text = foot_note

nslide += 1         # Diapositiva 7 - Consolidado s/HCITY + VC Funds MX | Core Asset Mix
slide_look(ppt, nslide, f'{path_pc}/Asset Allocation/Vistas/1 - ASSET MIX Consolidado sHcity + VCFunds.png', foot_note)

nslide += 1         # Dispositiva 8 - Mercados | Principales indices de Paises
slide_look(ppt, nslide, f'{path_pc}/DATOS/Graficas_PPT/REND_COUNTRIES.png', foot_note)

nslide += 1         # Diapositiva 9 - Mercados | Principales indices Sectores US
slide_look(ppt, nslide, f'{path_pc}/DATOS/Graficas_PPT/REND_SECTORS.png', foot_note)

nslide += 1         # Diapositiva 10 - Mercados | Principales indices Estrategia
slide_look(ppt, nslide, f'{path_pc}/DATOS/Graficas_PPT/REND_STRATEGY.png', foot_note)

nslide += 1         # Diapositiva 11 - Mercadso | Principales indices ETF y Deuda
slide_look(ppt, nslide, f'{path_pc}/DATOS/Graficas_PPT/REND_DEBT.png', foot_note)

nslide += 1         # Diapositiva 12 - Mercados | S&P vs Fondos JPM
slide_look(ppt, nslide, f'{path_pc}/DATOS/Graficas_PPT/REND_JPM_FUNDS.png', foot_note)

### ================================================ ###
###                      SC Holdings                 ###
### ================================================ ###
nslide += 1         # Diapositiva 13 - Separador SC Holdings

nslide += 1         # Diapositiva 14 - SCH | 2014-2024 (dolares)
df = get_df(path_to_db, 'SigCap.db', 'PORT_SCH_NAV')
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/SCHOLDINGS.png', txt = foot_note)

nslide += 1         # Diapositiva 15 - SCH | 2024-2025 (dolares)
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/SCHOLDINGS12M.png', txt = foot_note)

nslide += 1         # Diapositiva 16 - SCH | Contribuidores al rendimiento
lamina_n = ppt.slides[nslide]
placeholder = lamina_n.placeholders[13]                             # Contributors plot
image_path = f'{path_pc}/DATOS/Graficas_PPT/Top10SCH.png'
placeholder_picture = placeholder.insert_picture(image_path)
placeholder = lamina_n.placeholders[15]                             # Detributors plot
image_path = f'{path_pc}/DATOS/Graficas_PPT/Bot10SCH.png'
placeholder_picture = placeholder.insert_picture(image_path)
placeholder = lamina_n.placeholders[16]
placeholder.text = foot_note

nslide += 1         # Diapositiva 17 - SCH | Asset Mix
slide_look(ppt, nslide, f'{path_pc}/Asset Allocation/Vistas/2 - ASSET MIX SCH.png', foot_note)

nslide += 1         # Diapositiva 18 - SCH | Variacion semanal (dolares)
slide_look(ppt, nslide, f'{path_pc}/DATOS/Graficas_PPT/SCHOLDINGS_WATERFALL.png', foot_note)

nslide += 1         # Diapositiva 19 - SCH s/HCITY | 2014-2025 (dolares)
# Calculos de los valores para colocar en la tabla
hcity = get_df(path_to_db, 'SigCap.db', 'HCITY')
df_w_hcity = pd.merge(df, hcity[['DATE', 'HCITY_SCH', 'LOANS_SCH']], on = 'DATE')
df_w_hcity.columns = df.columns.to_list() + ['HCITY', 'HCITY_LOAN']
df_w_hcity['NAV WITHOUT HCITY'] = df_w_hcity.NAV - df_w_hcity.HCITY - df_w_hcity.HCITY_LOAN
slide_area(path_pc, ppt, nslide, df_w_hcity, f'{path_pc}/DATOS/Graficas_PPT/SCHOLDINGS_W_HCITY.png', serie = 'NAV WITHOUT HCITY', txt = foot_note)

nslide += 1         # Diapositiva 20 - SCH s/HCITY | Asset Mix
slide_look(ppt, nslide, f'{path_pc}/Asset Allocation/Vistas/3 - ASSET MIX SCH SIN HCITY.png', foot_note)

nslide += 1         # Diapositiva 21 - SCH | 2024-2026 (Retorno Efectivo Corregido)
slide_look(ppt, nslide, f'{path_pc}/DATOS/Graficas_PPT/REND_SCHvsSHV.png', foot_note)

nslide += 1         # Diapositiva 22 - SCH | YTD (Retorno Efectivo Corregido)
slide_look(ppt, nslide, f'{path_pc}/DATOS/Graficas_PPT/REND_SCHvsSHV_YTD.png', foot_note)

nslide += 1         # Diapositiva 23 - SCH | 2024-2026 (Retorno efectivo Corregido)
slide_look(ppt, nslide, f'{path_pc}/DATOS/Graficas_PPT/REND_SCH.png', foot_note)

nslide += 1         # Diapositiva 24 - SCH | YTD (Retorno efectivo Corregido)
slide_look(ppt, nslide, f'{path_pc}/DATOS/Graficas_PPT/REND_SCH_YTD.png', foot_note)

nslide += 1         # Diapositiva 25 - SCH | 2024-2026 (Retorno efectivo Corregido - sin HCITY)
slide_look(ppt, nslide, f'{path_pc}/DATOS/Graficas_PPT/REND_SCH_SIN_HCITY.png', foot_note)

nslide += 1         # Diapositiva 26 - SCH | YTD (Retorno efectivo Corregido - sin HCITY)
slide_look(ppt, nslide, f'{path_pc}/DATOS/Graficas_PPT/REND_SCH_SIN_HCITY_YTD.png', foot_note)

nslide += 1         # Diapositiva 27 - Separador Golgman Sachs
nslide += 1         # Diapositiva 28 - SCH | Goldman Sachs 2019-2025 (Dolares)
# df = get_df(path_to_db, 'SigCap_SCH.db', 'GSS')
# slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/GS.png', sheetname = 'SCHGSS', txt = foot_note)

nslide += 1         # Diapositiva 29 - SCH Goldman Sachs 2024-2025 (dolares)
# slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/GS12M.png', txt = foot_note)

nslide += 1         # Diapositiva 30 - Separador UBS
nslide += 1         # Diapositiva 31 - SCH | UBS 2019-2025 (dolares)
df = get_df(path_to_db, 'SigCap_SCH.db', 'UBS')
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/UBSSCH.png', sheetname = 'SCHUBS', txt = foot_note)

nslide += 1         # Dipaositiva 32 - SCH | UBS 2024-2025 (dolares)
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/UBSSCH12M.png', txt = foot_note)

nslide += 1         # Diapositiva 33 - Separador Morgan Stanley
nslide += 1         # Diapositiva 34 - SCH | Morgan Stanley 2014-2025 (dolares)
df = get_df(path_to_db, 'SigCap_SCH.db', 'MSY')
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/MS.png', sheetname = 'SCHMSY', txt = foot_note)

nslide += 1         # Diapositiva 35 - SCH | Morgan Stanley 2024-2025 (dolares)
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/MS12M.png', txt = foot_note)

nslide += 1         # Diapositiva 36 - Separador Santander Suiza
nslide += 1         # Diapositiva 37 - SCH | Santander Suiza 2019-2025 (dolares)
df = get_df(path_to_db, 'SigCap_SCH.db', 'SSZ')
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/SS12M.png', sheetname = 'SCHSSZ', txt = foot_note)

nslide += 1         # Diapositiva 38 - SCH | Santander Suiza 2024-2025 (dolares)
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/SS.png', txt = foot_note)

nslide += 1         # Diapositiva 39 - SCH | Santander Suiza Buy & Hold (dolares)
placeholder = lamina_n.placeholders[15]
placeholder.text = foot_note

nslide += 1         # Diapositiva 40 - Separador Deutsche Bank
nslide += 1         # Diapositiva 41 - SCH | Deutsche Bank 2019-2025 (dolares)
df = get_df(path_to_db, 'SigCap_SCH.db', 'DBK')
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/DB.png', sheetname = 'SCHDBK', txt = foot_note)

nslide += 1         # Diapositiva 42 - SCH | Deutsche Bank 2024-2025 (dolares)
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/DB12M.png', txt = foot_note)

nslide += 1         # Diapositiva 43 - Separador JP Morgan
nslide += 1         # Diapositiva 44 - SCH | JP Morgan 2019-2025 (dolares)
df = get_df(path_to_db, 'SigCap_SCH.db', 'JPM')
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/JPM.png', sheetname = 'SCHJPM', txt = foot_note)

nslide += 1         # Diapositiva 45 - SCH | JP Morgan 2024-2025 (dolares)
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/JPM12M.png', txt = foot_note)

### ================================================ ###
###                          NFD                     ###
### ================================================ ###
nslide += 1         # Diapositiva 46 - Separador NFD
nslide += 1         # Diapositiva 47 - NFD | Rendimiento 2024-2026
slide_look(ppt, nslide, f'{path_pc}/DATOS/Graficas_PPT/REND_NFD.png', foot_note)

nslide += 1         # Diapositiva 48 - NFD | Rendimiento YTD
slide_look(ppt, nslide, f'{path_pc}/DATOS/Graficas_PPT/REND_NFD_YTD.png', foot_note)

nslide += 1         # Diapositiva 49 - NFD | Rendimiento 2024-2026 (Sin Equities)
slide_look(ppt, nslide, f'{path_pc}/DATOS/Graficas_PPT/REND_NFD_SIN_EQTS.png', foot_note)

nslide += 1         # Diapositiva 50 - NFD | Rendimiento YTD (Sin Equities)
slide_look(ppt, nslide, f'{path_pc}/DATOS/Graficas_PPT/REND_NFD_SIN_EQTS_YTD.png', foot_note)

nslide += 1         # Diapositiva 51 - NFD | Asset Mix
slide_look(ppt, nslide, f'{path_pc}/Asset Allocation/Vistas/2 - ASSET MIX NFD.png', foot_note)

nslide += 1         # Diapositiva 52 - NFD | 2019-2025 (dolares)
df = get_df(path_to_db, 'SigCap.db', 'PORT_NFD_NAV')
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/NFD.png', txt = foot_note)

nslide += 1         # Diapositiva 53 - NFD | 2024-2025 (dolares)
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/NFD12M.png', txt = foot_note)

nslide += 1         # Diapositiva 54 - NFD | Variacion Semanal (dolares)
slide_look(ppt, nslide, f'{path_pc}/DATOS/Graficas_PPT/NFD_WATERFALL.png', foot_note)

nslide += 1         # Diapositiva 55 - NFD | Contribuidores al rendimiento
lamina_n = ppt.slides[nslide]
placeholder = lamina_n.placeholders[13]                             # Contributors plot
image_path = f'{path_pc}/DATOS/Graficas_PPT/Top10NFD.png'
placeholder_picture = placeholder.insert_picture(image_path)
placeholder = lamina_n.placeholders[15]                             # Detributors plot
image_path = f'{path_pc}/DATOS/Graficas_PPT/Bot10NFD.png'
placeholder_picture = placeholder.insert_picture(image_path)
placeholder = lamina_n.placeholders[16]
placeholder.text = foot_note

nslide += 1         # Diapositiva 56 - NFD | NFD s/HCITY | Asset Mix
slide_look(ppt, nslide, f'{path_pc}/Asset Allocation/Vistas/3 - ASSET MIX NFD SIN HCITY.png', foot_note)

nslide += 1         # Dipoasitiva 57 - NFD s/HCITY | 2019-2025 (dolares)
hcity = get_df(path_to_db, 'SigCap.db', 'HCITY')
df_w_hcity = pd.merge(df, hcity[['DATE', 'HCITY_NFD', 'LOANS_NFD']], on = 'DATE')
df_w_hcity.columns = df.columns.to_list() + ['HCITY', 'HCITY_LOAN']
df_w_hcity['NAV WITHOUT HCITY'] = df_w_hcity.NAV - df_w_hcity.HCITY - df_w_hcity.HCITY_LOAN
slide_area(path_pc, ppt, nslide, df_w_hcity, f'{path_pc}/DATOS/Graficas_PPT/NFD_W_HCITY.png', serie = 'NAV WITHOUT HCITY', txt = foot_note)

nslide += 1         # Diapositiva 58 - Separador CITIBANK
nslide += 1         # Diapositiva 59 - NFD | CITIBANK 2019-2025 (dolares)
df = get_df(path_to_db, 'SigCap_NFD.db', 'CTI')
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/CTI.png', sheetname = 'NFDCTI', txt = foot_note)

nslide += 1         # Dipoasotiva 60 - NFD | CITIBANK 2024-2025 (dolares)
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/CTI12M.png', txt = foot_note)

nslide += 1         # Diapositiva 61 - Separador Julius Baer
nslide += 1         # Diapositiva 62 - NFD | Julius Baer 2019-2025 (dolares)
df = get_df(path_to_db, 'SigCap_NFD.db', 'JBR')
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/JB.png', sheetname = 'NFDJBR', txt = foot_note)

nslide += 1         # Diapositiva 63 - NFD | Julius Baer 2024-2025 (dolares)
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/JB12M.png', txt = foot_note)

nslide += 1         # Diapositiva 64 - Separador Santander Suiza
nslide += 1         # Diapositiva 65 - NFD | Santander Suiza 2019-2025 (dolares)
df = get_df( path_to_db,'SigCap_NFD.db','SSZ' )
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/SSNFD.png', sheetname = 'NFDSSZ', txt = foot_note)

nslide += 1         # Diapositiva 66 - NFD | Santander Suiza 2024-2025 (dolares)
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/SSNFD12M.png', txt = foot_note)

nslide += 1         # Diapositiva 67 - Separador UBS
nslide += 1         # Diapositiva 68 - NFD | UBS 2019-2025 (dolares)
df = get_df( path_to_db,'SigCap_NFD.db','UBS' )
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/UBS.png', sheetname = 'NFDUBS', txt = foot_note)

nslide += 1         # Diapositiva 69 - NFD | UBS 2024-2025 (dolares)
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/UBS12M.png', txt = foot_note)

### ================================================ ###
###                          WK                      ###
### ================================================ ###
nslide += 1         # Diapositiva 70 - Separador WK
nslide += 1         # Diapositiva 71 - WK | 2019-2025 (dolares)
df = get_df(path_to_db, 'SigCap.db', 'PORT_WK_NAV')
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/WK.png', txt = foot_note)

nslide += 1         # Diapositiva 72 - WK | 2024.2025 (dolares)
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/WK12M.png', txt = foot_note)

nslide += 1         # Diapositiva 73 - WK | Contribuidores al rendimiento
lamina_n = ppt.slides[nslide]
placeholder = lamina_n.placeholders[13]                             # Contributors plot
image_path = f'{path_pc}/DATOS/Graficas_PPT/Top10WK.png'
placeholder_picture = placeholder.insert_picture(image_path)
placeholder = lamina_n.placeholders[15]                             # Detributors plot
image_path = f'{path_pc}/DATOS/Graficas_PPT/Bot10WK.png'
placeholder_picture = placeholder.insert_picture(image_path)
placeholder = lamina_n.placeholders[16]
placeholder.text = foot_note

nslide += 1         # Diapositiva 74 - WK | Asset Mix
slide_look(ppt, nslide, f'{path_pc}/Asset Allocation/Vistas/2 - ASSET MIX WK.png', foot_note)

nslide += 1         # Diapositiva 75 - Separador Morgan Stanley
nslide += 1         # Diapositiva 76 - WK | Morgan stanley 2019-2025 (dolares)
df = get_df(path_to_db, 'SigCap_WK.db', 'MSY')
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/MSWK.png', sheetname = 'WKMSY', txt = foot_note)

nslide += 1         # Diapositiva 77 - WK | Morgan Stanley 2024-2025 (dolares)
slide_area(path_pc, ppt, nslide, df, f'{path_pc}/DATOS/Graficas_PPT/MSWK12M.png', txt = foot_note)


### ================================================ ###
###                        NOTAS                     ###
### ================================================ ###
nslide += 1         # Diapositiva 78 - Separador Semaforo del portafolio
nslide += 1         # Diapositiva 79 - NFD | Asset Mix con notras estructuradas
slide_look(ppt, nslide, f'{path_pc}/Asset Allocation/Vistas/5 - ASSET MIX NFD STRUCTURED NOTES.png', foot_note)

nslide += 1         # Diapositiva 80 - Notas | Monitor Exposicion por Emisor en Indice
lamina_n = ppt.slides[nslide]
placeholder = lamina_n.placeholders[15]                             # Bar plot
placeholder.text = foot_note

#                   # Diapositivas 81-84 NFD | Monitor de Notas Estructuradsa (1-24)
imgs = listdir(f'{path_pc}/DATOS/Graficas_PPT/Notas/')
placehldrs = [13, 16, 17, 18, 19, 20]

for diap in range(4): # Son 4 laminas de notas
    nslide += 1
    lamina_n = ppt.slides[nslide]
    
    for img in range(6): # Son 6 imgs por lamina
        img_idx = diap * 6 + img
        if img_idx < len(imgs):
            placeholder = lamina_n.placeholders[placehldrs[img]]
            image_path = f'{path_pc}/DATOS/Graficas_PPT/Notas/{img_idx}.png'
            placeholder_picture = placeholder.insert_picture(image_path)
    
    placeholder = lamina_n.placeholders[15]
    placeholder.text = foot_note

nslide += 1         # Diapositiva 85 - Separador Cupones Semanales
nslide += 1         # Diapositiva 86 - Notas | VIX Corto Plazo vs Cupones Principales Indices
slide_look(ppt, nslide, f'{path_pc}/DATOS/Graficas_PPT/VIXXO.png', foot_note)

nslide += 1         # Diapositiva 87 - Notas | VIX vs Cupones Principales Indices
slide_look(ppt, nslide, f'{path_pc}/DATOS/Graficas_PPT/VIX.png', foot_note)

nslide += 1         # Diapositiva 88 - Consolidado | Asset Mix por Sponsor
slide_look(ppt, nslide, f'{path_pc}/Asset Allocation/Vistas/7 - ETF MIX BY SPONSOR.png', foot_note)

ppt.save(f'{path_pc}/DATOS/Presentaciones/PPT-{ppt_datetime.strftime("%Y-%m-%d")}.pptx')