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
from config import DATE, PARENT_PATH
from debugg.functionsV1 import vix_plot, str_notes, dxy_plot

path_notes = PARENT_PATH / 'Notas Estructuradas'
vix = pd.read_excel(path_notes / 'Notas Presentación (New).xlsx', sheet_name = 'VIX', skiprows = 1)
today = pd.to_datetime(DATE).date()

# =============================================================================
#    GRAFICA VIXXO 
# =============================================================================
cols = ['RTY', 'NDX', 'SPX', 'VIXXO', 'VIXXO3M', 'VIXXO6M', 'Max. Funding Rate']
vix_plot(PARENT_PATH, '/DATOS/Graficas_PPT/VIXXO', vix, cols)

# =============================================================================
# GRAFICA VIX
# =============================================================================
cols = ['RTY', 'NDX', 'SPX', 'VIX', 'Max. Funding Rate']
vix_plot(PARENT_PATH, '/DATOS/Graficas_PPT/VIX', vix, cols, yTicks = 10)


# =============================================================================
#  Graficas barreras
# =============================================================================
notes = pd.read_excel(f'{path_notes}/Graficas_Barreras - copia.xlsx', engine = 'openpyxl', sheet_name = 'NOTAS')
series = pd.read_excel(f'{path_notes}/Graficas_Barreras - copia.xlsx', engine = 'openpyxl', sheet_name = 'SERIES')
colors = pd.read_excel(f'{path_notes}/Graficas_Barreras - copia.xlsx', engine = 'openpyxl', sheet_name = 'COLORS')

notes = notes[(notes.Estatus == 'ON GOING') & (notes.IN_AA == 'YES') & (notes.SHORT_NAME != 'Nota DXY 08/27/2027')].reset_index(drop = True)
notes['Strike'] = 1
notes = notes[['ISIN', 'SHORT_NAME', 'Barrera', 'CCY', 'Nominal', 'Trade Date', 'Call_Period', 'CVD 1','CVD 2','CVD 3','CVD 4','CVD 5','CVD 6','CVD 7','CVD 8','Strike','BANK']]
notes['Barrera'] = (1 - notes.Barrera) * 100

imgs_dir = PARENT_PATH / 'DATOS' / 'Graficas_PPT' / 'NOTAS'
c = 0

# Elimina las imagenes previas del directorio
for file in os.listdir(imgs_dir):
    file_path = os.path.join(imgs_dir, file)
    if os.path.isfile(file_path):
        os.remove(file_path)
        c += 1
print(f'{c} images removed')

# Crea las nuevas imagenes de las notas
str_notes(imgs_dir, notes, series, colors, today)

# Crea la grafica de DXY
dxy_plot(imgs_dir, len(notes), series, colors)