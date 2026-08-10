import pandas as pd

path = 'C:/Users/DATOS-INVERSIONES/OneDrive/Escritorio/AGG SPX ACWI Series.xlsx'

data = pd.read_excel(path)
data['Fecha'] = data.Date.dt.day.astype(str) + data.Date.dt.month_name() + data.Date.dt.year.astype(str)
data = data.drop('Date', axis = 1)

print(data.head())
data.to_csv('C:/Users/DATOS-INVERSIONES/OneDrive/Escritorio/prueba.csv', index = False)