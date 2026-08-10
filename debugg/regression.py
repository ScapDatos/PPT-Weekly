import numpy as np
import pandas as pd
import statsmodels.api as sm
import matplotlib.pyplot as plt
from scipy.optimize import minimize
from sklearn.ensemble import RandomForestRegressor
from sklearn.preprocessing import PolynomialFeatures
from sklearn.linear_model import LinearRegression, Ridge

df = pd.read_excel(f'C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Contributtion/Dot plot.xlsx')

df['Fecha'] = pd.to_datetime(df['Fecha'], format = '%d/%m/%Y')
df['Año'] = df['Fecha'].dt.year
cols = [str(col) for col in df.columns]
df.columns = cols
df = df.drop(['Unnamed: 0', 'Unnamed: 21', 'Unnamed: 22', 'Date', 'Last Price'], axis = 1)
df = df.dropna(subset = 'Fecha')


# days_left = (pd.to_datetime('17/09/2027') - pd.to_datetime('31/12/2026')).days
# print(days_left/360)

def obtener_ffr_2y(row):
    year = row['Año'] + 2
    year2 = row.Fecha + pd.DateOffset(years = 2)
    year1 = pd.to_datetime(f'{int(year) - 1}-12-31', format = '%Y-%m-%d')
    delta_days = (year2 - year1).days
    percentage = delta_days / 360
    
    
    if str(int(year)) in df.columns:
        valores = df.loc[row.name, str(int(year) - 1)]
        diff = df.loc[row.name, str(int(year))] - df.loc[row.name, str(int(year) - 1)]
        diff_perc = percentage * diff
        valores += diff_perc
        
        if pd.notnull(valores):
            return valores  # si ya es mediana puedes dejarlo
    return np.nan

df['FFR_2Y'] = df.apply(obtener_ffr_2y, axis=1)
df['SPREAD'] = df.Long - df.FFR_2Y
# print(df)

# 1. Simulación de datos (reemplaza con tus datos reales)
data = df[['Fecha', 'Tasa de 10 AÑOS', 'FFR_2Y', 'Long']]

# 2. Modelo de regresión
X = df[['FFR_2Y', 'SPREAD']]
X = sm.add_constant(X)  # Agrega constante
y = df['Tasa de 10 AÑOS']

model = sm.OLS(y, X).fit()
df['Model_Pred'] = model.predict(X)

# 3. Calcular error estándar
residuals = df['Tasa de 10 AÑOS'] - df['Model_Pred']
std_error = residuals.std()
df['Upper'] = df['Model_Pred'] + std_error
df['Lower'] = df['Model_Pred'] - std_error

# 4. Proyección a futuro (2026)
future_ffr_2y = 3.125  # tomado del gráfico como "walks the walk"
future_lt = 3.0
fut = pd.DataFrame([[future_ffr_2y, future_lt]], columns=['FFR_2Y', 'Long'])
X_future = sm.add_constant(fut, has_constant = 'add')
future_pred = model.predict(X_future)[0]
print(f'OLS model: {future_pred}')


# 5. Gráfica
plt.figure(figsize=(12, 6))
plt.plot(df['Fecha'], df['Tasa de 10 AÑOS'], label='10-Year Treasury Yield (real)', color='navy')
plt.plot(df['Fecha'], df['Model_Pred'], label='Model Prediction', color='orange')
plt.fill_between(df['Fecha'], df['Lower'], df['Upper'], color='orange', alpha=0.2, label='±1 Std Error')

# Punto de predicción futura (2026)
plt.scatter(pd.to_datetime('17/09/2026', format = '%d/%m/%Y'), future_pred, color='darkorange', label='2026 Projection (dot plot walks)', zorder=5)
plt.axvline(x=pd.to_datetime('17/09/2026', format = '%d/%m/%Y'), color='gray', linestyle='--', alpha=0.5)

# Estética
plt.title("Modelo de tasa a 10 años basada en Dot Plot (FFR 2Y + LT) (Ordinary Least Squares)")
plt.xlabel("Fecha")
plt.yticks([0, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0])
plt.ylabel("Rendimiento (%)")
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()



### POLYNOMIAL ###
poly = PolynomialFeatures(degree = 3)
X_poly = poly.fit_transform(df[['FFR_2Y', 'SPREAD']])
X_poly = sm.add_constant(X_poly)

# model = LinearRegression().fit(X_poly, y)
model = sm.OLS(y, X_poly).fit()
df['Linear'] = model.predict(X_poly)

residuals = df['Tasa de 10 AÑOS'] - df['Linear']
std_error = residuals.std()
df['Upper'] = df['Linear'] + std_error
df['Lower'] = df['Linear'] - std_error

future_ffr_2y = 3.125  # tomado del gráfico como "walks the walk"
future_lt = 3.0
fut = pd.DataFrame([[future_ffr_2y, future_lt-future_ffr_2y]], columns=['FFR_2Y', 'SPREAD'])
X_future = poly.transform(fut)
future_pred = model.predict(X_future)[0]
print(f'Polynomial model: {future_pred}')

plt.figure(figsize=(12, 6))
plt.plot(df['Fecha'], df['Tasa de 10 AÑOS'], label='10-Year Treasury Yield (real)', color='navy')
plt.plot(df['Fecha']._append(pd.Series(pd.to_datetime('17/09/2026', format = '%d/%m/%Y'))), df['Linear']._append(pd.Series(future_pred)), label='Model Prediction', color='orange')
plt.fill_between(df['Fecha'], df['Lower'], df['Upper'], color='orange', alpha=0.2, label='±1 Std Error')

text = (
    f"""Dot plot in one-year time
    in the Fed 'walks the
    walk':
    Pred. = {future_pred:.3f}%
    End-27 = 3.125%
    End-28 = 3.125%
    -> 2Y ahead = 3.125%
    LT = 3%"""
    
)
ccy = (
    """Current dot plot:
    Pred. = 4.125%
    End-26 = 3.375%
    End-27 = 3.125%
    -> 2Y ahead = 3.188%
    LT = 3%"""
    
)
ffr = (
    f"""FFR: {df['Tasa de 10 AÑOS'].iloc[-1]:.2f}"""
)

plt.annotate(text, xy = (pd.to_datetime('17/09/2026', format = '%d/%m/%Y'), future_pred),
             xytext = (pd.to_datetime('17/09/2026', format = '%d/%m/%Y'), future_pred - 1.0),
             arrowprops = dict(arrowstyle = '->', color = 'black'), fontsize = 9, color = 'black',
             bbox = dict(boxstyle = 'round,pad=0.3', fc = 'white', ec = 'gray', lw = 1),
             ha = 'center', va = 'center')
plt.annotate(ccy, xy = (df.Fecha.iloc[-1], df.Linear.iloc[-1]),
             xytext = (df.Fecha.iloc[-8], df.Linear.iloc[-1] - 1.0),
             arrowprops = dict(arrowstyle = '->', color = 'black'), fontsize = 9, color = 'black',
             bbox = dict(boxstyle = 'round,pad=0.3', fc = 'white', ec = 'gray', lw = 1),
             ha = 'center', va = 'center')
plt.annotate(ffr, xy = (df.Fecha.iloc[-1], df['Tasa de 10 AÑOS'].iloc[-1]),
             xytext = (df.Fecha.iloc[-1], future_pred + 1.0),
             arrowprops = dict(arrowstyle = '->', color = 'black'), fontsize = 9, color = 'black',
             bbox = dict(boxstyle = 'round,pad=0.3', fc = 'white', ec = 'gray', lw = 1),
             ha = 'center', va = 'center')

# Punto de predicción futura (2026)
plt.scatter(pd.to_datetime('17/09/2026', format = '%d/%m/%Y'), future_pred, color = 'darkorange', label='2026 Projection (dot plot walks)', zorder=5)
plt.scatter(df.Fecha.iloc[-1], df.Linear.iloc[-1], c = 'k', zorder=5)
plt.axvline(x=pd.to_datetime('17/09/2026', format = '%d/%m/%Y'), color='gray', linestyle='--', alpha=0.5)

# Estética
# plt.title("Modelo de tasa a 10 años basada en Dot Plot (FFR 2Y + LT) (Polynomial Regession - 3th degree)")
plt.xlabel("Fecha")
plt.yticks([0, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0])
plt.ylabel("Rendimiento (%)")
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()


### RIDGE ###
# ridge_model = Ridge(alpha = 1.0, solver = 'sag').fit(df[['FFR_2Y', 'SPREAD']], y)
# df['Ridge'] = ridge_model.predict(df[['FFR_2Y', 'SPREAD']])
# future_pred = ridge_model.predict(fut)[0]
# print(f'Ridge model: {future_pred}')

# plt.figure(figsize=(12, 6))
# plt.plot(df['Fecha'], df['Tasa de 10 AÑOS'], label='10-Year Treasury Yield (real)', color='navy')
# plt.plot(df['Fecha'], df['Ridge'], label='Model Prediction', color='orange')
# # plt.fill_between(df['Fecha'], df['Lower'], df['Upper'], color='orange', alpha=0.2, label='±1 Std Error')

# # Punto de predicción futura (2026)
# plt.scatter(pd.to_datetime('17/09/2026', format = '%d/%m/%Y'), future_pred, color='darkorange', label='2026 Projection (dot plot walks)', zorder=5)
# plt.axvline(x=pd.to_datetime('17/09/2026', format = '%d/%m/%Y'), color='gray', linestyle='--', alpha=0.5)

# # Estética
# plt.title("Modelo de tasa a 10 años basada en Dot Plot (FFR 2Y + LT) (Ridge Regression)")
# plt.xlabel("Fecha")
# plt.yticks([0, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0])
# plt.ylabel("Rendimiento (%)")
# plt.legend()
# plt.grid(True)
# plt.tight_layout()
# plt.show()


### RANDOM FOREST ###
# rf = RandomForestRegressor(n_estimators = 100, max_depth = 4, random_state = 42)
# rf.fit(df[['FFR_2Y', 'SPREAD']], y)
# df['RF'] = rf.predict(df[['FFR_2Y', 'SPREAD']])

# future_pred = rf.predict(fut)[0]
# print(f'Random Forest model: {future_pred}')

# plt.figure(figsize=(12, 6))
# plt.plot(df['Fecha'], df['Tasa de 10 AÑOS'], label='10-Year Treasury Yield (real)', color='navy')
# plt.plot(df['Fecha'], df['RF'], label='Model Prediction', color='orange')
# # plt.fill_between(df['Fecha'], df['Lower'], df['Upper'], color='orange', alpha=0.2, label='±1 Std Error')

# # Punto de predicción futura (2026)
# plt.scatter(pd.to_datetime('17/09/2026', format = '%d/%m/%Y'), future_pred, color='darkorange', label='2026 Projection (dot plot walks)', zorder=5)
# plt.axvline(x=pd.to_datetime('17/09/2026', format = '%d/%m/%Y'), color='gray', linestyle='--', alpha=0.5)

# # Estética
# plt.title("Modelo de tasa a 10 años basada en Dot Plot (FFR 2Y + LT) (Random Forest)")
# plt.xlabel("Fecha")
# plt.yticks([0, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0])
# plt.ylabel("Rendimiento (%)")
# plt.legend()
# plt.grid(True)
# plt.tight_layout()
# plt.show()