import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.interpolate import interp1d
from statsmodels.tsa.arima.model import ARIMA
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import RandomForestRegressor
from sklearn.preprocessing import PolynomialFeatures

path = 'C:/Users/DATOS-INVERSIONES/OneDrive/Escritorio'

data = pd.read_excel(f'{path}/SOFR.xlsx', sheet_name = 'SOFR')#.dropna(ignore_index = True)
fr = pd.read_excel(f'{path}/SOFR.xlsx', sheet_name = 'Futures')
fr = fr[fr.Reset_Date < pd.to_datetime('01/01/2028')].reset_index(drop = True)

# plt.figure(figsize = (15, 6))
# plt.plot(data.Reset_Date, data['1MTSOFR']*100, c = 'g', label = '1M SOFR')
# plt.plot(fr.Reset_Date, fr['1mTSFR']*100, c = 'r', label = 'Forward 1M')
# plt.plot(data.Reset_Date, data['3MTSOFR']*100, c = 'b', label = '3M SOFR')
# plt.plot(fr.Reset_Date, fr['3mTSFR']*100, c = 'r', label = 'Forward 3M')

# plt.title('Curva SOFR Tasa 1 y 3 Meses')
# plt.xlabel('Fecha')
# plt.ylabel('Tasa SOFR (%)')
# plt.legend(loc = 'best')
# plt.grid(True)
# plt.tight_layout()
# plt.show()

# daily_dates = pd.date_range(start = fr.Reset_Date.min(), end = fr.Reset_Date.max(), freq = 'D')

# date_nums = (fr.Reset_Date - fr.Reset_Date.min()).dt.days
# inter_1m = interp1d(date_nums, fr['1mTSFR'], kind = 'cubic')
# inter_3m = interp1d(date_nums, fr['3mTSFR'], kind = 'cubic')

# daily_days = (daily_dates - fr.Reset_Date.min()).days
# sofr1mfwrd = inter_1m(daily_days)
# sofr3mfwrd = inter_3m(daily_days)

# curve = pd.DataFrame({
#     'Date': daily_dates,
#     'SOFR1MFWD': sofr1mfwrd,
#     'SOFR3MFWD': sofr3mfwrd
# })

# print(curve.head())


# plt.figure(figsize = (12, 6))
# plt.plot(curve.Date, curve.SOFR1MFWD*100, label = 'SOFR 1M Forward')
# plt.plot(curve.Date, curve.SOFR3MFWD*100, ls = '--', label = 'SOFR 3M Forward')
# plt.title('Curva Forward de la SOFR (Interpolada)')
# plt.xlabel('Fecha')
# plt.ylabel('Tasa (%)')
# plt.grid(True)
# plt.legend(loc = 'best')
# plt.tight_layout()
# plt.show()


### REGRESIONES ###
# data['days'] = (data.Reset_Date - data.Reset_Date.min()).dt.days
# X = data[['days']]
# y = data['1MTSOFR']

# ### Linear Regression ###
# lreg = LinearRegression()
# lreg.fit(X, y)
# data['Linear'] = lreg.predict(X)

# ### Polynomial Regression ###
# poly = PolynomialFeatures(degree = 5)
# X_poly = poly.fit_transform(X)
# polyreg = LinearRegression()
# polyreg.fit(X_poly, y)
# data['Poly'] = polyreg.predict(poly.transform(X))


# ### Predictions ###
# future_dates = pd.date_range(start = data.Reset_Date.max() + pd.Timedelta(days = 1), periods = 730, freq = 'D')
# future_days = (future_dates - data.Reset_Date.min()).days.values.reshape(-1, 1)

# linear = lreg.predict(future_days)
# polynomial = polyreg.predict(poly.transform(future_days))


# plt.figure(figsize = (14, 6))
# plt.plot(data.Reset_Date, y, label = 'SOFR 1M', alpha = 0.6)
# plt.plot(data.Reset_Date, data.Linear, ls = '--', label = 'Linear Regression')
# plt.plot(future_dates, linear, ls = '--', label = 'Linear Regression')
# plt.plot(data.Reset_Date, data.Poly, ls = '--', label = 'Poly Regression')
# plt.plot(future_dates, polynomial, ls = '-.', label = 'Polynomial Regression')
# plt.title('Predicción de la SOFR 1M vs Curva Forward del Mercado')
# plt.xlabel('Fecha')
# plt.ylabel('Tasa (%)')
# plt.legend()
# plt.grid(True)
# plt.tight_layout()
# plt.show()

### RANDOM FOREST ###
monthly = data.set_index('Reset_Date').resample('ME').last().reset_index()
monthly['Months'] = (monthly.Reset_Date - monthly.Reset_Date.min()).dt.days // 30

X = monthly[['Months']]
y = monthly['1MTSOFR']

rf = RandomForestRegressor(n_estimators = 10, random_state = 42)
rf.fit(X, y)
monthly['RF'] = rf.predict(X)

future_dates = pd.date_range("2025-09-30", "2027-09-30", freq="ME")
future_months = (future_dates - monthly.Reset_Date.min()).days // 30

future = pd.DataFrame({
    'Reset_Date': future_dates,
    'Months': future_months
})

future['FWD1M'] = rf.predict(future[['Months']])
print(future.tail())

plt.figure(figsize = (12, 6))
plt.plot(monthly.Reset_Date, monthly['1MTSOFR']*100, label = '1M Historic')
plt.plot(monthly.Reset_Date, monthly.RF*100, label = '1M Reg. RF')
plt.plot(future.Reset_Date, future.FWD1M*100, ls = '--', label = '1M SOFR RF')
plt.plot(fr.Reset_Date, fr['1mTSFR']*100, label = '1M SOFR Forward')
plt.xlabel("Fecha")
plt.ylabel("Tasa (%)")
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()