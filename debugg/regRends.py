import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from statsmodels.tsa.arima.model import ARIMA
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import RandomForestRegressor
from sklearn.preprocessing import PolynomialFeatures

series = pd.read_excel('C:/Users/DATOS-INVERSIONES/OneDrive/Escritorio/AGG SPX ACWI Series.xlsx')
series.index = series.Date
series['log_agg'] = np.log(series.AGG / series.AGG.shift(+1)).fillna(0)
series['log_spx'] = np.log(series.SPX / series.SPX.shift(+1)).fillna(0)
series['log_acwi'] = np.log(series.ACWI / series.ACWI.shift(+1)).fillna(0)

log_agg = series[series.log_agg != 0].log_agg
split = int(len(log_agg)*0.8)
train, test = log_agg[:split], log_agg[split:]


model = ARIMA(train, order = (1, 0, 1))
model_fit = model.fit()

forecast = model_fit.get_forecast(steps = len(test))
pred_mean = forecast.predicted_mean
conf_int = forecast.conf_int()

print(forecast)

plt.figure(figsize = (12, 6))

# plt.plot(train.index, train, label='Entrenamiento')
plt.plot(test.index, test, label='Prueba', color='gray')
plt.plot(pred_mean.index, pred_mean, label='Pronóstico', color='red')
plt.fill_between(conf_int.index, conf_int.iloc[:,0], conf_int.iloc[:,1], color='pink', alpha=0.3)

# plt.plot(series.Date, series.log_agg, 'g--', label = 'Log AGG')
# plt.plot(series.Date, series.log_spx, 'b--', label = 'Log SPX')
# plt.plot(series.Date, series.log_acwi, 'y--', label = 'Log ACWI')

# plt.plot(series.Date, series.AGG, 'g--', label = 'AGG')
# plt.plot(series.Date, series.SPX, 'b--', label = 'SPX')
# plt.plot(series.Date, series.ACWI, 'y--', label = 'ACWI')

plt.legend(loc = 'best')
plt.xlabel('Date')
plt.ylabel('Credit')
plt.grid(True)

plt.tight_layout()
plt.show()