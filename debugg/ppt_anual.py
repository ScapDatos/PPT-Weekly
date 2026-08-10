import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
import seaborn as sns

### ===================================== ###
# data = pd.read_excel('c:/Users/DATOS-INVERSIONES/OneDrive/Escritorio/PPT_Anual/Files/SPX Estimados (Data Bloomberg) Abril 2024.xlsx')
# data = data[data.Date <= '2025-02-06'].reset_index(drop = True)
# print(data.tail())

# data.plot(x = 'Date')
# plt.show()


### ===================================== ###
# indexes = {
#     'S5INFT':'TI Sector Relative PE vs SPX',
#     'S5FINL':'Financials Sector Relative PE vs SPX',
#     'S5COND':'C. Discretionary Sector Relative PE vs SPX',
#     'S5HLTH':'Healthcare Sector Relative PE vs SPX',
#     'S5TELS':'Comms Sector Relative PE vs SPX',
#     'S5INDU':'Industrials Sector Relative PE vs SPX',
#     'S5CONS':'C. Staples Sector Relative PE vs SPX',
#     'S5ENRS':'Energy Sector Relative PE vs SPX',
#     'S5UTIL':'Utilities Sector Relative PE vs SPX',
#     'S5RLST':'Real State Sector Relative PE vs SPX',
#     'S5MATR':'Materials Sector Relative PE vs SPX',
#     'NDX':'Nasdaq 100 Index Relative PE vs SPX',
#     'SPW':'S&P 500 EW Index Relative PE vs SPX',
#     'SML':'S&P 600 Small Cap Index Relative PE vs SPX'
# }
# data = pd.read_excel('C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Asset Allocation/Presentación Asset Allocation/Files_PPT_Anual/pp_anual_us_val_relativa.xlsx', sheet_name = 'Hoja1 (2)')

# for idx, (index, title) in enumerate(indexes.items()):
#     plt.rcParams['axes.xmargin'] = 0
#     fig, ax = plt.subplots(figsize = (10.77/2.54, 6.3/2.54))
    
#     ax.plot(
#         data.Date,
#         data[f'{index} Index'],
#         'k-',
#         label = f"Last: {data[f'{index} Index'].iloc[-1]:.1f}x"
#     )
#     ax.axhline(
#         y = data[f'Unnamed: {(idx+1) * 2}'].iloc[0],
#         c = '#0E8388',
#         ls = '--',
#         label = rf"$\mu$: {data[f'Unnamed: {(idx+1) * 2}'].iloc[0]:.1f}x"
#     )
    
#     ax.legend(
#         loc = 'best',
#         frameon = False,
#     )
#     ax.xaxis.set_major_locator(mdates.MonthLocator(interval = 12))
#     ax.xaxis.set_major_formatter(mdates.DateFormatter(r'$\bf{%b}$-$\bf{%y}$'))
    
#     ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: rf'$\bf{x:.1f}x$'))
    
#     ax.grid(axis = 'y', color = 'lightgray', lw = 2, alpha = 0.5)
#     ax.tick_params(axis = 'x', rotation = 90)
#     ax.tick_params(axis = 'both', labelsize = 8)
    
#     ax.spines['right'].set_visible(False)
#     ax.spines['left'].set_visible(False)
#     ax.spines['bottom'].set_visible(False)
#     ax.spines['top'].set_visible(False)
#     ax.set_axisbelow(True)
#     ax.set_title(title, size = 10, fontweight = 'bold', bbox = dict(
#                     facecolor = 'w',
#                     edgecolor = 'none'
#                 ),)
    
#     plt.tight_layout(pad = 0.09)
#     fig.savefig(f'C:/Users/DATOS-INVERSIONES/OneDrive/Escritorio/PPT_Anual/Plots/{idx+1}_{index}.png', dpi = 300)
#     plt.close()


### ===================================== ###
data = pd.read_excel('C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Asset Allocation/Presentación Asset Allocation/Files_PPT_Anual/pp_anual_bevens.xlsx', skiprows = 3)

plt.rcParams['axes.xmargin'] = 0.0 # PLOT1: 0.1
fig, ax = plt.subplots(figsize = (15.5/2.54, 12.5/2.54))

# PLOT 1
data = data[data.Date >= '30/12/2022'].reset_index(drop = True)
# data = data[['Date', '10BEvens', 'RHS0']].dropna(ignore_index = True)

ax2 = ax.twinx()
ax.plot(
    data.Date,
    data['10BEvens'],
    'k-',
    label = rf'10Y BE %: $\bf{data["10BEvens"].iloc[-1]:.2f}$'
)
ax2.plot(
    data.Date,
    data.RHS0,
    c = '#97B7B3',
    label = rf'[RHS] 10Y TIPS %: $\bf{data.RHS0.iloc[-1]:.2f}$'
)
ax.axhline(
    y = data['10BEvens'].mean(),
    # c = '#0E8388',
    c = 'k',
    ls = '--',
    label = rf'$\mu$ 10Y BE %: $\bf{data["10BEvens"].mean():.2f}$'
)
ax.axvline(
    x = pd.to_datetime('31/12/2023', format = '%d/%m/%Y'),
    c = 'k',
    ls = '--',
    lw = 1,
    ymax = 0.95
)
ax.axvline(
    x = pd.to_datetime('31/12/2024', format = '%d/%m/%Y'),
    c = 'k',
    ls = '--',
    lw = 1,
    ymax = 0.95
)
ax.axvline(
    x = pd.to_datetime('31/12/2025', format = '%d/%m/%Y'),
    c = 'k',
    ls = '--',
    lw = 1,
    ymax = 0.95,
)

ax.legend(
        loc = f'best',
        frameon = False,
        ncol = 3
)
ax2.legend(
        loc = f'lower right',
        # loc = 'upper left',
        frameon = False,
        ncol = 3
)

ax.set_title('10Y US BEs vs 10YTIPS', size = 14, fontweight = 'bold', family = 'Arial')

ax.xaxis.set_major_locator(mdates.MonthLocator(interval = 2))
ax.xaxis.set_major_formatter(mdates.DateFormatter(r'$\bf{%b}$-$\bf{%y}$'))
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: rf'$\bf{x:.1f}$'))

ax2.tick_params(axis = 'both', labelsize = 11, labelcolor = '#97B7B3', labelfontfamily = 'Arial')
ax2.spines['right'].set_visible(False)
ax2.spines['top'].set_visible(False)
ax2.spines['bottom'].set_visible(False)
ax2.spines['left'].set_visible(False)
ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: rf'$\bf{x:.1f}$'))

ax.grid(axis = 'y', color = 'lightgray', lw = 2, alpha = 0.5)
ax.tick_params(axis = 'x', rotation = 90)
ax.tick_params(axis = 'both', labelsize = 11, labelfontfamily = 'Arial')
ax.set_axisbelow(True)

ax.spines['right'].set_visible(False)
ax.spines['top'].set_visible(False)
ax.spines['bottom'].set_visible(False)
ax.spines['left'].set_visible(False)

plt.tight_layout(pad = 0.07)
fig.savefig(f'C:/Users/DATOS-INVERSIONES/OneDrive/Escritorio/PPT_Anual/Plots/BEsvsTIPS.png', dpi = 300)
plt.show()

# # PLOT 2
# data = data[data.Date >= '01/01/2013'].reset_index(drop = True)

# ax.plot(
#     data.Date,
#     data['10BEvens'],
#     'k-',
#     label = rf'Last: $\bf{data["10BEvens"].iloc[-1]:.2f}$'
# )
# ax.axhline(
#     y = data['10BEvens'].mean(),
#     c = '#0E8388',
#     label = rf'$\mu$: $\bf{data["10BEvens"].mean():.1f}$'
# )

# ax.legend(
#         loc = f'best',
#         frameon = False,
#         ncol = 3
# )

# ax.set_title('10Y US Breakevens 2013 - 2026', size = 14, fontweight = 'bold', family = 'Arial')

# ax.xaxis.set_major_locator(mdates.MonthLocator(interval = 6))
# ax.xaxis.set_major_formatter(mdates.DateFormatter(r'$\bf{%b}$-$\bf{%y}$'))
# ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: rf'$\bf{x:.1f}$'))

# ax.grid(axis = 'y', color = 'lightgray', lw = 2, alpha = 0.5)
# ax.tick_params(axis = 'x', rotation = 90)
# ax.tick_params(axis = 'both', labelsize = 11, labelfontfamily = 'Arial')
# ax.set_axisbelow(True)

# ax.spines['right'].set_visible(False)
# ax.spines['top'].set_visible(False)
# ax.spines['bottom'].set_visible(False)
# ax.spines['left'].set_visible(False)

# plt.tight_layout(pad = 0.05)
# fig.savefig(f'C:/Users/DATOS-INVERSIONES/OneDrive/Escritorio/PPT_Anual/Plots/Breakevens.png', dpi = 300)
# plt.show()

# # PLOT 3
# data = data[data.D1 >= '31/10/2022'].reset_index(drop = True)
# # data = data[['D1', 'WTI', 'RHS1']].dropna(ignore_index = True)
# ll = '\n'

# ax2 = ax.twinx()
# ax.plot(
#     data.D1,
#     data.WTI,
#     'k-',
#     label = rf'WTI(S):{ll}$\bf{data.WTI.iloc[-1]:.1f}$'
# )
# ax2.plot(
#     data.D1,
#     data.RHS1,
#     c = '#97B7B3',
#     label = rf'[RHS] 10Y(%):{ll}$\bf{data.RHS1.iloc[-1]:.1f}$'
# )
# ax.axvline(
#     x = pd.to_datetime('31/12/2023', format = '%d/%m/%Y'),
#     c = 'k',
#     ls = '--',
#     lw = 1,
#     ymax = 0.95
# )
# ax.axvline(
#     x = pd.to_datetime('31/12/2024', format = '%d/%m/%Y'),
#     c = 'k',
#     ls = '--',
#     lw = 1,
#     ymax = 0.95
# )
# ax.axvline(
#     x = pd.to_datetime('31/12/2025', format = '%d/%m/%Y'),
#     c = 'k',
#     ls = '--',
#     lw = 1,
#     ymax = 0.95
# )

# ax.legend(
#         loc = f'best',
#         frameon = False,
#         ncol = 3
# )
# ax2.legend(
#         loc = f'upper right',
#         frameon = False,
#         ncol = 3
# )

# ax.set_title('Crude Oil vs 10Y Rate', size = 14, fontweight = 'bold', family = 'Arial')

# ax.xaxis.set_major_locator(mdates.MonthLocator(interval = 2))
# ax.xaxis.set_major_formatter(mdates.DateFormatter(r'$\bf{%b}$-$\bf{%y}$'))
# ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: rf'$\bf{x:.0f}$'))

# ax2.tick_params(axis = 'both', labelsize = 11, labelcolor = '#97B7B3', labelfontfamily = 'Arial')
# ax2.spines['right'].set_visible(False)
# ax2.spines['top'].set_visible(False)
# ax2.spines['bottom'].set_visible(False)
# ax2.spines['left'].set_visible(False)
# ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: rf'$\bf{x:.1f}$'))

# ax.grid(axis = 'y', color = 'lightgray', lw = 2, alpha = 0.5)
# ax.tick_params(axis = 'x', rotation = 90)
# ax.tick_params(axis = 'both', labelsize = 11, labelfontfamily = 'Arial')
# ax.set_axisbelow(True)

# ax.spines['right'].set_visible(False)
# ax.spines['top'].set_visible(False)
# ax.spines['bottom'].set_visible(False)
# ax.spines['left'].set_visible(False)

# plt.tight_layout(pad = 0.07)
# fig.savefig(f'C:/Users/DATOS-INVERSIONES/OneDrive/Escritorio/PPT_Anual/Plots/Crude_Oil.png', dpi = 300)
# plt.show()




### ===================================== ###
# data = pd.read_excel('c:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Fixed Income/Spreads Fixed Income.xlsx', skiprows = 19, header = None)
# data = pd.read_excel('c:/Users/DATOS-INVERSIONES/OneDrive/Escritorio/PPT_Anual/Files/pp_anual_fixinc_spreads.xlsx', sheet_name = 'Sheet1', skiprows = 19, header = None)
# data.columns = [f'Col_{idx}' for idx in range(len(data.columns))]


# PLOT 1
# data = data.drop([f'Col_{idx}' for idx in range(8, 32)], axis = 1).reset_index(drop = True)

# col = 'Col_7' #col_3
# fliers = []

# for col in data.columns[1:]:
#     q1 = data[col].quantile(0.25)
#     q3 = data[col].quantile(0.75)
#     iqr = q3 - q1
#     max_val = data[col].max()
#     sup = q3 + 1.5*iqr if q3 + 1.5*iqr < max_val else max_val
    
#     outliers = data[data[col] > sup][col]
#     fliers.append(outliers.max() if len(outliers) > 0 else np.nan)

# plt.rcParams['axes.xmargin'] = 0.0 # PLOT1: 0.1
# fig, ax = plt.subplots(figsize = (15.5/2.54, 12.5/2.54))

# box_data = [data[col].dropna() for col in data.columns[1:]]

# ax.boxplot(
#     box_data,
#     tick_labels = ['US 10Yr', 'EUR 10\nYr', 'US Inv´t\nGrade', 'US High\n Yield', 'EM Bonds', 'Global\nHigh\nYield', 'US Lev\nLoans'],
#     showfliers = False
# )

# ax.scatter(
#     [i+1 for i in range(7)],
#     [data[col].mean() for col in data.columns[1:]],
#     marker = 'x',
#     c= 'gray',
#     label = 'Average'
# )

# ax.scatter(
#     [i+1 for i in range(7)],
#     [data[col].iloc[-1] for col in data.columns[1:]],
#     marker = 'D',
#     c= 'c',
#     label = 'Current'
# )

# ax.scatter(
#     [i+1 for i in range(7)],
#     fliers,
#     marker = 'o',
#     # c= 'k',
#     edgecolors = 'k',
#     facecolors = 'none',
#     label = 'Max. Outlier'
# )

# ax.legend(
#         loc = f'best',
#         frameon = False,
#         # ncol = 3
# )

# ax.set_title('Fixed Income Yield', size = 14, fontweight = 'bold', family = 'Arial')

# ax.yaxis.set_major_locator(plt.MultipleLocator(2))
# ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: rf'$\bf{x:.0f}\%$'))

# ax.grid(axis = 'y', color = 'lightgray', lw = 2, alpha = 0.5)
# ax.tick_params(axis = 'both', labelsize = 11, labelfontfamily = 'Arial')
# ax.set_xticklabels(ax.get_xticklabels(), fontweight='bold', fontsize = 12)
# ax.set_axisbelow(True)

# ax.spines['right'].set_visible(False)
# ax.spines['top'].set_visible(False)
# ax.spines['bottom'].set_visible(False)
# ax.spines['left'].set_visible(False)

# plt.tight_layout(pad = 0.05)
# fig.savefig(f'C:/Users/DATOS-INVERSIONES/OneDrive/Escritorio/PPT_Anual/Plots/FixIncYield.png', dpi = 300)
# plt.show()


# PLOT 2
# data = data[['Col_9', 'Col_10', 'Col_11', 'Col_12', 'Col_13']].reset_index(drop = True)
# fliers = []

# for col in data.columns:
#     q1 = data[col].quantile(0.25)
#     q3 = data[col].quantile(0.75)
#     iqr = q3 - q1
#     max_val = data[col].max()
#     sup = q3 + 1.5*iqr if q3 + 1.5*iqr < max_val else max_val
    
#     outliers = data[data[col] > sup][col]
#     fliers.append(outliers.max() if len(outliers) > 0 else np.nan)

# # tick_labels = ['US Invt\nGrade', 'US High\nYield', 'EM Bonds', 'Global High\nYield', 'US Lev\nLoans']
# plt.rcParams['axes.xmargin'] = 0.0 # PLOT1: 0.1
# fig, ax = plt.subplots(figsize = (15.5/2.54, 12.5/2.54))

# box_data = [data[col].dropna() for col in data.columns]

# ax.boxplot(
#     box_data,
#     tick_labels = ['US Invt\nGrade', 'US High\nYield', 'EM Bonds', 'Global High\nYield', 'US Lev\nLoans'],
#     showfliers = False
# )

# ax.scatter(
#     [i+1 for i in range(5)],
#     [data[col].mean() for col in data.columns],
#     marker = 'x',
#     c= 'gray',
#     label = 'Average'
# )

# ax.scatter(
#     [i+1 for i in range(5)],
#     [data[col].iloc[-1] for col in data.columns],
#     marker = 'D',
#     c= 'c',
#     label = 'Current'
# )

# ax.scatter(
#     [i+1 for i in range(5)],
#     fliers,
#     marker = 'o',
#     edgecolors = 'k',
#     facecolors = 'none',
#     label = 'Max. Outlier'
# )

# ax.legend(
#         loc = f'best',
#         frameon = False,
#         # ncol = 3
# )

# ax.set_title('Credit Spread (bp)', size = 14, fontweight = 'bold', family = 'Arial')

# ax.yaxis.set_major_locator(plt.MultipleLocator(300))
# ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: rf'$\bf{x:.0f}$'))

# ax.grid(axis = 'y', color = 'lightgray', lw = 2, alpha = 0.5)
# ax.tick_params(axis = 'both', labelsize = 10, labelfontfamily = 'Arial')
# ax.set_xticklabels(ax.get_xticklabels(), fontweight='bold', fontsize = 12)
# ax.set_axisbelow(True)

# ax.spines['right'].set_visible(False)
# ax.spines['top'].set_visible(False)
# ax.spines['bottom'].set_visible(False)
# ax.spines['left'].set_visible(False)

# plt.tight_layout(pad = 0.05)
# fig.savefig(f'C:/Users/DATOS-INVERSIONES/OneDrive/Escritorio/PPT_Anual/Plots/CreditSpread.png', dpi = 300)
# plt.show()



### ===================================== ###
# data = pd.read_excel('c:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Notas Estructuradas/Notas Presentación (New).xlsx', sheet_name = 'VIX', skiprows = 1)
# data = data[data.FECHA <= '03/02/2025']

# plt.scatter(data.VIX, data.SPX*100)
# plt.show()




### ===================================== ### GRAFICA DE CUPONES E INTERES (SOLO VERIFICAR DATOS)
# isin_gb = pd.read_excel('C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Notas Estructuradas/Graficas_Barreras.xlsx', sheet_name = 'NOTAS')
# isin_coupons = pd.read_excel('C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Notas Estructuradas/Graficas_Barreras.xlsx', sheet_name = 'CUPONES (2)')

# isin_gb_list = isin_gb[isin_gb.BANK == 'JB'].ISIN.unique()
# isin_coupons_list = isin_coupons[isin_coupons.BANK == 'JB'].ISIN.unique()

# nl = []
# for igb in isin_gb_list:
#     if igb not in isin_coupons_list:
#         vals = isin_gb[isin_gb.ISIN == igb][['ISIN', 'Estatus']].values.tolist()
#         nl.append(vals)

# print(nl)



### ===================================== ###
# data = pd.read_excel('C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Notas Estructuradas/Sharpe y Correlación.xlsx', sheet_name = 'Rendimiento', skiprows = 7)
# data = data.drop('Date', axis = 1)
# # print(data.head())
# # exit()

# corr = data[['SPX', 'SX5E', 'NKY', 'RTY', 'NDX', 'SX7E', 'SPW', 'MID']].corr()
# # corr = data.corr()
# mask = np.triu(np.ones_like(corr, dtype = bool))

# fig, ax = plt.subplots(figsize = (15.5/2.54, 12.5/2.54))
# # cmap = sns.diverging_palette(230, 20, as_cmap = True)
# # cmap = sns.diverging_palette(20, 230, as_cmap = True)
# cmap = sns.diverging_palette(100, 20, s = 100, l = 40, as_cmap = True)
# # cmap = sns.diverging_palette(100, 20, l = 40, as_cmap = True)

# sns.heatmap(corr, cmap = cmap, mask = mask, annot = True, fmt = '.2f', square = True)
# # sns.heatmap(corr, cmap = cmap, mask = mask, fmt = '.2f', square = True, xticklabels = 1, yticklabels = 1)

# ax.set_title('Indexes Correlation', size = 14, fontweight = 'bold', family = 'Arial')

# ax.tick_params(axis = 'both', labelsize = 12, labelfontfamily = 'Arial')
# ax.set_xticklabels(ax.get_xticklabels(), fontweight='bold')
# ax.set_yticklabels(ax.get_yticklabels(), fontweight='bold')
# # ax.set_axisbelow(True)

# ax.spines['right'].set_visible(False)
# ax.spines['top'].set_visible(False)
# ax.spines['bottom'].set_visible(False)
# ax.spines['left'].set_visible(False)

# plt.tight_layout(pad = 0.05)
# fig.savefig(f'C:/Users/DATOS-INVERSIONES/OneDrive/Escritorio/PPT_Anual/Plots/corr_current.png', dpi = 300)
# # fig.savefig(f'C:/Users/DATOS-INVERSIONES/OneDrive/Escritorio/PPT_Anual/Plots/corr_full.png', dpi = 300)
# plt.show()




### ===================================== ###
# import os
# files1 = os.listdir('C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Contributtion/Historicos/Contributtion')
# files2 = os.listdir('C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Contributtion')

# dates = pd.date_range('01/01/2025', '03/01/2026').strftime('%Y-%m-%d')

# assets = []
# for date in dates:
#     for file1 in files1:
#         if date in file1:
#             assets.append(file1)
#     # for file2 in files2:
#     #     assets.append(file2)

# # print(assets)

# navs, navs_wo_city, dts = [], [], []

# for asset in assets:
#     data = pd.read_excel(f'C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Contributtion/Historicos/Contributtion/{asset}', sheet_name = 'CONSOLIDADO')
#     n1 = data[data.ASSET_CLASS != 'COLATERAL'].MKT_VALUE_USD.sum()
#     n2 = data[(data.ASSET_CLASS != 'COLATERAL') & (data.SHORT_NAME != 'HCITY')].MKT_VALUE_USD.sum()
    
#     navs.append(n1)
#     navs_wo_city.append(n2)
#     dts.append(asset[:10])
    
# d = {
#     'Date': dts,
#     'NAV': navs,
#     'NAV WO HCITY': navs_wo_city
# }

# dd = pd.DataFrame(d)
# dd.to_excel('C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Contributtion/hist_navs.xlsx', index = False)
# print(dd)








### ===================================== ###
# path = 'C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Notas Estructuradas/Métricas Índices con Bancos Europa.xlsm'
# data = pd.read_excel(path, sheet_name = 'ACWI', skiprows = 2)[['Date', 'ACWI']]
# # data = data.dropna()

# data['ACWI'] = data.ACWI.replace('x', '').astype(float) # REVISAR LOS ERRORES CUANDO SON #NA #NA
# # data = data[(data.Date >= '15/01/2010')].reset_index(drop = True)
# # print(data[data.Date<='20/02/2023'].tail(15))

# indexes = ['ACWI']

# for index in indexes:
#     plt.rcParams['axes.xmargin'] = 0
#     # fig, ax = plt.subplots(figsize = (16.3/2.54, 16.3/2.54))
#     fig, ax = plt.subplots(figsize = (15.5/2.54, 12.5/2.54))
#     y = data[f'{index}']
#     y = y.replace('#N/A Requesting Data...', np.nan)

#     last_val = y.iloc[-1]
#     x = data.Date

#     ax.plot(x, y, 'k-')

#     mean = y.mean()
#     std = y.std()

#     ax.axhline(y = mean, ls = '--', c = '#333E50')
#     ax.axhline(y = mean + std, ls = '-.', c = '#0E8388')
#     ax.axhline(y = mean - std, ls = '-.', c = '#0E8388')
#     ax.axhline(y = mean + 2*std, ls = '-.', c = '#6DAAAA')
#     ax.axhline(y = mean - 2*std, ls = '-.', c = '#6DAAAA')

#     last_date = x.iloc[-1]

#     ax.axhline(y = last_val * 0.8, ls = '-', c = '#FF0000')
#     ax.axhline(y = last_val * 0.75, ls = '-', c = '#DC143C')
#     ax.axhline(y = last_val * 0.7, ls = '-', c = '#990000')
#     ax.axhline(y = last_val * 0.6, ls = '-', c = '#800000')

#     leg = {'$\mu$':mean, '-1$\sigma$':mean-std, '+1$\sigma$':mean+std, '-2$\sigma$':mean-2*std, '+2$\sigma$':mean+2*std,
#         '20%':last_val*0.8, '25%':last_val*0.75, '30%':last_val*0.7, '40%':last_val*0.6, index:last_val}
#     leg = dict(sorted(leg.items(), key = lambda item: item[1], reverse = True))
#     colors = {'$\mu$':'#333E50', '-1$\sigma$':'#0E8388', '+1$\sigma$':'#0E8388', '-2$\sigma$':'#6DAAAA', '+2$\sigma$':'#6DAAAA',
#               '20%':'#FF0000', '25%':'#DC143C', '30%':'#990000', '40%':'#800000', index:'white'}
    
#     ax.xaxis.set_major_locator(mdates.MonthLocator(interval = 12))
#     ax.xaxis.set_major_formatter(mdates.DateFormatter(r'$\bf{%b}$-$\bf{%y}$'))

#     ax.yaxis.set_major_locator(plt.MultipleLocator(5))
#     ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: rf'$\bf{x:.0f}.0x$'))

#     ax.spines['right'].set_visible(False)
#     ax.spines['top'].set_visible(False)
#     ax.set_axisbelow(True)
#     ax.tick_params(axis = 'x', rotation = 90)

#     y_step = 0.05
#     y_total = (len(leg) - 1) * y_step
#     y_start = 0.5 + y_total / 2
#     txts = []

#     plt.subplots_adjust(left=0.088, right=0.834, top=1, bottom=0.126)
#     for i, (key, val) in enumerate(leg.items()):
#         y_pos = y_start - i * y_step
#         t = ax.text(
#             x = 1.02,
#             y = y_pos,
#             s = f'{key}: {val:.1f}x',
#             transform = ax.transAxes,
#             color = colors[key],
#             bbox = dict(
#                 facecolor = 'none' if key != index else 'black',
#                 edgecolor = 'none',
#                 alpha = 0 if key != index else 1
#             ),
#             va = 'center', ha = 'left'
#         )
#         txts.append(t)


#     fig.canvas.draw()
#     renderer = fig.canvas.get_renderer()
#     max_w = max(t.get_window_extent(renderer = renderer).width for t in txts)
#     wfp = fig.get_size_inches()[0] * fig.dpi
#     margin_right = max_w / wfp
    
#     plt.show()

#     fig.savefig(f'C:/Users/DATOS-INVERSIONES/OneDrive/Escritorio/PPT_Anual/Plots/{index}.png', dpi = 300)





### ===================================== ###
# path = 'C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Notas Estructuradas/Métricas Índices con Bancos Europa.xlsm'

# data = pd.read_excel('C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Notas Estructuradas/Sharpe y Correlación.xlsx', sheet_name = 'PE', skiprows = 7)

# indexes = ['SPW', 'MID']

# for index in indexes:
#     plt.rcParams['axes.xmargin'] = 0
#     # fig, ax = plt.subplots(figsize = (16.3/2.54, 16.3/2.54))
#     fig, ax = plt.subplots(figsize = (15.5/2.54, 12.5/2.54))
#     y = data[f'{index} Index']
#     y = y.replace('#N/A #N/A', np.nan)

#     last_val = y.iloc[-1]
#     # y = winsorize(y.astype(float), limits = (0.05, 0.1)) if index == 'RTY' else y
#     x = data.Date

#     ax.plot(x, y, 'k-')

#     mean = y.mean()
#     std = y.std()

#     ax.axhline(y = mean, ls = '--', c = '#333E50')
#     ax.axhline(y = mean + std, ls = '-.', c = '#0E8388')
#     ax.axhline(y = mean - std, ls = '-.', c = '#0E8388')
#     ax.axhline(y = mean + 2*std, ls = '-.', c = '#6DAAAA')
#     ax.axhline(y = mean - 2*std, ls = '-.', c = '#6DAAAA')

#     last_date = x.iloc[-1]

#     ax.axhline(y = last_val * 0.8, ls = '-', c = '#FF0000')
#     ax.axhline(y = last_val * 0.75, ls = '-', c = '#DC143C')
#     ax.axhline(y = last_val * 0.7, ls = '-', c = '#990000')
#     ax.axhline(y = last_val * 0.6, ls = '-', c = '#800000')

#     leg = {'$\mu$':mean, '-1$\sigma$':mean-std, '+1$\sigma$':mean+std, '-2$\sigma$':mean-2*std, '+2$\sigma$':mean+2*std,
#         '20%':last_val*0.8, '25%':last_val*0.75, '30%':last_val*0.7, '40%':last_val*0.6, index:last_val}
#     leg = dict(sorted(leg.items(), key = lambda item: item[1], reverse = True))
#     colors = {'$\mu$':'#333E50', '-1$\sigma$':'#0E8388', '+1$\sigma$':'#0E8388', '-2$\sigma$':'#6DAAAA', '+2$\sigma$':'#6DAAAA',
#               '20%':'#FF0000', '25%':'#DC143C', '30%':'#990000', '40%':'#800000', index:'white'}
    
#     ax.xaxis.set_major_locator(mdates.MonthLocator(interval = 12))
#     ax.xaxis.set_major_formatter(mdates.DateFormatter(r'$\bf{%b}$-$\bf{%y}$'))

#     ax.yaxis.set_major_locator(plt.MultipleLocator(5))
#     ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: rf'$\bf{x:.0f}.0x$'))

#     ax.spines['right'].set_visible(False)
#     ax.spines['top'].set_visible(False)
#     ax.set_axisbelow(True)
#     ax.tick_params(axis = 'x', rotation = 90)

#     y_step = 0.05
#     y_total = (len(leg) - 1) * y_step
#     y_start = 0.5 + y_total / 2
#     txts = []

#     plt.subplots_adjust(left=0.088, right=0.858, top=1, bottom=0.126)
#     for i, (key, val) in enumerate(leg.items()):
#         y_pos = y_start - i * y_step
#         t = ax.text(
#             x = 1.02,
#             y = y_pos,
#             s = f'{key}: {val:.1f}x',
#             transform = ax.transAxes,
#             color = colors[key],
#             bbox = dict(
#                 facecolor = 'none' if key != index else 'black',
#                 edgecolor = 'none',
#                 alpha = 0 if key != index else 1
#             ),
#             va = 'center', ha = 'left'
#         )
#         txts.append(t)


#     fig.canvas.draw()
#     renderer = fig.canvas.get_renderer()
#     max_w = max(t.get_window_extent(renderer = renderer).width for t in txts)
#     wfp = fig.get_size_inches()[0] * fig.dpi
#     margin_right = max_w / wfp
    
#     # plt.show()

#     fig.savefig(f'C:/Users/DATOS-INVERSIONES/OneDrive/Escritorio/PPT_Anual/Plots/{index}_FwdPE.png', dpi = 300)








### ===================================== ###
# data = pd.read_excel('C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Asset Allocation/Presentación Asset Allocation/Files_PPT_Anual/pp_anual_macro_us.xlsx', skiprows = 5, sheet_name = 'Valores%')

# idx = 0
# for i, col in enumerate(data.columns):
#     if col == 'Term Premia':
#         idx = i
#         print(i, col)
        



# data = data[data.columns[idx-1:idx+1]].dropna(ignore_index = True)
# print(data.tail())