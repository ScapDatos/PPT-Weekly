### ============================================= ###
### Notas:
###     - Actualizar la fecha al viernes inmediato anterior.
###     - Actualizar la fecha del Asset anterior.
### ============================================= ###

import pandas as pd
from debugg.functions import process_portfs
from config import DATE, ytdDATE, PARENT_PATH

dates = [ytdDATE, DATE]

for idx, date in enumerate(dates):
    if idx == 1:
        vsDate = ((pd.to_datetime(date, format = '%Y-%m-%d') - pd.Timedelta(weeks = 1))).strftime('%Y-%m-%d')
        fname = 'Weekly'
    else:
        vsDate = date
        fname = 'YTD'
    
    # =========================== Ruta de los archivos ============================
    data = pd.read_excel(PARENT_PATH / 'Contributtion' / f'{DATE}-ASSET_ALLOC.xlsx', sheet_name = 'CONSOLIDADO')
    vsData = pd.read_excel(PARENT_PATH / 'Contributtion' / f'{vsDate}-ASSET_ALLOC.xlsx', sheet_name = 'CONSOLIDADO')

    data['QUANTITY'] = data.QUANTITY.astype(float).clip(lower=1)
    vsData['QUANTITY'] = vsData.QUANTITY.astype(float).clip(lower=1)

    data = data[data.MKT_VALUE_USD != 0].fillna(0)
    vsData = vsData[vsData.MKT_VALUE_USD != 0].fillna(0)

    data['PRICE'] = data.MKT_VALUE_USD / data.QUANTITY
    vsData['PRICE'] = vsData.MKT_VALUE_USD / vsData.QUANTITY

    # =============================================================================
    sch = process_portfs(data, vsData, 'SCH', PARENT_PATH / 'DATOS' / 'Graficas_PPT', 'Top10SCH', 'Bot10SCH')
    nfd = process_portfs(data, vsData, 'NFD', PARENT_PATH / 'DATOS' / 'Graficas_PPT', 'Top10NFD', 'Bot10NFD')
    wk = process_portfs(data, vsData, 'WK', PARENT_PATH / 'DATOS' / 'Graficas_PPT', 'Top10WK', 'Bot10WK')

    with pd.ExcelWriter(PARENT_PATH / 'Contributtion' / f'Contribution_{fname}.xlsx', engine = 'openpyxl') as writer:
        for dat in [sch, nfd, wk]:
            for key, val in dat.items():
                val.to_excel(writer, sheet_name = key, index = False)