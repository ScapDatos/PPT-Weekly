import argparse
from config import PARENT_PATH
from debugg.functionsV1 import un_leveraged, get_lev, build_total

path_pc = PARENT_PATH / 'DATOS'
path_to_db = path_pc / 'Bases de datos'

port_banks = {
    'SCH': [{'GSS':'GS'}, {'UBS':'UBS'}, {'MSY':'MS'}, {'SSZ':'SS'}, {'DBK':'DB'}, {'JPM':'JPM'}],
    'NFD': [{'CTI':'CT'}, {'JBR':'JB'}, {'SSZ':'SS'}, {'UBS':'UBS'}],
    'WK': [{'MSY':'MS'}]
}

# def parse_args(argv = None):
#     p = argparse.ArgumentParser(description = 'Compute portfolio leverage')
#     p.add_argument('--data-path', default = PARENT_PATH / 'DATOS')
#     p.add_argument('--db-path', default = 'Bases de datos',
#                    help = 'Database path')
#     p.add_argument('--start-date', default = '2020-01-01',
#                    help = 'Start date for the main run (YYYY-MM-DD)')
#     p.add_argument()

### ============================================================================= ###
###                                      PORTS                                    ###
### ============================================================================= ###
sch, nfd, wk = get_lev(path_to_db, port_banks, '2020-01-01')
sch_2024, nfd_2024, wk_2024 = get_lev(path_to_db, port_banks, '2024-01-01')

### ============================================================================= ###
###                                      TOTAL                                    ###
### ============================================================================= ###
total = build_total(sch, nfd, wk)
total_2024 = build_total(sch_2024, nfd_2024, wk_2024)

### ============================================================================= ###
###                                      PLOTS                                    ###
### ============================================================================= ###
un_leveraged(total, path_pc, 'Graficas_PPT/LEVERAGE_LV', [33, 7], pct = 'PCT_')
un_leveraged(total, path_pc, 'Graficas_PPT/LEVERAGE_LV_SCH', [16.5, 7], ['SCH'], 5.6, True, 2, 1, right = 0.92)
un_leveraged(total, path_pc, 'Graficas_PPT/LEVERAGE_LV_NFD', [16.5, 7], ['NFD'], 25, True, 2, 5, right = 0.92)

un_leveraged(total_2024, path_pc, 'Graficas_PPT/LEVERAGE_LV_2024', [33, 7], pct = 'PCT_')
un_leveraged(total_2024, path_pc, 'Graficas_PPT/LEVERAGE_LV_SCH_2024', [16.5, 7], ['SCH'], 5.6, True, 2, 1, right = 0.92)
un_leveraged(total_2024, path_pc, 'Graficas_PPT/LEVERAGE_LV_NFD_2024', [16.5, 7], ['NFD'], 25, True, 2, 5, right = 0.92)