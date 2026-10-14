import argparse
from config import PARENT_PATH
from debugg.functionsV1 import un_leveraged, get_lev, build_total

port_banks = {
    'SCH': [{'GSS':'GS'}, {'UBS':'UBS'}, {'MSY':'MS'}, {'SSZ':'SS'}, {'DBK':'DB'}, {'JPM':'JPM'}],
    'NFD': [{'CTI':'CT'}, {'JBR':'JB'}, {'SSZ':'SS'}, {'UBS':'UBS'}],
    'WK': [{'MSY':'MS'}]
}

def parse_args():
    p = argparse.ArgumentParser(description = 'Compute portfolio leverage')
    p.add_argument('--data-path', default = PARENT_PATH / 'DATOS')
    p.add_argument('--db-path', default = 'Bases de datos',
                   help = 'Database path')
    p.add_argument('--start-date', type = str, default = '2020-01-01',
                   help = 'Start date for the main run (YYYY-MM-DD)')
    p.add_argument('--plots-dir', default = 'Graficas_PPT',
                   help = 'Output folder for plots')
    p.add_argument('--plots', action = 'store_false',
                   help = 'Plot all figures')
    p.add_argument('--export-xlsx', default = None,
                   help = 'If set, export the "total" DataFrame to this .xlsx path')
    return p.parse_args()

def main():
    args = parse_args()
    path_pc = args.data_path
    path_to_db = path_pc / args.db_path

    ### ============================================================ ###
    ###                          PORTS/TOTAL                         ###
    ### ============================================================ ###
    sch, nfd, wk = get_lev(path_to_db, port_banks, args.start_date)
    total = build_total(sch, nfd, wk)

    if args.export_xlsx:
        total.to_excel(args.export_xlsx, index = False)

    ### ============================================================ ###
    ###                             PLOTS                            ###
    ### ============================================================ ###
    if args.plots:
        subidr = args.plots_dir
        sfx = '' if args.start_date == '2020-01-01' else '_YTD'
        un_leveraged(total, path_pc, f'{subidr}/LEVERAGE_LV{sfx}', [33, 7], pct = 'PCT_')
        un_leveraged(total, path_pc, f'{subidr}/LEVERAGE_LV_SCH{sfx}', [16.5, 7], ['SCH'], 5.6, True, 2, 1, right = 0.92)
        un_leveraged(total, path_pc, f'{subidr}/LEVERAGE_LV_NFD{sfx}', [16.5, 7], ['NFD'], 25, True, 2, 5, right = 0.92)

if __name__ == '__main__':
    main()