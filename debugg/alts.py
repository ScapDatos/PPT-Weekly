import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.cbook import boxplot_stats
from matplotlib.lines import Line2D

path = 'C:/Users/DATOS-INVERSIONES/OneDrive/0. Nube Asset Mgmt/Alternativos/Macros Prequin'
files = os.listdir(os.path.join(path, 'Datos crudos'))

def boxx_plott(data:pd.DataFrame, name:str, path:str):
    fig, ax = plt.subplots()
    
    fliers_max, fliers_min = [], []

    for col in data.columns:
        q1 = data[col].quantile(0.25)
        q3 = data[col].quantile(0.75)
        iqr = q3 - q1
        max_val = data[col].max()
        sup = q3 + 1.5*iqr if q3 + 1.5*iqr < max_val else max_val
        min_val = data[col].min()
        inf = q1 - 1.5*iqr if q1 - 1.5*iqr > min_val else min_val
        
        outliers_max = data[data[col] > sup][col]
        outliers_min = data[data[col] < inf][col]
        fliers_max.append(outliers_max.max() if len(outliers_max) > 0 else np.nan)
        fliers_min.append(outliers_min.min() if len(outliers_min) > 0 else np.nan)
    
    vals = [data[col].dropna() for col in data.columns]
    bp = ax.boxplot(vals, tick_labels = data.columns, showmeans = True, showfliers = False)
    stats = boxplot_stats(vals)
    
    ax.scatter(
        [i+1 for i in range(4)],
        fliers_max,
        marker = 'o',
        edgecolors = 'k',
        facecolors = 'none',
        # label = 'Max. Outlier'
    )
    ax.scatter(
        [i+1 for i in range(4)],
        fliers_min,
        marker = 'o',
        edgecolors = 'k',
        facecolors = 'none',
        # label = 'Max. Outlier'
    )
    
    handles = [
        Line2D([0], [0], c = 'orange', label = 'Median'),
        Line2D([0], [0], marker = '^', c = 'g', ls = 'none', label = 'Mean'),
        Line2D([0], [0], marker = 'o', markeredgecolor = 'k', markerfacecolor = 'none', ls = 'none', label = 'Min./Max. value')
    ]
    
    # for i,s in enumerate(stats, start = 1):
    #     ax.text(i+0.25, s['q1'], f'Q1: {s["q1"]:.2f}', ha = 'left', va = 'top')
    #     ax.text(i+0.25, s['med'], f'Q2: {s["med"]:.2f}', ha = 'left', va = 'top')
    #     ax.text(i+0.25, s['q3'], f'Q3: {s["q3"]:.2f}', ha = 'left', va = 'top')
        
    #     ax.text(i+0.25, s['whislo'], f"Min: {s['whislo']:.2f}", ha='left', va='top')
    #     ax.text(i+0.25, s['whishi'], f"Max: {s['whishi']:.2f}", ha='left', va='bottom')
        
    #     ax.text(i+0.25, s['mean'], f"μ: {s['mean']:.2f}", ha='left', va='bottom')
    
    ax.legend(
            handles = handles,
            loc = f'best',
            frameon = False,
            # ncol = 3
    )
    
    ax.set_title(name, size = 14, fontweight = 'bold', family = 'Arial')
    
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: rf'$\bf{x:.1f}$'))
    
    ax.grid(axis = 'y', color = 'lightgray', lw = 2, alpha = 0.5)
    ax.tick_params(axis = 'both', labelsize = 10, labelfontfamily = 'Arial')
    ax.set_xticklabels(ax.get_xticklabels(), fontweight='bold', fontsize = 12)
    ax.set_axisbelow(True)
    
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    ax.spines['left'].set_visible(False)
    
    plt.tight_layout(pad = 0.05)
    fig.savefig(f'{path}/Stats/{name}.png', dpi = 300)
    # plt.show()

for file in files:
    
    data = pd.read_csv(f'{path}/Datos crudos/{file}')
    # data['NET IRR (%)'] = data['NET IRR (%)'].replace('n/m', np.nan)
    data = data.replace(['n/m', 'n/a'], np.nan)
    
    data[['NET IRR (%)', 'RVPI (%)', 'DPI (%)']] = data[['NET IRR (%)', 'RVPI (%)', 'DPI (%)']].astype(str).replace(',', '')
    data[['NET IRR (%)', 'RVPI (%)', 'DPI (%)']] = data[['NET IRR (%)', 'RVPI (%)', 'DPI (%)']].astype(float) / 100
    data = data[data['VINTAGE / INCEPTION YEAR'] >= 2015].reset_index(drop = True)

    desc = data[['NET IRR (%)', 'RVPI (%)', 'DPI (%)', 'NET MULTIPLE (X)']].describe()
    # desc.to_csv(f'{path}/Stats/Stats_{file}')
    # print(f'{file} DONE')
    
    boxx_plott(data[['NET IRR (%)', 'RVPI (%)', 'DPI (%)', 'NET MULTIPLE (X)']], file[:-4], path)