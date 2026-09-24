#!/usr/bin/env python3
"""Figure 1. Nigerian vehicle-stock anchors (2010 estimate; NBS 2017-2018 counts), the World Bank (2013) projection, and the three scenarios."""
from common import *
wbp = np.array([fp.cell(row=fpr(y), column=10).value or np.nan for y in YEARS], float)          # World Bank projection (col J)
obs = np.array([fp.cell(row=fpr(y), column=2).value or np.nan for y in YEARS], float)            # anchored / bridged (col B)
pts = [(2010, 8.671), (2017, 11.458), (2017.75, 11.547), (2018.25, 11.654), (2018, 11.826)]     # anchors and NBS quarterly counts
fig, ax = plt.subplots(figsize=(8, 4.8))
ax.plot(YEARS, wbp, '--', color='grey', lw=2, label='World Bank (2013) projection, Table 22.4')
ax.plot(YEARS, obs, '-', color='black', lw=2.2, label='Anchors (interpolated) and 2019-24 bridge (anchor-derived growth)')
ax.scatter([p[0] for p in pts], [p[1] for p in pts], color='black', zorder=5, s=36, label='Anchors: WB 2010 base-year estimate; NBS 2017-2018 counts')
ax.scatter([2020], [11.605], marker='s', color='grey', zorder=5, s=36, label='CEIC/OICA 2020 (modelled)')
m = np.array(YEARS) >= 2024
for s, lab in [('A', 'Scenario A (WB growth rate)'), ('B', 'Scenario B (Gompertz, central)'), ('C', 'Scenario C (declining CAGR)')]:
    ax.plot(np.array(YEARS)[m], fleet(s)[m], color=COL[s], lw=2, label=lab)
ax.set_xlim(2010, 2040); ax.set_ylim(0, 45); ax.set_xlabel('Year'); ax.set_ylabel('Vehicle stock (millions)'); ax.legend(fontsize=8.5, loc='upper left'); ax.grid(alpha=.3)
save(fig, 'fig1_validation')
