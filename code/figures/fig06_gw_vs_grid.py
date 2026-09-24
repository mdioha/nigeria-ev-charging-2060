#!/usr/bin/env python3
"""Figure 6. Managed-charging generation requirement (Results column U) and uncontrolled peak (column K) vs grid benchmarks."""
from common import *
LF = lp['B32'].value; fig, ax = plt.subplots(figsize=(8, 5))
ax.fill_between(RY, rcol('C', 21), rcol('A', 21), color='#1f77b4', alpha=.15); ax.plot(RY, rcol('B', 21), color='#1f77b4', lw=2.4, label=f'Dispatched generation, managed charging (gross-up 1.20; LF by year, {LF:.2f} in 2060)')
ax.fill_between(RY, rcol('C', 11), rcol('A', 11), color='#d62728', alpha=.15); ax.plot(RY, rcol('B', 11), color='#d62728', lw=2.4, label='Peak charging load, uncontrolled (90% coincidence)')
for yv, lab, c, ls in [(4.286, 'Average available dispatch capacity, 4.286 GW (NERC, Apr. 2026)', 'black', '-'), (13.625, 'Installed capacity, 13.625 GW (NERC, Apr. 2026)', 'grey', ':'), (277, 'ETP 2.0 target: 277 GW installed by 2060', '#bcbd22', '-.')]:
    ax.axhline(yv, color=c, ls=ls, lw=1.4); ax.text(2025.5, yv * 1.12, lab, color=c, fontsize=8.5)
ax.set_yscale('log'); ax.set_ylim(1, 1200); ax.set_xlim(2025, 2060); ax.set_xlabel('Year'); ax.set_ylabel('Power requirement (GW, log scale)'); ax.legend(loc='lower right', fontsize=8.5); ax.grid(alpha=.3, which='both')
save(fig, 'fig6_gw_vs_grid')
