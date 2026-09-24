#!/usr/bin/env python3
"""Figure 2. Nigeria vehicle stock by class, 2010 to 2060, three scenarios."""
from common import *
fig, axs = plt.subplots(1, 3, figsize=(13, 5), sharey=True)
for ax, s in zip(axs, 'ABC'):
    cs = cls_stock(s); ax.stackplot(YEARS, cs, labels=CLS, colors=CCOL, alpha=.85); tot = cs.sum(0)
    ax.plot(YEARS, tot, color='black', lw=1.5); ax.text(2059.5, tot[-1] + 3, f'{tot[-1]:.1f}M', ha='right', fontweight='bold')
    ax.set_title(f'Scenario {s}'); ax.set_xlabel('Year'); ax.set_xlim(2010, 2060); ax.grid(alpha=.3)
axs[0].set_ylabel('Vehicle stock (millions)'); axs[2].legend(loc='upper left', fontsize=9)
save(fig, 'fig2_vehicle_stock')
