#!/usr/bin/env python3
"""Figure 5. Annual charging demand by vehicle class, Scenario B (EV stock x VKT x kWh/km x charging-loss uplift)."""
from common import *
loss = asm['B61'].value
twh = np.array([rcol('B', 3 + i) * asm.cell(row=6 + i, column=2).value * asm.cell(row=6 + i, column=3).value * loss / 1e9 for i in range(6)])
fig, ax = plt.subplots(figsize=(8, 4.8)); ax.stackplot(RY, twh, labels=CLS, colors=CCOL, alpha=.85)
ax.set_xlabel('Year'); ax.set_ylabel('Annual at-wall charging demand (TWh)'); ax.set_xlim(2025, 2060); ax.legend(loc='upper left', fontsize=9); ax.grid(alpha=.3)
save(fig, 'fig5_demand_by_class')
