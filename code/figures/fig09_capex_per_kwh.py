#!/usr/bin/env python3
"""Figure 9. Charger CAPEX per delivered kWh over a 10-year life, Scenario B 2060 (Delivery check sheet, Scenario B block)."""
from common import *
names = ['Home L1', 'Home L2', 'Public AC', 'Public DC fast', 'Depot HD', 'Battery swap']
unit = [asm['B53'].value, asm['B35'].value, asm['B36'].value, asm['B37'].value, asm['B38'].value, asm['B39'].value]      # USD per port/station
cnt = [dc.cell(row=16 + i, column=2).value for i in range(6)]; assigned = [dc.cell(row=16 + i, column=6).value for i in range(6)]   # units; TWh/yr assigned
cents = [unit[i] * cnt[i] / (assigned[i] * 1e9 * 10) * 100 for i in range(6)]
fig, ax = plt.subplots(figsize=(8, 4.6)); bars = ax.bar(names, cents, color=CCOL)
for b, v in zip(bars, cents): ax.text(b.get_x() + b.get_width() / 2, v * 1.03, f'{v:.2f}', ha='center', fontsize=9)
ax.set_ylabel('Charger CAPEX per delivered kWh\n(US cents, 10-year charger life)'); ax.grid(alpha=.3, axis='y'); plt.setp(ax.get_xticklabels(), rotation=15)
save(fig, 'fig9_capex_per_kwh')
