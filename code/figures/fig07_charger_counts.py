#!/usr/bin/env python3
"""Figure 7. Charging units required in 2060 by type and scenario (Results columns L-Q, row 2060)."""
from common import *
types = ['Home L1\nsockets', 'Home L2', 'Public AC', 'Public DC\nfast', 'Depot HD', 'Battery swap\nstations']
fig, ax = plt.subplots(figsize=(9, 5)); x = np.arange(6); w = .26
for j, s in enumerate('ABC'):
    vals = [r60(s, c) / 1e6 for c in range(12, 18)]; ax.bar(x + (j - 1) * w, vals, w, color=COL[s], label=f'Scenario {s}')
    for xi, v in zip(x + (j - 1) * w, vals): ax.text(xi, v * 1.05, f'{v:.2f}M' if v >= 0.1 else f'{v*1000:.0f}k', ha='center', fontsize=7.5, rotation=90)
ax.set_yscale('log'); ax.set_ylim(0.01, 200); ax.set_xticks(x); ax.set_xticklabels(types); ax.set_ylabel('Units in 2060 (millions, log scale)'); ax.legend(); ax.grid(alpha=.3, axis='y', which='both')
save(fig, 'fig7_charger_counts_2060')
