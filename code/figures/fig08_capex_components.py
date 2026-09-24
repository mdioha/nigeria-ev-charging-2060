#!/usr/bin/env python3
"""Figure 8. Net first-installation CAPEX 2025-2060 by component and scenario (Results columns R, S, T, row 2060)."""
from common import *
fig, ax = plt.subplots(figsize=(8, 5)); x = np.arange(3); w = .26
for j, s in enumerate('ABC'):
    vals = [r60(s, c) for c in (18, 19, 20)]; ax.bar(x + (j - 1) * w, vals, w, color=COL[s], label=f'Scenario {s}')
    for xi, v in zip(x + (j - 1) * w, vals): ax.text(xi, v + 8, f'${v:.0f}B', ha='center', fontsize=9, fontweight='bold')
ax.set_xticks(x); ax.set_xticklabels(['Charger\nCAPEX', 'Grid upgrade\nCAPEX', 'Total\nCAPEX']); ax.set_ylabel('Net first-installation build-out (USD billions, 2025$)'); ax.legend(); ax.grid(alpha=.3, axis='y')
save(fig, 'fig8_capex_components')
