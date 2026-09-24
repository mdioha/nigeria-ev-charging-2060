#!/usr/bin/env python3
"""Figure 4. Annual at-wall charging electricity demand, 2025 to 2060, three scenarios (Results column J)."""
from common import *
fig, ax = plt.subplots(figsize=(8, 4.8))
for s, lab in [('A', 'Scenario A, upper bound'), ('B', 'Scenario B, central case'), ('C', 'Scenario C, lower bound')]:
    v = rcol(s, 10); ax.plot(RY, v, color=COL[s], lw=2.2, marker='o', markevery=5, label=lab); ax.text(2060.3, v[-1], f'{v[-1]:.0f} TWh', color=COL[s], fontweight='bold', va='center')
ax.axhline(36, ls='--', color='grey')
ax.text(2063.5, 26, 'Current Nigeria grid generation,\nabout 36 TWh (2024)', color='grey', fontsize=9, ha='right', va='top', linespacing=1.1)
ax.set_xlabel('Year'); ax.set_ylabel('Annual at-wall charging demand (TWh)'); ax.set_xlim(2025, 2064); ax.set_ylim(-60, None); ax.legend(loc='upper left'); ax.grid(alpha=.3)
save(fig, 'fig4_annual_twh')
