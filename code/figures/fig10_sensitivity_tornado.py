#!/usr/bin/env python3
"""Figure 10. One-at-a-time sensitivity of 2060 Scenario B CAPEX (Sensitivity sheet perturbation table)."""
from common import *
from matplotlib.patches import Patch
base = sens[SENS_BASE].value; rows = []
for r in SENS_LOW_ROWS:
    lo, hi = sens.cell(row=r, column=4).value, sens.cell(row=r + 1, column=4).value
    if abs(hi - lo) > 0.05: rows.append((sens.cell(row=r, column=1).value, lo, hi))        # VKT and kWh/km rows have zero effect and are omitted
rows.sort(key=lambda t: abs(t[2] - t[1])); fig, ax = plt.subplots(figsize=(9, 5))
for i, (lab, lo, hi) in enumerate(rows): ax.barh(i, lo, color='#e15759'); ax.barh(i, hi, color='#4e79a7')
ax.set_yticks(range(len(rows))); ax.set_yticklabels([r[0] for r in rows], fontsize=9); ax.axvline(0, color='black', lw=1)
ax.set_xlabel(f'Change in 2060 net build-out CAPEX (USD billions)\nScenario B baseline = USD {base:.0f} billion')
ax.legend(handles=[Patch(color='#e15759', label='Downward perturbation'), Patch(color='#4e79a7', label='Upward perturbation')], loc='lower center', bbox_to_anchor=(0.5, 1.01), ncol=2, frameon=False); ax.grid(alpha=.3, axis='x')
save(fig, 'fig10_sensitivity_tornado')
