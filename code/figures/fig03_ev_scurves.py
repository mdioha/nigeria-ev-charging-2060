#!/usr/bin/env python3
"""Figure 3. Class-specific logistic S-curves for EV adoption (t50 and k from Assumptions rows 16-21)."""
from common import *
fig, ax = plt.subplots(figsize=(8, 4.8)); t = np.linspace(2020, 2060, 401)
for i in range(6):
    t50 = asm.cell(row=16 + i, column=2).value; k = asm.cell(row=16 + i, column=3).value
    ax.plot(t, np.where(t < 2025, 0, 1 / (1 + np.exp(-k * (t - t50)))) * 100, color=CCOL[i], lw=2.2, label=f'{CLS[i]} (t50={t50:.0f}, k={k:.2f})')
ax.axhline(60, ls=':', color='grey'); ax.text(2021, 61.5, 'Author calibration target: about 60% of stock by 2050\n(informed by ETP 2.0; not an ETP class milestone)', color='grey', fontsize=8.5, linespacing=1.1)
ax.set_xlabel('Year'); ax.set_ylabel('EV share of vehicle stock (%)'); ax.set_ylim(0, 105); ax.legend(fontsize=9); ax.grid(alpha=.3)
save(fig, 'fig3_ev_scurves')
