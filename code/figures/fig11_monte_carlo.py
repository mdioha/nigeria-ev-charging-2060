#!/usr/bin/env python3
"""Figure 11. Surrogate Monte Carlo distribution of 2060 Scenario B CAPEX (5,000 trials; see ../monte_carlo.py)."""
import sys; from pathlib import Path; sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common import *
from monte_carlo import load_spec, run
base, inputs = load_spec(WB); total, _ = run(base, inputs)
p5, p50, p95 = np.percentile(total, [5, 50, 95]); fig, ax = plt.subplots(figsize=(8, 4.8))
ax.hist(total, bins=60, color='#4e79a7', alpha=.85, edgecolor='white', lw=.4)
ax.axvline(p5, color='darkred', ls='--', label=f'5th %ile: ${p5:.0f}B'); ax.axvline(p50, color='black', lw=2, label=f'Median: ${p50:.0f}B')
ax.axvline(total.mean(), color='navy', ls=':', label=f'Mean: ${total.mean():.0f}B'); ax.axvline(p95, color='darkred', ls='--', label=f'95th %ile: ${p95:.0f}B')
ax.set_xlabel('Total 2060 net build-out CAPEX (USD billions, 2025$)'); ax.set_ylabel('Frequency (5,000 surrogate-analysis draws)'); ax.legend(); ax.grid(alpha=.3)
save(fig, 'fig11_monte_carlo')
