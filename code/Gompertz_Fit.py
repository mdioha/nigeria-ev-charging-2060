"""Gompertz_Fit.py -- nonlinear least-squares cross-check of the Excel LINEST fit (revision 2).
Fits V/P = gamma * exp(alpha * exp(beta * GDP_pc)) to the three OBSERVED Nigerian anchors
(World Bank 2010 base year; NBS Q1-2017 and Q4-2018 vehicle-population counts).
The printed future stocks are raw curve outputs for fit diagnostics. The master workbook scales
Scenario B's post-2024 curve to the shared 2024 fleet anchor before reporting scenario results.
Usage: python3 Gompertz_Fit.py [gamma]   (default gamma = 300)
"""
import sys, numpy as np
from scipy.optimize import curve_fit
gamma = float(sys.argv[1]) if len(sys.argv) > 1 else 300.0
anchors = {  # year: (vehicle stock M, population M, GDP per capita PPP $)
    2010: (8.671474, 164.3043, 2348.423), 2017: (11.458370, 197.8621, 2471.237), 2018: (11.826033, 202.6471, 2459.279)}  # WPP 2024 (1 Jan); WDI GDP const 2015 US$
X = np.array([v[2] for v in anchors.values()], float); Y = np.array([v[0] * 1000 / v[1] for v in anchors.values()])
f = lambda x, a, b: gamma * np.exp(a * np.exp(b * x))
(a, b), cov = curve_fit(f, X, Y, p0=[-5, -4.8e-4], maxfev=20000)
# linearised OLS (what Excel LINEST does)
ylin = np.log(-np.log(Y / gamma)); slope, intercept = np.polyfit(X, ylin, 1)
print(f"gamma = {gamma:.0f}")
print(f"Nonlinear LSQ : alpha = {a:.4f}  beta = {b:.7f}")
print(f"Linearised OLS: alpha = {-np.exp(intercept):.4f}  beta = {slope:.7f}  (Excel LINEST equivalent)")
pop = {2030: 259.866, 2040: 310.248, 2050: 357.001, 2060: 397.912}  # UN WPP 2024 medium (1 Jan)
for y, p in pop.items():
    gdp = 2348.909 * 1.025 ** (y - 2024); v = f(gdp, a, b)
    print(f"  {y}: V/1000 = {v:6.1f}  stock = {v * p / 1000:6.1f} M")
