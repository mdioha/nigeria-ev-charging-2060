"""data_vintage_sensitivity.py -- Scenario B 2060 fleet under alternative GDP/population input vintages.
Reproduces the independent audit's figures (60.22 M current-price refit; 97.56 M constant-2021-PPP refit) from the
archived files in raw_data/, using the model logic: linearised Gompertz fit (gamma = 300) on the 2010/2017/2018
observed anchors, projection at 2.5%/yr real GDP-per-capita growth, and scaling to the common 2024 fleet anchor.
Usage: python3 data_vintage_sensitivity.py  (writes 6_audit_response/data_vintage_sensitivity.csv)"""
import csv, numpy as np
from pathlib import Path
PKG = Path(__file__).resolve().parents[1]; RD = PKG / 'raw_data'
def load(fn, col):
    with open(RD / fn) as f: return {int(r['year']): float(r[col]) for r in csv.DictReader(f) if r[col]}
legacy_gdp = load('legacy_workbook_macro_series.csv', 'gdp_pc_used_R1_R2'); legacy_pop = load('legacy_workbook_macro_series.csv', 'population_M_used_R1_R2')
kd = load('wdi_gdp_pc_ppp_constant2021_NGA.csv', 'gdp_pc_ppp_const2021_intl_usd'); cd = load('wdi_gdp_pc_ppp_current_NGA.csv', 'gdp_pc_ppp_current_intl_usd')
wpop = {y: v / 1e6 for y, v in load('wdi_population_NGA.csv', 'population_persons').items()}
fpop = {y: v / 1e6 for y, v in load('macro_inputs_NGA.csv', 'population_1jan_persons').items()}
fgdp = load('macro_inputs_NGA.csv', 'gdp_pc_const2015_usd')
anch = {2010: 8.671474, 2017: 11.458370, 2018: 11.826033}; cagr = (anch[2018] / anch[2010]) ** (1 / 8) - 1; F24 = anch[2018] * (1 + cagr) ** 6
def run(gdp, pop_hist, gamma=300.0, g=0.025, pop_proj=None, gdp_path=False):
    X = np.array([gdp[y] for y in anch]); V = np.array([anch[y] * 1000 / pop_hist[y] for y in anch])
    b, i = np.polyfit(X, np.log(-np.log(V / gamma)), 1); a = -np.exp(i); G = lambda x: gamma * np.exp(a * np.exp(b * x))
    x24 = gdp[2024]; e24 = a * b * x24 * np.exp(b * x24)          # income elasticity of ownership at 2024
    pp = pop_proj or legacy_pop; popg = lambda y: pp[y] / pp[2024]
    xg = (lambda y: gdp[y]) if gdp_path else (lambda y: x24 * (1 + g) ** (y - 2024))
    F = {y: F24 * G(xg(y)) * popg(y) / G(x24) for y in (2030, 2040, 2050, 2060)}
    return a, b, e24, F
cases = [('Earlier draft (legacy GDP and population series)', legacy_gdp, legacy_pop, {}, 'superseded; not reproducible from WDI/WPP'),
         ('Central case: WDI GDP const. 2015 US$ / WPP 2024 population, 2.5%/yr per-capita growth', fgdp, fpop, dict(pop_proj=fpop), 'adopted central case'),
         ('Same data, source-file GDP path (5%/yr GDP from 2026)', fgdp, fpop, dict(pop_proj=fpop, gdp_path=True), 'not adopted: implies ownership above the 300/1000 ceiling by 2060'),
         ('WDI constant-2021 PPP GDP per capita, WPP 2024 population', kd, fpop, dict(pop_proj=fpop), 'alternative real-income measure (PPP)'),
         ('WDI current-price PPP GDP per capita, WPP 2024 population', cd, fpop, dict(pop_proj=fpop), 'conceptually inconsistent: nominal anchors with real growth')]
out = PKG / 'outputs' / 'data_vintage_sensitivity.csv'; out.parent.mkdir(exist_ok=True)
with open(out, 'w', newline='') as f:
    w = csv.writer(f); w.writerow(['case', 'alpha', 'beta', 'income_elasticity_2024', 'fleet_2030_M', 'fleet_2040_M', 'fleet_2050_M', 'fleet_2060_M', 'vehicles_per_1000_2060', 'note'])
    for lab, gdp, ph, kw, note in cases:
        a, b, e, F = run(gdp, ph, **kw); w.writerow([lab, round(a, 4), f'{b:.4e}', round(e, 2)] + [round(F[y], 2) for y in (2030, 2040, 2050, 2060)] + [round(F[2060] * 1000 / (kw.get('pop_proj') or legacy_pop)[2060])] + [note])
        print(f'{lab[:70]:<72} elast={e:4.2f}  2060 fleet={F[2060]:6.2f} M ({F[2060]*1000/(kw.get("pop_proj") or legacy_pop)[2060]:.0f}/1000)')
print('wrote', out)
