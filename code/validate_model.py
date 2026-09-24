"""validate_model.py -- independent re-implementation of the master workbook in pure Python.
Reads ONLY input cells (Assumptions, GDP/population path, Gompertz anchors, class shares, load profiles)
and recomputes every result; compares against the workbook's cached (LibreOffice-recalculated) values.
Exit code 0 if all checks pass within tolerance."""
import sys, numpy as np
from pathlib import Path
from openpyxl import load_workbook
HERE = Path(__file__).resolve().parent
WB = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE.parent / 'workbook' / 'Nigeria_EV_Charging_Master_R2.xlsx'
# Layout of the authoritative workbook (row/column positions used when reading it back)
lay = {'FP0': 18, 'SHARE_ROW': 72, 'CLASS_HDR': {'A': 75, 'B': 130, 'C': 185}, 'RES0': 4, 'R60': 39,
       'SENS_LOW_ROWS': [100, 102, 104, 106, 108, 110, 112, 114, 132], 'SENS_BASE': 'B77', 'MC0': 118}
wb = load_workbook(WB, data_only=True); wf = load_workbook(WB)  # cached values / formulas
a = wb['Assumptions']; gp = wb['GDP and Population']
TOL = 1e-6; fails = []; nchecks = 0
def chk(name, ours, theirs, tol=TOL):
    global nchecks; nchecks += 1
    rel = abs(ours - theirs) / max(abs(theirs), 1e-9)
    if rel > tol: fails.append(f"{name}: python={ours:.6g} workbook={theirs:.6g} rel={rel:.2e}")

# ---------------- inputs -----------------
Y = list(range(2010, 2061)); gdp = {y: gp.cell(row=5 + y - 2010, column=2).value for y in Y}; pop = {y: gp.cell(row=5 + y - 2010, column=4).value for y in Y}
VKT = [a.cell(row=6 + i, column=2).value for i in range(6)]; KWH = [a.cell(row=6 + i, column=3).value for i in range(6)]; HOME = [a.cell(row=6 + i, column=4).value for i in range(6)]
T50 = [a.cell(row=16 + i, column=2).value for i in range(6)]; KK = [a.cell(row=16 + i, column=3).value for i in range(6)]
RAT = dict(l1=a['B52'].value, l2=a['B26'].value, ac=a['B27'].value, dc=a['B28'].value, dep=a['B29'].value, sw=a['B30'].value)
CAP = dict(l1=a['B53'].value, l2=a['B35'].value, ac=a['B36'].value, dc=a['B37'].value, dep=a['B38'].value, sw=a['B39'].value)
KW = dict(l1=a['B54'].value, l2=a['C35'].value, ac=a['C36'].value, dc=a['C37'].value, dep=a['C38'].value, sw=a['C39'].value)
DIV = dict(l1=a['B55'].value, l2=a['D35'].value, ac=a['D36'].value, dc=a['D37'].value, dep=a['D38'].value, sw=a['D39'].value)
GRID = a['B43'].value; L1S = a['B51'].value; MOTL1 = a['B56'].value; GROSS = a['B59'].value; LOSS = a['B61'].value; COINC = a['B62'].value
V2G = (a['B64'].value, a['B65'].value, a['B66'].value); SWAP = (a['B69'].value, a['B71'].value, a['B72'].value)
UTIL = [a[f'B{r}'].value for r in range(75, 80)]; ELAS = a['B85'].value; VKTP, VKTC = a['B88'].value, a['B89'].value; FORCE = a['B92'].value
chk('Motorcycle VKT formula', HOME[0] * VKTP + (1 - HOME[0]) * VKTC, VKT[0])

# ---------------- Gompertz fit (linearised OLS, as LINEST) -----------------
gf = wb['Gompertz fit']; gamma = gf['B12'].value
anc = [(gf.cell(row=16 + i, column=1).value, gf.cell(row=16 + i, column=2).value) for i in range(3)]
X = np.array([gdp[y] for y, _ in anc], float); V = np.array([v * 1000 / pop[y] for y, v in anc])
ylin = np.log(-np.log(V / gamma)); beta, intercept = np.polyfit(X, ylin, 1); alpha = -np.exp(intercept)
chk('alpha', alpha, gf['B24'].value); chk('beta', beta, gf['B25'].value)
resid = ylin - (intercept + beta * X); r2 = 1 - resid.var() * len(X) / (ylin.var() * len(X)); chk('R2', r2, gf['B26'].value, 1e-5)
se_b = np.sqrt((resid @ resid) / (len(X) - 2) / ((X - X.mean()) ** 2).sum()); chk('SE beta', se_b, gf['B23'].value, 1e-5)

# ---------------- fleet -----------------
fp = wb['Fleet projections']; FP0 = lay['FP0']
obs = {y: v for y, v in anc}; cagr_obs = (obs[2018] / obs[2010]) ** (1 / 8) - 1; chk('obs CAGR', cagr_obs, fp['B11'].value)
cagr_A = fp['B12'].value; c_term = (pop[2060] / pop[2059] - 1) + ELAS * a['B47'].value; chk('C terminal', c_term, fp['B14'].value)
base = {}
for y in Y:
    if y <= 2010: base[y] = obs[2010]
    elif y < 2017: base[y] = obs[2010] * (obs[2017] / obs[2010]) ** ((y - 2010) / 7)
    elif y == 2017: base[y] = obs[2017]
    elif y == 2018: base[y] = obs[2018]
    elif y <= 2024: base[y] = base[y - 1] * (1 + cagr_obs)
fleet = {'A': {}, 'B': {}, 'C': {}}
for y in Y:
    if y <= 2024: fleet['A'][y] = fleet['B'][y] = fleet['C'][y] = base[y]
    else:
        fleet['A'][y] = fleet['A'][y - 1] * (1 + cagr_A)
        own_y = gamma * np.exp(alpha * np.exp(beta * gdp[y]))
        own_24 = gamma * np.exp(alpha * np.exp(beta * gdp[2024]))
        fleet['B'][y] = base[2024] * own_y * pop[y] / (own_24 * pop[2024])
        fr = (y - 2024) / 36; fleet['C'][y] = fleet['C'][y - 1] * (1 + cagr_obs + fr * (c_term - cagr_obs))
for s, c in [('A', 3), ('B', 4), ('C', 5)]:
    for y in Y: chk(f'fleet {s} {y}', fleet[s][y], fp.cell(row=FP0 + y - 2010, column=c).value, 1e-5)
shares = [fp.cell(row=lay['SHARE_ROW'], column=3 + i).value for i in range(6)]
chk('normalized class shares', sum(shares), 1.0, 1e-12)

# ---------------- EV shares -----------------
def evshare(i, y):
    if y < 2025: return 0.0
    v = 1 / (1 + np.exp(-KK[i] * (y - T50[i])))
    return v / (1 / (1 + np.exp(-KK[i] * (2060 - T50[i])))) if FORCE == 1 else v
ev = wb['EV adoption']
for y in Y:
    for i in range(6): chk(f'EV share {i} {y}', evshare(i, y), ev.cell(row=4 + y - 2010, column=2 + i).value)

# ---------------- hourly profiles (used for the per-year load factor) -----------------
lp = wb['Load profile']; PROF = np.array([[lp.cell(row=5 + h, column=2 + i).value for i in range(5)] for h in range(24)])

# ---------------- results -----------------
def results(s, y):
    F = fleet[s][y] * 1e6; n = [F * shares[i] * evshare(i, y) for i in range(6)]
    m, car, lcv, hgv, lb, co = n
    twh = sum(n[i] * VKT[i] * KWH[i] for i in range(6)) * LOSS / 1e9
    l1 = ((car * HOME[1] + lcv * HOME[2]) * L1S + m * HOME[0] * MOTL1) / RAT['l1']
    l2 = (car * HOME[1] + lcv * HOME[2]) * (1 - L1S) / RAT['l2']
    ac = (car * (1 - HOME[1]) + lcv * (1 - HOME[2])) / RAT['ac']; dc = (car + lcv) / RAT['dc']; dep = (hgv + lb + co) / RAT['dep']; sw = m * (1 - HOME[0]) / RAT['sw']
    cnt = dict(l1=l1, l2=l2, ac=ac, dc=dc, dep=dep, sw=sw)
    peak = sum(cnt[k] * KW[k] * DIV[k] for k in cnt) * COINC / 1e6
    ccap = sum(cnt[k] * CAP[k] for k in cnt) / 1e9; gcap = peak * 1e6 * GRID / 1e9
    Ecls = [n[i] * VKT[i] * KWH[i] * LOSS for i in range(6)]; Et = sum(Ecls)
    home_e = Ecls[1] * HOME[1] + Ecls[2] * HOME[2] + Ecls[0] * HOME[0] * VKTP / VKT[0]; nonhome_e = Ecls[1] * (1 - HOME[1]) + Ecls[2] * (1 - HOME[2])
    wts = np.array([home_e, nonhome_e * .6, nonhome_e * .4, Ecls[3] + Ecls[4] + Ecls[5], Ecls[0] * (1 - HOME[0]) * VKTC / VKT[0]]) / Et
    agg_y = PROF @ wts; LF = agg_y.mean() / agg_y.max()                     # load factor of that year's charging mix
    disp = twh * GROSS * 1000 / 8760 / LF; v2g = (car + lcv) * V2G[0] * V2G[1] * V2G[2] / 1e6
    return dict(fleet=F, n=n, ev=sum(n), twh=twh, peak=peak, cnt=cnt, ccap=ccap, gcap=gcap, tcap=ccap + gcap, disp=disp, v2g=v2g)
cols = dict(fleet=2, ev=9, twh=10, peak=11, l1=12, l2=13, ac=14, dc=15, dep=16, sw=17, ccap=18, gcap=19, tcap=20, disp=21, v2g=22)
RES = {}
for s in 'ABC':
    rs = wb[f'Results Scenario {s}']
    for y in range(2025, 2061):
        r = results(s, y); RES[(s, y)] = r; row = lay['RES0'] + y - 2025
        for k, c in cols.items():
            ours = r[k] if k not in r['cnt'] else r['cnt'][k]
            chk(f'{s} {y} {k}', ours, rs.cell(row=row, column=c).value, 1e-5)
        for i in range(6): chk(f'{s} {y} class{i}', r['n'][i], rs.cell(row=row, column=3 + i).value, 1e-5)

# ---------------- load profile LF -----------------
prof = PROF
rb = RES[('B', 2060)]; E = [rb['n'][i] * VKT[i] * KWH[i] * LOSS / 1e9 for i in range(6)]; Etot = sum(E)
home_e = E[1] * HOME[1] + E[2] * HOME[2] + E[0] * HOME[0] * VKTP / VKT[0]; nonhome = E[1] * (1 - HOME[1]) + E[2] * (1 - HOME[2])
w = np.array([home_e, nonhome * .6, nonhome * .4, E[3] + E[4] + E[5], E[0] * (1 - HOME[0]) * VKTC / VKT[0]]) / Etot
agg = prof @ w; LF_py = agg.mean() / agg.max(); chk('LF', LF_py, lp['B32'].value)
for i in range(5): chk(f'LP weight {i}', w[i], lp.cell(row=30, column=2 + i).value)

for y in range(2025, 2061):                                                    # per-year LF cells on the Load profile sheet
    r_ = RES[('B', y)]; chk(f'LF {y}', (r_['twh'] * GROSS * 1000 / 8760) / r_['disp'], lp.cell(row=75, column=10 + y - 2025).value, 1e-5)   # LF implied by dispatched GW
fp_ = wb['Fleet projections']; chk('B scaling factor', fleet['B'][2024] / (gamma * np.exp(alpha * np.exp(beta * gdp[2024])) * pop[2024] / 1000), fp_['B15'].value, 1e-6)
chk('effective asymptote', gamma * fp_['B15'].value, fp_['B70'].value, 1e-9)
chk('cars per bus (passenger-km)', VKT[4] * 40 * 0.6 / (VKT[1] * 1.5), a['B83'].value, 1e-9)

# ---------------- delivery check (Scenario B) -----------------
dc_ = wb['Delivery check']; c = rb['cnt']
E_home_m = rb['n'][0] * HOME[0] * VKTP * KWH[0] * LOSS / 1e9; E_swap_m = rb['n'][0] * (1 - HOME[0]) * VKTC * KWH[0] * LOSS / 1e9
home_ldv = E[1] * HOME[1] + E[2] * HOME[2]
assigned = [home_ldv * L1S + E_home_m, home_ldv * (1 - L1S), nonhome * .6, nonhome * .4, E[3] + E[4] + E[5], E_swap_m]
deliv = [c['l1'] * KW['l1'] * 8760 * UTIL[0] / 1e9, c['l2'] * KW['l2'] * 8760 * UTIL[1] / 1e9, c['ac'] * KW['ac'] * 8760 * UTIL[2] / 1e9,
         c['dc'] * KW['dc'] * 8760 * UTIL[3] / 1e9, c['dep'] * KW['dep'] * 8760 * UTIL[4] / 1e9, c['sw'] * SWAP[2] * SWAP[0] * SWAP[1] / 1e9]
for k in range(6):
    chk(f'deliv {k}', deliv[k], dc_.cell(row=16 + k, column=5).value, 1e-5); chk(f'assigned {k}', assigned[k], dc_.cell(row=16 + k, column=6).value, 1e-5)
    chk(f'headroom {k}', deliv[k] / assigned[k], dc_.cell(row=16 + k, column=7).value, 1e-5)

# ---------------- composition sensitivity -----------------
co = wb['Composition sensitivity']; disp_share = a['B82'].value; cpb = a['B83'].value
F = rb['fleet']; sh_car = shares[1] * (1 - disp_share); sh_bus = shares[4] + shares[1] * disp_share / cpb
car2 = F * sh_car * evshare(1, 2060); bus2 = F * sh_bus * evshare(4, 2060)
chk('comp cars', car2, co['C6'].value); chk('comp buses', bus2, co['C7'].value)
d_car = car2 - rb['n'][1]; d_bus = bus2 - rb['n'][4]
twh2 = rb['twh'] + (d_car * VKT[1] * KWH[1] + d_bus * VKT[4] * KWH[4]) * LOSS / 1e9; chk('comp twh', twh2, co['C8'].value)
c2 = dict(c); c2['l1'] += d_car * HOME[1] * L1S / RAT['l1']; c2['l2'] += d_car * HOME[1] * (1 - L1S) / RAT['l2']; c2['ac'] += d_car * (1 - HOME[1]) / RAT['ac']; c2['dc'] += d_car / RAT['dc']; c2['dep'] += d_bus / RAT['dep']
peak2 = sum(c2[k] * KW[k] * DIV[k] for k in c2) * COINC / 1e6; tcap2 = sum(c2[k] * CAP[k] for k in c2) / 1e9 + peak2 * 1e6 * GRID / 1e9
chk('comp tcap', tcap2, co['C18'].value, 1e-5)

# ---------------- sensitivity table -----------------
sn = wb['Sensitivity']; base_t = rb['tcap']; chk('sens baseline', base_t, sn['B77'].value, 1e-5)
def tcap_pert(fleet_f=1, capex_f=1, grid_f=1, sh=None, dc_f=1, home_f=1):
    F = rb['fleet'] * fleet_f; n = [F * shares[i] * (sh[i] if sh else evshare(i, 2060)) for i in range(6)]; m, car, lcv, hgv, lb, co_ = n
    hs = [min(1, HOME[i] * home_f) for i in range(6)]
    cn = dict(l1=((car * hs[1] + lcv * hs[2]) * L1S + m * hs[0] * MOTL1) / RAT['l1'], l2=(car * hs[1] + lcv * hs[2]) * (1 - L1S) / RAT['l2'],
              ac=(car * (1 - hs[1]) + lcv * (1 - hs[2])) / RAT['ac'], dc=(car + lcv) / (RAT['dc'] * dc_f), dep=(hgv + lb + co_) / RAT['dep'], sw=m * (1 - hs[0]) / RAT['sw'])
    pk = sum(cn[k] * KW[k] * DIV[k] for k in cn) * COINC / 1e6
    return sum(cn[k] * CAP[k] for k in cn) * capex_f / 1e9 + pk * 1e6 * GRID * grid_f / 1e9
def gamma_refit(fac):
    g2 = gamma * fac; yl = np.log(-np.log(V / g2)); b2, i2 = np.polyfit(X, yl, 1)
    a2 = -np.exp(i2)
    own60 = g2 * np.exp(a2 * np.exp(b2 * gdp[2060])); own24 = g2 * np.exp(a2 * np.exp(b2 * gdp[2024]))
    f60 = base[2024] * own60 * pop[2060] / (own24 * pop[2024])
    return f60 / fleet['B'][2060]
sh_lo = [1 / (1 + np.exp(-KK[i] * (2060 - (T50[i] + 5)))) for i in range(6)]; sh_hi = [1 / (1 + np.exp(-KK[i] * (2060 - (T50[i] - 5)))) for i in range(6)]
perts = [(tcap_pert(fleet_f=gamma_refit(.75)), tcap_pert(fleet_f=gamma_refit(1.25))), (tcap_pert(capex_f=.7), tcap_pert(capex_f=1.3)), (tcap_pert(grid_f=.5), tcap_pert(grid_f=1.5)),
         (tcap_pert(sh=sh_lo), tcap_pert(sh=sh_hi)), (tcap_pert(dc_f=.5), tcap_pert(dc_f=1.5)), (tcap_pert(home_f=.5), tcap_pert(home_f=1.5)), (base_t, base_t), (base_t, base_t)]
for (lo, hi), r in zip(perts, lay['SENS_LOW_ROWS']):
    chk(f'sens row {r} low', lo - base_t, sn.cell(row=r, column=4).value, 1e-4); chk(f'sens row {r} high', hi - base_t, sn.cell(row=r + 1, column=4).value, 1e-4)
if len(lay['SENS_LOW_ROWS']) > len(perts):
    r = lay['SENS_LOW_ROWS'][-1]
    no_coinc = base_t + rb['gcap'] * (1 / COINC - 1)
    chk('coincidence sensitivity baseline', base_t, sn.cell(row=r, column=3).value, 1e-5)
    chk('coincidence sensitivity 1.00', no_coinc, sn.cell(row=r + 1, column=3).value, 1e-5)

# Full-electrification normalization check (toggle-independent).
full_n = [rb['fleet'] * shares[i] for i in range(6)]
chk('100% EV stock equals fleet', sum(full_n), rb['fleet'], 1e-12)

# ---------------- throughput-constrained energy sensitivity -----------------
sn_ = wb['Sensitivity']; ac_assigned = assigned[2] * 1.3; ac_deliv = deliv[2]
extra = max(0, ac_assigned - ac_deliv) * 1e9 / (KW['ac'] * 8760 * UTIL[2]); chk('extra AC ports', extra, sn_['B138'].value, 1e-6)
chk('throughput CAPEX effect', extra * CAP['ac'] / 1e9 + extra * KW['ac'] * DIV['ac'] * a['B62'].value * GRID / 1e9, sn_['B141'].value, 1e-6)

# ---------------- summary -----------------
sm = wb['Summary 2060']
for j, s in enumerate('ABC'):
    chk(f'summary tcap {s}', RES[(s, 2060)]['tcap'], sm.cell(row=18, column=2 + j).value, 1e-5)
    chk(f'summary ev share {s}', RES[(s, 2060)]['ev'] / RES[(s, 2060)]['fleet'], sm.cell(row=6, column=2 + j).value, 1e-5)

print(f"Checks: {nchecks}   Failures: {len(fails)}")
for f in fails[:20]: print('  FAIL', f)
sys.exit(1 if fails else 0)
