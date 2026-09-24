#!/usr/bin/env python3
"""
build_workbook.py -- rebuilds the master workbook of
"Electric Vehicle Charging Infrastructure Requirements for Nigeria to 2060" (Utilities Policy, JUIP-D-26-01375)
from scratch, cell for cell.

    python3 code/build_workbook.py                    -> workbook/Nigeria_EV_Charging_Master_R2.xlsx
    python3 code/verify_workbook.py                   -> compares it with raw_data/authoritative_workbook_R2.xlsx

What comes from where
  * Every formula in the workbook is written by the functions below (the model logic).
  * Labels, notes and input values: code/workbook_static.py (auto-extracted, one line per cell).
  * GDP and population series: raw_data/macro_inputs_NGA.csv (World Bank WDI GDP, constant 2015 US$; UN WPP 2024).
  * Parameter provenance table: raw_data/parameter_dictionary.csv.
  * Fonts, fills, number formats, column widths, row heights, merged cells: raw_data/workbook_format.json.
The workbook is written with formulas only; open it in Excel (or run LibreOffice headless) to evaluate them.

Sheet map (row numbers used by the formulas)
  Assumptions          inputs in blocks 1-15, rows 3-92 (B6, B46, B48, B60 are formulas)
  GDP and Population   row 5 = 2010 ... row 55 = 2060; B = GDP/cap, C = GDP, D = population (M), E/F = growth
  Gompertz fit         anchors rows 16-18; gamma B12; alpha B24; beta B25; R^2 B26; SE(beta) B23; CI B30:C30
  Fleet projections    anchors B6:B8; growth rates B11:B14; annual table row 18 = 2010 ... row 68 = 2060 (B anchored/bridged,
                       C Scenario A, D Scenario B, E Scenario C, F GDP/cap, G population, H Gompertz V/1000, I C growth rate,
                       J World Bank 2013 projection); class shares row 72; class tables headed at rows 75 (A), 130 (B), 185 (C)
  Validation           anchor/estimate points rows 5-11; CAGRs B13:B14; 2060 endpoints rows 17-19
  EV adoption          row 4 = 2010 ... row 54 = 2060, columns B-G = six classes
  Load profile         hourly profile rows 5-28 (columns B-F), weights row 30, LF B32, energy by class B36:B41
  Results Scenario A/B/C  header row 3; row 4 = 2025 ... row 39 = 2060; 22 columns
  Delivery check       one block per scenario starting at rows 4, 14, 24
  Sensitivity          shadow inputs rows 4-75; baseline B77; gamma refit 80-87; t50 helpers 90-95; perturbations 100-115;
                       Monte Carlo specification 118-131; coincidence-factor test 132-133
  Composition sensitivity, Summary 2060, Parameter provenance
"""
import csv, json
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L
from workbook_static import STATIC

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'workbook' / 'Nigeria_EV_Charging_Master_R2.xlsx'
YEARS = list(range(2010, 2061))
CLASSES = ['Motorcycle', 'Private car', 'Light commercial', 'Heavy goods', 'Light bus', 'Coach']
GP = "'GDP and Population'"          # sheet-name prefixes used inside formulas
GF = "'Gompertz fit'"
FP = "'Fleet projections'"
EV = "'EV adoption'"
def gp_row(y): return 5 + (y - 2010)          # GDP and Population
def fp_row(y): return 18 + (y - 2010)         # Fleet projections annual table
def ev_row(y): return 4 + (y - 2010)          # EV adoption
def res_row(y): return 4 + (y - 2025)         # Results sheets
CLASS_HDR = {'A': 75, 'B': 130, 'C': 185}     # class-disaggregated tables in Fleet projections
def cls_row(scen, y): return CLASS_HDR[scen] + 1 + (y - 2010)
def num(x): return ('%.6f' % x).rstrip('0').rstrip('.')    # 0.25, not 0.250000: matches how Excel/LibreOffice store the constant
SHARE_ROW = 72
R60 = res_row(2060)                           # = 39

wb = Workbook(); wb.remove(wb.active)
def sheet(name):
    ws = wb.create_sheet(name)
    for coord, v in STATIC.get(name, {}).items(): ws[coord] = v
    return ws

# ---------------------------------------------------------------- Assumptions ---------------------------------------
asm = sheet('Assumptions')
asm['B6'] = '=D6*B88+(1-D6)*B89'                       # motorcycle/tricycle VKT = home share x private VKT + commercial share x commercial VKT
asm['B46'] = f"={GP}!B19"                              # 2024 GDP per capita
asm['B48'] = f"={GP}!D55"                              # 2060 population
asm['B60'] = "='Load profile'!B32"                     # smart-charging load factor, 2060 (by-year values on the Load profile sheet)
asm['B83'] = '=(B10*40*0.6)/(B7*1.5)'                  # cars replaced per bus on an annual passenger-km basis: bus VKT x 40 seats x 60% load / (car VKT x 1.5 occupants)

# ---------------------------------------------------------------- GDP and Population --------------------------------
gp = sheet('GDP and Population')
with open(ROOT / 'raw_data' / 'macro_inputs_NGA.csv', newline='') as f:
    MACRO = {int(r['year']): r for r in csv.DictReader(f)}
for y in YEARS:
    r = gp_row(y); gp.cell(row=r, column=1, value=y)
    gp.cell(row=r, column=4, value=float(MACRO[y]['population_1jan_persons']) / 1e6)
    if y <= 2024:
        gp.cell(row=r, column=3, value=float(MACRO[y]['gdp_const2015_usd']))
        gp[f'B{r}'] = f'=C{r}/(D{r}*1000000)'                       # GDP per capita, observed
    else:
        gp[f'B{r}'] = f'=B{r-1}*(1+Assumptions!$B$47)'              # projected at the assumed real growth rate
    if y > 2010:
        gp[f'E{r}'] = f'=B{r}/B{r-1}-1'; gp[f'F{r}'] = f'=D{r}/D{r-1}-1'

# ---------------------------------------------------------------- Gompertz fit --------------------------------------
gf = sheet('Gompertz fit')
for i, y in enumerate((2010, 2017, 2018)):                 # anchors (2010 estimate; 2017, 2018 NBS counts), rows 16-18 (stock values are static inputs)
    r = 16 + i
    gf[f'C{r}'] = f"={GP}!D{gp_row(y)}"; gf[f'D{r}'] = f"={GP}!B{gp_row(y)}"
    gf[f'E{r}'] = f'=B{r}*1000/C{r}'; gf[f'F{r}'] = f'=E{r}/$B$12'; gf[f'G{r}'] = f'=LN(F{r})'; gf[f'H{r}'] = f'=LN(-G{r})'
LIN = 'LINEST(H16:H18,D16:D18,TRUE(),TRUE())'
gf['B21'] = f'=INDEX({LIN},1,1)'; gf['B22'] = f'=INDEX({LIN},1,2)'; gf['B23'] = '=STEYX(H16:H18,D16:D18)/SQRT(DEVSQ(D16:D18))'
gf['B24'] = '=-EXP(B22)'; gf['B25'] = '=B21'; gf['B26'] = f'=INDEX({LIN},3,1)'; gf['B27'] = f'=INDEX({LIN},2,2)'; gf['B28'] = f'=INDEX({LIN},4,2)'
gf['B29'] = '=B21/B23'; gf['B30'] = '=B21-TINV(0.05,B28)*B23'; gf['C30'] = '=B21+TINV(0.05,B28)*B23'
for i in range(3):                                          # fit quality at the anchors, rows 36-38
    r, s = 36 + i, 16 + i
    gf[f'A{r}'] = f'=A{s}'; gf[f'B{r}'] = f'=E{s}'; gf[f'C{r}'] = f'=$B$12*EXP($B$24*EXP($B$25*D{s}))'; gf[f'D{r}'] = f'=B{r}-C{r}'; gf[f'E{r}'] = f'=D{r}/B{r}'

# ---------------------------------------------------------------- Fleet projections ---------------------------------
fp = sheet('Fleet projections')
for r, y in ((6, 2010), (7, 2017), (8, 2018)): fp[f'A{r}'] = y                     # anchor years
fp['B11'] = '=(B8/B6)^(1/8)-1'                                                     # anchor-derived CAGR 2010-2018
fp['B12'] = '=(44.79147/8.671474)^(1/25)-1'                                          # Scenario A growth rate implied by World Bank (2013) Table 22.4
fp['B13'] = '=B11'                                                                   # Scenario C starting rate
fp['B14'] = f"=({GP}!D55/{GP}!D54-1)+Assumptions!$B$85*Assumptions!$B$47"           # Scenario C terminal rate
fp['A15'] = 'Scenario B anchor scaling factor (bridged 2024 fleet / Gompertz fit at 2024):'; fp['B15'] = '=D32/(H32*G32/1000)'
fp['A70'] = 'Effective Scenario B asymptote after anchoring (gamma x scaling factor, vehicles per 1,000):'; fp['B70'] = f"={GF}!B12*B15"
WB2013 = {2010: 8.671474, 2020: 17.647653, 2030: 33.900720, 2035: 44.791470}       # World Bank (2013) Table 22.4
for y in YEARS:
    r = fp_row(y); fp[f'A{r}'] = y
    if y <= 2024:                                                                    # anchors, interpolation and bridge
        if y == 2010: fp[f'B{r}'] = '=$B$6'
        elif y < 2017: fp[f'B{r}'] = f'=$B$6*($B$7/$B$6)^({(y-2010)/7:.6f})'
        elif y == 2017: fp[f'B{r}'] = '=$B$7'
        elif y == 2018: fp[f'B{r}'] = '=$B$8'
        else: fp[f'B{r}'] = f'=B{r-1}*(1+$B$11)'
        fp[f'C{r}'] = f'=B{r}'; fp[f'D{r}'] = f'=B{r}'; fp[f'E{r}'] = f'=B{r}'
    else:
        fp[f'C{r}'] = f'=C{r-1}*(1+$B$12)'                                           # A: constant World Bank growth rate
        fp[f'D{r}'] = f'=$D$32*(H{r}*G{r})/($H$32*$G$32)'                            # B: Gompertz path scaled to the 2024 anchor
        fp[f'I{r}'] = f'=$B$13+({num((y-2024)/36)})*($B$14-$B$13)'                    # C: linearly declining growth rate
        fp[f'E{r}'] = f'=E{r-1}*(1+I{r})'
    fp[f'F{r}'] = f"={GP}!B{gp_row(y)}"; fp[f'G{r}'] = f"={GP}!D{gp_row(y)}"
    fp[f'H{r}'] = f"={GF}!$B$12*EXP({GF}!$B$24*EXP({GF}!$B$25*F{r}))"               # Gompertz vehicles per 1,000
    if y <= 2035:                                                                    # World Bank projection, geometric interpolation
        ks = sorted(WB2013); lo = max(k for k in ks if k <= y); hi = min(k for k in ks if k >= y)
        fp[f'J{r}'] = WB2013[lo] if lo == hi else round(WB2013[lo] * (WB2013[hi] / WB2013[lo]) ** ((y - lo) / (hi - lo)), 4)
for scen, col in (('A', 'C'), ('B', 'D'), ('C', 'E')):                               # class-disaggregated tables
    for y in YEARS:
        r = cls_row(scen, y); fp[f'A{r}'] = y; fp[f'B{r}'] = f'={col}{fp_row(y)}'
        for i in range(6): fp[f'{L(3+i)}{r}'] = f'=$B{r}*${L(3+i)}${SHARE_ROW}'

# ---------------------------------------------------------------- Validation ----------------------------------------
va = sheet('Validation')
for r, y in zip(range(5, 12), (2010, 2017, 2018, 2018, 2018, 2020, 2023)):          # anchor/estimate points (values are static inputs)
    va[f'D{r}'] = f"={FP}!J{fp_row(y)}"; va[f'E{r}'] = f'=IF(D{r}="","",D{r}/B{r})'; va[f'F{r}'] = f"={FP}!D{fp_row(y)}"
va['B13'] = f"={FP}!B11"; va['B14'] = f"={FP}!B12"
for k, col in enumerate('CDE'):
    va[f'B{17+k}'] = f"={FP}!{col}{fp_row(2060)}"; va[f'C{17+k}'] = f"=B{17+k}*1000/{GP}!D55"

# ---------------------------------------------------------------- EV adoption ---------------------------------------
ev = sheet('EV adoption')
for y in YEARS:
    r = ev_row(y); ev[f'A{r}'] = y
    for i in range(6):
        k, t = f'Assumptions!$C${16+i}', f'Assumptions!$B${16+i}'     # logistic k and t50; B92 = 1 normalises each curve to 100% in 2060
        ev[f'{L(2+i)}{r}'] = f'=IF(A{r}<2025,0,(1/(1+EXP(-{k}*(A{r}-{t}))))/IF(Assumptions!$B$92=1,1/(1+EXP(-{k}*(2060-{t}))),1))'

# ---------------------------------------------------------------- Load profile --------------------------------------
lp = sheet('Load profile')
for h in range(24):
    r = 5 + h; lp[f'A{r}'] = h; lp[f'G{r}'] = f'=SUMPRODUCT(B{r}:F{r},$B$30:$F$30)'
for c in 'BCDEF': lp[f'{c}29'] = f'=SUM({c}5:{c}28)'
lp['B30'] = '=(B37*Assumptions!$D$7+B38*Assumptions!$D$8+B36*Assumptions!$D$6*Assumptions!$B$88/Assumptions!$B$6)/$B$42'   # home segment weight
lp['C30'] = '=((B37*(1-Assumptions!$D$7)+B38*(1-Assumptions!$D$8))*0.6)/$B$42'                                             # workplace / public AC
lp['D30'] = '=((B37*(1-Assumptions!$D$7)+B38*(1-Assumptions!$D$8))*0.4)/$B$42'                                             # public DC fast
lp['E30'] = '=(B39+B40+B41)/$B$42'                                                                                          # depot
lp['F30'] = '=B36*(1-Assumptions!$D$6)*Assumptions!$B$89/Assumptions!$B$6/$B$42'                                            # battery swap
lp['B32'] = '=AVERAGE(G5:G28)/MAX(G5:G28)'; lp['B33'] = '=INDEX(A5:A28,MATCH(MAX(G5:G28),G5:G28,0))'
for i, col in enumerate('CDEFGH'):                                                  # 2060 Scenario B energy by class (TWh)
    lp[f'B{36+i}'] = f"='Results Scenario B'!{col}{R60}*Assumptions!$B${6+i}*Assumptions!$C${6+i}*Assumptions!$B$61/1000000000"
lp['B42'] = '=SUM(B36:B41)'
# Load factor by year: segment energy weights for each year (rows 45-49), aggregate hourly profile (rows 50-73), LF (row 75); columns J.. = 2025..2060
lp['A44'] = 'Load factor by year (same hourly profiles, each year\'s charging mix; Scenario B energy by class)'
for k, lab in enumerate(['Home weight', 'Workplace/public AC weight', 'Public DC fast weight', 'Depot weight', 'Battery swap weight']): lp[f'I{45+k}'] = lab
for h in range(24): lp[f'I{50+h}'] = f'Hour {h} aggregate share'
lp['I75'] = 'Load factor LF = mean / max'; lp['I44'] = 'Year'
RB_ = "'Results Scenario B'"
for y in range(2025, 2061):
    c = L(10 + y - 2025); r = res_row(y); lp[f'{c}44'] = y
    E = lambda col, i: f"{RB_}!{col}{r}*Assumptions!$B${6+i}*Assumptions!$C${6+i}*Assumptions!$B$61"          # class energy (any unit; weights are ratios)
    tot = '(' + '+'.join(E(col, i) for i, col in enumerate('CDEFGH')) + ')'
    lp[f'{c}45'] = f'=({E("D",1)}*Assumptions!$D$7+{E("E",2)}*Assumptions!$D$8+{E("C",0)}*Assumptions!$D$6*Assumptions!$B$88/Assumptions!$B$6)/{tot}'
    lp[f'{c}46'] = f'=(({E("D",1)}*(1-Assumptions!$D$7)+{E("E",2)}*(1-Assumptions!$D$8))*0.6)/{tot}'
    lp[f'{c}47'] = f'=(({E("D",1)}*(1-Assumptions!$D$7)+{E("E",2)}*(1-Assumptions!$D$8))*0.4)/{tot}'
    lp[f'{c}48'] = f'=({E("F",3)}+{E("G",4)}+{E("H",5)})/{tot}'
    lp[f'{c}49'] = f'={E("C",0)}*(1-Assumptions!$D$6)*Assumptions!$B$89/Assumptions!$B$6/{tot}'
    for h in range(24): lp[f'{c}{50+h}'] = f'=$B{5+h}*{c}$45+$C{5+h}*{c}$46+$D{5+h}*{c}$47+$E{5+h}*{c}$48+$F{5+h}*{c}$49'   # plain formula: no array functions, evaluates in Excel and LibreOffice
    lp[f'{c}75'] = f'=AVERAGE({c}50:{c}73)/MAX({c}50:{c}73)'
def lf_ref(y): return f"'Load profile'!{L(10 + y - 2025)}75"                  # per-year load factor

# ---------------------------------------------------------------- Results Scenario A/B/C ----------------------------
for scen, fcol in (('A', 'C'), ('B', 'D'), ('C', 'E')):
    rs = sheet(f'Results Scenario {scen}')
    for y in range(2025, 2061):
        r = res_row(y); rs[f'A{r}'] = y
        rs[f'B{r}'] = f"={FP}!{fcol}{fp_row(y)}*1000000"                                                            # total fleet
        for i in range(6):                                                                                         # EVs by class
            rs[f'{L(3+i)}{r}'] = f"={FP}!{L(3+i)}{cls_row(scen, y)}*1000000*{EV}!{L(2+i)}{ev_row(y)}"
        rs[f'I{r}'] = f'=SUM(C{r}:H{r})'
        rs[f'J{r}'] = ('=(C{r}*Assumptions!$B$6*Assumptions!$C$6+D{r}*Assumptions!$B$7*Assumptions!$C$7+E{r}*Assumptions!$B$8*Assumptions!$C$8'
                       '+F{r}*Assumptions!$B$9*Assumptions!$C$9+G{r}*Assumptions!$B$10*Assumptions!$C$10+H{r}*Assumptions!$B$11*Assumptions!$C$11)'
                       '*Assumptions!$B$61/1000000000').format(r=r)                                                # at-wall energy, TWh
        rs[f'L{r}'] = f'=((D{r}*Assumptions!$D$7+E{r}*Assumptions!$D$8)*Assumptions!$B$51+C{r}*Assumptions!$D$6*Assumptions!$B$56)/Assumptions!$B$52'  # home L1 sockets
        rs[f'M{r}'] = f'=(D{r}*Assumptions!$D$7+E{r}*Assumptions!$D$8)*(1-Assumptions!$B$51)/Assumptions!$B$26'                                        # home L2
        rs[f'N{r}'] = f'=(D{r}*(1-Assumptions!$D$7)+E{r}*(1-Assumptions!$D$8))/Assumptions!$B$27'                                                       # public AC
        rs[f'O{r}'] = f'=(D{r}+E{r})/Assumptions!$B$28'                                                                                                  # public DC fast
        rs[f'P{r}'] = f'=(F{r}+G{r}+H{r})/Assumptions!$B$29'                                                                                             # depot heavy-duty
        rs[f'Q{r}'] = f'=C{r}*(1-Assumptions!$D$6)/Assumptions!$B$30'                                                                                    # battery-swap stations
        rs[f'K{r}'] = (f'=(L{r}*Assumptions!$B$54*Assumptions!$B$55+M{r}*Assumptions!$C$35*Assumptions!$D$35+N{r}*Assumptions!$C$36*Assumptions!$D$36'
                       f'+O{r}*Assumptions!$C$37*Assumptions!$D$37+P{r}*Assumptions!$C$38*Assumptions!$D$38+Q{r}*Assumptions!$C$39*Assumptions!$D$39)'
                       f'*Assumptions!$B$62/1000000')                                                                                                    # uncontrolled peak, GW
        rs[f'R{r}'] = (f'=(L{r}*Assumptions!$B$53+M{r}*Assumptions!$B$35+N{r}*Assumptions!$B$36+O{r}*Assumptions!$B$37'
                       f'+P{r}*Assumptions!$B$38+Q{r}*Assumptions!$B$39)/1000000000')                                                                   # charger CAPEX, $bn
        rs[f'S{r}'] = f'=K{r}*1000000*Assumptions!$B$43/1000000000'                                                                                      # grid CAPEX, $bn
        rs[f'T{r}'] = f'=R{r}+S{r}'
        rs[f'U{r}'] = f'=J{r}*Assumptions!$B$59*1000/8760/{lf_ref(y)}'                                                                                  # dispatched generation, GW (load factor of that year)
        rs[f'V{r}'] = f'=(D{r}+E{r})*Assumptions!$B$64*Assumptions!$B$65*Assumptions!$B$66/1000000'                                                     # V2G capacity, GW

# ---------------------------------------------------------------- Delivery check ------------------------------------
dc = sheet('Delivery check')
for k, scen in enumerate('ABC'):
    row = 4 + 10 * k; sn = f"'Results Scenario {scen}'"; eb = row + 1
    for i, col in enumerate('CDEFGH'):                                              # energy by class helpers, columns L-Q
        dc[f'{L(12+i)}{eb}'] = f"={sn}!{col}{R60}*Assumptions!$B${6+i}*Assumptions!$C${6+i}*Assumptions!$B$61/1000000000"
    dc[f'R{eb}'] = f"={sn}!C{R60}*Assumptions!$D$6*Assumptions!$B$88*Assumptions!$C$6*Assumptions!$B$61/1000000000"          # home-charged motorcycles
    dc[f'S{eb}'] = f"={sn}!C{R60}*(1-Assumptions!$D$6)*Assumptions!$B$89*Assumptions!$C$6*Assumptions!$B$61/1000000000"      # swap-served motorcycles
    home_ldv = f'($M${eb}*Assumptions!$D$7+$N${eb}*Assumptions!$D$8)'; nonhome = f'($M${eb}*(1-Assumptions!$D$7)+$N${eb}*(1-Assumptions!$D$8))'
    segs = [('L', '$B$54', '$B$75', f'={home_ldv}*Assumptions!$B$51+$R${eb}'), ('M', '$C$35', '$B$76', f'={home_ldv}*(1-Assumptions!$B$51)'),
            ('N', '$C$36', '$B$77', f'={nonhome}*0.6'), ('O', '$C$37', '$B$78', f'={nonhome}*0.4'), ('P', '$C$38', '$B$79', f'=$O${eb}+$P${eb}+$Q${eb}'),
            ('Q', None, None, f'=$S${eb}')]
    for j, (col, kw, ut, assigned) in enumerate(segs):
        r = row + 2 + j; dc[f'B{r}'] = f"={sn}!{col}{R60}"
        if kw:
            dc[f'C{r}'] = f'=Assumptions!{kw}'; dc[f'D{r}'] = f'=Assumptions!{ut}'; dc[f'E{r}'] = f'=B{r}*C{r}*8760*D{r}/1000000000'
        else:                                                                       # battery swap: kWh per swap x max swaps/day x days
            dc[f'C{r}'] = '=Assumptions!$B$69'; dc[f'D{r}'] = '=Assumptions!$B$72'; dc[f'E{r}'] = f'=B{r}*D{r}*C{r}*Assumptions!$B$71/1000000000'
        dc[f'F{r}'] = assigned; dc[f'G{r}'] = f'=E{r}/F{r}'

# ---------------------------------------------------------------- Sensitivity ---------------------------------------
s = sheet('Sensitivity')
for i in range(6):                                                                  # shadow inputs referenced from the master sheets
    s[f'B{4+i}'] = f"={FP}!{L(3+i)}{cls_row('B', 2060)}"; s[f'B{10+i}'] = f"={EV}!{L(2+i)}{ev_row(2060)}"
    s[f'B{16+i}'] = f'=Assumptions!B{6+i}'; s[f'B{22+i}'] = f'=Assumptions!C{6+i}'; s[f'B{28+i}'] = f'=Assumptions!D{6+i}'
for i, (ratio, capex, kw, div) in enumerate([('B52', 'B53', 'B54', 'B55'), ('B26', 'B35', 'C35', 'D35'), ('B27', 'B36', 'C36', 'D36'),
                                              ('B28', 'B37', 'C37', 'D37'), ('B29', 'B38', 'C38', 'D38'), ('B30', 'B39', 'C39', 'D39')]):
    s[f'B{34+i}'] = f'=Assumptions!{ratio}'; s[f'B{40+i}'] = f'=Assumptions!{capex}'; s[f'B{46+i}'] = f'=Assumptions!{kw}'; s[f'B{52+i}'] = f'=Assumptions!{div}'
s['B58'] = '=Assumptions!B43'; s['B59'] = '=Assumptions!B51'; s['B60'] = '=Assumptions!B56'; s['B61'] = f"={GF}!B12"; s['B62'] = '=Assumptions!B62'
s['B63'] = f"={GP}!B55"; s['B64'] = f"={GP}!D55"
for i in range(6): s[f'B{65+i}'] = f'=Assumptions!B{16+i}'; s[f'C{65+i}'] = f'=Assumptions!C{16+i}'
for i in range(3): s[f'B{72+i}'] = f"={GF}!E{16+i}"; s[f'C{72+i}'] = f"={GF}!D{16+i}"
for i in range(6): s[f'{L(2+i)}75'] = f"={FP}!{L(3+i)}{SHARE_ROW}"

def chain(fleet_factor='1', capex_f='1', grid_f='1', share_expr=None, dc_f='1', home_f='1'):
    """2060 Scenario B total CAPEX ($bn) recomputed from the shadow inputs, with one perturbation applied.
    Mirrors the Results-sheet chain: EV stock by class -> chargers by type -> charger CAPEX + grid CAPEX."""
    ev = [f'(B{4+i}*1000000*({fleet_factor})*({share_expr[i] if share_expr else f"B{10+i}"}))' for i in range(6)]
    m, car, lcv, hgv, lb, co = ev
    hs = lambda i: f'MIN(1,B{28+i}*({home_f}))'
    l1 = f'((({car})*{hs(1)}+({lcv})*{hs(2)})*B59+({m})*{hs(0)}*B60)/B34'
    l2 = f'((({car})*{hs(1)}+({lcv})*{hs(2)})*(1-B59))/B35'
    ac = f'((({car})*(1-{hs(1)})+({lcv})*(1-{hs(2)})))/B36'
    dcf = f'(({car})+({lcv}))/(B37*({dc_f}))'
    dep = f'(({hgv})+({lb})+({co}))/B38'
    sw = f'(({m})*(1-{hs(0)}))/B39'
    ccap = f'(({l1})*B40+({l2})*B41+({ac})*B42+({dcf})*B43+({dep})*B44+({sw})*B45)*({capex_f})/1000000000'
    peak = f'(({l1})*B46*B52+({l2})*B47*B53+({ac})*B48*B54+({dcf})*B49*B55+({dep})*B50*B56+({sw})*B51*B57)*B62/1000000'
    gcap = f'({peak})*1000000*B58*({grid_f})/1000000000'
    return f'=({ccap})+({gcap})'
s['B77'] = chain(); s['C77'] = f"='Results Scenario B'!T{R60}"
for col, fac in (('B', 0.75), ('C', 1.25)):                                         # gamma refit helpers (rows 80-87)
    s[f'{col}80'] = f'=$B$61*{fac}'
    for i in range(3): s[f'{col}{81+i}'] = f'=LN(-LN($B${72+i}/({col}$80)))'
    s[f'{col}84'] = f'=SLOPE({col}81:{col}83,$C$72:$C$74)'; s[f'{col}85'] = f'=-EXP(INTERCEPT({col}81:{col}83,$C$72:$C$74))'
    s[f'{col}86'] = f"={FP}!$D$32*({col}80*EXP({col}85*EXP({col}84*$B$63))*$B$64)/({col}80*EXP({col}85*EXP({col}84*{GP}!$B$19))*{GP}!$D$19)"
    s[f'{col}87'] = f'={col}86/(SUM($B$4:$B$9))'
for i in range(6):                                                                  # t50 shift helpers (rows 90-95)
    s[f'B{90+i}'] = f'=1/(1+EXP(-$C${65+i}*(2060-($B${65+i}+5))))'; s[f'C{90+i}'] = f'=1/(1+EXP(-$C${65+i}*(2060-($B${65+i}-5))))'
PERT = [(chain(fleet_factor='$B$87'), chain(fleet_factor='$C$87')), (chain(capex_f='0.7'), chain(capex_f='1.3')), (chain(grid_f='0.5'), chain(grid_f='1.5')),
        (chain(share_expr=[f'$B${90+i}' for i in range(6)]), chain(share_expr=[f'$C${90+i}' for i in range(6)])), (chain(dc_f='0.5'), chain(dc_f='1.5')),
        (chain(home_f='0.5'), chain(home_f='1.5')), (chain(), chain()), (chain(), chain())]   # VKT and kWh/km: no effect on CAPEX by construction
LOW_ROWS = []
for k, (lo, hi) in enumerate(PERT):
    r = 100 + 2 * k; LOW_ROWS.append(r)
    s[f'C{r}'] = lo; s[f'D{r}'] = f'=C{r}-$B$77'; s[f'C{r+1}'] = hi; s[f'D{r+1}'] = f'=C{r+1}-$B$77'
    s[f'E{r}'] = f'=ABS(C{r+1}-C{r})'; s[f'F{r}'] = f'=IF(E{r}>0.25*$B$77,"HIGH",IF(E{r}>0.08*$B$77,"MEDIUM","LOW"))'
for k in range(6):                                                                  # Monte Carlo input ranges (rows 120-125)
    lr = LOW_ROWS[k]; s[f'C{120+k}'] = f'=MIN(D{lr},D{lr+1})'; s[f'D{120+k}'] = f'=MAX(D{lr},D{lr+1})'
s['C132'] = '=B77'; s['C133'] = f"='Results Scenario B'!T{R60}+'Results Scenario B'!S{R60}*(1/Assumptions!B62-1)"   # coincidence factor 0.90 -> 1.00
s['D132'] = '=C132-$B$77'; s['D133'] = '=C133-$B$77'; s['E132'] = '=ABS(C133-C132)'; s['F132'] = '=IF(E132>0.25*$B$77,"HIGH",IF(E132>0.08*$B$77,"MEDIUM","LOW"))'
# Throughput-constrained energy sensitivity: the fixed-ratio +30% VKT (or kWh/km) case leaves charger counts unchanged, but at +30% energy the
# public AC segment exceeds its throughput capacity (Delivery check); the ports needed to restore headroom of 1.0 and their capital are computed here.
s['A135'] = 'Throughput-constrained energy sensitivity (+30% VKT or +30% kWh/km, 2060 Scenario B)'
s['A136'] = 'Public AC energy assigned at +30% (TWh/yr):'; s['B136'] = "='Delivery check'!F18*1.3"
s['A137'] = 'Public AC deliverable energy (TWh/yr):'; s['B137'] = "='Delivery check'!E18"
s['A138'] = 'Additional public AC ports required (units):'; s['B138'] = '=MAX(0,B136-B137)*1000000000/(Assumptions!$C$36*8760*Assumptions!$B$77)'
s['A139'] = 'Additional charger CAPEX ($bn):'; s['B139'] = '=B138*Assumptions!$B$36/1000000000'
s['A140'] = 'Additional grid CAPEX ($bn):'; s['B140'] = '=B138*Assumptions!$C$36*Assumptions!$D$36*Assumptions!$B$62*Assumptions!$B$43/1000000000'   # kW of new peak x USD/kW
s['A141'] = 'Throughput-constrained CAPEX effect of +30% energy ($bn):'; s['B141'] = '=B139+B140'; s['C141'] = '=B141/$B$77'
s['A142'] = 'Lowest headroom of the other segments at +30% energy (x):'; s['B142'] = "=MIN('Delivery check'!G16,'Delivery check'!G17,'Delivery check'!G19,'Delivery check'!G20,'Delivery check'!G21)/1.3"

# ---------------------------------------------------------------- Composition sensitivity --------------------------
co = sheet('Composition sensitivity')
RB = "'Results Scenario B'"
co['B4'] = f"={FP}!D{SHARE_ROW}"; co['C4'] = '=B4*(1-Assumptions!$B$82)'                       # private-car share displaced
co['B5'] = f"={FP}!G{SHARE_ROW}"; co['C5'] = '=B5+B4*Assumptions!$B$82/Assumptions!$B$83'      # buses added, capacity-neutral
co['B6'] = f"={RB}!D{R60}"; co['C6'] = f"={RB}!B{R60}*C4*{EV}!C{ev_row(2060)}"
co['B7'] = f"={RB}!G{R60}"; co['C7'] = f"={RB}!B{R60}*C5*{EV}!F{ev_row(2060)}"
co['B8'] = f"={RB}!J{R60}"; co['C8'] = '=B8+((C6-B6)*Assumptions!$B$7*Assumptions!$C$7+(C7-B7)*Assumptions!$B$10*Assumptions!$C$10)*Assumptions!$B$61/1000000000'
for r, col, extra in ((9, 'L', '=B9+(C6-B6)*Assumptions!$D$7*Assumptions!$B$51/Assumptions!$B$52'), (10, 'M', '=B10+(C6-B6)*Assumptions!$D$7*(1-Assumptions!$B$51)/Assumptions!$B$26'),
                      (11, 'N', '=B11+(C6-B6)*(1-Assumptions!$D$7)/Assumptions!$B$27'), (12, 'O', '=B12+(C6-B6)/Assumptions!$B$28'), (13, 'P', '=B13+(C7-B7)/Assumptions!$B$29'), (14, 'Q', '=B14')):
    co[f'B{r}'] = f"={RB}!{col}{R60}"; co[f'C{r}'] = extra
co['B15'] = f"={RB}!K{R60}"
co['C15'] = ('=(C9*Assumptions!$B$54*Assumptions!$B$55+C10*Assumptions!$C$35*Assumptions!$D$35+C11*Assumptions!$C$36*Assumptions!$D$36'
             '+C12*Assumptions!$C$37*Assumptions!$D$37+C13*Assumptions!$C$38*Assumptions!$D$38+C14*Assumptions!$C$39*Assumptions!$D$39)*Assumptions!$B$62/1000000')
co['B16'] = f"={RB}!R{R60}"; co['C16'] = '=(C9*Assumptions!$B$53+C10*Assumptions!$B$35+C11*Assumptions!$B$36+C12*Assumptions!$B$37+C13*Assumptions!$B$38+C14*Assumptions!$B$39)/1000000000'
co['B17'] = f"={RB}!S{R60}"; co['C17'] = '=C15*1000000*Assumptions!$B$43/1000000000'
co['B18'] = f"={RB}!T{R60}"; co['C18'] = '=C16+C17'
for r in range(8, 19): co[f'D{r}'] = f'=C{r}-B{r}'; co[f'E{r}'] = f'=IF(B{r}=0,"",D{r}/B{r})'

# ---------------------------------------------------------------- Summary 2060 ---------------------------------------
sm = sheet('Summary 2060')
for i, col in enumerate(['B', 'I', None, 'J', 'U', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'V']):
    for j, scen in enumerate('ABC'):
        sn = f"'Results Scenario {scen}'"
        sm.cell(row=4 + i, column=2 + j, value=f"={sn}!I{R60}/{sn}!B{R60}" if col is None else f"={sn}!{col}{R60}")

# ---------------------------------------------------------------- Parameter provenance -------------------------------
pv = sheet('Parameter provenance')
with open(ROOT / 'raw_data' / 'parameter_dictionary.csv', newline='') as f:
    for i, row in enumerate(csv.reader(f)):
        for j, v in enumerate(row):
            if v != '': pv.cell(row=3 + i, column=1 + j, value=v)

# ---------------------------------------------------------------- formatting -----------------------------------------
FMT = json.load((ROOT / 'raw_data' / 'workbook_format.json').open())
for ws in wb.worksheets:
    f = FMT[ws.title]
    for coord, st in f['cells'].items():
        c = ws[coord]; name, sz, b, i, col = st['font']
        c.font = Font(name=name, size=sz, bold=b, italic=i, color=col)
        if st['fill']: c.fill = PatternFill('solid', fgColor=st['fill'])
        c.number_format = st['nf']
        h, v, w = st['align']
        if h or v or w: c.alignment = Alignment(horizontal=h, vertical=v, wrap_text=w)
    for m in f['merged']: ws.merge_cells(m)
    for k, w in f['widths'].items(): ws.column_dimensions[k].width = w
    for k, hgt in f['heights'].items(): ws.row_dimensions[int(k)].height = hgt
    if f['freeze']: ws.freeze_panes = f['freeze']
SEC = PatternFill('solid', fgColor='FF2E75B6'); SECF = Font(name='Calibri', size=11, bold=True, color='FFFFFFFF'); RESF = PatternFill('solid', fgColor='FFFFF2CC')
def section(ws, coord): ws[coord].fill = SEC; ws[coord].font = SECF
def result(ws, coord, nf): ws[coord].fill = RESF; ws[coord].font = Font(name='Calibri', size=11, bold=True); ws[coord].number_format = nf
result(fp, 'B15', '0.0000'); result(fp, 'B70', '0'); asm['B83'].number_format = '0.0'
section(lp, 'A44'); lp['I44'].font = Font(name='Calibri', size=11, bold=True)
for y in range(2025, 2061):
    c = L(10 + y - 2025); lp[f'{c}44'].font = Font(name='Calibri', size=11, bold=True)
    for r in range(45, 74): lp[f'{c}{r}'].number_format = '0.0%'
    result(lp, f'{c}75', '0.000')
section(s, 'A135'); s.merge_cells('A135:H135')
for r, nf in ((136, '0.0'), (137, '0.0'), (138, '#,##0'), (139, '0.00'), (140, '0.00'), (141, '0.00'), (142, '0.00')): s[f'B{r}'].number_format = nf
result(s, 'B141', '0.00'); s['C141'].number_format = '0.0%'
OUT.parent.mkdir(exist_ok=True); wb.save(OUT)
print(f'wrote {OUT}')
