# Electric Vehicle Charging Infrastructure Requirements for Nigeria to 2060 — model, data and code

Code and data for the paper *Electric Vehicle Charging Infrastructure Requirements for Nigeria to 2060*
(manuscript JUIP-D-26-01375, *Utilities Policy*). The supplementary Excel workbook, the Monte Carlo analysis and every
figure in the paper can be regenerated from this repository.

## What is here

| Path | Contents |
|---|---|
| `code/build_workbook.py` | Rebuilds the supplementary workbook from scratch. **Every formula in the workbook is written here**; labels and input values come from `code/workbook_static.py`, the GDP/population series from `raw_data/macro_inputs_NGA.csv`, the parameter provenance table from `raw_data/parameter_dictionary.csv`, and formatting from `raw_data/workbook_format.json`. |
| `code/verify_workbook.py` | Compares the rebuilt workbook with the authoritative file cell by cell (values, formulas, number formats, fonts, fills, merged cells, column widths, row heights). Expected result: `differences: 0`. |
| `code/validate_model.py` | Independent re-implementation of the whole model in Python; recomputes 2,830 cells from the inputs (including the year-specific load factor, the anchored Scenario B asymptote and the throughput-constrained sensitivity) and compares them with the workbook. Expected result: `Failures: 0`. |
| `code/monte_carlo.py` | The surrogate Monte Carlo analysis of Section 4.6 (5,000 draws, seed 42). Writes `outputs/monte_carlo_draws.csv` and `outputs/monte_carlo_summary.csv`. |
| `code/figures/fig01_…py` – `fig11_…py` | One standalone script per figure in the paper; `common.py` holds the shared workbook access. |
| `code/data_vintage_sensitivity.py`, `code/Gompertz_Fit.py` | Sensitivity of the central case to GDP/population data vintage; nonlinear cross-check of the Gompertz fit. |
| `raw_data/` | Input extracts with sources and retrieval dates; the authoritative workbook used for verification. |
| `workbook/` | The rebuilt (and recalculated) supplementary workbook. |
| `figures/` | Figures 1–11 (PNG 300 dpi and SVG). |

## Reproduce everything

```bash
python3 -m pip install -r requirements.txt      # exact versions used: requirements-lock.txt
bash code/run_all.sh
```

`run_all.sh` rebuilds the workbook, verifies it against `raw_data/authoritative_workbook_R2.xlsx`, recalculates it with
LibreOffice, runs the independent validator and the Monte Carlo analysis, and regenerates all figures. Without
LibreOffice, open and save the rebuilt workbook in Excel (so formulas evaluate), then run the remaining scripts directly:

```bash
python3 code/validate_model.py workbook/Nigeria_EV_Charging_Master_R2.xlsx
python3 code/monte_carlo.py    workbook/Nigeria_EV_Charging_Master_R2.xlsx
python3 code/figures/fig04_annual_twh.py          # or any other figure
```

## The model in brief

Vehicle stock is projected 2010–2060 from three anchors (one estimated: the World Bank 2010 base year; two observed: NBS 2017 and 2018 counts) with an anchor-derived 2010-2018 growth rate used to bridge to 2024, under three
scenarios: **A** constant growth at the World Bank (2013) implied rate (upper bound for 2060), **B** a Gompertz
motorisation model driven by real income and population, fitted to the anchors and scaled to the common 2024 fleet
(central case; anchoring the fitted curve to the bridged 2024 fleet multiplies it by 1.23, so its effective asymptote is about
368 vehicles per 1,000 rather than the nominal γ = 300), **C** growth declining to population growth plus income elasticity ×
GDP-per-capita growth (lower bound). The managed-charging load factor is computed for each year from that year's charging mix. Class-specific logistic EV-adoption curves, vehicle-kilometres, energy intensities and charger ratios give
annual charging electricity, charger and battery-swap counts, peak load, generation requirement and net first-installation
capital. `Assumptions!B92 = 1` forces every class to 100% EV stock share in 2060 (the complete-electrification variant).

Row/column positions used by the scripts are listed at the top of `code/build_workbook.py`.

## Data sources

Vehicle stock: National Bureau of Statistics (Nigeria) Road Transport Data; Cervigni et al. (2013), World Bank.
GDP: World Bank WDI, GDP in constant 2015 US$ (`NY.GDP.MKTP.KD`). Population: UN World Population Prospects 2024
(1 January). Charging parameters: see `raw_data/parameter_dictionary.csv` (74 inputs with value, unit, source, locator
and type). Source data remain subject to their publishers' terms; the code in this repository is released under the MIT
licence.
