#!/usr/bin/env bash
# Rebuild the workbook from code, verify it against the authoritative file, recalculate it, run the independent
# validator, the Monte Carlo analysis and all eleven figure scripts.  Requires: python3 (+ requirements.txt) and
# LibreOffice (soffice) for formula recalculation; without LibreOffice, open/save the workbook in Excel instead.
set -euo pipefail
ROOT="$(CDPATH= cd -- "$(dirname -- "$0")/.." && pwd)"; cd "$ROOT"
PY="${PYTHON_BIN:-python3}"; WB="$ROOT/workbook/Nigeria_EV_Charging_Master_R2.xlsx"
"$PY" code/build_workbook.py
"$PY" code/verify_workbook.py                       # structural comparison with raw_data/authoritative_workbook_R2.xlsx
if command -v soffice >/dev/null 2>&1; then
  TMP="$(mktemp -d)"; trap 'rm -rf "$TMP"' EXIT
  soffice --headless --convert-to xlsx --outdir "$TMP" "$WB" >/dev/null && cp "$TMP/$(basename "$WB")" "$WB"
else
  echo "soffice not found: open and save $WB in Excel so the formulas evaluate, then run the remaining steps by hand." >&2; exit 2
fi
"$PY" code/validate_model.py "$WB"                  # independent re-implementation, expect "Failures: 0"
"$PY" code/monte_carlo.py "$WB"
for f in code/figures/fig*.py; do "$PY" "$f"; done
