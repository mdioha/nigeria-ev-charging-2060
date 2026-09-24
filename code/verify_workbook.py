#!/usr/bin/env python3
"""verify_workbook.py -- cell-by-cell comparison of the rebuilt workbook with the authoritative one.
Compares every cell's value or formula, number format, font, fill and alignment, plus merged ranges, column widths,
row heights and freeze panes. Exit code 0 only if nothing differs.
Usage: python3 code/verify_workbook.py [rebuilt.xlsx] [authoritative.xlsx]"""
import sys
from pathlib import Path
from openpyxl import load_workbook
ROOT = Path(__file__).resolve().parents[1]
A = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / 'workbook' / 'Nigeria_EV_Charging_Master_R2.xlsx'
B = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / 'raw_data' / 'authoritative_workbook_R2.xlsx'
def rgb(c):
    try: return str(c.rgb) if c is not None and c.type == 'rgb' else None
    except Exception: return None
def fmt(c):
    f, fl, a = c.font, c.fill, c.alignment
    return (f.name, float(f.sz) if f.sz else None, bool(f.bold), bool(f.italic), rgb(f.color), rgb(fl.fgColor) if fl is not None and fl.fill_type == 'solid' else None,
            c.number_format, a.horizontal or 'general', a.vertical or 'bottom', bool(a.wrap_text))   # unset == Excel defaults
wa, wb_ = load_workbook(A), load_workbook(B); diffs = []; n = 0
if wa.sheetnames != wb_.sheetnames: diffs.append(('sheets', wa.sheetnames, wb_.sheetnames))
for name in wb_.sheetnames:
    if name not in wa.sheetnames: continue
    x, y = wa[name], wb_[name]
    rows, cols = max(x.max_row, y.max_row), max(x.max_column, y.max_column)
    for r in range(1, rows + 1):
        for col in range(1, cols + 1):
            cx, cy = x.cell(row=r, column=col), y.cell(row=r, column=col); n += 1
            vx, vy = cx.value, cy.value
            if isinstance(vx, (int, float)) and isinstance(vy, (int, float)) and not isinstance(vx, bool):
                same = abs(vx - vy) <= 1e-12 * max(1, abs(vy))
            else: same = vx == vy
            if not same: diffs.append((name, cx.coordinate, 'value', str(vx)[:70], str(vy)[:70]))
            if (cx.value is not None or cy.value is not None) and fmt(cx) != fmt(cy): diffs.append((name, cx.coordinate, 'format', fmt(cx), fmt(cy)))
    if sorted(map(str, x.merged_cells.ranges)) != sorted(map(str, y.merged_cells.ranges)): diffs.append((name, '', 'merged', len(x.merged_cells.ranges), len(y.merged_cells.ranges)))
    for k, v in y.column_dimensions.items():
        if v.width and abs((x.column_dimensions[k].width or 0) - v.width) > 1e-6: diffs.append((name, k, 'width', x.column_dimensions[k].width, v.width))
    for k, v in y.row_dimensions.items():
        if v.height and abs((x.row_dimensions[k].height or 0) - v.height) > 1e-6: diffs.append((name, k, 'height', x.row_dimensions[k].height, v.height))
    if x.freeze_panes != y.freeze_panes: diffs.append((name, '', 'freeze', x.freeze_panes, y.freeze_panes))
print(f'cells compared: {n}   differences: {len(diffs)}')
for d in diffs[:40]: print('  ', d)
sys.exit(1 if diffs else 0)
