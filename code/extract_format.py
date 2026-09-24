#!/usr/bin/env python3
"""extract_format.py -- capture cell formatting, merged ranges, column widths, row heights and freeze panes from the
authoritative workbook into raw_data/workbook_format.json, so build_workbook.py can reproduce the workbook's appearance
exactly. Only needed again if the authoritative workbook's formatting changes.
Usage: python3 extract_format.py [authoritative.xlsx] [out.json]"""
import json, sys
from pathlib import Path
from openpyxl import load_workbook
ROOT = Path(__file__).resolve().parents[1]
SRC = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / 'raw_data' / 'authoritative_workbook_R2.xlsx'
OUT = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / 'raw_data' / 'workbook_format.json'
def rgb(c):
    try: return str(c.rgb) if c is not None and c.type == 'rgb' else None
    except Exception: return None
def style(cell):
    f, fl, a = cell.font, cell.fill, cell.alignment
    s = {'font': [f.name, float(f.sz) if f.sz else None, bool(f.bold), bool(f.italic), rgb(f.color)],
         'fill': rgb(fl.fgColor) if fl is not None and fl.fill_type == 'solid' else None,
         'nf': cell.number_format, 'align': [None if a.horizontal == 'general' else a.horizontal, None if a.vertical == 'bottom' else a.vertical, bool(a.wrap_text)]}   # Excel's explicit defaults == unset
    return s
wb = load_workbook(SRC); fmt = {}
for ws in wb.worksheets:
    cells = {}
    for row in ws.iter_rows():
        for c in row:
            s = style(c)
            if c.value is None and s['font'][0] in ('Calibri', None) and s['font'][1] in (11.0, None) and not s['font'][2] and not s['font'][3] and s['fill'] is None and s['nf'] == 'General' and s['align'] == [None, None, False]:
                continue
            cells[c.coordinate] = s
    fmt[ws.title] = {'cells': cells, 'merged': sorted(str(m) for m in ws.merged_cells.ranges),
                     'widths': {k: v.width for k, v in ws.column_dimensions.items() if v.width},
                     'heights': {k: v.height for k, v in ws.row_dimensions.items() if v.height}, 'freeze': ws.freeze_panes}
json.dump(fmt, OUT.open('w'), indent=0)
print(f'wrote {OUT} ({sum(len(v["cells"]) for v in fmt.values())} styled cells across {len(fmt)} sheets)')
