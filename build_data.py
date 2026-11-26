"""Convert the NJ Clean Energy SSB rebate workbook into data.json for the lookup site.
Usage: python build_data.py path/to/SSB-Projects.xlsx
Reads the Retro, New Construction and P4P EB/NC tabs (by name prefix), writes data.json beside this script."""
import openpyxl, json, sys, os, re
src = sys.argv[1]
wb = openpyxl.load_workbook(src, read_only=True)
def num(v): return v if isinstance(v, (int, float)) else 0
def iso(d):
    if hasattr(d, 'strftime'): return d.strftime('%Y-%m-%d')
    d = str(d or '').strip()
    if len(d) == 10 and d[2] == '/': return f'{d[6:]}-{d[:2]}-{d[3:5]}'
    return d
tabs = [('Retro', 'Retrofit'), ('New Construction', 'New construction'), ('P4P EB', 'Pay for Performance (existing bldg)'), ('P4P NC', 'Pay for Performance (new constr.)')]
techs, utils, rows, sources = [], [], [], []
def idx(lst, v):
    v = (v or '').strip()
    if v not in lst: lst.append(v)
    return lst.index(v)
for prefix, label in tabs:
    ws = next((w for w in wb.worksheets if w.title.startswith(prefix)), None)
    if ws is None: print('missing tab', prefix); continue
    si = len(sources); sources.append({'label': label, 'tab': ws.title}); n = 0
    for i, r in enumerate(ws.iter_rows(values_only=True)):
        if i == 0: continue
        r = list(r)[:12] + [None] * 12
        if not any(r[:8]): continue
        p4p = prefix.startswith('P4P')
        amt = round(sum(num(x) for x in r[9:12]), 2) if p4p else num(r[8])
        date = '' if p4p else iso(r[9])
        rows.append([si, idx(techs, r[0]), (r[1] or '').strip(), (r[2] or '').strip(), (r[3] or '').strip(),
                     (r[4] or '').strip(), str(r[6] or '').strip(), idx(utils, r[7]), amt, date])
        n += 1
    sources[si]['count'] = n
    print(ws.title, n)
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data.json')
with open(out, 'w') as f:
    f.write('{"file":' + json.dumps(re.sub(r'^[0-9a-f]{8}-', '', os.path.basename(src))) + ',"sources":' + json.dumps(sources) +
            ',"techs":' + json.dumps(techs) + ',"utils":' + json.dumps(utils) + ',"rows":[\n')
    f.write(',\n'.join(json.dumps(r, separators=(',', ':')) for r in rows))
    f.write('\n]}\n')
print(len(rows), 'rows ->', out)
