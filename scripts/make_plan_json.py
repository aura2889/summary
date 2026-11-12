"""data/plan_budget_2026.xlsx (лист «ТС КБМ по цехам» + «Список (для фильтра)») -> plan.json: план по бюджету 2026 по типам, цехам и автоколоннам.
Запуск: python scripts/make_plan_json.py  (после замены xlsx)"""
import openpyxl, json, os, re
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
wb = openpyxl.load_workbook(os.path.join(ROOT, 'data', 'plan_budget_2026.xlsx'), data_only=True)
ws = wb['ТС КБМ по цехам']; lst = wb['Список (для фильтра)']
DEPS = json.load(open(os.path.join(ROOT, 'data.json'), encoding='utf-8')).get('depOrder', [])
ALIAS = {'Эл.Цех': 'ЭЦ', 'ППД': 'ЦППД', 'ОКС, ДРиСУ': 'ОКС', 'МГ': 'МГС', 'ГАПиТ': 'ГАПиИТ', 'оппоIR': 'IR'}
def clean(s): return re.sub(r'\s+', ' ', str(s or '')).strip().replace('Cедельный', 'Седельный')
def pk(v):
    s = str(v); return s[:-2] if s.endswith('.0') else s
shops = []
for c in range(6, 84, 2):
    n = ws.cell(3, c).value
    if n is None: continue
    n = clean(n); dep = ALIAS.get(n, n if n in DEPS else None)
    shops.append(dict(n=n, c=c, dep=dep))
rows = []; prev = None; prevname = ''
for r in range(6, 53):
    a = ws.cell(r, 1).value; nm = clean(ws.cell(r, 2).value)
    if a is not None: prev = pk(a); key = prev; p = prev; prevname = nm; name = nm
    else:
        key = prev + '.1'; p = ''; name = re.sub(r'\s*с водителем$', '', prevname) + ' ' + nm
    d = n = 0; sh = {}
    for s in shops:
        x, y = ws.cell(r, s['c']).value or 0, ws.cell(r, s['c'] + 1).value or 0
        if x or y: sh[s['n']] = [x, y]; d += x; n += y
    rows.append(dict(k=key, p=p, name=name, d=d, n=n, shops=sh, cols={}, _r=r))
idx = {}
for i, r in enumerate(rows): idx[(r['k'])] = r
# Список: (префикс, марка, цех, смена, колонна, гараж, кол-во) - привязка к строке по порядку: префикс, а «без водителя» -> k.1
for pre, mk, shop, sm, col, gar, q in [[c.value for c in x] for x in lst.iter_rows(min_row=2)]:
    k = pk(pre) + ('.1' if str(mk).strip() == 'без водителя' else '')
    r = idx.get(k)
    if not r: continue
    b = r['cols'].setdefault(str(col), [0, 0]); b[0 if sm == 'День' else 1] += q
out = dict(year=2026, title='План по бюджету 2026 г. (ТС ПУ)', shops=[dict(n=s['n'], dep=s['dep']) for s in shops],
           rows=[{k: v for k, v in r.items() if k != '_r'} for r in rows])
json.dump(out, open(os.path.join(ROOT, 'plan.json'), 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
print(len(rows), sum(r['d'] + r['n'] for r in rows), [s['n'] for s in shops if not s['dep']])
