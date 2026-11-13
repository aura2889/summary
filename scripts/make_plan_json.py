"""data/plan_budget_2026.xlsx (лист «ТС КБМ по цехам» + «Список (для фильтра)») -> plan.json: план по бюджету 2026 по типам, цехам и автоколоннам.
Запуск: python scripts/make_plan_json.py  (после замены xlsx)"""
import openpyxl, json, os, re
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
wb = openpyxl.load_workbook(os.path.join(ROOT, 'data', 'plan_budget_2026.xlsx'), data_only=True)
ws = wb['ТС КБМ по цехам']; lst = wb['Список (для фильтра)']
DEPS = json.load(open(os.path.join(ROOT, 'data.json'), encoding='utf-8')).get('depOrder', [])
# названия цехов как на сайте (в Excel они иначе); несколько столбцов Excel -> один цех суммируются
ALIAS = {'Эл.Цех': 'ЭЦ', 'ППД': 'ЦППД', 'ОКС, ДРиСУ': 'ОКС', 'МГ': 'МГС', 'ГАПиТ': 'ГАПиИТ', 'оппоIR': 'IR',
         'Диспетчерская группа': 'ЦИТС', 'СӨС': 'ОБиОТ', 'Нейрон': 'ОБиОТ', 'БД': 'ОБиОТ', 'Чакур': 'ОБиОТ', 'СМУ УТТ': 'ПРЦЭО', 'ОРП': 'Отделкадр', 'ДТТ-1': 'Гараж-1', 'ДТТ-2': 'Гараж-2'}
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
        if x or y:
            k = s['dep'] or s['n']; b = sh.setdefault(k, [0, 0]); b[0] += x; b[1] += y; d += x; n += y
    rows.append(dict(k=key, p=p, name=name, d=d, n=n, shops=sh, cols={}, _r=r))
idx = {}
for i, r in enumerate(rows): idx[(r['k'])] = r
# Список: (префикс, марка, цех, смена, колонна, гараж, кол-во) - привязка к строке по порядку: префикс, а «без водителя» -> k.1
for pre, mk, shop, sm, col, gar, q in [[c.value for c in x] for x in lst.iter_rows(min_row=2)]:
    k = pk(pre) + ('.1' if str(mk).strip() == 'без водителя' else '')
    r = idx.get(k)
    if not r: continue
    b = r['cols'].setdefault(str(col), [0, 0]); b[0 if sm == 'День' else 1] += q
    dp_ = ALIAS.get(clean(shop), clean(shop)); i_ = 0 if sm == 'День' else 1
    r.setdefault('sg', {}).setdefault(dp_, {}).setdefault(str(gar), [0, 0])[i_] += q      # цех -> гараж -> [день, ночь]
    r.setdefault('sc', {}).setdefault(dp_, {}).setdefault(str(col), [0, 0])[i_] += q      # цех -> автоколонна -> [день, ночь]
order = []
for s in shops:
    k = s['dep'] or s['n']
    if k not in order: order.append(k)
# Мобил Сервис (третья сторона): зафиксированный план по цехам и типам
mob = json.load(open(os.path.join(ROOT, 'data', 'plan_mobil.json'), encoding='utf-8'))['rows']
known = {r['k'] for r in rows}
for k in mob:
    if k not in known:      # тип есть только у Мобила (например, автокран 100т)
        nm = {'63': 'Автокран 100т', '40': 'Прицеп'}.get(k, 'Тип ' + k)
        rows.append(dict(k=k, p=k, name=nm, d=0, n=0, shops={}, cols={}, _r=0))
mobil = {k: dict(d=sum(v[0] for v in sh.values()), n=sum(v[1] for v in sh.values()), shops=sh) for k, sh in mob.items()}
out = dict(year=2026, mobil=mobil, title='План по бюджету 2026 г. (ТС ПУ)', shops=[dict(n=k, dep=k) for k in order],
           rows=[{k: v for k, v in r.items() if k != '_r'} for r in rows])
json.dump(out, open(os.path.join(ROOT, 'plan.json'), 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
print(len(rows), sum(r['d'] + r['n'] for r in rows), [s['n'] for s in shops if not s['dep']])

# план по цехам для свода Excel (scripts/svod/resources/plan_dept.json): КБМ (файл по колоннам) + Мобил Сервис
pdp = os.path.join(ROOT, 'scripts', 'svod', 'resources', 'plan_dept.json')
old = json.load(open(pdp, encoding='utf-8'))
plan = {}; total = {k: 0 for k in old['total']}
for r in rows:
    k = r['k']; sh = {d: list(v) for d, v in r['shops'].items()}
    for d, v in (mob.get(k) or {}).items():
        x = sh.setdefault(d, [0, 0]); x[0] += v[0]; x[1] += v[1]
    sh = {d: v for d, v in sh.items() if v[0] or v[1]}
    if sh: plan[k] = sh
    total[k] = sum(v[0] + v[1] for v in sh.values())
old['plan'] = plan; old['total'] = total
old['depts'] = [d for d in old['depts'] if d not in ('БД', 'СМУ УТТ', 'Нейрон')]
json.dump(old, open(pdp, 'w', encoding='utf-8'), ensure_ascii=False)
print('plan_dept', sum(total.values()))
