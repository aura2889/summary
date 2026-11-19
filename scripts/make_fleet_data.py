"""Разовая подготовка статичных данных по парку КБМ (запуск вручную):
   python3 scripts/make_fleet_data.py <база_техники.xlsx> <путевые_листы_месяц.xlsx> <период, напр. 2026-10> <дата выгрузки базы ГГГГ-ММ-ДД>
   -> data/fleet_base.json   (префикс 13: владелец, модель, год выпуска, № автоколонны, пробег по одометру на дату выгрузки)
   -> data/mileage_oct2026.json (сумма поля «Пробег» из путевых листов за месяц по каждому госномеру)"""
import sys, os, json, collections, re
import openpyxl
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
base_f, mil_f, period, asof = sys.argv[1:5]

wb = openpyxl.load_workbook(base_f, data_only=True)
ws = wb['Лист1']; rows = list(ws.iter_rows(values_only=True))
hn = [re.sub(r'\s+', ' ', str(x or '')).strip().lower() for x in rows[0]]
col = lambda s: next(i for i, x in enumerate(hn) if x.startswith(s))
c_o, c_p, c_n, c_m, c_t, c_y, c_a, c_k = col('владелец'), col('префикс'), col('гос'), col('марка'), col('тип'), col('год'), col('№ а/к'), col('пробег')
base = {}
for r in rows[1:]:
    if not r[c_n]: continue
    base[str(r[c_n]).strip()] = {'o': str(r[c_o] or '').strip(), 'm': re.sub(r'\s+', ' ', str(r[c_m] or '')).strip(), 'y': r[c_y], 'ak': str(r[c_a] or '').replace('№', '').strip(), 'km': r[c_k]}
json.dump({'asof': asof, 'plates': base}, open(os.path.join(ROOT, 'data', 'fleet_base.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=0)

wb = openpyxl.load_workbook(mil_f, data_only=True, read_only=True); ws = wb.active
it = ws.iter_rows(values_only=True); h = next(it)
hn = [re.sub(r'\s+', ' ', str(x or '')).strip().lower() for x in h]
c_pl = hn.index('государственный №'); c_km = hn.index('пробег'); c_dt = hn.index('дата путевого листа')
acc = collections.defaultdict(lambda: [0, 0])
for r in it:
    pl = str(r[c_pl] or '').strip()
    if not pl or not isinstance(r[c_km], (int, float)): continue
    d = r[c_dt]
    if hasattr(d, 'strftime') and d.strftime('%Y-%m') != period: continue
    acc[pl][0] += r[c_km]
    if r[c_km] > 0: acc[pl][1] += 1
out = {'period': period, 'plates': {k: {'km': int(v[0]), 'n': v[1]} for k, v in acc.items()}}
json.dump(out, open(os.path.join(ROOT, 'data', 'mileage_oct2026.json'), 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
print('база', len(base), 'номеров; пробег за', period, len(out['plates']), 'номеров, всего км', sum(v['km'] for v in out['plates'].values()))
