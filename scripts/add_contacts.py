"""Добавляет в data/contacts.json водителей из листа «Водители» (ФИО, колонка, телефон) — только ФИО и телефоны.
Запуск: python3 add_contacts.py файл.xlsx [data/contacts.json]"""
import sys, os, json, openpyxl
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import *
src = sys.argv[1]; out = sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data', 'contacts.json')
c = json.load(open(out, encoding='utf-8'))
ws = openpyxl.load_workbook(src, data_only=True)['Водители']
new = []
for r in ws.iter_rows(values_only=True):
    nm = r[2] if len(r) > 2 else None
    if not nm or str(nm).strip().lower() in ('фио', 'вакансия'): continue
    ph = best2(phones(r[4] if len(r) > 4 else None))
    if ph: new.append((str(nm).strip(), ph))
have = {}
for n, p in c['rows1']: have[nname(n)] = (n, p)
added = upd = 0
for n, p in new:
    k = nname(n)
    if k in have:
        o = have[k][1]; m = o + [x for x in p if x not in o]
        if m[:2] != o:
            have[k][1][:] = m[:2]; upd += 1
    else:
        c['rows1'].append([n, p]); have[k] = (n, p); added += 1
sh = shared_phones([(n, p) for n, p in c['rows1']] + [(n, p) for n, p in c['rows2']])
c['junk'] = sorted(set(c['junk']) | set(sh))
json.dump(c, open(out, 'w', encoding='utf-8'), ensure_ascii=False)
print('добавлено', added, 'обновлено', upd, 'мусорных номеров', len(c['junk']))
