"""Делает data/contacts.json из файла «КБМ Водители» (только ФИО и телефоны, без зарплат и должностей).
Запуск: python3 make_contacts.py contacts.xlsx out_contacts.json export1.xlsx [export2.xlsx ...]
В базу попадают: весь лист «Водители» + те, кто из КБМ1-7 встречается в переданных экспортах (чтобы не публиковать всех сотрудников)."""
import sys, json, openpyxl
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import *
src, out, exports = sys.argv[1], sys.argv[2], sys.argv[3:]
wb = openpyxl.load_workbook(src, data_only=True)
rows1 = []
for r in wb['Водители'].iter_rows(min_row=2, values_only=True):
    if r[2] and str(r[2]).strip().lower() not in ('вакансия', 'фио'): rows1.append((str(r[2]).strip(), best2(phones(r[4]))))
rows2 = []
for t in ['КБМ1','КБМ2','КБМ3','КБМ4','КБМ5','КБМ6','КБМ7']:
    for r in wb[t].iter_rows(min_row=2, values_only=True):
        if len(r) > 5 and r[1]:
            ph = phones(r[4]) + [p for p in phones(r[5]) if p not in phones(r[4])]
            rows2.append((str(r[1]).strip(), best2(ph)))
junk_map = {**shared_phones(rows1), **shared_phones(rows2)}      # считаем по ПОЛНЫМ данным
# имена из экспортов
names = set()
for f in exports:
    ws = openpyxl.load_workbook(f, data_only=True).active
    for r in ws.iter_rows(min_row=2, values_only=True):
        for x in (r[7], r[9]):
            if x and str(x).strip(): names.add(nname(x))
def same(a, b):
    ta, tb = a.split(), b.split()
    sa, sb = set(ta), set(tb)
    small, big = (sa, sb) if len(sa) <= len(sb) else (sb, sa)
    if len(small) >= 2 and small <= big: return True
    return len(ta) >= 2 and len(tb) >= 2 and (ta[:2] == tb[:2] or ta[:2][::-1] == tb[:2])
byfirst = collections.defaultdict(list)
for n in names:
    for t in n.split(): byfirst[t].append(n)
def relevant(nm):
    n = nname(nm)
    cand = set()
    for t in n.split(): cand.update(byfirst.get(t, ()))
    return any(same(n, c) for c in cand)
r2 = [(n, p) for n, p in rows2 if p and relevant(n)]
r1 = [(n, p) for n, p in rows1 if p]
json.dump({'rows1': r1, 'rows2': r2, 'junk': sorted(junk_map)}, open(out, 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
print('Водители:', len(r1), ' КБМ (только из экспортов):', len(r2), ' общих номеров:', len(junk_map))
