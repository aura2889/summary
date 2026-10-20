"""Добавляет в data/manual.json водителей, вписанных вручную для машин без путевых листов.
Файл: колонки A госномер (как в своде), G водитель, H телефон, I второй водитель, J его телефон.
Запуск: python3 add_manual.py файл.xlsx ГГГГ-ММ-ДД  (дата, за которую путевые листы; водители подставляются только в данные этого дня)"""
import sys, os, re, json, openpyxl
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import phones, best2
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data', 'manual.json')
day = sys.argv[2]
M = json.load(open(out, encoding='utf-8')) if os.path.exists(out) else {}
m = M.setdefault(day, {})
ws = openpyxl.load_workbook(sys.argv[1], data_only=True).active
def drv(n, t):
    if not n or not str(n).strip(): return None
    n = re.sub(r'^\d+\s+', '', str(n).strip())
    e = {'n': n}; ph = best2(phones(t))
    if ph:
        e['t'] = ph[0]
        if len(ph) > 1: e['t2'] = ph[1]
    return e
add = 0
for r in ws.iter_rows(min_row=2, values_only=True):
    if not r[0]: continue
    l = [x for x in (drv(r[6], r[7]), drv(r[8] if len(r) > 8 else None, r[9] if len(r) > 9 else None)) if x]
    if l: m[str(r[0]).strip()] = l; add += 1
json.dump(M, open(out, 'w', encoding='utf-8'), ensure_ascii=False, indent=0)
print('машин с водителями:', add, 'за', day, 'всего:', len(m))
