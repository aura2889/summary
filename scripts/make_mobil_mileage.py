"""Разовая подготовка: пробег пикапов Мобил Сервис (МСГ) помесячно за янв-окт 2026 (колонки 1..10 = месяцы), лист «Лист1».
   python3 scripts/make_mobil_mileage.py <файл.xlsx>  ->  data/mileage_mobil.json (ключ: серия+цифры+регион без пробелов, латиница)"""
import sys, os, json, re
import openpyxl
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CY = str.maketrans('АВЕКМНОРСТУХ', 'ABEKMHOPCTYX')
ws = openpyxl.load_workbook(sys.argv[1], data_only=True)['Лист1']
out = {}
for r in ws.iter_rows(min_row=4, values_only=True):
    if not r[1]: continue
    k = re.sub(r'^R(?=\d{3}[A-Z]{2,3}$)', '', re.sub(r'\s+', '', str(r[1]).upper().translate(CY)))
    d = [x if isinstance(x, (int, float)) else 0 for x in r[3:13]]
    if k in out: out[k] = [a + b for a, b in zip(out[k], d)]     # машину переводили между цехами - месяцы не пересекаются
    else: out[k] = d
res = {'period': '2026-10', 'plates': {k: {'km': int(d[9]), 'm': d} for k, d in out.items()}}   # km = октябрь, m = янв..окт
json.dump(res, open(os.path.join(ROOT, 'data', 'mileage_mobil.json'), 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
print(len(out), 'номеров, всего км', sum(v['km'] for v in res['plates'].values()))
