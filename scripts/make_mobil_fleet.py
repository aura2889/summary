"""Разовая подготовка: техника Мобил Сервис (марка, тип, госномер, год, состояние) из файла «Техникалар жайы»
   python3 scripts/make_mobil_fleet.py <файл.xlsx> [дата выгрузки ГГГГ-ММ-ДД]  ->  data/mobil_fleet.json (ключ = серия+цифры+регион, латиница)
   Колонку «Цех» не берём - цех сейчас считается по своду."""
import sys, os, re, json
import openpyxl
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CY = str.maketrans('АВЕКМНОРСТУХҮҒӨҚҺ', 'ABEKMHOPCTYXYFOKH')
def key(s):
    s = re.sub(r'\([^)]*\)', '', str(s)).upper().translate(CY)
    s = re.sub(r'[\s|_\-]+', '', s)
    return re.sub(r'^R(?=\d{3}[A-Z]{2,3}$)', '', s)      # «R 417 BХ» - номер без региона
ST = {'жақсы': 'хорошее', 'орташа': 'среднее', 'нашар': 'плохое'}
ws = openpyxl.load_workbook(sys.argv[1], read_only=True, data_only=True)['Лист1']
out = {}; empty = 0
for r in ws.iter_rows(min_row=3, max_col=7, values_only=True):
    if r[3] is None:
        empty += 1
        if empty > 50: break
        continue
    empty = 0
    k = key(r[3]); mk = re.sub(r'\s+', ' ', str(r[1] or '')).strip(); tp = str(r[2] or '').strip().lower()
    m = mk if tp in ('пикап', '') else (mk + ' ' + tp).strip()
    e = {'m': m, 'y': r[4] if isinstance(r[4], (int, float)) else None, 'st': ST.get(str(r[5] or '').strip().lower(), str(r[5] or '').strip())}
    out.setdefault(k, e)
json.dump({'asof': sys.argv[2] if len(sys.argv) > 2 else '', 'plates': out}, open(os.path.join(ROOT, 'data', 'mobil_fleet.json'), 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
print(len(out), 'номеров')
