"""Список вахты Мобил Сервис -> data/mobil.json (ключ: цифры+серия без региона).
Запуск: python3 make_mobil.py roster.xlsx data/mobil.json [старый_лист_Мобил.xlsx]
Новый список ПЕРЕКРЫВАЕТ старые записи, остальные машины остаются как были."""
import sys, os, json, openpyxl
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import *
PHONE_RE = re.compile(r'(?:\+?7|8)[\s\-()]*\d{3}[\s\-()]*\d{3}[\s\-]*\d{2}[\s\-]*\d{2}')
SKIP = re.compile(r'резерв|ремонт|место|^\s*$', re.I)
LETT = r'[A-Za-zА-Яа-яӘәҒғҚқҢңӨөҰұҮүІіҺһ]'
def split_cell(c):
    c = str(c or '').strip()
    m = PHONE_RE.search(c)
    name = c[:m.start()].strip() if m else c
    return name, phones(c[m.start():]) if m else []
def key_noreg(p): return re.sub(r'(?<=[A-Z])\d{2}$', '', p)
def read_roster(path):
    ws = openpyxl.load_workbook(path, data_only=True)
    ws = ws['Мобил Сервис'] if 'Мобил Сервис' in ws.sheetnames else ws.active
    out = collections.OrderedDict()
    for r in ws.iter_rows(min_row=3, values_only=True):
        c = str(r[2] or '').strip()
        if not r[1] or not re.search(r'\d{3}', c): continue
        plate = re.sub(r'^R(?=\d)', '', lat(re.sub(r'\s+', '', c).upper()))
        drivers = []
        for ci in range(3, min(len(r), 8)):
            t = r[ci]
            if t is None or not str(t).strip(): continue
            t = ' '.join(str(t).split())
            if SKIP.search(t) and not PHONE_RE.search(t): continue
            name, ph = split_cell(t)
            if name and re.search(LETT, name): drivers.append([name, ph])
            elif ph and drivers: drivers[-1][1] += [p for p in ph if p not in drivers[-1][1]]
        if drivers: out.setdefault(key_noreg(plate), []).extend(drivers)
    return out
def read_old(path):
    ws = openpyxl.load_workbook(path, data_only=True)['Мобил']
    mob = collections.OrderedDict()
    for r in ws.iter_rows(min_row=3, values_only=True):
        e, f = (str(r[4] or '').strip(), str(r[5] or '').strip())
        if not e: continue
        plate = lat((e + f).upper().replace(' ', ''))
        drivers = []
        for ci in range(7, len(r)):
            c = r[ci]
            if c in (None, ''): continue
            t = str(c).strip()
            if not t: continue
            if t.lower().startswith('вр.'):
                if drivers: drivers.pop()
                continue
            if SKIP.search(t) and not PHONE_RE.search(t): continue
            name, ph = split_cell(t)
            if name and re.search(LETT, name): drivers.append([name, ph])
            elif ph and drivers: drivers[-1][1] += [p for p in ph if p not in drivers[-1][1]]
        mob.setdefault(plate, []).extend(drivers)
    return {key_noreg(k): v for k, v in mob.items() if v}
if __name__ == '__main__':
    roster, out = sys.argv[1], sys.argv[2]
    res = {}
    if len(sys.argv) > 3: res.update(read_old(sys.argv[3]))
    elif os.path.exists(out): res.update(json.load(open(out, encoding='utf-8')))
    new = read_roster(roster); res.update(new)
    json.dump(res, open(out, 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
    print('машин в новом списке:', len(new), ' всего в базе:', len(res))
