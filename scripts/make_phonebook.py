# -*- coding: utf-8 -*-
"""Основной список телефонов (файл КБМ: листы КБМ1…КБМ7, ФИО в B, мобильный в E, домашний в F) -> rows0 для data/contacts.json.
В базу попадают только те, кого видели среди водителей (экспорты/drivers.json) или у кого подразделение «Автоколонна» - чтобы не публиковать всех сотрудников."""
import os, re, glob, json, collections, openpyxl
from common import nname, phones, best2, shared_phones

def is_phonebook(path):
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        if not wb.sheetnames or not wb.sheetnames[0].upper().startswith('КБМ'): return False
        h = [str(x or '') for x in next(wb[wb.sheetnames[0]].iter_rows(max_row=1, values_only=True), ())]
        return 'Полное имя' in h and any('Мобильный' in x for x in h)
    except Exception: return False

def known_names(root):
    names = set()
    for f in glob.glob(os.path.join(root, 'archive', '*', 'drivers.json')) + [os.path.join(root, 'drivers.json')]:
        try:
            for lst in json.load(open(f, encoding='utf-8'))['map'].values():
                for d in lst: names.add(nname(d['n']))
        except Exception: pass
    exp = os.path.join(root, 'raw', 'export.xlsx')
    if os.path.exists(exp):
        ws = openpyxl.load_workbook(exp, data_only=True).active
        for r in ws.iter_rows(min_row=2, values_only=True):
            for i in (7, 9):
                if len(r) > i and r[i] and str(r[i]).strip(): names.add(nname(r[i]))
    return names

def read_phonebook(path, root):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    rows = []
    for ws in wb.worksheets:
        if not ws.title.upper().startswith('КБМ'): continue
        for r in ws.iter_rows(min_row=2, values_only=True):
            if len(r) < 5 or not r[1]: continue
            nm = re.sub(r'\s+', ' ', str(r[1])).strip()
            ph = phones(r[4]) + ([p for p in phones(r[5]) if p not in phones(r[4])] if len(r) > 5 else [])
            ph = best2(ph)
            if ph: rows.append((nm, ph, 'автоколонна' in str(r[2] or '').lower()))
    return rows

def build_rows0(path, root):
    allrows = read_phonebook(path, root)
    junk = shared_phones([(n, p) for n, p, _ in allrows])
    names = known_names(root)
    byfirst = collections.defaultdict(list)
    for n in names:
        for t in n.split(): byfirst[t].append(n)
    def same(a, b):
        ta, tb = a.split(), b.split(); sa, sb = set(ta), set(tb)
        small, big = (sa, sb) if len(sa) <= len(sb) else (sb, sa)
        if len(small) >= 2 and small <= big: return True
        return len(ta) >= 2 and len(tb) >= 2 and (ta[:2] == tb[:2] or ta[:2][::-1] == tb[:2])
    def relevant(nm):
        n = nname(nm); cand = set()
        for t in n.split(): cand.update(byfirst.get(t, ()))
        return any(same(n, c) for c in cand)
    rows0 = [[n, p] for n, p, ak in allrows if ak or relevant(n)]
    return rows0, sorted(junk)
