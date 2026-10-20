"""Сборка drivers.json: водители из экспорта путевых листов + телефоны из data/contacts.json и data/mobil.json."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import *

junk, junk_lost = set(), set()
idx1, idx2, cache, amb = {}, {}, {}, {}

def build_idx(rows):
    idx = collections.defaultdict(list)
    for nm, ph in rows:
        n = nname(nm)
        if n: idx[n].append((str(nm).strip(), ph or []))
    return idx

def lookup(idx, name):
    n = nname(name); toks = n.split()
    def same_person(a, b):
        ta, tb = set(nname(a).split()), set(nname(b).split())
        small, big = (ta, tb) if len(ta) <= len(tb) else (tb, ta)
        return len(small) >= 2 and small <= big
    def pick(cands, how):
        cands = [c for c in cands if c[1]]
        if not cands: return None
        kept = [(nm, [p for p in ph if p not in junk]) for nm, ph in cands]
        kept = [c for c in kept if c[1]]
        if not kept:
            junk_lost.add(name)
            if how == 'точно' and len({tuple(c[1]) for c in cands}) == 1:
                return (cands[0][1][:2], 'точно ОБЩИЙ', cands[0][0])
            return None
        uniq = {tuple(c[1]) for c in kept}
        if len(uniq) == 1: return (kept[0][1], how, kept[0][0])
        if all(same_person(kept[0][0], c[0]) for c in kept):      # одна запись с должностью, другая без
            u = []
            for c in kept:
                for p in c[1]:
                    if p not in u: u.append(p)
            return (u[:2], how + ' (2 записи)', kept[0][0])
        return ('AMBIG', how, [c[0] for c in kept])
    if n in idx:
        r = pick(idx[n], 'точно')
        if r: return r
    ts = set(toks)
    cand = []
    if len(ts) >= 2:
        for k, v in idx.items():
            kt = set(k.split()); small, big = (ts, kt) if len(ts) <= len(kt) else (kt, ts)
            if len(small) >= 2 and small <= big: cand += v
    if cand:
        r = pick(cand, 'по частям имени')
        if r: return r
    if len(toks) >= 2:
        cand = []
        for k, v in idx.items():
            kt = k.split()
            if len(kt) >= 2 and (kt[:2] == toks[:2] or (kt[:2][::-1] == toks[:2])): cand += v
        if cand:
            r = pick(cand, 'фамилия+имя')
            if r: return r
    return None

cache = {}
junk_lost = set()
def phones_for(name):
    if name in cache: return cache[name]
    res = None
    for label, idx in (('Водители', idx1), ('КБМ', idx2)):
        r = lookup(idx, name)
        if r and r[0] != 'AMBIG': res = (r[0], label, r[1], r[2]); break
        if r and r[0] == 'AMBIG' and res is None: amb[name] = (label, r[2])
    cache[name] = res
    return res
amb = {}

ALIAS = {'56-396ARE': '56-AKER396', '56-395ARE': '56-AKER395'}   # опечатка диспетчеров, подтверждено пользователем

def key_noreg(p): return re.sub(r'(?<=[A-Z])\d{2}$', '', p)

def export_cols(header):
    """номера колонок экспорта по заголовкам: (гос№, ФИО1, ФИО2, вид оплаты) или None. Формат выгрузки менялся, поэтому ищем по названиям"""
    h = [str(x or '').replace('\xa0', ' ').strip() for x in header]
    plate = next((i for i, x in enumerate(h) if x.startswith('Государственный') or x == 'Гос №'), None)
    fio = [i for i, x in enumerate(h) if x.startswith('Водитель Ф.И.О')]
    pay = next((i for i, x in enumerate(h) if x.startswith('Вид оплаты')), None)
    if plate is None or len(fio) < 1: return None
    return plate, fio[0], (fio[1] if len(fio) > 1 else None), pay

def build(data, export_path, contacts, mobil, manual=None):
    """data - dict из svod2json; contacts - содержимое contacts.json; mobil - содержимое mobil.json
    -> (drivers_dict, список строк отчёта, список имён без телефона)"""
    global junk, idx1, idx2
    junk_lost.clear(); cache.clear(); amb.clear()
    junk = set(contacts['junk'])
    idx1 = build_idx([(n, p) for n, p in contacts['rows1']])
    idx2 = build_idx([(n, p) for n, p in contacts['rows2']])
    ex = openpyxl.load_workbook(export_path, data_only=True).active
    allrows = list(ex.iter_rows(values_only=True))
    cp, c1, c2, cs = export_cols(allrows[0])
    hh = [str(x or '').strip() for x in allrows[0]]
    cd = next((i for i, x in enumerate(hh) if x == 'Дата' or x.startswith('Дата путевого')), None)
    exrows = [r for r in allrows[1:] if len(r) > cp and r[cp]]
    E = collections.defaultdict(list)
    EP = collections.OrderedDict()
    for r in exrows:
        plate = str(r[cp]).strip(); pp = parse_plate(plate)
        names = [str(x).strip() for x in (r[c1], r[c2] if c2 is not None else None) if x and str(x).strip()]
        inf = EP.setdefault(plate, {'pp': pp, 'status': set(), 'names': []})
        inf['status'].add(r[cs] if cs is not None else None)
        for n in names:
            if n not in inf['names']: inf['names'].append(n)
    for plate, inf in EP.items():
        if inf['pp']:
            pref, dg, ser, _ = inf['pp']
            E[(pref, dg)].append((ser, plate))
    res, ambig_plates = {}, []
    for v in data['veh']:
        num = v['num']
        if num.startswith('60-') or '-' not in num: continue
        plate = ALIAS.get(num)
        if not plate:
            pp = parse_plate(num)
            if not pp: continue
            pref, dg, ser, _ = pp
            c = E.get((pref, dg), [])
            if ser: c = [x for x in c if x[0] == ser]
            c = list(dict.fromkeys(x[1] for x in c))
            if len(c) > 1: ambig_plates.append((num, c)); continue
            if not c: continue
            plate = c[0]
        if plate not in EP: continue
        res[num] = plate
    drv = {'date': data['date'], 'map': {}}
    miss_phone = []
    for num, plate in res.items():
        lst = []
        for nm in EP[plate]['names']:
            e = {'n': nm}
            r = phones_for(nm)
            if r:
                if r[0]: e['t'] = r[0][0]
                if len(r[0]) > 1: e['t2'] = r[0][1]
                if 'ОБЩИЙ' in r[2]: e['s'] = 1
            else: miss_phone.append(nm)
            lst.append(e)
        drv['map'][num] = lst
    # --- Мобил Сервис (60-XX-цифрыСЕРИЯ)
    mob_added, mob_miss = 0, []
    for v in data['veh']:
        num = v['num']
        if not num.startswith('60-'): continue
        m = re.match(r'60-[\d.]+-(.+)$', num)
        if not m: mob_miss.append(num); continue
        rest = lat(re.sub(r'(д/н|д|н)$', '', m.group(1)).upper()).replace(' ', '')
        k = key_noreg(rest)
        if k not in mobil: mob_miss.append(num); continue
        out = []
        for nm, ph in mobil[k]:
            d = {'n': nm}; ph = best2(ph)
            if ph: d['t'] = ph[0]
            if len(ph) > 1: d['t2'] = ph[1]
            out.append(d)
        drv['map'][num] = out; mob_added += 1
    for num, lst in (manual or {}).items():          # водители, введённые вручную (data/manual.json): для машин, которых нет в экспорте путевых листов
        if num in drv['map'] or not lst: continue
        if any(v['num'] == num for v in data['veh']): drv['map'][num] = [dict(x) for x in lst]
    veh = data['veh']; mp = drv['map']
    allrec = sum(len(l) for l in mp.values()); withp = sum(1 for l in mp.values() for d in l if d.get('t'))
    nofind = [v['num'] for v in veh if v['num'] not in mp and '-' in v['num']]
    dates = set()
    for r in exrows:
        d = r[cd] if cd is not None else None
        if hasattr(d, 'strftime'): dates.add(d.strftime('%Y-%m-%d'))
    L = [f'Машин в своде: {len(veh)}; водитель найден у {len(mp)} (из них Mobil {mob_added})',
         f'Записей водителей: {allrec}, с телефоном: {withp}',
         f'Дата путевых листов в экспорте: {", ".join(sorted(dates)) or "?"}; дата свода: {data["date"]}']
    if ambig_plates: L.append('Неоднозначные номера: ' + str(ambig_plates))
    L.append(f'Не найдены в экспорте ({len(nofind)}): ' + ', '.join(nofind))
    if mob_miss: L.append('Mobil не найдены: ' + ', '.join(mob_miss))
    miss = sorted(set(miss_phone) | {d['n'] for l in mp.values() for d in l if not d.get('t')})
    return drv, L, miss
