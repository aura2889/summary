"""Сборка drivers.json: водители из экспорта путевых листов + телефоны из data/contacts.json и data/mobil.json."""
import datetime, sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import *

junk, junk_lost = set(), set()
idx0, idx1, idx2, cache, amb = {}, {}, {}, {}, {}

def is_mob(num):
    """сотовый ли номер: 8 7XX…, где код после 87 начинается на 0, 4, 5, 6 или 7 (городские 87 1xx/2xx/3xx - не берём)"""
    d = re.sub(r'\D', '', str(num))
    return len(d) == 11 and d[0] == '8' and d[1] == '7' and d[2] in '04567'

OVR = {}
def load_override(root):
    """data/phone_override.json: {"ФИО": ["8 7XX XXX XX XX", ...]} - эти номера ставятся первыми"""
    OVR.clear()
    try:
        for k, v in json.load(open(os.path.join(root, 'data', 'phone_override.json'), encoding='utf-8')).items(): OVR[nname(k)] = v
    except Exception: pass

def build_idx(rows):
    idx = collections.defaultdict(list)
    for nm, ph in rows:
        ph = [x for x in (ph or []) if is_mob(x)]
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
    """главный справочник (rows0) - первым; второй номер, если его нет, берём из старой базы"""
    if name in cache: return cache[name]
    res = None; first = None
    for label, idx in (('Справочник КБМ', idx0), ('Водители', idx1), ('КБМ', idx2)):
        r = lookup(idx, name)
        if r and r[0] != 'AMBIG':
            if first is None:
                first = (list(r[0]), label, r[1], r[2])
                if label != 'Справочник КБМ' or len(r[0]) >= 2: break
            elif first[1] == 'Справочник КБМ':
                for p in r[0]:
                    if p not in first[0] and len(first[0]) < 2: first[0].append(p)
                break
        if r and r[0] == 'AMBIG' and first is None: amb[name] = (label, r[2])
    o = OVR.get(nname(name))
    if o:
        base = list(first[0]) if first else []
        first = ([x for x in o + base if is_mob(x)][:2], 'вручную', 'вручную', name) if not first else (list(dict.fromkeys(o + base))[:2], first[1], first[2], first[3])
    cache[name] = first
    return first
amb = {}

ALIAS = {'56-396ARE': '56-AKER396', '56-395ARE': '56-AKER395', '35-903AC12': '35-903AG12', '20-586AG12': '20-586AI12'}   # опечатка диспетчеров, подтверждено пользователем

def key_noreg(p): return re.sub(r'(?<=[A-Z])\d{2}$', '', p)

def export_cols(header):
    """номера колонок экспорта по заголовкам: (гос№, ФИО1, ФИО2, вид оплаты) или None. Формат выгрузки менялся, поэтому ищем по названиям"""
    h = [str(x or '').replace('\xa0', ' ').strip() for x in header]
    plate = next((i for i, x in enumerate(h) if x.startswith('Государственный') or x == 'Гос №'), None)
    fio = [i for i, x in enumerate(h) if x.startswith('Водитель Ф.И.О')]
    pay = next((i for i, x in enumerate(h) if x.startswith('Вид оплаты')), None)
    if plate is None or len(fio) < 1: return None
    return plate, fio[0], (fio[1] if len(fio) > 1 else None), pay

def fleet_block(data):
    """статичные данные парка: база КБМ (префикс 13: год выпуска, № а/к, пробег по одометру) и пробег по путевым листам за октябрь 2026.
       -> {num: {'y','ak','km','o','m'(база), 'ok','on'(октябрь)}} по номерам из свода"""
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    def ld(n):
        try: return json.load(open(os.path.join(root, 'data', n), encoding='utf-8'))
        except Exception: return None
    base, mil = ld('fleet_base.json'), ld('mileage_oct2026.json')
    if not base and not mil: return {}
    B = (base or {}).get('plates', {}); M = (mil or {}).get('plates', {})
    R, Ef = {}, collections.defaultdict(list)          # R: (префикс, цифры, серия) -> запись; в базе и в путевых листах номер может быть написан по-разному (кириллица/латиница)
    for pl in set(B) | set(M):
        pp = parse_plate(pl)
        if not pp: continue
        k = pp[:3]; r = R.setdefault(k, {})
        b = B.get(pl)
        if b: r.update({'y': b.get('y'), 'ak': b.get('ak'), 'km': b.get('km'), 'o': b.get('o'), 'm': b.get('m')})
        m = M.get(pl)
        if m: r.update({'ok': m['km'], 'on': m['n']})
        if k not in Ef[(pp[0], pp[1])]: Ef[(pp[0], pp[1])].append(k)
    out = {}
    for v in data['veh']:
        num = v['num']
        if num.startswith('60-') or '-' not in num: continue
        al = ALIAS.get(num)
        pp = parse_plate(al or num)
        if not pp: continue
        c = Ef.get((pp[0], pp[1]), [])
        if pp[2]: c = [x for x in c if x[2] == pp[2]]
        if len(c) != 1: continue
        if R[c[0]]: out[num] = R[c[0]]
    out['_'] = {'base': (base or {}).get('asof'), 'mil': (mil or {}).get('period')}
    return out

def build(data, export_path, contacts, mobil):
    """data - dict из svod2json; contacts - содержимое contacts.json; mobil - содержимое mobil.json
    -> (drivers_dict, список строк отчёта, список имён без телефона)"""
    global junk, idx0, idx1, idx2
    junk_lost.clear(); cache.clear(); amb.clear()
    junk = set(contacts['junk']); load_override(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    idx0 = build_idx([(n, p) for n, p in contacts.get('rows0', [])])
    idx1 = build_idx([(n, p) for n, p in contacts['rows1']])
    idx2 = build_idx([(n, p) for n, p in contacts['rows2']])
    ex = openpyxl.load_workbook(export_path, data_only=True).active
    allrows = list(ex.iter_rows(values_only=True))
    cp, c1, c2, cs = export_cols(allrows[0])
    hh = [str(x or '').strip() for x in allrows[0]]
    cd = next((i for i, x in enumerate(hh) if x == 'Дата' or x.startswith('Дата путевого')), None)
    exrows = [r for r in allrows[1:] if len(r) > cp and r[cp]]
    # --- какая из строк по одному госномеру актуальна (после пересменки/перевахтовки номер встречается дважды+).
    # Смотрим не на положение строки, а на «Дату ввода операции» (если её нет - на «Время ввода операции»); «Дата путевого листа» не используется.
    # Колонки ищем по заголовкам: порядок в выгрузке меняется.
    hn = [str(x or '').replace('\xa0', ' ').strip().lower() for x in allrows[0]]
    c_dt = next((i for i, x in enumerate(hn) if x.startswith('дата ввода')), None)
    c_tm = next((i for i, x in enumerate(hn) if x.startswith('время ввода')), None)
    c_jr = next((i for i, x in enumerate(hn) if x.startswith('№ журнала') or x == 'номер журнала'), None)
    def _tm(v):
        if v is None or v == '': return None
        if hasattr(v, 'hour') and not hasattr(v, 'year'): return v.hour * 3600 + v.minute * 60 + getattr(v, 'second', 0)
        if hasattr(v, 'year'): return (v.hour * 3600 + v.minute * 60 + v.second) if (v.hour or v.minute or v.second) else None
        m = re.match(r'(\d{1,2})[:.](\d{2})(?::(\d{2}))?', str(v).strip())
        return int(m.group(1)) * 3600 + int(m.group(2)) * 60 + int(m.group(3) or 0) if m else None
    def _dt(v):
        if hasattr(v, 'toordinal'): return v.toordinal()
        m = re.match(r'(\d{4})-(\d{2})-(\d{2})', str(v or ''))
        if m: return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3))).toordinal()
        m = re.match(r'(\d{1,2})[./](\d{1,2})[./](\d{4})', str(v or ''))
        return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1))).toordinal() if m else None
    def entry_key(r):
        d = _dt(r[c_dt]) if c_dt is not None and len(r) > c_dt else None
        t = _tm(r[c_tm]) if c_tm is not None and len(r) > c_tm else (_tm(r[c_dt]) if d is not None else None)
        if d is None and t is not None:       # даты нет, только время: обед 11:00-15:00 (перевахтовка) - свежее, утро до 11 - дальше, после 15:00 - вчерашний вечер
            tier = 3 if 11 * 3600 <= t <= 15 * 3600 else (2 if t < 11 * 3600 else 1)
            return (0, tier, t, jr_of(r))
        return (d or 0, 0, t or 0, jr_of(r))
    def jr_of(r):
        try: return int(r[c_jr]) if c_jr is not None and r[c_jr] not in (None, '') else 0
        except (TypeError, ValueError): return 0
    E = collections.defaultdict(list)
    EP = collections.OrderedDict(); best = {}
    import datetime as _d
    _now = _d.datetime.utcnow() + _d.timedelta(hours=5)           # Актау/Актобе UTC+5
    morning = (_now.hour < 12) if not os.environ.get('EXPORT_PART') else os.environ['EXPORT_PART'] == 'morning'
    for ri, r in enumerate(exrows):
        plate = str(r[cp]).strip(); pp = parse_plate(plate)
        names = [str(x).strip() for x in (r[c1], r[c2] if c2 is not None else None) if x and str(x).strip()]
        ek = entry_key(r)
        pos = -ri if morning else ri       # одна и та же дата: утром (до 12:00) верхняя строка файла, после обеда - нижняя
        k = (ek[0], ek[1], 1 if names else 0, pos) if ek[0] == 0 else (ek[0], 0, 1 if names else 0, pos)      # свежая дата, потом время; при равенстве - строка с водителем, потом больший № журнала
        if plate in best and best[plate] > k: continue
        best[plate] = k
        inf = EP[plate] = {'pp': pp, 'status': set(), 'names': []}
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
            d = {'n': nm}; ph = [x for x in best2(ph) if is_mob(x)]
            if ph: d['t'] = ph[0]
            if len(ph) > 1: d['t2'] = ph[1]
            out.append(d)
        drv['map'][num] = out; mob_added += 1
    try: drv['fleet'] = fleet_block(data)
    except Exception as e: print('fleet:', e)
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
