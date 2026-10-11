import sys, re, json, collections, difflib, openpyxl

# ---------------- нормализация ----------------
KZ = str.maketrans({'ә':'а','ғ':'г','қ':'к','ң':'н','ө':'о','ұ':'у','ү':'у','һ':'х','і':'и','ё':'е','й':'и'})
def nname(s):
    s = str(s or '').lower().translate(KZ)
    s = re.sub(r'[^а-яa-z\s]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()

CYR2LAT = {'А':'A','В':'B','Е':'E','К':'K','М':'M','Н':'H','О':'O','Р':'P','С':'C','Т':'T','Х':'X','У':'Y',
           'Ү':'Y','Ұ':'Y','Қ':'K','Ә':'A','Ө':'O','І':'I','Ң':'H','Ғ':'G'}
def lat(s): return ''.join(CYR2LAT.get(c, c) for c in s)

def parse_plate(raw):
    """-> (prefix, digits, series, rest_norm) или None"""
    s = re.sub(r'\s+', '', str(raw or ''))
    if '-' not in s: return None
    pref, rest = s.split('-', 1)
    rest = re.sub(r'(d/n|д/н|д|н|d|n)$', '', rest)          # пометки смены у диспетчеров (строчные)
    rest = lat(rest.upper())
    rest = re.sub(r'[/\-_.]', '', rest)
    m = re.fullmatch(r'[HR]?(\d+)([A-Z]*?)(\d{2})?', rest)  # цифры [серия] [регион]
    if m and (m.group(2) or True):
        digits, ser, reg = m.group(1), m.group(2), m.group(3)
        if len(ser) == 4 and ser[0] == 'R': ser = ser[1:]      # 465RAKD -> AKD
        if len(ser) == 4 and ser[3] == 'R': ser = ser[:3]      # 221AKDR -> AKD
        if not ser:                                          # H2826/12 -> 282612, H2826 -> 2826
            full = digits + (reg or '')
            if len(full) >= 6 and full.endswith('12') and not (reg is None and len(digits) == 6 and not rest.startswith(('H',))):
                digits = full[:-2]
            else:
                digits = full
        return pref, digits, ser, rest
    m = re.fullmatch(r'([A-Z]{2,4})(\d+)R?', rest)           # AKD677R, AKER396 (серия впереди)
    if m:
        L = m.group(1)
        if len(L) == 4 and L[3] == 'R': L = L[:3]
        return pref, m.group(2), L, rest
    m = re.fullmatch(r'R?(\d+)([A-Z]+)(\d{2})?', rest)
    if m: return pref, m.group(1), m.group(2), rest
    return pref, rest, '', rest

# ---------------- телефоны ----------------
MOB = {'700','701','702','705','707','708','747','771','775','776','777','778'}
def fmt(d):
    """-> (строка, is_mobile) или None, если номер не похож на номер"""
    if len(d) == 11 and d[0] in '78' and d[1:4] in MOB: d = d[1:]
    if len(d) == 10 and d[:3] in MOB:
        return ('8 %s %s %s %s' % (d[:3], d[3:6], d[6:8], d[8:]), True)
    if len(d) == 10 and d[0] == '7' and d[1] in '23':             # городской: 7292-401592, 7272-900229
        return ('8 %s %s' % (d[:4], d[4:]), False)
    if len(d) == 11 and d[0] == '8' and d[1] == '7' and d[2] in '23':
        return ('8 %s %s' % (d[1:5], d[5:]), False)
    return None
def split_digits(d):
    if not d: return []
    for n in (11, 10):
        if len(d) >= n and fmt(d[:n]):
            rest = split_digits(d[n:])
            if rest is not None: return [d[:n]] + rest
    return None
def phones(cell):
    out = []
    txt = re.sub(r'(?<=\d)\s*[\-–]\s*(?=\d)', '', str(cell or ''))
    for chunk in re.split(r'[/;,\n]+|\s{2,}|\t', txt):
        d = re.sub(r'\D', '', chunk)
        if len(d) < 10: continue
        parts = split_digits(d)
        if parts is None: continue                       # мусор вместо номера - не берём
        for p in parts:
            f = fmt(p)
            if f and f not in out: out.append(f)
    return out
def best2(plist):
    mob = [p for p, m in plist if m]; land = [p for p, m in plist if not m]
    r = mob + land
    seen = []
    for x in r:
        if x not in seen: seen.append(x)
    return seen[:2]


# номера, повторяющиеся у разных людей, считаем мусором
def shared_phones(rows):
    m = collections.defaultdict(set)
    for nm, ph in rows:
        for p in ph: m[p].add(nname(nm))
    out = {}
    for p, names in m.items():
        clusters = []
        for n in names:
            ts = set(n.split())
            for c in clusters:
                cs = c[0]
                small, big = (ts, cs) if len(ts) <= len(cs) else (cs, ts)
                if len(small) >= 2 and small <= big:      # то же ФИО, но с должностью в конце
                    c[0] = big; break
            else: clusters.append([ts, n])
        if len(clusters) > 1: out[p] = sorted(c[1] for c in clusters)
    return out
