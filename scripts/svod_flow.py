"""Сборка «Свода диспетчерских файлов» из 7 файлов диспетчеров (правила - в svod/parse_current.py и svod/build_current.py).
Вызывается из build.py. Недостающие/необновлённые файлы заменяются последним листом из прошлых данных и помечаются."""
import os, re, sys, json, shutil, pickle, subprocess, tempfile, datetime
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__)); SV = os.path.join(HERE, 'svod')
sys.path.insert(0, SV)
import parse_current as PC

# роль -> (основа названия в «Исходных данных», строки, столбцы, прочее)
ROLES = {
    'D3':   dict(base='Диспетчер №3',                 rows=(3, 60), cols=(1, 13)),
    'MS':   dict(base='Мобил Сервис (60-…)',          rows=(3, 40), cols=(2, 19)),
    'D7':   dict(base='Диспетчер №7 (ПКРС/ТМС/УТТ)',  rows=(2, 60), cols=(1, 10), hrow=1),
    'SVOD': dict(base='Сводка от (27-…, 28-…)',       rows=(3, 25), cols=(1, 15)),
    'BUS':  dict(base='Автобусы / АЦН',               rows=(3, 25), cols=(1, 8)),
    'K8':   dict(base='К-8 (месячный файл)',          rows=(2, 29), cols=(1, 11), hrow=1),
    'SB':   dict(base='Ақпар (С13…)',                 rows=(3, 40), cols=(1, 11)),
}
ROLE_DISP = {'SB': 'Диспетчер 1', 'SVOD': 'Диспетчер 2', 'D3': 'Диспетчер 3', 'BUS': 'Диспетчер 4',
             'D7': 'Диспетчер 7', 'K8': 'Диспетчер 8', 'MS': 'Мобил Сервис'}
ROLE_TITLE = {k: v for k, v in ROLE_DISP.items()}
LAST_TRAIN = []

def _n(x): return re.sub(r'[\s.]', '', str(x or '')).upper()

# ---------- объединённый файл «Гараж-2 техника шығу ақпары» (диспетчеры 7 и 8 вместе, с 23.10) ----------
GAR2 = dict(rows=(2, 66), cols=(1, 13), hrow=1,
            alias={'ПКРС(2)': 'ЦПКРС-2', 'БПОКРС-2': 'ЦПКРС-2', 'ОБИКРС': 'ОБиПКРС', 'ГАЗУЧАСТОК': 'ГУ', 'ЦТТГАРАЖ2': 'Гараж-2', 'РМУГАРАЖ2': 'Гараж-2', 'ДБИПКРС(1)': 'ЦПКРС-1'})

def plate_key(raw):
    """ключ машины для сопоставления между файлами (без «С»-префикса, пометок смены, регистра/раскладки)"""
    from common import parse_plate
    from build_drivers import ALIAS
    t = re.sub(r'\s+', '', str(raw or ''))
    t = ALIAS.get(t.upper(), t)
    r = parse_plate(t)
    if not r: return t.upper()
    pref, digits, ser, rest = r
    return re.sub(r'^[СC]', '', pref) + '|' + digits + ser

def _split_path(raw_dir): return os.path.join(raw_dir, '..', 'data', 'gar2_split.json')

def _load_split(raw_dir, ref):
    """{ключ машины: 'D7'|'K8'} - кто из двух диспетчеров вёл машину; при первом разе берётся из последних отдельных файлов Д7 и К-8"""
    p = _split_path(raw_dir)
    if os.path.exists(p): return json.load(open(p, encoding='utf-8'))
    m = {}
    for role in ('D7', 'K8'):
        f = os.path.join(raw_dir, role + '.xlsx')
        if not os.path.exists(f): continue
        wbk = openpyxl.load_workbook(f); sh = latest_before(wbk.sheetnames, ref) or pick_sheet(wbk.sheetnames, ref)[0]
        spec = ROLES[role]; cfg = dict(key=role, name='x', path=f, sheet=sh, rows=spec['rows'], cols=spec['cols'], mode='day')
        if 'hrow' in spec: cfg['hrow'] = spec['hrow']
        _adapt(cfg, wbk[sh]); out, exc, _ = PC.parse(cfg)
        for x in out: m[plate_key(x['plate'])] = role
        for e in exc: m.setdefault(plate_key(re.split(r'\s{2,}', str(e[1]).strip())[0]), role)
    return m

def parse_gar2(raw_dir, ref, today):
    """-> (out_D7, out_K8, exc_D7, exc_K8, warns, sheet) или None, если листа за сегодня нет/пусто"""
    f = os.path.join(raw_dir, 'GAR2.xlsx')
    if not os.path.exists(f): return None
    wbk = openpyxl.load_workbook(f); sheet, exact, w = pick_sheet(wbk.sheetnames, ref)
    if sheet is None or not exact: return None
    cfg = dict(key='GAR2', name='x', path=f, sheet=sheet, rows=GAR2['rows'], cols=GAR2['cols'], mode='day', hrow=GAR2['hrow'], alias=GAR2['alias'])
    out, exc, notes = PC.parse(cfg)
    if not out: return None
    warns = [f'Гараж-2 (Д7+К-8): {x}' for x in w]
    split = _load_split(raw_dir, ref)
    by = {'D7': [], 'K8': []}
    for k, r in split.items(): by[r].append(k)
    cnt = {}                                              # по типу машины (префиксу) - чьи они чаще
    for k, r in split.items(): cnt.setdefault(k.split('|')[0], {'D7': 0, 'K8': 0})[r] += 1
    new = []
    def who(plate):
        k = plate_key(plate)
        if k in split: return split[k]
        if not re.search(r'\d{3}', k.split('|')[-1]): return 'K8'           # ячейка без номера (только тип) - ни в какие списки не попадает
        c = cnt.get(k.split('|')[0]); r = ('D7' if c['D7'] > c['K8'] else 'K8') if c else 'K8'
        split[k] = r; new.append((plate, r)); return r
    res = {'D7': [], 'K8': []}; rexc = {'D7': [], 'K8': []}
    for x in out:
        r = who(x['plate']); x['src'] = r; res[r].append(x)
    for e in exc: rexc[who(re.split(r'\s{2,}', str(e[1]).strip())[0])].append(e)
    for plate, r in new: warns.append(f'Гараж-2: новая машина {plate} - отнесена к {ROLE_DISP[r]} по типу (раньше в отдельных файлах не встречалась)')
    if new: json.dump(split, open(_split_path(raw_dir), 'w', encoding='utf-8'), ensure_ascii=False, sort_keys=True)
    elif not os.path.exists(_split_path(raw_dir)): json.dump(split, open(_split_path(raw_dir), 'w', encoding='utf-8'), ensure_ascii=False, sort_keys=True)
    return res, rexc, warns, sheet

# ---------- какой это файл (по содержимому) ----------
def _sheet_role(rows):
    r1 = [_n(c) for c in rows[0]] if rows else []
    r2 = [_n(c) for c in rows[1]] if len(rows) > 1 else []
    if r1 and r1[0].startswith('ДИСПЕТЧЕР№3'): return 'D3'
    if r1 and r1[0].startswith('ДИСПЕТЧЕРА') and r2 and r2[0] == 'ПКРС' and any(c in ('РЕЗЕРВ', 'СПИСАНИЕ') for c in r2): return 'D7'   # новый вид файла диспетчера 7 (с 16.10)
    if r1 and r1[0].startswith('ДИСПЕТЧЕРА') and any('ЦПП-МПГУ' in c for c in r2): return 'K8'                                          # новый вид файла К-8 (с 16.10)
    if r1 and r1[0] == 'ЦДН-1' and any('ДБИПКРС(1)' in c for c in r1) and 'РСУ' in r1: return 'GAR2'      # объединённый файл Д7+К-8 (с 23.10)
    if any('ДБИПКРС' in c for c in r1): return 'D7'
    if r2 and r2[0] == '№': return 'MS'
    if r2 and r2[0] == 'СБ': return 'SB'
    if 'СМЦ' in r2: return 'SVOD'
    if any('ЦПП-МПГУ' in c for c in r1): return 'K8'
    if 'ТМС' in r2 and 'ДТТ' in r2: return 'BUS'
    return None

def detect_role(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    votes, seen = {}, 0
    for sn in reversed(wb.sheetnames):
        if not re.match(r'^\s*\d', sn): continue
        rows = list(wb[sn].iter_rows(min_row=1, max_row=3, max_col=16, values_only=True))
        if not any(c is not None for r in rows for c in r): continue
        role = _sheet_role(rows); seen += 1
        if role: votes[role] = votes.get(role, 0) + 1
        if seen >= 4: break
    wb.close()
    if not votes: return None
    best = sorted(votes.items(), key=lambda kv: -kv[1])
    if len(best) > 1 and best[0][1] == best[1][1]: return None
    return best[0][0]

# ---------- какой лист брать ----------
_SH = re.compile(r'^\s*(\d{1,2})(?:\s*[.,]\s*(\d{1,2}))?')
def sheet_info(name):
    m = _SH.match(name)
    if not m: return None
    day = int(m.group(1)); mon = int(m.group(2)) if m.group(2) else None
    rest = re.sub(r'[\s.,]|20\d\d', '', name[m.end():])
    if not (1 <= day <= 31) or (mon is not None and not 1 <= mon <= 12): return None
    return dict(name=name, day=day, mon=mon, extra=bool(rest), xl=len(rest))

def pick_sheet(names, ref):
    """-> (имя листа, точный ли (за ref), предупреждения)"""
    infos = [dict(i=i, **si) for i, n in enumerate(names) if (si := sheet_info(n))]
    warns = []
    cands = [x for x in infos if x['day'] == ref.day and x['mon'] in (None, ref.month)]
    if any(x['mon'] for x in cands): cands = [x for x in cands if x['mon']]     # лист с явной датой «14,10» важнее старого листа «14»
    if cands:
        use = next(x for x in cands if x['name'] == today_sheets(names, ref)[0])          # лист с припиской («пер», «пересменка»…) свежее обычного; из равных - самый правый; пустые листы пропускает сборка
        if len(cands) > 1: warns.append('несколько листов за сегодня (' + ', '.join(repr(c['name']) for c in cands) + f'), взят {use["name"]!r}')
        return use['name'], True, warns
    lb = latest_before(names, ref)
    if lb: return lb, False, warns
    if infos: return infos[-1]['name'], False, warns
    return None, False, warns

def today_sheets(names, ref):
    """все листы за сегодня справа налево (самый правый, т.е. самый свежий, первым)"""
    infos = [dict(i=i, **si) for i, n in enumerate(names) if (si := sheet_info(n))]
    cands = [x for x in infos if x['day'] == ref.day and x['mon'] in (None, ref.month)]
    if any(x['mon'] for x in cands): cands = [x for x in cands if x['mon']]
    return [x['name'] for x in sorted(cands, key=lambda x: (-x['xl'], -x['i']))]      # чем длиннее приписка к названию («пер», «ауысым», «пересменка»…), тем свежее лист; при равной длине - правый

def latest_before(names, ref):
    """самый свежий по ДАТЕ лист раньше ref (порядок листов в книге не важен: книга годовая, листы могут идти не по порядку); при равных датах - последний по порядку (напр. «17-пер» после «17,10»)"""
    best = None
    for i, n in enumerate(names):
        si = sheet_info(n)
        if not si: continue
        d = sheet_date(n, ref)
        if d is None or d >= ref: continue
        if best is None or d >= best[0]: best = (d, n)
    return best[1] if best else None

def sheet_date(si_name, ref):
    si = sheet_info(si_name)
    if not si: return None
    mon = si['mon'] or ref.month; year = ref.year
    if si['mon'] is None and si['day'] > ref.day: mon = ref.month - 1 or 12; year = ref.year if ref.month > 1 else ref.year - 1
    try: return datetime.date(year, mon, si['day'])
    except ValueError: return None

# ---------- сборка ----------
def _adapt(cfg, ws):
    """новые виды файлов (с 16.10): сверху добавлена строка-заголовок «Диспетчер ақпары…», у диспетчера 7 другая раскладка колонок"""
    r1 = [_n(ws.cell(1, c).value) for c in range(1, 4)]; r2 = [_n(ws.cell(2, c).value) for c in range(1, 15)]
    if not r1[0].startswith('ДИСПЕТЧЕРА'): return
    if cfg['key'] == 'D7' and r2[0] == 'ПКРС':
        cfg.update(rows=(3, 60), cols=(1, 14), hrow=2)
    elif cfg['key'] == 'K8':
        cfg.update(rows=(cfg['rows'][0] + 1, cfg['rows'][1] + 1), hrow=cfg.get('hrow', 1) + 1)


KNOWN = json.load(open(os.path.join(SV, 'known.json'), encoding='utf-8')) if os.path.exists(os.path.join(SV, 'known.json')) else None

def unknown_items(path, role, ref):
    """новые (неизвестные правилам) отделы/префиксы в файле роли - чтобы не принять за файл диспетчера чужой файл похожего вида"""
    if not KNOWN: return set()
    wbk = openpyxl.load_workbook(path); sh = pick_sheet(wbk.sheetnames, ref)[0]
    if sh is None: return set()
    spec = ROLES[role]
    cfg = dict(key=role, name='x', path=path, sheet=sh, rows=spec['rows'], cols=spec['cols'], mode='day')
    if 'hrow' in spec: cfg['hrow'] = spec['hrow']
    _adapt(cfg, wbk[sh])
    out = PC.parse(cfg)[0]
    return ({x['dept'] for x in out} - set(KNOWN['depts']) - {'ПЧ'}) | ({x['code'] for x in out} - set(KNOWN['codes']))

def build(raw_dir, ref, fresh, out_dir, via=()):
    """raw_dir/<ROLE>.xlsx; fresh = {ROLE: True если файл пришёл сегодня}. -> (путь к xlsx, статусы ролей, предупреждения, ошибки)"""
    ds = ref.strftime('%d.%m.%Y'); short = ref.strftime('%d.%m')
    warns, errs, status, ALL, EXC, NOTES = [], [], {}, [], [], []
    gar = None
    if via:                                                 # диспетчеры 7 и 8 пришли одним файлом (Гараж-2)
        try: gar = parse_gar2(raw_dir, ref, True)
        except Exception as e: warns.append(f'Гараж-2 (Д7+К-8): не удалось разобрать объединённый файл: {e}')
        if gar: warns += gar[2]
        else: warns.append('Гараж-2 (Д7+К-8): в объединённом файле нет листа за сегодня - беру отдельные файлы Д7/К-8')
    for role, spec in ROLES.items():
        if gar and role in via:
            out = gar[0][role]; exc = gar[1][role]
            for x in out: x['srcname'] = f'{spec["base"]} ({short})'
            status[role] = dict(state='ok', sheet=gar[3], date=ref.isoformat())
            ALL += out; EXC += [(role,) + x for x in exc]; continue
        path = os.path.join(raw_dir, role + '.xlsx')
        if not os.path.exists(path):
            status[role] = dict(state='absent'); warns.append(f'{ROLE_DISP[role]}: файла ещё не было ни разу - в своде его нет'); continue
        wbk = openpyxl.load_workbook(path); names = wbk.sheetnames
        sheet, exact, w = pick_sheet(names, ref)
        warns += [f'{ROLE_DISP[role]}: {x}' for x in w]
        if sheet is None:
            status[role] = dict(state='absent'); warns.append(f'{ROLE_DISP[role]}: не нашёл листов с датами'); continue
        plans = []
        if exact and fresh.get(role, False): plans += [(c, True) for c in today_sheets(names, ref)]          # справа налево: пустой лист пропускаем, берём следующий слева
        elif exact: warns.append(f'{ROLE_DISP[role]}: файл за {short} ещё не приходил (в старом файле лист за сегодня не считается)')
        if exact:
            plans.append((latest_before(names, ref), False))
        else: plans.append((sheet, False))
        used = None
        for sh, today in plans:
            if sh is None: continue
            cfg = dict(key=role, name=f'{spec["base"]} ({short})', path=path, sheet=sh, rows=spec['rows'], cols=spec['cols'], mode='day')
            if 'hrow' in spec: cfg['hrow'] = spec['hrow']
            _adapt(cfg, wbk[sh])
            out, exc, notes = PC.parse(cfg)
            if today and not out:
                warns.append(f'{ROLE_DISP[role]}: лист {sh!r} за сегодня пустой - смотрю лист левее/прошлый'); continue
            used = (cfg, out, exc, notes, sh, today); break
        if not used:
            status[role] = dict(state='absent'); warns.append(f'{ROLE_DISP[role]}: не удалось взять ни один лист'); continue
        cfg, out, exc, notes, sheet, today = used
        d = ref if today else (sheet_date(sheet, ref) or ref - datetime.timedelta(days=1))
        if not today:
            for x in out: x['srcname'] = f'{spec["base"]} (данные за {d.strftime("%d.%m")})'
        status[role] = dict(state='ok' if today else 'stale', sheet=sheet, date=d.isoformat())
        if not today: warns.append(f'{ROLE_DISP[role]}: нет свежего файла за {short}, взят лист {sheet!r} ({d.strftime("%d.%m")})')
        ALL += out; EXC += [(role,) + x for x in exc]; NOTES += [(role,) + x for x in notes]
    global LAST_TRAIN
    LAST_TRAIN = [dict(num=re.split(r'\s{2,}', e[2].strip())[0].strip(), dep=e[3], disp=ROLE_DISP[e[0]]) for e in EXC if 'обучение' in str(e[4])]   # машины на обучении (жёлтая заливка при «Оқу»): в выход не входят
    if not ALL: return None, status, warns, ['нет ни одного распознанного файла диспетчера']
    # ---- проверки (строгий режим): новое незнакомое - не публиковать
    if KNOWN:
        nd = sorted({x['dept'] for x in ALL} - set(KNOWN['depts']) - {'ПЧ'})
        nc = sorted({x['code'] for x in ALL} - set(KNOWN['codes']))
        if nd: errs.append('новый(е) отдел(ы): ' + ', '.join(nd) + ' (файлы: ' + ', '.join(sorted({x['src'] for x in ALL if x['dept'] in nd})) + ')')
        if nc: errs.append('новый(е) префикс(ы): ' + ', '.join(nc) + ' (' + ', '.join(sorted({x['plate'] for x in ALL if x['code'] in nc})[:6]) + ')')
    for n in NOTES:
        if n[1] == 'нет шапки': warns.append(f'ячейка {n[2]} ({n[0]}) без шапки: {n[3]}')
    if errs: return None, status, warns, errs
    # ---- build_current.py работает в своей рабочей папке (res.pkl, resources/, out/)
    wd = tempfile.mkdtemp(prefix='svod_')
    try:
        shutil.copytree(os.path.join(SV, 'resources'), os.path.join(wd, 'resources')); os.makedirs(os.path.join(wd, 'out'))
        pickle.dump((ALL, EXC, NOTES), open(os.path.join(wd, 'res.pkl'), 'wb'))
        r = subprocess.run([sys.executable, '-I', os.path.join(SV, 'build_current.py'), ds], cwd=wd, capture_output=True, text=True)
        if r.returncode: return None, status, warns, ['сбой build_current: ' + r.stderr[-600:]]
        raw = os.path.join(wd, 'out', f'Свод_Диспетчер_{ds}.xlsx')
        rc = os.path.join(wd, 'rc'); os.makedirs(rc)
        r = subprocess.run(['soffice', '--headless', '--convert-to', 'xlsx', '--outdir', rc, raw], capture_output=True, text=True, timeout=300)
        res = os.path.join(rc, os.path.basename(raw))
        if not os.path.exists(res): return None, status, warns, ['LibreOffice не пересчитал свод: ' + (r.stderr or r.stdout)[-300:]]
        os.makedirs(out_dir, exist_ok=True)
        dst = os.path.join(out_dir, 'svod.xlsx'); shutil.copy(res, dst)
    finally:
        shutil.rmtree(wd, ignore_errors=True)
    # ---- контроль: в пересчитанном своде ошибки формул = 0, итог = сумма «Исходных данных»
    wb = openpyxl.load_workbook(dst, data_only=True); bad = 0
    for ws in wb:
        for row in ws.iter_rows(values_only=True):
            bad += sum(1 for c in row if isinstance(c, str) and c.startswith('#'))
    tot = sum(int(r[8] or 0) + int(r[9] or 0) for r in wb['Исходные данные'].iter_rows(min_row=2, values_only=True) if r[4])
    if tot != sum(x['day'] + x['night'] for x in ALL) + 6: errs.append('в «Исходных данных» не столько строк, сколько разобрано из файлов')
    ws = wb['Выход ТС']; fact = None
    for row in ws.iter_rows(min_row=3, max_col=5, values_only=True):
        if row[1] == 'ИТОГО': fact = row[4]
    if bad: errs.append(f'в своде ошибок формул: {bad}')
    if fact != tot: errs.append(f'итог на листе «Выход ТС» ({fact}) не равен сумме исходных данных ({tot})')
    if errs: return None, status, warns, errs
    return dst, status, warns, []
