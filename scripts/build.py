"""Робот: берёт файлы из input/, собирает data.json и drivers.json, кладёт в архив, чистит старое.
Запуск из корня репозитория: python3 scripts/build.py"""
import sys, os, re, json, glob, shutil, datetime
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import openpyxl
from svod2json import convert
import build_drivers, make_mobil, repair

KEEP_DAYS = 365
INP = os.path.join(ROOT, 'input'); ARC = os.path.join(ROOT, 'archive')
summary = []
def say(s): print(s); summary.append(s)

def classify(path):
    """свод / экспорт / вахта / ремонт автоколонны / None - по содержимому, а не по имени файла"""
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        names = wb.sheetnames
        if 'Исходные данные' in names and 'Выход ТС' in names: return 'svod'
        ws = wb[names[0]]
        first = next(ws.iter_rows(max_row=1, values_only=True), ())
        if len(first) > 16 and 'Государственный' in str(first[16] or ''): return 'export'
        a1 = next(ws.iter_rows(max_row=1, values_only=True), (None,))[0]
        if 'Мобил Сервис' in names or 'Вахта' in str(a1 or ''): return 'roster'
        if repair.is_repair_file(path): return 'repair'
    except Exception as e:
        say(f'⚠ не удалось прочитать {os.path.basename(path)}: {e}')
    return None

def load(p, default=None):
    return json.load(open(p, encoding='utf-8')) if os.path.exists(p) else default
def dump(obj, p):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    json.dump(obj, open(p, 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))

def main():
    files = [f for f in glob.glob(os.path.join(INP, '*')) if f.lower().endswith(('.xlsx', '.xlsm', '.xls')) and not os.path.basename(f).startswith('~$')]
    kinds, conv, tmps = {}, {}, []
    for f in files:
        try:
            xp, tmp = repair.to_xlsx(f)          # старые .xls -> .xlsx
        except Exception as e:
            say(f'⚠ не удалось открыть {os.path.basename(f)}: {e}'); continue
        conv[f] = xp
        if tmp: tmps.append(tmp)
        k = classify(xp)
        if k: kinds.setdefault(k, []).append(f)
        else: say(f'⚠ файл {os.path.basename(f)} не распознан (не свод, не экспорт, не список вахты, не отчёт автоколонны)')
    # --- Мобил: новый список вахты
    mob_path = os.path.join(ROOT, 'data', 'mobil.json')
    mobil = load(mob_path, {})
    for f in kinds.get('roster', []):
        new = make_mobil.read_roster(conv[f]); mobil.update(new); dump(mobil, mob_path)
        say(f'Список вахты Mobil обновлён: {len(new)} машин')
    svods, exports = kinds.get('svod', []), kinds.get('export', [])
    if svods or exports:
        if len(svods) != 1 or len(exports) != 1:
            say(f'❌ нужен ровно один свод и один экспорт, найдено: свод {len(svods)}, экспорт {len(exports)}')
            return 1
        data = convert(conv[svods[0]])
        date = data['date']
        say(f'Свод на {date}: план {data["totals"]["plan"]}, факт {data["totals"]["fact"]}, машин {len(data["veh"])}')
        contacts = load(os.path.join(ROOT, 'data', 'contacts.json'))
        if not contacts: say('❌ нет data/contacts.json'); return 1
        drv, rep, miss = build_drivers.build(data, conv[exports[0]], contacts, mobil)
        for l in rep: say(l)
        if miss: say(f'Без телефона ({len(miss)}): ' + '; '.join(miss))
        # архив дня
        day = os.path.join(ARC, date)
        dump(data, os.path.join(day, 'data.json')); dump(drv, os.path.join(day, 'drivers.json'))
        # индекс дат + чистка
        dates = sorted((d for d in os.listdir(ARC) if re.fullmatch(r'\d{4}-\d{2}-\d{2}', d)), reverse=True) if os.path.isdir(ARC) else []
        limit = (datetime.date.today() - datetime.timedelta(days=KEEP_DAYS)).isoformat()
        for d in [d for d in dates if d < limit]:
            shutil.rmtree(os.path.join(ARC, d)); say(f'Удалён старый день {d}')
        dates = [d for d in dates if d >= limit]
        dump({'dates': dates, 'latest': dates[0]}, os.path.join(ARC, 'index.json'))
        # «сегодня» для сайта: только если это самый свежий день
        if dates[0] == date:
            dump(data, os.path.join(ROOT, 'data.json')); dump(drv, os.path.join(ROOT, 'drivers.json'))
            say(f'Сайт обновлён на {date}')
        else:
            say(f'Загружен более старый день {date}; на сайте остаётся {dates[0]}')
        say(f'В архиве дней: {len(dates)}')
    # --- ремонт по автоколоннам
    if kinds.get('repair'):
        tz = datetime.timezone(datetime.timedelta(hours=5))
        ref = datetime.date.fromisoformat(data['date']) if (svods or exports) else datetime.datetime.now(tz).date()
        day = os.path.join(ARC, ref.isoformat()); rp = os.path.join(day, 'repair.json')
        cur = load(rp, {'aks': []})
        aks = {a['n']: a for a in cur['aks'] if a.get('n')}
        warns = []
        for f in kinds['repair']:
            try:
                res, w = repair.read_ak(conv[f], ref)
            except Exception as e:
                say(f'⚠ {os.path.basename(f)}: ошибка чтения отчёта автоколонны: {e}'); continue
            warns += w
            if not res or not res.get('n'):
                say(f'⚠ {os.path.basename(f)}: не определён номер автоколонны'); continue
            aks[res['n']] = res
            say(f'АК №{res["n"]}: машин в ремонте {len(res["vehs"])} (лист {res["sheet"]})')
        for w in warns: say('⚠ ' + w)
        out = {'date': ref.isoformat(), 'aks': [aks[k] for k in sorted(aks)]}
        dump(out, rp)
        say(f'Автоколонн в отчёте на {ref.isoformat()}: {len(aks)} из 8' + ('' if len(aks) == 8 else ' (не хватает: ' + ', '.join(str(i) for i in range(1, 9) if i not in aks) + ')'))
        idx = load(os.path.join(ARC, 'index.json'), {'dates': []})
        if not idx['dates'] or ref.isoformat() >= idx['dates'][0]:
            dump(out, os.path.join(ROOT, 'repair.json'))
    elif not (kinds.get('roster') or svods or exports):
        say('В input/ нет подходящих файлов - ничего не делаю'); return 0
    for t in tmps: shutil.rmtree(t, ignore_errors=True)
    for ks in kinds.values():          # обработанные файлы убираем из папки
        for f in ks: os.remove(f)
    return 0

if __name__ == '__main__':
    rc = main()
    out = os.environ.get('GITHUB_STEP_SUMMARY')
    if out: open(out, 'a', encoding='utf-8').write('\n'.join('- ' + s for s in summary) + '\n')
    sys.exit(rc)
