import openpyxl, re, pickle
from openpyxl.styles.colors import COLOR_INDEX
D="input/"   # файлы дня кладутся в папку input/ под именами D3.xlsx, MS.xlsx, D7.xlsx, SVOD.xlsx, BUS.xlsx, K8.xlsx, SB.xlsx
FILES=[
 dict(key='D3', name='Диспетчер №3 (15.10)', path=D+"D3.xlsx", sheet='15.10.', rows=(3,40), cols=(1,13), mode='day'),
 dict(key='MS', name='Мобил Сервис (60-…)', path=D+"MS.xlsx", sheet='15', rows=(3,40), cols=(2,19), mode='day'),
 dict(key='D7', name='Диспетчер №7 (ПКРС/ТМС/УТТ, 15.10)', path=D+"D7.xlsx", sheet='15.', rows=(2,60), cols=(1,10), mode='day', hrow=1),
 dict(key='SVOD', name='Сводка от 15.10.2026 (27-…, 28-…)', path=D+"SVOD.xlsx", sheet='15,10', rows=(3,25), cols=(1,15), mode='day'),
 dict(key='BUS', name='Автобусы / АЦН (15.10)', path=D+"BUS.xlsx", sheet='15', rows=(3,25), cols=(1,8), mode='day'),
 dict(key='K8', name='К-8 (месячный файл, лист 8)', path=D+"K8.xlsx", sheet='15', rows=(2,29), cols=(1,11), mode='day', hrow=1),
 dict(key='SB', name='Ақпар 01-31.10 (С13…)', path=D+"SB.xlsx", sheet='15', rows=(3,40), cols=(1,11), mode='day'),
]
def rgb(c):
    if c is None: return None
    try:
        if c.type=='rgb': return c.rgb[-6:].upper() if isinstance(c.rgb,str) else None
        if c.type=='indexed':
            v=COLOR_INDEX[c.indexed] if c.indexed<len(COLOR_INDEX) else None
            return v[-6:].upper() if v else None
    except Exception: return None
    return None
RED={'FF0000'}; GREEN={'00B050','008000','009900','006600','339966'}; MAROON={'660033'}
def fontkind(cell):
    c=cell.font.color
    r=rgb(c)
    if r in RED: return 'red'
    if r in GREEN: return 'green'
    if r in MAROON: return 'maroon'
    return None
def is_gray(cell):
    f=cell.fill
    if not f or f.fill_type!='solid': return False
    c=f.fgColor
    if c.type=='theme': return c.theme==0 and (c.tint or 0)<-0.05
    r=rgb(c)
    return r in ('C0C0C0','969696','808080','D9D9D9','BFBFBF','A6A6A6','7F7F7F','999999','CCCCCC')
def fillhex(cell):
    f=cell.fill
    if not f or f.fill_type!='solid': return ''
    c=f.fgColor
    if c.type=='theme': return f"t{c.theme}{c.tint:+.1f}"
    return rgb(c) or ''
PLATE=re.compile(r'^\s*([СCсc]?)\s*(?:(60)\s*-\s*(?=\d{2}\s*-))?(\d{2}(?:[.,]\d)?)\s*[-\s]\s*(?=[A-Za-zА-Яа-я/]*\d)')
NUMLESS=re.compile(r'^\s*[СCсc]?\s*\d{2}\s*[-]?\s*(д\s*/\s*н|д|н)?\s*$', re.I)
IGN=re.compile(r'(выход|ремонт|резерв|списан|^рем$|^рез$|^спис$|жөндеу|общи|на линии|день/ночь|в ожидании|тех простой|префикс|уақыты|оқу|есептен|желіге|шықты|план|факт|разница|причина|всего)', re.I)
FINAL={'ПЧ-Семсер':'ПЧ','ЦИТС.':'ЦИТС','Газ.цех':'ГЦ','РСУ':'РСЦ','Экология':'ГООС','оппоIR':'IR'}
# объединение отделов (правила 23.10): итоговое название отдела в своде
MERGE={'УПТОК':'БМТС','ГУ':'ГЦ','БПО2':'ЦПКРС-2','РМУ2':'Гараж-2','ПЧ':'ОБиОТ','ПЧСЕМСЕР':'ОБиОТ','ЧАКУРМЕД':'ОБиОТ'}
def merge_dept(d):
    return MERGE.get(re.sub(r'[\s\-.]','',str(d)).upper(),d)
def norm_dept(raw, gar2=False):
    d=_norm_dept(raw,gar2)
    return merge_dept(FINAL.get(d,d))
def _norm_dept(raw, gar2=False):
    s=re.sub(r'\s+',' ',str(raw)).strip()
    U0=s.upper()
    if 'ПКРС' in U0 and 'УТТ' in U0: return 'ЦПКРС-2'
    if 'ПКРС' in U0 and 'ТМС' in U0: return 'ЦПКРС-1'
    if 'ДРИСУ' in U0.replace(' ',''): return 'ДРиСУ УТТ'
    if U0=='ЦТТ': return 'ЦТТ'
    if U0.startswith('СМУ'): return 'СМУ УТТ'
    if U0.replace(' ','') in ('ГАЗ.УЧАСТОК','ГАЗУЧАСТОК'): return 'ГЦ'
    u=s.upper().replace(' ','').replace('№','').replace('.','')
    u=u.replace('ЦДН-','ЦДН').replace('ЦДН','ЦДН')
    m=re.match(r'^ЦДН(\d)$',u)
    if m: return 'ЦДН-'+m.group(1)
    if u in ('ЭЛЦЕХ','ЭЛЕКТРОЦЕХ','ЭЦ'): return 'ЭЦ'
    if u in ('ППД','ЦППД'): return 'ЦППД'
    if u in ('БУР','БУРЕНИЕ'): return 'Бурение'
    if u in ('СМЦ','ПРЦЭО'): return 'ПРЦЭО'
    if u in ('ЭКОЛОГИЯ','ЭКОЛ','ЭКОЛОГ'): return 'Экология'
    if u in ('ТМС','ПКРС/ТМС','ДБИПКРС(КРС-1)'): return 'ЦПКРС-1'
    if u=='ПКРС': return 'ЦПКРС-2' if gar2 else 'ЦПКРС-1'
    if u in ('ЦПКРС-1',): return 'ЦПКРС-1'
    if 'УТТ' in u and 'ДТТ' not in u: return 'ЦПКРС-2'
    if u=='КРС-2': return 'ЦПКРС-2'
    if u in ('ОБИПКРС','ОБПКРС'): return 'ОБиПКРС'
    if u in ('ОТДКАДР','ОТКАДР','ОДТЕЛКАДР'): return 'Отделкадр'
    if u in ('ЦТВС(РЭО)','ЦТВС'): return 'ЦТВС'
    if u=='ЦИТСДИСПЕТЧЕР': return 'ЦИТС.'
    if u=='ПЧ': return 'ПЧ-Семсер'
    if u=='ОКС': return 'ОКС'
    if u in ('ЦПС','ЦППН'): return 'ЦППН'
    if u=='ЦИТС': return 'ЦИТС'
    if u in ('OKC',): return 'ОКС'
    if u=='ПКРС-2': return 'ЦПКРС-2'
    if u=='ГАРАЖ-2': return 'Гараж-2'
    if u in ('ДТТ','ГАРАЖ-1'): return 'Гараж-1'
    if u in ('ИГУ','ИГУОТР','ОГИР'): return 'ОГиР'
    if u.startswith('ЦПП-МПГУ') or u.startswith('МПГУ') or u=='ЦПП': return 'ЦПП'
    return s
def parse(cfg):
    wb=openpyxl.load_workbook(cfg['path']); ws=wb[cfg['sheet']]
    gar2=cfg.get('garazh2',False) or ('ГАРАЖ-2' in str(ws['L1'].value or '').upper().replace(' ',''))
    r0,r1=cfg['rows']; c0,c1=cfg['cols']
    merged={}
    for mr in ws.merged_cells.ranges:
        for c in range(mr.min_col,mr.max_col+1):
            if mr.min_row<=cfg.get('hrow',2)<=mr.max_row: merged[c]=mr.min_col
    # night block start row
    night_start=None
    if cfg['mode']=='block':
        for r in range(1,60):
            v=ws.cell(r,1).value
            if v and 'очь' in str(v) and 'выход' in str(v): night_start=r;break
    out=[];excl=[];notes=[]
    YEL=('FFFF00','FFFF66','FFFF99','FFFFCC')
    def _isoqu(t): t=str(t).lower(); return 'оқу' in t or 'оку' in t
    def _lg(c): return ws.cell(c.row, c.column-1) if c.column>1 else None
    # легенда «Оқу»: сама подпись жёлтая, либо жёлтая ячейка-образец слева от подписи (диспетчер 7)
    def _up(c): return ws.cell(c.row-1, c.column) if c.row>1 else None
    oqu=any(_isoqu(c.value) and (fillhex(c) in YEL or (_lg(c) is not None and fillhex(_lg(c)) in YEL) or (_up(c) is not None and fillhex(_up(c)) in YEL)) for row in ws.iter_rows() for c in row if c.value)   # + жёлтый образец НАД подписью (объединённый файл Гаража-2)
    def _cmt(c):   # примечание к ячейке «сабақ»/«оқу» (диспетчер 2)
        t=(c.comment.text if c.comment else '').lower(); return 'сабақ' in t or 'сабак' in t or 'оқу' in t or 'оку' in t
    # header row 2
    hdr={}
    for c in range(c0,c1+1):
        src=merged.get(c,c)
        v=ws.cell(cfg.get('hrow',2),src).value
        hdr[c]=v
    cur={}   # column -> (dept raw)
    for c in range(c0,c1+1):
        v=hdr[c]
        if v is None: cur[c]=None
        else:
            s=str(v)
            if cfg['key']=='MS' and c==2: s='ТБ'
            cur[c]=('IGN' if IGN.search(s) else s)
    rawhdr=dict(cur)
    for r in range(r0,r1+1):
        if night_start and r==night_start: 
            # reset to row below headers
            pass
        for c in range(c0,c1+1):
            cell=ws.cell(r,c); v=cell.value
            if v is None: continue
            s=str(v)
            if not s.strip(): continue
            if night_start and r==night_start: 
                continue
            # night block header row: row after night_start
            isnight_blk = bool(night_start and r>night_start)
            m=PLATE.match(s)
            if not m:
                t=s.strip()
                if NUMLESS.match(s):
                    excl.append((cell.coordinate,s.strip(),cur.get(c),'нет номера (только тип)'))
                    continue
                if not re.search(r'[A-Za-zА-Яа-яӘәҒғҚқҢңӨөҰұҮүІіҺһ]',t): continue  # pure numbers/counts like 1, 5+2
                if re.fullmatch(r'\d+\s*\+\s*\d+',t): continue
                if t.startswith('Для ISR'): continue
                if t.lower()=='хитачи': notes.append((cell.coordinate,t,cur.get(c),'пометка типа техники, отдел не меняется')); continue
                if re.search(r'(көлік|жалпы|менеджер|нач участка|жоспар|нақты|айырма|себеб|мекеме|түр|санi|саны)',t,re.I): continue
                # label: new dept for this column
                cur[c]='IGN' if IGN.search(t) else t
                if r>=3 and isnight_blk: pass
                continue
            d=cur.get(c)
            if d is None or d=='IGN':
                if d=='IGN': excl.append((cell.coordinate,s.strip(),rawhdr.get(c),'колонка «выход/ремонт/резерв/списание»'))
                else: notes.append(('нет шапки',cell.coordinate,s.strip()))
                continue
            if (oqu or _cmt(cell)) and fillhex(cell) in YEL: excl.append((cell.coordinate,s.strip(),d,'жёлтая заливка при «Оқу» на листе (обучение)')); continue
            fk=fontkind(cell)
            if fk=='red': excl.append((cell.coordinate,s.strip(),d,'красный шрифт (ремонт)')); continue
            if fk=='green' and not (cfg['key']=='K8' and c==5 and r<=13): excl.append((cell.coordinate,s.strip(),d,'зелёный шрифт (резерв)')); continue
            if fk=='maroon': excl.append((cell.coordinate,s.strip(),d,'списание')); continue
            low=s.strip().lower()
            parts=re.split(r'\s{2,}',s.strip()); main=parts[0]
            cpref,mob,code=m.group(1),m.group(2),m.group(3).replace(',','.')
            if code in ('12','13') and cpref and not mob: code+='.1'
            if code=='57' and len(parts)>1 and parts[1].strip().startswith('25'): code='57.1'
            tail=s.strip()
            both=bool(re.search(r'(д\s*/\s*н|d\s*/\s*n)\s*$',tail,re.I))
            nightm=False; daym=False
            if not both:
                if re.search(r'(?:[нН]|(?<=[\d\s])[nN])\s*$',tail): nightm=True
                elif re.search(r'(?:д|(?<=[\d\s])[dD])\s*$',tail): daym=True
            gray=is_gray(cell)
            day=night=0; note=''
            if both: day=night=1
            elif nightm: night=1
            elif daym: day=1
            elif gray: night=1; note='серая заливка без буквы → ночь'
            elif isnight_blk: night=1; note='ночной блок'
            else: day=1
            dept=norm_dept(d,gar2)
            if cfg.get('alias'): dept=cfg['alias'].get(re.sub(r'[\s.№]','',str(d)).upper(),dept)
            if c in cfg.get('dept_over',{}) and d==hdr[c]: dept=cfg['dept_over'][c]
            dept=merge_dept(dept)
            if fk is None and fillhex(cell) in ('FF66FF',): note=(note+'; ' if note else '')+'розовая заливка (одна машина в двух отделах)'
            if re.search(r'[Хх]\s*$',tail): pass
            # plate id
            plate=re.sub(r'\s+','',main).upper()
            plate=re.sub(r'(Д/Н|Н|Д)$','',plate) if not re.search(r'\d$',plate) else plate
            out.append(dict(cell=cell.coordinate,src=cfg['key'],srcname=cfg['name'],hdr=str(d).strip(),dept=dept,plate=main.strip(),code=code,day=day,night=night,note=note,raw=s.strip(),pid=plate))
    return out,excl,notes
if __name__=='__main__':
    ALL=[];EXC=[];NOTES=[]
    for cfg in FILES:
        o,e,n=parse(cfg)
        print(cfg['key'],len(o),'day',sum(x['day'] for x in o),'night',sum(x['night'] for x in o),'excl',len(e),'notes',len(n))
        ALL+=o;EXC+=[(cfg['key'],)+x for x in e];NOTES+=[(cfg['key'],)+x for x in n]
    pickle.dump((ALL,EXC,NOTES),open('res.pkl','wb'))
    print(sum(x['day'] for x in ALL),sum(x['night'] for x in ALL))
    for n in NOTES: print(n)
