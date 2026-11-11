"""Goals export + employee list -> DATA / ORG / HEAD_TITLE used by the dashboard template."""
import re, json
from collections import defaultdict
import openpyxl

OTHER = 'Other Reporting Lines'
# Dept Heads now come from "Dashboard Access.xlsx" (see load_access); these are only the defaults for local tests
DEFAULT_HEADS = [
    dict(name='Amy Nguyen', title='Director, Services & Support', emails=[], root='Nguyen, Thi Anh Nguyet', above=False),
    dict(name='Toan Nguyen', title='Director, Services', emails=[], root='Nguyen, Thien Toan', above=False),
    dict(name='Ping Guo', title='APAC Operations Director', emails=[], root='Thach, Soc Kha', above=True),
    dict(name='Thanh Thai', title='Manager, Engineering', emails=[], root='Thai, Bao Thanh', above=False),
    dict(name='Anh Nguyen', title='HR Director', emails=[], root='', above=False),
]


def load_access(path):
    """Read Dashboard Access.xlsx -> (heads, master_emails)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    split = lambda v: [x.strip().lower() for x in re.split(r'[;,\n]', str(v or '')) if '@' in x]
    heads = []
    for r in wb['Dept Heads'].iter_rows(min_row=2, values_only=True):
        if not r[0] or not str(r[0]).strip(): continue
        heads.append(dict(name=str(r[0]).strip(), title=str(r[1] or '').strip(), emails=split(r[2]),
                          root=str(r[3] or '').strip(), above=str(r[4] or '').strip().lower() in ('yes', 'y', 'true', '1')))
    master = []
    for r in wb['Master access'].iter_rows(min_row=2, values_only=True): master += split(r[0])
    if not heads: raise SystemExit('Dashboard Access.xlsx: sheet "Dept Heads" is empty')
    return heads, master


NAME_FIX = {'Wang Selina': 'Selina Wang', 'Guo Ping': 'Ping Guo', 'Charlie W Tracy': 'Charlie Tracy'}


def flip(n):
    p = [x.strip() for x in str(n).split(',')]
    return (p[1] + ' ' + p[0]) if len(p) == 2 else str(n).strip()


def load_goals(goals_xlsx):
    wb = openpyxl.load_workbook(goals_xlsx, data_only=True, read_only=True)

    def rows(sheet):
        it = wb[sheet].iter_rows(values_only=True)
        hdr = next(it); ix = {h: i for i, h in enumerate(hdr)}
        return ix, [r for r in it if r[ix['Country']] == 'Vietnam']
    gi, goals = rows('Goals')
    ki, krs = rows('Goals and Key Results')
    kr_by = defaultdict(list)
    for r in krs:
        n = r[ki['Key Result']]
        if isinstance(n, str) and r[ki['KR Progress']] is not None:
            kr_by[int(r[ki['GoalID']])].append([n.strip(), round(float(r[ki['KR Progress']]), 2)])
    return gi, goals, kr_by


def load_roster(emp_xlsx):
    wb = openpyxl.load_workbook(emp_xlsx, data_only=True)
    recs = []

    def add(eng, legal, sup, dept, title, src, uid=None):
        recs.append(dict(uid=str(uid or '').strip().lstrip('0'), eng=flip(eng) if ',' in str(eng) else str(eng).strip(), legal=str(legal).strip(),
                         sup=(sup or '').strip(), dept=(dept or '').strip(), title=title, src=src))
    for r in wb['List of employee'].iter_rows(min_row=3, values_only=True):
        if r[0]: add(r[2], r[1], r[14], r[18], r[12], 'List', r[0])
    for r in wb['List of employee (Toan Nguyen)'].iter_rows(min_row=2, values_only=True):
        if r[0] and r[2]: add(r[2], r[1], r[5], r[3], r[4], 'Toan', r[0])
    for sh in wb.sheetnames:
        if sh.strip() in ('Main sheet', '_PA_Data'):
            for r in wb[sh].iter_rows(min_row=2, values_only=True):
                if r[0] and str(r[0]).strip() and r[2]: add(r[2], r[1], r[5], r[3], r[4], sh.strip(), r[0])
    # status: active if ANY Main-sheet row is Active, or the ID is in the current roster ('List of employee')
    active_main, roster_ids, seen_main = set(), set(), set()
    for r in wb['Main sheet'].iter_rows(min_row=2, values_only=True):
        if r[0] and str(r[0]).strip():
            uid = str(r[0]).strip().lstrip('0'); seen_main.add(uid)
            if str(r[17]).strip().lower() == 'active': active_main.add(uid)
    for r in wb['List of employee'].iter_rows(min_row=3, values_only=True):
        if r[0]: roster_ids.add(str(r[0]).strip().lstrip('0'))
    return recs, active_main | roster_ids, seen_main


def build(goals_xlsx, emp_xlsx, heads_cfg=None):
    heads_cfg = heads_cfg or DEFAULT_HEADS
    HEAD_LEGAL = {h['root']: h['name'] for h in heads_cfg if h['root']}
    ABOVE = {h['name'] for h in heads_cfg if h['above']}
    NAMES = {h['name'] for h in heads_cfg}
    gi, goals, kr_by = load_goals(goals_xlsx)
    recs, active_ids, seen_main = load_roster(emp_xlsx)
    legal2rec = {}
    for r in recs: legal2rec.setdefault(r['legal'], r)

    def find(name, dept, uid=None):
        by = [r for r in recs if uid and r['uid'] == str(uid).lstrip('0')]
        if by: return next((r for r in by if r['src'] == 'List'), by[0])
        num = dept.split('-')[0]
        c = [r for r in recs if r['eng'] == name]
        c1 = [r for r in c if r['dept'].split('-')[0] == num]
        return (c1 or c)[0] if (c1 or c) else None

    def chain(r):
        c = [r]
        while True:
            s = legal2rec.get(c[-1]['sup'])
            if s and s not in c: c.append(s)
            else: return c, c[-1]['sup']

    # employees (first appearance order of owner id)
    emps, idx = [], {}
    for r in goals:
        uid = int(r[gi['Owner EmployeeId']])
        if uid not in idx:
            idx[uid] = len(emps)
            emps.append([uid, f"{r[gi['Owner First Name']]} {r[gi['Owner Last Name']]}", r[gi['Department']]])
    out_goals = []
    for r in goals:
        gid = int(r[gi['GoalID']]); krs = kr_by.get(gid, [])
        p = round(float(r[gi['Progress (%)']] or 0), 2)   # the system's own goal %
        out_goals.append([idx[int(r[gi['Owner EmployeeId']])], gid, r[gi['Goal Name']] or '',
                          r[gi['Goal Progress Status']], p, krs])

    title_by_id = {int(r[gi['Owner EmployeeId']]): (r[gi['Title']] or '') for r in goals}
    # org
    ORG, issues = [], []
    contrib = {int(r[gi['GoalID']]): (r[gi['Owner EmployeeId']], (r[gi['Goal Name']] or '').strip().lower(),
                                      r[gi['Goal Start Date']], r[gi['Goal Due Date']],
                                      (r[gi['Contributor(s)']] or '').strip()) for r in goals}
    for e in emps:
        rec = find(e[1], e[2], e[0])
        if rec and rec['sup'].startswith('#'): rec = None   # broken lookup in the sheet (#N/A)
        if not rec:
            gt = title_by_id.get(e[0], '')
            gone = str(e[0]) not in active_ids
            ORG.append([OTHER, 'No longer active' if gone else 'Not in org file', gt, '', 1 if gone else 0])
            issues.append(('not in employee list' + (' (treated as no longer active)' if gone else ''), e[1], e[0])); continue
        c, ext = chain(rec)
        hd = next((x for x in c if x['legal'] in HEAD_LEGAL), None)
        head = HEAD_LEGAL[hd['legal']] if hd else None
        supr = c[1] if len(c) > 1 else None
        sup = supr['eng'] if supr else (flip(rec['sup']) if rec['sup'] else '')
        sup = NAME_FIX.get(sup, sup)
        st = supr['title'] if supr and supr['title'] and not str(supr['title']).startswith('#') else ''
        title = rec['title'] if rec['title'] and not str(rec['title']).startswith('#') else ''
        if head and hd is rec: sup, st = (head, '') if head in ABOVE else ('(Dept Head)', '')
        if head is None and sup in NAMES: head = sup           # reports directly to a Dept Head who is outside the roster chain
        if rec['eng'] in NAMES and head is None: head, sup, st = rec['eng'], '(Dept Head)', ''
        if head is None:
            head = OTHER
            if not sup: sup = 'Not in org file'
        ORG.append([head, sup or 'Not in org file', title, st, 0 if str(e[0]) in active_ids else 1])
    # contributor families: same owner + same goal name + same dates, every record has a distinct contributor
    def cname(email):
        return ' '.join(w[:1].upper() + w[1:] for w in re.split(r'[._]', email.split('@')[0]) if w)
    fam = defaultdict(list)
    for r in goals:
        fam[(r[gi['Owner EmployeeId']], (r[gi['Goal Name']] or '').strip().lower(),
             r[gi['Goal Start Date']], r[gi['Goal Due Date']])].append(r)
    goalfam, fid = {}, 0
    for v in fam.values():
        cs = [(r[gi['Contributor(s)']] or '').strip() for r in v]
        if len(v) >= 2 and all(cs) and len({cname(c).lower() for c in cs}) == len(cs):
            fid += 1
            for r, c in zip(v, cs): goalfam[int(r[gi['GoalID']])] = (fid, cname(c))
    for g in out_goals:
        if g[1] in goalfam: g.extend(goalfam[g[1]])
    return dict(emps=emps, goals=out_goals), ORG, issues, recs, active_ids, goalfam
