#!/usr/bin/env python3
"""Build the Goal Dashboard site (one folder per Dept Head + master) from the latest Excel exports.

usage: python build.py --goals GoalsExport_2026-11-10.xlsx --emps "Danh sach nhan vien - updated 09.06.2025.xlsx" --out site
"""
import argparse, json, os, re, shutil, sys, datetime
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_data as bd

slug = lambda n: re.sub(r'[^a-z0-9]+', '-', n.lower()).strip('-')


def data_date(goals_path):
    m = re.search(r'(\d{4})-(\d{2})-(\d{2})', os.path.basename(goals_path))
    if not m: return None
    d = datetime.date(int(m[1]), int(m[2]), int(m[3]))
    return d.strftime('%d %b %Y').lstrip('0')


def subset(DATA, ORG, heads):
    keep = [i for i in range(len(DATA['emps'])) if heads is None or ORG[i][0] in heads]
    m = {old: new for new, old in enumerate(keep)}
    d = {'emps': [DATA['emps'][i] for i in keep], 'goals': [[m[g[0]]] + g[1:] for g in DATA['goals'] if g[0] in m]}
    return d, [ORG[i] for i in keep]


def roster_report(DATA, ORG, issues, prev):
    cur = {str(e[0]): dict(name=e[1], head=ORG[i][0], sup=ORG[i][1], gone=ORG[i][4]) for i, e in enumerate(DATA['emps'])}
    L = ['# Roster check', '']
    if prev:
        new = [v['name'] + f" (ID {k}, {v['head']})" for k, v in cur.items() if k not in prev]
        left = [v['name'] + f" (ID {k})" for k, v in cur.items() if k in prev and v['gone'] and not prev[k]['gone']]
        back = [v['name'] + f" (ID {k})" for k, v in cur.items() if k in prev and not v['gone'] and prev[k]['gone']]
        moved = [f"{v['name']}: {prev[k]['head']} / {prev[k]['sup']} -> {v['head']} / {v['sup']}"
                 for k, v in cur.items() if k in prev and (v['head'], v['sup']) != (prev[k]['head'], prev[k]['sup'])]
        dropped = [f"{v['name']} (ID {k})" for k, v in prev.items() if k not in cur]
        for title, items in (('New employees (have goals, not in last run)', new), ('Now no longer active', left),
                             ('Active again', back), ('Changed leader / Dept Head', moved),
                             ('No goals any more (were in last run)', dropped)):
            L += [f'## {title} ({len(items)})'] + [f'- {x}' for x in items] + ['']
    else:
        L += ['_No previous run to compare with (first run)._', '']
    L += [f'## Needs a look in the employee list ({len(issues)})'] + [f'- {a}: {b} (ID {c})' for a, b, c in issues]
    L += ['', f'## Totals', f'- Employees with goals: {len(cur)}', f"- No longer active: {sum(v['gone'] for v in cur.values())}",
          f"- Not mapped to a Dept Head (Other Reporting Lines): {sum(v['head'] == bd.OTHER for v in cur.values())}"]
    return '\n'.join(L), cur


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--goals', required=True); ap.add_argument('--emps', required=True)
    ap.add_argument('--access', default=None)
    ap.add_argument('--out', default='site'); ap.add_argument('--api', default='api')
    ap.add_argument('--state', default='state/roster.json'); ap.add_argument('--report', default='report/roster_changes.md')
    ap.add_argument('--frame-ancestors', default='https://*.sharepoint.com')
    a = ap.parse_args()

    cfg, master = bd.load_access(a.access) if a.access else (bd.DEFAULT_HEADS, [])
    HT = {h['name']: h['title'] for h in cfg}
    DATA, ORG, issues, *_ = bd.build(a.goals, a.emps, cfg)
    assert len(DATA['emps']) > 0 and len(DATA['goals']) > 0, 'no Vietnam goals found - wrong file?'
    meta = {'updated': data_date(a.goals)}

    # 1) the page: ONE html, no data inside
    for d in (a.out, a.api):
        if os.path.isdir(d): shutil.rmtree(d)
    os.makedirs(a.out)
    shutil.copy(os.path.join(HERE, 'index.html'), os.path.join(a.out, 'index.html'))   # the page you edit, published as-is

    # 2) the API: one data file per Dept Head + master (only served to the right signed-in person)
    os.makedirs(os.path.join(a.api, 'data'))
    def put(key, heads_set, ht):
        d, o = subset(DATA, ORG, heads_set)
        json.dump({'DATA': d, 'ORG': o, 'HEAD_TITLE': ht, 'META': meta}, open(os.path.join(a.api, 'data', key + '.json'), 'w', encoding='utf-8'), ensure_ascii=False)
        print(key, len(d['emps']), 'employees', len(d['goals']), 'goals')
    put('master', None, HT)
    access = {}
    for h in cfg:
        if not any(o[0] == h['name'] for o in ORG): print('warning: no employees found for', h['name']); continue
        put(slug(h['name']), {h['name']}, {h['name']: h['title']})
        for e in h['emails']: access.setdefault(e, []).append(slug(h['name']))
    for e in master: access.setdefault(e, []).append('master')
    json.dump(access, open(os.path.join(a.api, 'access.json'), 'w'), indent=1)
    for fn in ('host.json', 'package.json'): shutil.copy(os.path.join(HERE, 'api_src', fn), os.path.join(a.api, fn))
    os.makedirs(os.path.join(a.api, 'src', 'functions'))
    shutil.copy(os.path.join(HERE, 'api_src', 'src', 'functions', 'data.js'), os.path.join(a.api, 'src', 'functions', 'data.js'))
    auth = json.load(open('inputs/auth.json')) if os.path.exists('inputs/auth.json') else {
        'clientId': os.environ.get('CLIENT_ID', ''), 'tenantId': os.environ.get('TENANT_ID', '')}
    if not auth.get('clientId') or not auth.get('tenantId'): raise SystemExit('missing clientId/tenantId (inputs/auth.json)')
    json.dump(auth, open(os.path.join(a.api, 'config.json'), 'w'))                   # used by the API to verify tokens
    json.dump(auth, open(os.path.join(a.out, 'auth-config.json'), 'w'))              # public ids, used by the page to sign in
    print('access:', len(access), 'sign-in emails;', [h['name'] for h in cfg if not h['emails']], 'have no email yet')

    swa = {
        'platform': {'apiRuntime': 'node:20'},   # the API uses the Azure Functions v4 model (needs Node 18+)
        'globalHeaders': {'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff',
                          'Content-Security-Policy': f"frame-ancestors 'self' {a.frame_ancestors}"},
    }
    json.dump(swa, open(os.path.join(a.out, 'staticwebapp.config.json'), 'w'), indent=2)

    prev = json.load(open(a.state, encoding='utf-8')) if os.path.exists(a.state) else None
    rep, cur = roster_report(DATA, ORG, issues, prev)
    os.makedirs(os.path.dirname(a.report), exist_ok=True); open(a.report, 'w', encoding='utf-8').write(rep)
    os.makedirs(os.path.dirname(a.state), exist_ok=True); json.dump(cur, open(a.state, 'w', encoding='utf-8'))
    print(rep)


main()
