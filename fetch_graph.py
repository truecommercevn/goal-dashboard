#!/usr/bin/env python3
"""Download the newest Excel inputs from SharePoint (Microsoft Graph, app-only) into ./inputs.

Env (all required unless noted):
  TENANT_ID, CLIENT_ID, CLIENT_SECRET
  GOALS_SITE_URL   e.g. https://truecommerce.sharepoint.com/sites/PeopleInsightsHub
  GOALS_FOLDER     default "Goal Dashboard"        (folder inside the Documents library)
  EMP_SITE_URL     e.g. https://truecommerce.sharepoint.com/sites/VNPeopleTeam
  EMP_FOLDER       default "General/Employee Information"
  EMP_FILE         default "Danh sach nhan vien - updated 09.06.2025.xlsx"  (one fixed file, updated in place)
  LIBRARY          default "Documents"
Also downloads "Dashboard Access.xlsx" (Dept Head names + emails) from the goals folder.
Writes inputs/GoalsExport_*.xlsx, inputs/employees.xlsx, inputs/access.xlsx, inputs/names.json and prints "changed=true|false"
(compared with the fingerprint stored in state/fingerprint.txt).
"""
import os, re, sys, json, urllib.parse
import requests

G = 'https://graph.microsoft.com/v1.0'
env = lambda k, d=None: os.environ.get(k, d) or (sys.exit(f'missing env {k}') if d is None else d)


def token():
    r = requests.post(f"https://login.microsoftonline.com/{env('TENANT_ID')}/oauth2/v2.0/token", data={
        'client_id': env('CLIENT_ID'), 'client_secret': env('CLIENT_SECRET'),
        'scope': 'https://graph.microsoft.com/.default', 'grant_type': 'client_credentials'}, timeout=60)
    r.raise_for_status(); return r.json()['access_token']


def get(h, url, **kw):
    r = requests.get(url, headers=h, timeout=120, **kw); r.raise_for_status(); return r


def drive_id(h, site_url, library):
    u = urllib.parse.urlparse(site_url)
    site = get(h, f"{G}/sites/{u.netloc}:{u.path}").json()['id']
    for d in get(h, f"{G}/sites/{site}/drives").json()['value']:
        if d['name'].lower() == library.lower(): return d['id']
    sys.exit(f'library {library!r} not found on {site_url}')


def children(h, drive, folder):
    url = f"{G}/drives/{drive}/root:/{urllib.parse.quote(folder.strip('/'))}:/children?$top=200"
    out = []
    while url:
        j = get(h, url).json(); out += j['value']; url = j.get('@odata.nextLink')
    return [x for x in out if 'file' in x]


def download(h, drive, item, dest):
    r = get(h, f"{G}/drives/{drive}/items/{item['id']}/content", stream=True)
    with open(dest, 'wb') as f:
        for c in r.iter_content(1 << 20): f.write(c)


def main():
    h = {'Authorization': 'Bearer ' + token()}
    lib = env('LIBRARY', 'Documents')
    gd = drive_id(h, env('GOALS_SITE_URL'), lib)
    cands = []
    for x in children(h, gd, env('GOALS_FOLDER', 'Goal Dashboard')):
        m = re.fullmatch(r'GoalsExport_(\d{4}-\d{2}-\d{2})\.xlsx', x['name'], re.I)
        if m: cands.append((m[1], x))
    if not cands: sys.exit('no GoalsExport_YYYY-MM-DD.xlsx found in the goals folder')
    day, goals = max(cands, key=lambda t: t[0])           # newest date IN THE FILE NAME (not modified time)
    ed = drive_id(h, env('EMP_SITE_URL'), lib)
    want = env('EMP_FILE', 'Danh sach nhan vien - updated 09.06.2025.xlsx')
    emp = next((x for x in children(h, ed, env('EMP_FOLDER', 'General/Employee Information')) if x['name'] == want), None)
    if not emp: sys.exit(f'employee file {want!r} not found')

    acc = next((x for x in children(h, gd, env('GOALS_FOLDER', 'Goal Dashboard')) if x['name'].lower() == 'dashboard access.xlsx'), None)
    if not acc: sys.exit('Dashboard Access.xlsx not found in the goals folder')
    fp = f"{goals['name']}|{goals.get('eTag')}|{emp['name']}|{emp.get('eTag')}|{acc.get('eTag')}"
    old = open('state/fingerprint.txt').read().strip() if os.path.exists('state/fingerprint.txt') else ''
    changed = fp != old or os.environ.get('FORCE') == 'true'
    print('goals file :', goals['name'], '| employee file:', emp['name'], '| changed' if changed else '| unchanged', file=sys.stderr)
    os.makedirs('inputs', exist_ok=True); os.makedirs('state', exist_ok=True)
    if changed:
        download(h, gd, goals, f"inputs/{goals['name']}")
        download(h, ed, emp, 'inputs/employees.xlsx')
        download(h, gd, acc, 'inputs/access.xlsx')
        open('state/fingerprint.txt', 'w').write(fp)
        json.dump({'goals': goals['name']}, open('inputs/names.json', 'w'))
    print('changed=' + ('true' if changed else 'false'))


main()
