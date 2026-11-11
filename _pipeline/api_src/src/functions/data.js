// GET /api/data  -> returns ONLY the dashboard data the signed-in person is allowed to see.
// Identity comes from Azure Static Web Apps (x-ms-client-principal); permissions from access.json,
// both built hourly from "Dashboard Access.xlsx". The page itself contains no data.
const { app } = require('@azure/functions');
const fs = require('fs'); const path = require('path');
const root = path.join(__dirname, '..', '..');

app.http('data', {
  methods: ['GET'], authLevel: 'anonymous', route: 'data',
  handler: async (request) => {
    const hdr = request.headers.get('x-ms-client-principal');
    if (!hdr) return { status: 401, body: 'not signed in' };
    let email = '';
    try { email = String(JSON.parse(Buffer.from(hdr, 'base64').toString('utf8')).userDetails || '').trim().toLowerCase(); } catch (e) {}
    if (!email) return { status: 401, body: 'not signed in' };
    let key = null;
    try {
      const keys = JSON.parse(fs.readFileSync(path.join(root, 'access.json'), 'utf8'))[email] || [];
      key = keys.includes('master') ? 'master' : keys[0];
    } catch (e) {}
    if (!key || !/^[a-z0-9-]+$/.test(key)) return { status: 403, body: 'no access' };
    try {
      const body = fs.readFileSync(path.join(root, 'data', key + '.json'), 'utf8');
      return { status: 200, headers: { 'content-type': 'application/json', 'cache-control': 'no-store' }, body };
    } catch (e) { return { status: 404, body: 'no data' }; }
  }
});
