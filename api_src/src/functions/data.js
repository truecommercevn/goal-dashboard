// GET /api/data  -> returns ONLY the dashboard data the signed-in person is allowed to see.
// The browser sends its Microsoft (Entra ID) sign-in token in X-User-Token. We verify the signature, tenant and
// audience here, take the email from the token, look it up in access.json (built hourly from "Dashboard Access.xlsx")
// and return that person's slice only. The page itself contains no data.
const { app } = require('@azure/functions');
const { createRemoteJWKSet, jwtVerify } = require('jose');
const fs = require('fs'); const path = require('path');
const root = path.join(__dirname, '..', '..');
const cfg = JSON.parse(fs.readFileSync(path.join(root, 'config.json'), 'utf8'));
const jwks = createRemoteJWKSet(new URL(cfg.jwksUrl || `https://login.microsoftonline.com/${cfg.tenantId}/discovery/v2.0/keys`));

async function emailFrom(token) {
  const { payload } = await jwtVerify(token, jwks, {
    issuer: cfg.issuer || `https://login.microsoftonline.com/${cfg.tenantId}/v2.0`,
    audience: cfg.clientId,
    algorithms: ['RS256'],
  });
  if (payload.tid !== cfg.tenantId) throw new Error('wrong tenant');
  return String(payload.preferred_username || payload.email || payload.upn || '').trim().toLowerCase();
}

app.http('data', {
  methods: ['GET'], authLevel: 'anonymous', route: 'data',
  handler: async (request) => {
    const token = request.headers.get('x-user-token');
    if (!token) return { status: 401, body: 'not signed in' };
    let email = '';
    try { email = await emailFrom(token); } catch (e) { return { status: 401, body: 'invalid sign-in' }; }
    if (!email) return { status: 401, body: 'no email in token' };
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
