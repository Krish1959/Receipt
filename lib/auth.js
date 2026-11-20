'use strict';

// A single shared password for the whole app.
//
// The password is checked on the SERVER and never appears in any page sent to
// the browser. A correct answer sets a signed, http-only cookie that the
// browser cannot read from script; every later request is judged on that
// cookie alone.
//
// What this is: a gate that keeps the practice's forms off the open internet,
// so a stray link does not let a passer-by file invoices.
// What it is not: per-person accounts. Everyone shares one password, so there
// is no record of WHO filed something, and changing it signs everybody out.

const crypto = require('crypto');

const PASSWORD = process.env.APP_PASSWORD || 'Raj&Raj';
// Set APP_SECRET in Render to keep people signed in across restarts.
const SECRET = process.env.APP_SECRET || crypto.randomBytes(32).toString('hex');
const COOKIE = 'rg_session';
const MAX_AGE_MS = 12 * 60 * 60 * 1000;          // a working day

// Paths reachable without the password.
const OPEN_PATHS = new Set(['/login.html', '/api/login', '/api/logout', '/healthz', '/favicon.ico']);

const b64 = s => Buffer.from(s).toString('base64url');
const sign = data => crypto.createHmac('sha256', SECRET).update(data).digest('base64url');

function makeToken() {
  const payload = b64(JSON.stringify({ exp: Date.now() + MAX_AGE_MS }));
  return `${payload}.${sign(payload)}`;
}

function validToken(token) {
  const [payload, mac] = String(token || '').split('.');
  if (!payload || !mac) return false;
  const expected = sign(payload);
  // Constant-time compare, so a wrong signature cannot be found a byte at a time.
  if (mac.length !== expected.length) return false;
  if (!crypto.timingSafeEqual(Buffer.from(mac), Buffer.from(expected))) return false;
  try {
    return JSON.parse(Buffer.from(payload, 'base64url').toString()).exp > Date.now();
  } catch (_) { return false; }
}

function passwordMatches(given) {
  const a = Buffer.from(String(given || ''));
  const b = Buffer.from(PASSWORD);
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}

const readCookie = (req, name) => {
  const found = String(req.headers.cookie || '')
    .split(';').map(s => s.trim()).find(s => s.startsWith(name + '='));
  return found ? decodeURIComponent(found.slice(name.length + 1)) : '';
};

function setCookie(res, value, maxAgeMs) {
  const bits = [
    `${COOKIE}=${value}`, 'Path=/', 'HttpOnly', 'SameSite=Lax',
    `Max-Age=${Math.floor(maxAgeMs / 1000)}`
  ];
  if (process.env.NODE_ENV === 'production' || process.env.RENDER) bits.push('Secure');
  res.setHeader('Set-Cookie', bits.join('; '));
}

// Blocks everything but the open paths. A page request is sent to the login
// screen; an API request gets a 401 the page can act on.
function requireLogin(req, res, next) {
  const path = req.path.replace(/\/+$/, '') || '/';
  if (OPEN_PATHS.has(req.path) || OPEN_PATHS.has(path)) return next();
  if (validToken(readCookie(req, COOKIE))) return next();
  if (req.path.startsWith('/api/')) {
    return res.status(401).json({ ok: false, error: 'Please sign in again.', login: true });
  }
  res.redirect(`/login.html?next=${encodeURIComponent(req.originalUrl)}`);
}

function mount(app) {
  app.post('/api/login', (req, res) => {
    if (!passwordMatches((req.body || {}).password)) {
      return res.status(401).json({ ok: false, error: 'Wrong password.' });
    }
    setCookie(res, makeToken(), MAX_AGE_MS);
    res.json({ ok: true });
  });

  app.post('/api/logout', (_req, res) => {
    setCookie(res, '', 0);
    res.json({ ok: true });
  });

  app.use(requireLogin);
}

module.exports = { mount, requireLogin, makeToken, validToken, passwordMatches, COOKIE };
