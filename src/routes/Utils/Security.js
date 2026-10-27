const helmet = require('helmet');
const hpp = require('hpp');
const proxyaddr = require('proxy-addr'); // sudah terpasang sebagai dependensi Express
const { rateLimit } = require('express-rate-limit');
const { tooManyRequests, badRequest } = require('./Errors');

// Semua limiter meneruskan error 429 ke errorHandler -> format JSON seragam
const makeLimiter = ({ windowMs, limit, message, skipSuccessfulRequests = false, skip }) =>
  rateLimit({
    windowMs,
    limit,
    standardHeaders: 'draft-7', // header RateLimit-* untuk klien
    legacyHeaders: false,
    skipSuccessfulRequests,
    skip,
    handler: (req, res, next) => next(tooManyRequests(message)),
  });

// 1) Limiter umum: semua request per IP
const globalLimiter = makeLimiter({
  windowMs: 60 * 1000,
  limit: 120,
  message: 'Terlalu banyak permintaan. Tunggu sebentar lalu coba lagi.',
});

// 2) Limiter login: anti brute-force password (hanya request GAGAL yang dihitung)
const authLimiter = makeLimiter({
  windowMs: 15 * 60 * 1000,
  limit: 10,
  skipSuccessfulRequests: true,
  message: 'Terlalu banyak percobaan login. Coba lagi dalam 15 menit.',
});

// 3) Limiter tulis: POST/PUT/PATCH/DELETE lebih ketat (anti spam input data)
const writeLimiter = makeLimiter({
  windowMs: 60 * 1000,
  limit: 30,
  skip: (req) => req.method === 'GET' || req.method === 'HEAD' || req.method === 'OPTIONS',
  message: 'Terlalu banyak perubahan data. Pelan-pelan ya, coba lagi sebentar.',
});

// Timeout request: jangan biarkan request menggantung selamanya.
// Setelah timeout, respons terlambat dari handler diabaikan supaya tidak
// memicu "Cannot set headers after they are sent".
const noop = function () { return this; };
const requestTimeout = (ms = 30000) => (req, res, next) => {
  const timer = setTimeout(() => {
    if (res.headersSent) return;
    req.timedOut = true;
    res.status(503).set('Retry-After', '10').json({
      success: false,
      message: 'Permintaan terlalu lama diproses, coba lagi.',
      code: 'REQUEST_TIMEOUT',
    });
    res.json = res.send = res.sendStatus = res.redirect = noop;
  }, ms);
  const clear = () => clearTimeout(timer);
  res.on('finish', clear);
  res.on('close', clear);
  next();
};

// ---- Sanitasi input untuk Google Sheets (anti formula injection) ----
// Nama field (huruf kecil) yang TIDAK boleh diubah: kata sandi & token.
const SKIP_KEYS = new Set([
  'password', 'newpassword', 'oldpassword', 'currentpassword',
  'confirmpassword', 'passwordconfirm', 'token', 'refreshtoken',
]);
const BLOCKED_KEYS = new Set(['__proto__', 'constructor', 'prototype']); // anti prototype pollution
const MAX_DEPTH = 10;

// Sel yang diawali = @ TAB CR, atau + / - yang bukan angka murni, dianggap formula
const dangerous = (v) => /^[=@\t\r]/.test(v) || /^[+-](?![\d.,\s]*$)/.test(v);

const cleanValue = (val, key = '', depth = 0) => {
  if (depth > MAX_DEPTH) throw badRequest('Struktur data terlalu dalam.');
  if (typeof val === 'string') {
    return !SKIP_KEYS.has(String(key).toLowerCase()) && dangerous(val) ? `'${val}` : val;
  }
  if (Array.isArray(val)) return val.map((v) => cleanValue(v, key, depth + 1));
  if (val && typeof val === 'object') {
    const out = {};
    for (const k of Object.keys(val)) {
      if (BLOCKED_KEYS.has(k)) continue;
      out[k] = cleanValue(val[k], k, depth + 1);
    }
    return out;
  }
  return val;
};

const sanitizeSheetInput = (req, res, next) => {
  try {
    if (req.body && typeof req.body === 'object') req.body = cleanValue(req.body);
    next();
  } catch (err) {
    next(err);
  }
};

// TRUST_PROXY di .env: angka (jumlah proxy: 0, 1, 2), 'false', atau daftar IP/CIDR/'loopback'.
// 'true' sengaja dipetakan ke 1 karena trust proxy = true tidak aman untuk rate limit.
// Nilai kosong -> 1. Nilai tidak valid -> berhenti dengan pesan jelas (bukan menebak diam-diam).
const parseTrustProxy = (raw) => {
  // Buang spasi, kutip, dan kurung yang tidak sengaja ikut tertulis di .env (mis. {"0"})
  const v = String(raw ?? '').trim().replace(/^[\s"'{[]+|[\s"'}\]]+$/g, '');

  if (v === '') return 1;
  if (v === 'true') return 1;
  if (v === 'false') return false;
  if (/^\d+$/.test(v)) return Number(v);

  try {
    proxyaddr.compile(v.split(',').map((s) => s.trim())); // validasi IP/CIDR/'loopback'
    return v;
  } catch {
    throw new Error(
      `TRUST_PROXY tidak valid: ${JSON.stringify(raw)}. ` +
      'Isi dengan angka (0, 1, 2), "false", atau daftar IP/CIDR.'
    );
  }
};

// Pasang semua pengaman umum ke app (urutan penting)
function applySecurity(app) {
  app.disable('x-powered-by');

  // Di belakang Nginx / proxy / hosting, IP asli dibaca dari X-Forwarded-For.
  app.set('trust proxy', parseTrustProxy(process.env.TRUST_PROXY));

  app.use(helmet({ crossOriginResourcePolicy: { policy: 'cross-origin' } }));
  app.use(globalLimiter);          // limiter dipasang SEBELUM parse body (spam ditolak lebih awal)
  app.use(requestTimeout(30000));
  app.use(hpp());                  // anti HTTP parameter pollution (?a=1&a=2)
}

module.exports = {
  applySecurity,
  authLimiter,
  writeLimiter,
  sanitizeSheetInput,
};