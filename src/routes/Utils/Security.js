// ==========================================================
// src/middleware/security.js
// Keamanan server: rate limiter, anti spam/brute-force, header aman,
// batas ukuran body, timeout request, anti HTTP parameter pollution,
// dan pembersih "formula injection" untuk Google Sheets.
//
// Install dulu:  npm install helmet express-rate-limit hpp
// ==========================================================
const helmet = require('helmet');
const hpp = require('hpp');
const { rateLimit } = require('express-rate-limit');
const { tooManyRequests } = require('./Errors');

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

// 2) Limiter login: anti brute-force password (hanya login GAGAL yang dihitung)
const authLimiter = makeLimiter({
  windowMs: 15 * 60 * 1000,
  limit: 10,
  skipSuccessfulRequests: true,
  message: 'Terlalu banyak percobaan login. Coba lagi dalam 15 menit.',
});

// 3) Limiter tulis: POST/PUT/DELETE lebih ketat (anti spam input data)
const writeLimiter = makeLimiter({
  windowMs: 60 * 1000,
  limit: 30,
  skip: (req) => req.method === 'GET' || req.method === 'HEAD' || req.method === 'OPTIONS',
  message: 'Terlalu banyak perubahan data. Pelan-pelan ya, coba lagi sebentar.',
});

// Timeout request: jangan biarkan request menggantung selamanya
const requestTimeout = (ms = 30000) => (req, res, next) => {
  const timer = setTimeout(() => {
    if (!res.headersSent) {
      res.status(503).json({
        success: false,
        message: 'Permintaan terlalu lama diproses, coba lagi.',
        code: 'REQUEST_TIMEOUT',
      });
    }
  }, ms);
  const clear = () => clearTimeout(timer);
  res.on('finish', clear);
  res.on('close', clear);
  next();
};

// Cegah formula injection di Google Sheets (valueInputOption USER_ENTERED
// akan menjalankan teks yang diawali = + - @ sebagai rumus).
// Teks berbahaya diberi awalan ' supaya dianggap teks biasa. Angka negatif (-5) tetap aman.
const SKIP_KEYS = new Set(['password']); // password jangan diubah
const dangerous = (v) => /^[=@]/.test(v) || /^[+-](?![\d.,\s]*$)/.test(v);
const cleanValue = (val, key) => {
  if (typeof val === 'string') {
    return !SKIP_KEYS.has(key) && dangerous(val) ? `'${val}` : val;
  }
  if (Array.isArray(val)) return val.map((v) => cleanValue(v, key));
  if (val && typeof val === 'object') {
    const out = {};
    for (const k of Object.keys(val)) out[k] = cleanValue(val[k], k);
    return out;
  }
  return val;
};
const sanitizeSheetInput = (req, res, next) => {
  if (req.body && typeof req.body === 'object') req.body = cleanValue(req.body);
  next();
};

// Pasang semua pengaman umum ke app (urutan penting)
function applySecurity(app) {
  app.disable('x-powered-by');

  // Di belakang Nginx / proxy / hosting, IP asli dibaca dari X-Forwarded-For.
  // Atur lewat .env: TRUST_PROXY=1 (default 1). Isi 0 jika tanpa proxy.
  app.set('trust proxy', Number(process.env.TRUST_PROXY ?? 1));

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