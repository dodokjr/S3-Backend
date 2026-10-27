require('dotenv').config();
const http = require('http');
const express = require('express');
const cors = require('cors');
const apiRoutes = require('./src/routes/api');
const authRoutes = require('./src/routes/auth');
const { notFoundHandler, errorHandler } = require('./src/routes/Utils/Errors');
const { applySecurity, authLimiter, writeLimiter, sanitizeSheetInput } = require('./src/routes/Utils/Security');

const app = express();
const PORT = process.env.PORT || 3001;
const isProd = process.env.NODE_ENV === 'production';

// ==========================================
// CORS (dipasang PALING AWAL supaya preflight OPTIONS tidak kena rate limit
// dan respons 429/error tetap membawa header CORS)
// ==========================================
// Bersihkan format .env: buang spasi, kutip, kurung, dan garis miring di akhir
const cleanOrigin = (o) =>
  o.trim().replace(/^[\s"'{[]+|[\s"'}\]]+$/g, '').replace(/\/+$/, '');

const allowedOrigins = (process.env.CORS_ORIGIN || '')
  .split(',').map(cleanOrigin).filter(Boolean);

if (!allowedOrigins.length && isProd) {
  console.warn('PERINGATAN: CORS_ORIGIN kosong di production, request lintas-origin ditolak.');
}

app.use(cors({
  origin: (origin, cb) => {
    if (!origin) return cb(null, true);                            // Postman/curl/server-ke-server
    if (allowedOrigins.includes(origin)) return cb(null, true);
    if (!allowedOrigins.length && !isProd) return cb(null, true);  // dev tanpa CORS_ORIGIN
    console.warn(`CORS ditolak: origin="${origin}" | diizinkan=${JSON.stringify(allowedOrigins)}`);
    return cb(null, false);
  },
  methods: ['GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
  // credentials: true, // aktifkan HANYA jika frontend memakai cookie / withCredentials
  maxAge: 600,
}));

// ==========================================
// KEAMANAN (trust proxy, helmet, rate limit umum, timeout, hpp)
// ==========================================
applySecurity(app);

app.use(express.json({ limit: '100kb' }));
app.use(sanitizeSheetInput); // field password/token dikecualikan di dalam Security.js

// ==========================================
// RATE LIMIT PER JALUR
// ==========================================
const isAuthPath = (p) => p === '/auth' || p.startsWith('/auth/');

app.use('/s3/api/auth', authLimiter);
app.use('/s3/api', (req, res, next) =>
  isAuthPath(req.path) ? next() : writeLimiter(req, res, next)
);

// ==========================================
// ROUTES
// authRoutes dipasang SEBELUM apiRoutes: kalau apiRoutes punya middleware
// "wajib login" di level router, /auth/login tidak akan terblokir 401.
// ==========================================
app.use('/s3/api/auth', authRoutes);
app.use('/s3/api', apiRoutes);

app.use(notFoundHandler);
app.use(errorHandler);

// ==========================================
// SERVER (timeout diset SEBELUM listen)
// ==========================================
const server = http.createServer(app);
server.requestTimeout = 30000;   // batas menerima seluruh request
server.headersTimeout = 20000;   // harus <= requestTimeout
server.keepAliveTimeout = 5000;  // naikkan (mis. 65000) jika proxy memakai keep-alive ke Node dan muncul 502 acak

server.on('error', (err) => {
  console.error('Gagal start server:', err.message);
  process.exit(1);
});

server.listen(PORT, () => {
  console.log(`Server berjalan di http://localhost:${PORT}`);
});

// ==========================================
// SHUTDOWN
// ==========================================
let shuttingDown = false;
const shutdown = (reason, code = 0) => {
  if (shuttingDown) return;
  shuttingDown = true;
  console.log(`Menutup server (${reason})...`);
  server.close(() => process.exit(code));
  server.closeIdleConnections?.(); // Node 18.2+: koneksi keep-alive menganggur tidak menahan close()
  setTimeout(() => {
    server.closeAllConnections?.(); // paksa tutup sisa koneksi
    process.exit(code || 1);
  }, 10000).unref();
};

process.on('SIGTERM', () => shutdown('SIGTERM'));
process.on('SIGINT', () => shutdown('SIGINT'));
process.on('unhandledRejection', (reason) => {
  console.error('Unhandled Rejection:', reason);
});
process.on('uncaughtException', (err) => {
  console.error('Uncaught Exception:', err);
  shutdown('uncaughtException', 1);
});