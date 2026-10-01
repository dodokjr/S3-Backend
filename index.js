require("dotenv").config();
const express = require('express');
const cors = require('cors');
const apiRoutes = require('./src/routes/api');
const authRoutes = require('./src/routes/auth');
const { notFoundHandler, errorHandler } = require('./src/routes/Utils/Errors');
const { applySecurity, authLimiter, writeLimiter, sanitizeSheetInput } = require('./src/routes/Utils/Security');

const app = express();

// ==========================================
// KEAMANAN (helmet, rate limit umum, timeout, hpp)
// ==========================================
applySecurity(app);

// CORS: isi CORS_ORIGIN di .env, pisahkan koma. Contoh:
// CORS_ORIGIN=https://webkamu.com,https://admin.webkamu.com
// Jika kosong, semua origin diizinkan (hanya cocok untuk development).
const allowedOrigins = (process.env.CORS_ORIGIN || '')
  .split(',').map((o) => o.trim()).filter(Boolean);
app.use(cors({
  origin: allowedOrigins.length ? allowedOrigins : true,
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
  maxAge: 600,
}));

// Batasi ukuran body supaya server tidak dibanjiri data besar
app.use(express.json({ limit: '100kb' }));
app.use(sanitizeSheetInput);

const PORT = process.env.PORT || 3001;

// Limiter khusus: login (anti brute-force) dan tulis data (anti spam)
app.use('/s3/api/auth', authLimiter);
app.use('/s3/api', writeLimiter);

// Mendaftarkan Routes
app.use('/s3/api', apiRoutes);       // Mengakses /s3/api/stock, /s3/api/sales, dst
app.use('/s3/api/auth', authRoutes); // Mengakses /s3/api/auth/signin dan /s3/api/auth/logout

// ==========================================
// ERROR HANDLER GLOBAL (berlaku untuk SEMUA halaman / route)
// WAJIB di paling bawah, setelah semua app.use(route)
// ==========================================
app.use(notFoundHandler); // URL yang tidak ada -> 404 (JSON)
app.use(errorHandler);    // semua error (400, 401, 403, 404, 429, 500, 503) -> JSON seragam

const server = app.listen(PORT, () => {
  console.log(`Server berjalan di http://localhost:${PORT}`);
});

// Batas waktu koneksi (anti slowloris / koneksi menggantung)
server.requestTimeout = 30000;
server.headersTimeout = 35000;
server.keepAliveTimeout = 5000;

// ==========================================
// Mati dengan rapi & anti crash diam-diam
// ==========================================
let shuttingDown = false;
const shutdown = (reason, code = 0) => {
  if (shuttingDown) return;
  shuttingDown = true;
  console.log(`Menutup server (${reason})...`);
  server.close(() => process.exit(code));
  setTimeout(() => process.exit(code || 1), 10000).unref(); // paksa keluar jika macet
};

process.on('SIGTERM', () => shutdown('SIGTERM'));
process.on('SIGINT', () => shutdown('SIGINT'));
process.on('unhandledRejection', (reason) => {
  console.error('Unhandled Rejection:', reason);
});
// Setelah uncaughtException proses tidak aman dilanjutkan:
// catat, tutup rapi, lalu biarkan PM2 / Docker menyalakan ulang.
process.on('uncaughtException', (err) => {
  console.error('Uncaught Exception:', err);
  shutdown('uncaughtException', 1);
});