class AppError extends Error {
  constructor(status, message, code) {
    super(message);
    this.name = 'AppError';
    this.status = status;
    this.code = code;
    this.isOperational = true; // error yang kita lempar sendiri (bukan bug)
  }
}

// ---- Factory per kode status ----
const badRequest         = (msg = 'Permintaan tidak valid.')              => new AppError(400, msg, 'BAD_REQUEST');
const unauthorized       = (msg = 'Silakan login terlebih dahulu.')       => new AppError(401, msg, 'UNAUTHORIZED');
const forbidden          = (msg = 'Anda tidak punya akses.')              => new AppError(403, msg, 'FORBIDDEN');
const notFound           = (msg = 'Data tidak ditemukan.')                => new AppError(404, msg, 'NOT_FOUND');
const conflict           = (msg = 'Data sudah ada / bentrok.')            => new AppError(409, msg, 'CONFLICT');
const tooManyRequests    = (msg = 'Terlalu banyak permintaan, coba lagi nanti.') => new AppError(429, msg, 'TOO_MANY_REQUESTS');
const internalError      = (msg = 'Terjadi kesalahan pada server.')       => new AppError(500, msg, 'INTERNAL_ERROR');
const serviceUnavailable = (msg = 'Layanan sedang tidak tersedia, coba lagi nanti.') =>
  new AppError(503, msg, 'SERVICE_UNAVAILABLE');

// Bungkus handler async supaya error otomatis masuk ke errorHandler (tanpa try/catch)
const asyncHandler = (fn) => (req, res, next) => Promise.resolve(fn(req, res, next)).catch(next);

// ---- 404 untuk route yang tidak ada (taruh SETELAH semua route) ----
const notFoundHandler = (req, res, next) => {
  next(notFound(`Route ${req.method} ${req.originalUrl} tidak ditemukan.`));
};

const NET_CODES = ['ECONNRESET', 'ETIMEDOUT', 'ENOTFOUND', 'EAI_AGAIN', 'ECONNREFUSED', 'ESOCKETTIMEDOUT'];
const BUSY_STATUS = [429, 500, 502, 503, 504];

// Ubah error dari library (Google Sheets, body-parser, jaringan) jadi AppError yang tepat
function normalizeError(err) {
  if (err instanceof AppError) return err;
  if (!err || typeof err !== 'object') return null;

  // JSON body rusak / terlalu besar (body-parser)
  if (err.type === 'entity.parse.failed') return badRequest('Format JSON tidak valid.');
  if (err.type === 'entity.too.large') return new AppError(413, 'Data terlalu besar.', 'PAYLOAD_TOO_LARGE');

  // Jaringan putus -> 503
  if (NET_CODES.includes(err.code)) {
    return serviceUnavailable('Tidak dapat terhubung ke Google Sheets, coba lagi nanti.');
  }

  // Error dari API upstream (Google Sheets): ada err.response
  if (err.response) {
    const upstream = Number(err.code || err.status || err.response.status);
    const reason = String(
      err.errors?.[0]?.reason || err.response?.data?.error?.errors?.[0]?.reason || ''
    );
    // Kuota habis / Google bermasalah -> 503 (kuota Sheets kadang datang sebagai 403 rateLimitExceeded)
    if (BUSY_STATUS.includes(upstream) || /ratelimit|quota/i.test(reason)) {
      return serviceUnavailable('Layanan Google Sheets sedang sibuk, coba lagi sebentar.');
    }
    // 400/401/403/404 dari Google = masalah konfigurasi server, BUKAN kesalahan user.
    // Jangan diteruskan apa adanya (bisa membocorkan detail & membuat frontend mengira sesi habis).
    return null; // -> 500 generik
  }

  // Error lain yang membawa status HTTP valid (mis. dari middleware lain)
  const st = Number(err.status || err.statusCode);
  if (st >= 400 && st < 500) return new AppError(st, err.message, 'ERROR');
  if (st >= 500 && st < 600) return new AppError(st, 'Terjadi kesalahan pada server.', 'ERROR');

  return null; // bug tak terduga -> 500
}

// ---- Handler error GLOBAL (taruh PALING AKHIR, setelah notFoundHandler) ----
// Wajib 4 argumen supaya Express mengenalinya sebagai error handler.
// eslint-disable-next-line no-unused-vars
const errorHandler = (err, req, res, next) => {
  const known = normalizeError(err);
  const status = known ? known.status : 500;
  const isProd = process.env.NODE_ENV === 'production';

  // Log hanya error server (5xx); 4xx adalah kesalahan klien
  if (status >= 500) console.error(`[${status}] ${req.method} ${req.originalUrl}`, err);

  // Respons sudah terkirim (mis. setelah timeout): jangan kirim lagi.
  // Jika respons sudah selesai, cukup berhenti; jika masih streaming, serahkan ke Express (tutup koneksi).
  if (res.headersSent) return res.writableEnded ? undefined : next(err);

  // Di production, pesan error 500 yang tak terduga disembunyikan dari klien
  const message = known
    ? known.message
    : (isProd ? 'Terjadi kesalahan pada server.' : err?.message || 'Terjadi kesalahan pada server.');

  if (status === 503) res.set('Retry-After', '10');

  res.status(status).json({
    success: false,
    message,
    ...(known?.code ? { code: known.code } : {}),
    ...(!isProd && status >= 500 && err?.stack ? { stack: err.stack } : {}),
  });
};

module.exports = {
  AppError,
  badRequest,
  unauthorized,
  forbidden,
  notFound,
  conflict,
  tooManyRequests,
  internalError,
  serviceUnavailable,
  asyncHandler,
  notFoundHandler,
  errorHandler,
};