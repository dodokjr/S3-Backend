const jwt = require('jsonwebtoken');
const { OFFICIAL_DEV_EMAIL } = require('../../config/googleSheets');

const JWT_SECRET = process.env.JWT_SECRET;

if (!JWT_SECRET) {
  // Gagal cepat saat startup daripada diam-diam menandatangani token dengan
  // secret undefined (yang bisa ditebak/dipalsukan siapa pun).
  throw new Error('JWT_SECRET belum diset di environment variables!');
}

// ==========================================
// MIDDLEWARE: VERIFIKASI TOKEN JWT
// ==========================================
// Membaca token dari header "Authorization: Bearer <token>", memverifikasi
// tanda tangannya, lalu menaruh payload yang sudah tervalidasi di req.user.
// Role di sini TIDAK BISA dipalsukan oleh client, karena token hanya bisa
// dibuat oleh server saat signin (menggunakan JWT_SECRET yang hanya diketahui server).
const verifyToken = (req, res, next) => {
  try {
    const authHeader = req.headers['authorization'] || '';
    const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : null;

    if (!token) {
      return res.status(401).json({
        success: false,
        message: 'Akses ditolak! Token tidak ditemukan. Silakan login kembali.',
      });
    }

    const decoded = jwt.verify(token, JWT_SECRET); // throws jika invalid/expired

    // Lapisan pertahanan tambahan: kalau token mengklaim role 'developer',
    // pastikan emailnya tetap cocok dengan email developer resmi.
    if (decoded.role === 'developer' && decoded.email !== OFFICIAL_DEV_EMAIL.toLowerCase()) {
      return res.status(403).json({
        success: false,
        message: 'Token tidak valid untuk hak akses Developer.',
      });
    }

    req.user = decoded; // { email, role, iat, exp }
    next();
  } catch (error) {
    if (error.name === 'TokenExpiredError') {
      return res.status(401).json({ success: false, message: 'Sesi telah berakhir, silakan login kembali.' });
    }
    return res.status(401).json({ success: false, message: 'Token tidak valid.' });
  }
};

// ==========================================
// MIDDLEWARE: HAK AKSES DEVELOPER & ADMIN
// ==========================================
// Dipasang SETELAH verifyToken. Mengecek req.user.role yang sudah tervalidasi
// dari JWT, bukan dari header/body yang bisa diedit bebas oleh client.
const allowDeveloperAndAdmin = (req, res, next) => {
  const role = (req.user?.role || '').toLowerCase();

  if (role === 'developer' || role === 'admin') {
    return next();
  }

  return res.status(403).json({
    success: false,
    message: 'Akses ditolak! Hanya role Developer dan Admin yang diizinkan mengubah data.',
  });
};

// ==========================================
// MIDDLEWARE: HAK AKSES PER MODUL
// ==========================================
// Dipasang SETELAH verifyToken, sama seperti allowDeveloperAndAdmin.
// Lokasi: routes/middleware/auth.js
//
// Aturan akses per modul (berlaku untuk GET, POST, PUT, dan DELETE):
//   - /sales    -> Sales, Developer, Admin
//   - /finance  -> Finance, Developer, Admin
//   - /stock    -> Developer, Admin
//   - /gudang   -> Gudang, Sales, Developer, Admin
//   - /users    -> Developer, Admin
//   - role lain (mis. karyawan) -> ditolak dari endpoint yang pakai middleware ini
//
// Mau tambah/ubah akses? Cukup edit tabel MODULE_ROLES di bawah.
//
// Pemakaian: router.post('/sales', verifyToken, checkModuleAccess('sales'), handler)
const MODULE_ROLES = {
  sales:   ['developer', 'admin', 'sales'],
  finance: ['developer', 'admin', 'finance'],
  stock:   ['developer', 'admin'],
  gudang:  ['developer', 'admin', 'sales', 'gudang'],
  users:   ['developer', 'admin'],
};

const checkModuleAccess = (moduleName) => (req, res, next) => {
  const role = (req.user?.role || '').toLowerCase();
  const allowedRoles = MODULE_ROLES[moduleName] || ['developer', 'admin'];

  if (allowedRoles.includes(role)) {
    return next();
  }

  return res.status(403).json({
    success: false,
    message: `Akses ditolak! Endpoint /${moduleName} hanya bisa diakses oleh role: ${allowedRoles.join(', ')}.`,
  });
};

module.exports = { verifyToken, allowDeveloperAndAdmin, checkModuleAccess, JWT_SECRET };