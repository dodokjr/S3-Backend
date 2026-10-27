const jwt = require('jsonwebtoken');
const { OFFICIAL_DEV_EMAIL } = require('../../config/googleSheets');

const JWT_SECRET = process.env.JWT_SECRET;

if (!JWT_SECRET) {
  throw new Error('JWT_SECRET belum diset di environment variables!');
}

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

    const decoded = jwt.verify(token, JWT_SECRET); 
    if (decoded.role === 'developer' && decoded.email !== OFFICIAL_DEV_EMAIL.toLowerCase()) {
      return res.status(403).json({
        success: false,
        message: 'Token tidak valid untuk hak akses Developer.',
      });
    }

    req.user = decoded; // staff: { email, role, iat, exp } | customer: { nama, email, role, iat, exp }
    next();
  } catch (error) {
    if (error.name === 'TokenExpiredError') {
      return res.status(401).json({ success: false, message: 'Sesi telah berakhir, silakan login kembali.' });
    }
    return res.status(401).json({ success: false, message: 'Token tidak valid.' });
  }
};

const CUSTOMER_TOKEN_EXPIRES_IN = '7d';

const signCustomerToken = ({ nama, email }) =>
  jwt.sign(
    { nama, email: email.toLowerCase(), role: 'customer' },
    JWT_SECRET,
    { expiresIn: CUSTOMER_TOKEN_EXPIRES_IN }
  );

// Dipasang SETELAH verifyToken: hanya izinkan token customer
const allowCustomer = (req, res, next) => {
  if ((req.user?.role || '').toLowerCase() === 'customer') return next();

  return res.status(403).json({
    success: false,
    message: 'Akses ditolak! Endpoint ini hanya untuk customer.',
  });
};

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
//   - /popchat  -> Developer, Admin, CS (balas chat customer)
//   - role lain (mis. karyawan, customer) -> ditolak dari endpoint yang pakai middleware ini
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
  popchat: ['developer', 'admin', 'cs'],
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

module.exports = {
  verifyToken,
  allowDeveloperAndAdmin,
  checkModuleAccess,
  signCustomerToken,
  allowCustomer,
  JWT_SECRET,
};