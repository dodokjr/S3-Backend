const express = require('express');
const jwt = require('jsonwebtoken');
const router = express.Router();
const { getSheetClient, SPREADSHEET_ID, OFFICIAL_DEV_EMAIL } = require('../config/googleSheets');
const { verifyToken, JWT_SECRET } = require('../routes/middleware/auth');

// ==========================================
// HELPER STATUS LOGIN (is_login)
// ==========================================
// Nilai di sheet bisa berupa TRUE / true / 1 (sedang login)
// atau FALSE / false / 0 / kosong (tidak login).
const isLoggedInValue = (value) => {
  const v = (value === undefined || value === null ? '' : value).toString().trim().toLowerCase();
  return v === 'true' || v === '1';
};

// Menulis nilai status mengikuti format yang sudah dipakai di sheet:
// kalau nilai sebelumnya angka 0/1 -> tulis 1/0, selain itu -> tulis true/false.
const buildStatusValue = (currentValue, isLogin) => {
  const v = (currentValue === undefined || currentValue === null ? '' : currentValue).toString().trim();
  if (v === '0' || v === '1') return isLogin ? 1 : 0;
  return isLogin;
};

const STATUS_HEADERS = ['is_login', 'status', 'is_active'];

// Endpoint SIGN IN (Login)
router.post('/signin', async (req, res) => {
  try {
    const { email, password } = req.body;
    const identifier = email ? email.trim() : '';

    if (!identifier || !password) {
      return res.status(400).json({ success: false, message: 'Nama/Email dan password wajib diisi!' });
    }

    const sheets = await getSheetClient();

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A:F',
    });

    const rows = response.data.values;

    if (!rows || rows.length <= 1) {
      return res.status(401).json({ success: false, message: 'Data user tidak ditemukan di database.' });
    }

    const headers = rows[0];
    const dataRows = rows.slice(1);

    let matchedUser = null;
    let rowIndex = -1;

    for (let i = 0; i < dataRows.length; i++) {
      const row = dataRows[i];
      let userObj = {};
      headers.forEach((header, index) => {
        const cleanHeader = header.trim().toLowerCase();
        userObj[cleanHeader] = row[index] || '';
      });

      const dbEmail = userObj['email'] || '';
      const dbName = userObj['name'] || userObj['nama'] || '';
      const dbPassword = userObj['password'] || '';

      const isIdentifierMatch = (dbEmail.toLowerCase() === identifier.toLowerCase()) ||
                                (dbName.toLowerCase() === identifier.toLowerCase());

      // CATATAN: password masih dibandingkan sebagai plaintext di sini.
      // Ini di luar cakupan perubahan kali ini — beri tahu saya kalau
      // mau saya lanjutkan dengan migrasi ke bcrypt juga.
      if (isIdentifierMatch && dbPassword === password) {
        matchedUser = userObj;
        rowIndex = i + 2;
        break;
      }
    }

    if (!matchedUser) {
      return res.status(401).json({ success: false, message: 'Nama/Email atau password yang Anda masukkan salah!' });
    }

    // Default role ke yang paling rendah kalau kolom role kosong.
    const userRole = (matchedUser['role'] || 'karyawan').trim().toLowerCase();
    const userEmail = (matchedUser['email'] || '').trim().toLowerCase();
    const isRoleDeveloper = userRole === 'developer';

    // Filter keamanan ketat untuk Developer
    if (isRoleDeveloper) {
      if (userEmail !== OFFICIAL_DEV_EMAIL.toLowerCase()) {
        return res.status(403).json({
          success: false,
          message: 'Akses ditolak! Email tidak valid untuk hak akses Developer.'
        });
      }
    }

    if (!isRoleDeveloper && identifier.toLowerCase() === OFFICIAL_DEV_EMAIL.toLowerCase()) {
      return res.status(403).json({
        success: false,
        message: 'Akses ditolak! Email developer resmi tidak diizinkan menggunakan role lain.'
      });
    }

    // ==========================================
    // CEK SESI GANDA: tolak login kalau akun sedang aktif (is_login TRUE / 1)
    // ==========================================
    // Dicek SETELAH password terverifikasi, supaya orang yang tidak punya
    // password tidak bisa mengetahui status login akun tertentu.
    const statusColIndex = headers.findIndex(h => STATUS_HEADERS.includes(h.trim().toLowerCase()));
    const currentStatusValue = statusColIndex !== -1 ? (dataRows[rowIndex - 2][statusColIndex] || '') : '';

    if (statusColIndex !== -1 && isLoggedInValue(currentStatusValue)) {
      return res.status(409).json({
        success: false,
        message: 'Akun ini sedang login di perangkat lain. Silakan logout terlebih dahulu atau hubungi Admin.'
      });
    }

    // UPDATE STATUS DI GOOGLE SHEETS JADI TRUE / 1 (LOGIN)
    if (statusColIndex !== -1 && rowIndex !== -1) {
      const columnLetter = String.fromCharCode(65 + statusColIndex);
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `Users!${columnLetter}${rowIndex}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: [[buildStatusValue(currentStatusValue, true)]]
        }
      });
    }

    // Token JWT ditandatangani server dengan JWT_SECRET, berisi email + role,
    // dan punya masa berlaku (expiry).
    const token = jwt.sign(
      { email: userEmail, role: userRole },
      JWT_SECRET,
      { expiresIn: '8h' }
    );

    return res.json({
      success: true,
      message: 'Sign In berhasil!',
      role: userRole,
      token: token,
      user: {
        name: matchedUser['name'] || matchedUser['nama'] || '',
        email: matchedUser['email'] || '',
        role: userRole
      }
    });

  } catch (error) {
    console.error('Sign In Error:', error);
    res.status(500).json({ success: false, message: 'Terjadi kesalahan pada server: ' + error.message });
  }
});

// Endpoint LOGOUT
// Dilindungi verifyToken; email diambil dari token (req.user.email) yang sudah
// tervalidasi — bukan dari req.body.email.
router.post('/logout', verifyToken, async (req, res) => {
  try {
    const email = req.user.email;

    const sheets = await getSheetClient();

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A:F',
    });

    const rows = response.data.values;
    if (!rows || rows.length <= 1) {
      return res.status(404).json({ success: false, message: 'Data user tidak ditemukan.' });
    }

    const headers = rows[0];
    const dataRows = rows.slice(1);
    const emailColIndex = headers.findIndex(h => h.trim().toLowerCase() === 'email');
    let rowIndex = -1;

    for (let i = 0; i < dataRows.length; i++) {
      const dbEmail = dataRows[i][emailColIndex] || '';

      if (dbEmail.toLowerCase() === email.toLowerCase()) {
        rowIndex = i + 2;
        break;
      }
    }

    if (rowIndex === -1) {
      return res.status(404).json({ success: false, message: 'User dengan email tersebut tidak ditemukan.' });
    }

    const statusColIndex = headers.findIndex(h => STATUS_HEADERS.includes(h.trim().toLowerCase()));

    // UPDATE STATUS DI GOOGLE SHEETS JADI FALSE / 0 (LOGOUT)
    if (statusColIndex !== -1) {
      const currentStatusValue = dataRows[rowIndex - 2][statusColIndex] || '';
      const columnLetter = String.fromCharCode(65 + statusColIndex);
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `Users!${columnLetter}${rowIndex}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: [[buildStatusValue(currentStatusValue, false)]]
        }
      });
    }

    return res.json({
      success: true,
      message: 'Logout berhasil!'
    });

  } catch (error) {
    console.error('Logout Error:', error);
    res.status(500).json({ success: false, message: 'Terjadi kesalahan pada server: ' + error.message });
  }
});

module.exports = router;