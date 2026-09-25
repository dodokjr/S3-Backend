const express = require('express');
const jwt = require('jsonwebtoken');
const router = express.Router();
const { getSheetClient, SPREADSHEET_ID, OFFICIAL_DEV_EMAIL } = require('../config/googleSheets');
const { verifyToken, JWT_SECRET } = require('../routes/middleware/auth');

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
      // Ini di luar cakupan perbaikan JWT kali ini — beri tahu saya kalau
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

    // PERBAIKAN: default role sebelumnya 'admin' jika kolom role kosong — ini
    // berbahaya karena baris user yang tidak lengkap otomatis jadi admin.
    // Default sekarang ke role paling rendah.
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

    // UPDATE STATUS DI GOOGLE SHEETS JADI TRUE (LOGIN)
    const statusColIndex = headers.findIndex(h => ['is_login', 'status', 'is_active'].includes(h.trim().toLowerCase()));
    
    if (statusColIndex !== -1 && rowIndex !== -1) {
      const columnLetter = String.fromCharCode(65 + statusColIndex);
      // PERBAIKAN: range sebelumnya salah ditulis dengan \( \) alih-alih ${ }.
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `Users!${columnLetter}${rowIndex}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: [[true]]
        }
      });
    }

    // PERBAIKAN UTAMA: token sekarang adalah JWT sungguhan yang ditandatangani
    // server dengan JWT_SECRET, berisi email + role, dan punya masa berlaku (expiry).
    // Ini menggantikan token string timestamp sebelumnya yang tidak diverifikasi
    // sama sekali oleh route lain. SEMUA role (termasuk developer) sekarang
    // mendapat token, karena middleware di route lain akan memverifikasi token
    // ini untuk menentukan role — bukan lagi header yang bisa dipalsukan.
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
// PERBAIKAN: sekarang dilindungi verifyToken, dan email diambil dari token
// (req.user.email) yang sudah tervalidasi — bukan dari req.body.email yang
// sebelumnya bisa diisi bebas oleh siapa pun untuk memaksa logout user lain.
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
    let rowIndex = -1;

    for (let i = 0; i < dataRows.length; i++) {
      const row = dataRows[i];
      const dbEmail = row[headers.findIndex(h => h.trim().toLowerCase() === 'email')] || '';
      
      if (dbEmail.toLowerCase() === email.toLowerCase()) {
        rowIndex = i + 2;
        break;
      }
    }

    if (rowIndex === -1) {
      return res.status(404).json({ success: false, message: 'User dengan email tersebut tidak ditemukan.' });
    }

    const statusColIndex = headers.findIndex(h => ['is_login', 'status', 'is_active'].includes(h.trim().toLowerCase()));

    if (statusColIndex !== -1) {
      const columnLetter = String.fromCharCode(65 + statusColIndex);
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `Users!${columnLetter}${rowIndex}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: [[false]]
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