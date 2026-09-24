const express = require('express');
const router = express.Router();
const { getSheetClient, SPREADSHEET_ID, OFFICIAL_DEV_EMAIL } = require('../config/googleSheets');

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
      range: 'Users!A:E', 
    });

    const rows = response.data.values;
    
    if (!rows || rows.length <= 1) {
      return res.status(401).json({ success: false, message: 'Data user tidak ditemukan di database.' });
    }

    const headers = rows[0]; 
    const dataRows = rows.slice(1);

    let matchedUser = null;

    for (let row of dataRows) {
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

      if (isIdentifierMatch && dbPassword === password) {
        matchedUser = userObj;
        break;
      }
    }

    if (!matchedUser) {
      return res.status(401).json({ success: false, message: 'Nama/Email atau password yang Anda masukkan salah!' });
    }

    const userRole = (matchedUser['role'] || 'admin').trim();
    const userEmail = (matchedUser['email'] || '').trim().toLowerCase();
    const isRoleDeveloper = userRole.toLowerCase() === 'developer';

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

    const token = isRoleDeveloper ? null : 'token_sess_' + Date.now();

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
router.post('/logout', (req, res) => {
  try {
    // Karena aplikasi menggunakan sistem token stateless / session sederhana,
    // proses logout biasanya ditangani di sisi klien (menghapus token/localStorage).
    // Di sisi server, kita bisa mengirimkan respons sukses untuk mengonfirmasi aksi.
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