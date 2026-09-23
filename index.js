require("dotenv").config();
const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');

const app = express();
app.use(express.json());
app.use(cors());

const PORT = process.env.PORT || 3001;
const SPREADSHEET_ID = '1RkRKEt8AW0CNArv-FqEM9JJYUox1fGHZkV9JLuQgqXw';

// Email resmi yang diizinkan khusus untuk Developer
const OFFICIAL_DEV_EMAIL = 's3-600@appss-3c587.iam.gserviceaccount.com';

async function getSheetClient() {
  const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS);
  
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  
  const client = await auth.getClient();
  return google.sheets({ version: 'v4', auth: client });
}

// Endpoint untuk MEMBACA data Users
app.get('/api/data', async (req, res) => {
  try {
    const sheets = await getSheetClient();
    
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A:C', 
    });

    const rows = response.data.values;
    
    if (!rows || rows.length === 0) {
      return res.json({ success: true, data: [] });
    }

    const headers = rows[0]; 
    const dataRows = rows.slice(1); 

    const formattedData = dataRows.map((row) => {
      let obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index] || ''; 
      });
      return obj;
    });

    res.json({
      success: true,
      data: formattedData,
    });

  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: error.message });
  }
});

// Endpoint untuk MEMBACA data Real Stock
app.get('/api/stock', async (req, res) => {
  try {
    const sheets = await getSheetClient();
    
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Real_Stock!A:H', 
    });

    const rows = response.data.values;
    
    if (!rows || rows.length === 0) {
      return res.json({ success: true, data: [] });
    }

    const headers = rows[0]; 
    const dataRows = rows.slice(1); 

    const formattedData = dataRows.map((row) => {
      let obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index] || ''; 
      });
      return obj;
    });

    res.json({
      success: true,
      data: formattedData,
    });

  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: error.message });
  }
});

// ==========================================
// ENDPOINT LOGIN DENGAN FILTER KEAMANAN DEVELOPER
// ==========================================
app.post('/api/login', async (req, res) => {
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
      message: 'Login berhasil!',
      role: userRole,
      token: token, 
      user: {
        name: matchedUser['name'] || matchedUser['nama'] || '',
        email: matchedUser['email'] || '',
        role: userRole
      }
    });

  } catch (error) {
    console.error('Login Error:', error);
    res.status(500).json({ success: false, message: 'Terjadi kesalahan pada server: ' + error.message });
  }
});

app.listen(PORT, () => {
  console.log(`Server berjalan di http://localhost:${PORT}`);
});