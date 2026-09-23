const express = require('express');
const { google } = require('googleapis');

const app = express();
app.use(express.json());

const PORT = 3001;
const SPREADSHEET_ID = '1RkRKEt8AW0CNArv-FqEM9JJYUox1fGHZkV9JLuQgqXw';

async function getSheetClient() {
  const auth = new google.auth.GoogleAuth({
    keyFile: 'credentials.json',
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  const client = await auth.getClient();
  return google.sheets({ version: 'v4', auth: client });
}

// Endpoint untuk MEMBACA data dan mengubahnya menjadi objek berbasis header sheet
app.get('/api/data', async (req, res) => {
  try {
    const sheets = await getSheetClient();
    
    // Ambil data (sesuaikan range, misal A:C untuk mengambil seluruh kolom yang terisi)
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A:C', // Ganti 'Sheet1' jika nama tab Anda berbeda
    });

    const rows = response.data.values;
    
    if (!rows || rows.length === 0) {
      return res.json({ success: true, data: [] });
    }

    // Baris pertama dianggap sebagai Header (Nama Kolom)
    const headers = rows[0]; 
    const dataRows = rows.slice(1); // Baris data mulai dari baris ke-2

    // Mapping array 2D menjadi array of objects berdasarkan header
    const formattedData = dataRows.map((row) => {
      let obj = {};
      headers.forEach((header, index) => {
        // Masukkan nilai kolom berdasarkan nama header-nya
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

app.get('/api/stock', async (req, res) => {
  try {
    const sheets = await getSheetClient();
    
    // Ambil data (sesuaikan range, misal A:C untuk mengambil seluruh kolom yang terisi)
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Real_Stock!A:H', // Ganti 'Sheet1' jika nama tab Anda berbeda
    });

    const rows = response.data.values;
    
    if (!rows || rows.length === 0) {
      return res.json({ success: true, data: [] });
    }

    // Baris pertama dianggap sebagai Header (Nama Kolom)
    const headers = rows[0]; 
    const dataRows = rows.slice(1); // Baris data mulai dari baris ke-2

    // Mapping array 2D menjadi array of objects berdasarkan header
    const formattedData = dataRows.map((row) => {
      let obj = {};
      headers.forEach((header, index) => {
        // Masukkan nilai kolom berdasarkan nama header-nya
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

app.listen(PORT, () => {
  console.log(`Server berjalan di http://localhost:${PORT}`);
});