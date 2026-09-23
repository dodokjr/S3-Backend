require("dotenv").config();
const express = require('express');
const { google } = require('googleapis');

const app = express();
app.use(express.json());

const PORT = process.env.PORT || 3001;
const SPREADSHEET_ID = '1RkRKEt8AW0CNArv-FqEM9JJYUox1fGHZkV9JLuQgqXw';

async function getSheetClient() {
  // Pastikan process.env.GOOGLE_CREDENTIALS sudah di-parse menjadi objek
  const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS);
  
  const auth = new google.auth.GoogleAuth({
    credentials, // Menggunakan 'credentials' (objek), bukan 'keyFile' (path file)
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

app.listen(PORT, () => {
  console.log(`Server berjalan di http://localhost:${PORT}`);
});