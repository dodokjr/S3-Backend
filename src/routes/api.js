const express = require('express');
const router = express.Router();
const { getSheetClient, SPREADSHEET_ID } = require('../config/googleSheets');

// Endpoint untuk MEMBACA data Users
router.get('/users', async (req, res) => {
  try {
    const sheets = await getSheetClient();
    
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A:F', 
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
router.get('/stock', async (req, res) => {
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

module.exports = router;