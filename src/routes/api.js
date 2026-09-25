const express = require('express');
const router = express.Router();
const { getSheetClient, SPREADSHEET_ID } = require('../config/googleSheets');
const { verifyToken, allowDeveloperAndAdmin } = require('../routes/middleware/auth');

// ==========================================
// 1. ENDPOINT USERS
// ==========================================

// Membaca data Users (butuh login, tapi tidak harus admin/dev — sesuaikan
// jika ingin GET ini juga dibatasi hanya admin/dev dengan menambahkan
// allowDeveloperAndAdmin setelah verifyToken)
router.get('/users', verifyToken, async (req, res) => {
  try {
    const sheets = await getSheetClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A:F', 
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) return res.json({ success: true, data: [] });

    const headers = rows[0].map(h => h.trim().toLowerCase());
    const dataRows = rows.slice(1); 

    const formattedData = dataRows.map((row) => {
      let obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index] || ''; 
      });
      delete obj.password; // Password tidak pernah dikirim ke client
      return obj;
    });

    res.json({ success: true, data: formattedData });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: error.message });
  }
});

// Tambah User Baru (Hanya Developer & Admin)
// Tambah User Baru (Hanya Developer & Admin)
router.post('/users', verifyToken, allowDeveloperAndAdmin, async (req, res) => {
  try {
    const { name, email, password, role, is_login } = req.body;

    if (!name || !email) {
      return res.status(400).json({ success: false, message: 'Nama dan Email wajib diisi!' });
    }

    const sheets = await getSheetClient();

    // Ambil data yang sudah ada untuk menentukan id berikutnya (auto-increment).
    // Kolom A sekarang berisi id, sehingga range diperluas dari A:E menjadi A:F.
    const existingResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A:F',
    });

    const rows = existingResponse.data.values || [];
    const dataRows = rows.slice(1); // lewati baris header

    let maxId = 0;
    dataRows.forEach((row) => {
      const idNumber = Number(row[0]);
      if (!Number.isNaN(idNumber) && idNumber > maxId) {
        maxId = idNumber;
      }
    });
    const newId = (maxId + 1).toString();

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A:F',
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[newId, name, email, password || '', role || 'karyawan', is_login || false]]
      }
    });

    res.json({
      success: true,
      message: 'User berhasil ditambahkan!',
      data: { id: newId, name, email, role: role || 'karyawan', status: is_login ? 'TRUE' : 'FALSE' }
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: error.message });
  }
});

// Update User (Hanya Developer & Admin)
router.put('/users', verifyToken, allowDeveloperAndAdmin, async (req, res) => {
  try {
    const { email, name, password, role, is_login } = req.body;
    if (!email) return res.status(400).json({ success: false, message: 'Email wajib disertakan untuk update user.' });

    const sheets = await getSheetClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A:F', // PERBAIKAN: A:E -> A:F karena kolom id ditambahkan di posisi A
    });

    const rows = response.data.values;
    if (!rows || rows.length <= 1) return res.status(404).json({ success: false, message: 'Data user kosong.' });

    const headers = rows[0].map(h => h.trim().toLowerCase());
    let rowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      const dbEmail = rows[i][headers.indexOf('email')] || '';
      if (dbEmail.toLowerCase() === email.toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return res.status(404).json({ success: false, message: 'User tidak ditemukan.' });

    const oldRow = rows[rowIndex - 1];
    // PERBAIKAN: id (kolom pertama) dipertahankan apa adanya dari baris lama —
    // id tidak boleh berubah saat update, hanya field lain yang bisa diperbarui.
    const updatedRow = [
      oldRow[headers.indexOf('id')],
      name !== undefined ? name : oldRow[headers.indexOf('name')],
      email,
      password !== undefined ? password : oldRow[headers.indexOf('password')],
      role !== undefined ? role : oldRow[headers.indexOf('role')],
      is_login !== undefined ? is_login : oldRow[headers.indexOf('is_login')]
    ];

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Users!A${rowIndex}:F${rowIndex}`, // PERBAIKAN: A:E -> A:F
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [updatedRow] }
    });

    res.json({ success: true, message: 'User berhasil diperbarui!' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: error.message });
  }
});

// Hapus User (Hanya Developer & Admin)
router.delete('/users', verifyToken, allowDeveloperAndAdmin, async (req, res) => {
  try {
    const { email } = req.body;
    if (!email) return res.status(400).json({ success: false, message: 'Email wajib disertakan untuk menghapus user.' });

    const sheets = await getSheetClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A:F', // PERBAIKAN: A:E -> A:F
    });

    const rows = response.data.values;
    if (!rows || rows.length <= 1) return res.status(404).json({ success: false, message: 'Data user kosong.' });

    const headers = rows[0].map(h => h.trim().toLowerCase());
    let rowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      const dbEmail = rows[i][headers.indexOf('email')] || '';
      if (dbEmail.toLowerCase() === email.toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return res.status(404).json({ success: false, message: 'User tidak ditemukan.' });

    const sheetInfo = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const sheetId = sheetInfo.data.sheets.find(s => s.properties.title === 'Users').properties.sheetId;

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [{
          deleteDimension: {
            range: {
              sheetId: sheetId,
              dimension: 'ROWS',
              startIndex: rowIndex - 1,
              endIndex: rowIndex
            }
          }
        }]
      }
    });

    res.json({ success: true, message: 'User berhasil dihapus!' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: error.message });
  }
});

// ==========================================
// 2. ENDPOINT REAL STOCK
// ==========================================

// Membaca data Real Stock (publik/tanpa proteksi, sesuai perilaku asli)
router.get('/stock', async (req, res) => {
  try {
    const sheets = await getSheetClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Real_Stock!A:H', 
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) return res.json({ success: true, data: [] });

    const headers = rows[0]; 
    const dataRows = rows.slice(1); 

    const formattedData = dataRows.map((row) => {
      let obj = {};
      headers.forEach((header, index) => {
        obj[header.trim()] = row[index] || ''; 
      });
      return obj;
    });

    res.json({ success: true, data: formattedData });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: error.message });
  }
});

// Tambah Stock Barang (Hanya Developer & Admin)
router.post('/stock', verifyToken, allowDeveloperAndAdmin, async (req, res) => {
  try {
    const { No_ID, Nama_Barang, Box, PerPcs, PerDus, Harga, Satuan, Gambar } = req.body;
    const sheets = await getSheetClient();

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Real_Stock!A:H',
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[No_ID, Nama_Barang, Box, PerPcs, PerDus, Harga, Satuan, Gambar]]
      }
    });

    res.json({ success: true, message: 'Stock barang berhasil ditambahkan!' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: error.message });
  }
});

// Update Stock Barang (Hanya Developer & Admin)
router.put('/stock', verifyToken, allowDeveloperAndAdmin, async (req, res) => {
  try {
    const { No_ID, Nama_Barang, Box, PerPcs, PerDus, Harga, Satuan, Gambar } = req.body;
    if (!No_ID) return res.status(400).json({ success: false, message: 'No_ID wajib disertakan untuk update stock.' });

    const sheets = await getSheetClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Real_Stock!A:H',
    });

    const rows = response.data.values;
    if (!rows || rows.length <= 1) return res.status(404).json({ success: false, message: 'Data stock kosong.' });

    const headers = rows[0].map(h => h.trim());
    let rowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      const dbId = rows[i][headers.indexOf('No_ID')] || '';
      if (dbId.toString() === No_ID.toString()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return res.status(404).json({ success: false, message: 'Barang dengan No_ID tersebut tidak ditemukan.' });

    const oldRow = rows[rowIndex - 1];
    const updatedRow = [
      No_ID,
      Nama_Barang !== undefined ? Nama_Barang : oldRow[headers.indexOf('Nama_Barang')],
      Box !== undefined ? Box : oldRow[headers.indexOf('Box')],
      PerPcs !== undefined ? PerPcs : oldRow[headers.indexOf('PerPcs')],
      PerDus !== undefined ? PerDus : oldRow[headers.indexOf('PerDus')],
      Harga !== undefined ? Harga : oldRow[headers.indexOf('Harga')],
      Satuan !== undefined ? Satuan : oldRow[headers.indexOf('Satuan')],
      Gambar !== undefined ? Gambar : oldRow[headers.indexOf('Gambar')]
    ];

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Real_Stock!A${rowIndex}:H${rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [updatedRow] }
    });

    res.json({ success: true, message: 'Stock barang berhasil diperbarui!' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: error.message });
  }
});

// Hapus Stock Barang (Hanya Developer & Admin)
router.delete('/stock', verifyToken, allowDeveloperAndAdmin, async (req, res) => {
  try {
    const { No_ID } = req.body;
    if (!No_ID) return res.status(400).json({ success: false, message: 'No_ID wajib disertakan untuk menghapus stock.' });

    const sheets = await getSheetClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Real_Stock!A:H',
    });

    const rows = response.data.values;
    if (!rows || rows.length <= 1) return res.status(404).json({ success: false, message: 'Data stock kosong.' });

    const headers = rows[0].map(h => h.trim());
    let rowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      const dbId = rows[i][headers.indexOf('No_ID')] || '';
      if (dbId.toString() === No_ID.toString()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return res.status(404).json({ success: false, message: 'Barang tidak ditemukan.' });

    const sheetInfo = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const sheetId = sheetInfo.data.sheets.find(s => s.properties.title === 'Real_Stock').properties.sheetId;

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [{
          deleteDimension: {
            range: {
              sheetId: sheetId,
              dimension: 'ROWS',
              startIndex: rowIndex - 1,
              endIndex: rowIndex
            }
          }
        }]
      }
    });

    res.json({ success: true, message: 'Stock barang berhasil dihapus!' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: error.message });
  }
});


// ==========================================
// 3. ENDPOINT KEUANGAN (FINANCIAL)
// ==========================================

// Membaca data Keuangan (publik/tanpa proteksi, sesuai perilaku asli)
router.get('/finance', async (req, res) => {
  try {
    const sheets = await getSheetClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Keuangan!A:H', 
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) return res.json({ success: true, data: [] });

    const headers = rows[0]; 
    const dataRows = rows.slice(1); 

    const formattedData = dataRows.map((row) => {
      let obj = {};
      headers.forEach((header, index) => {
        obj[header.trim()] = row[index] || ''; 
      });
      return obj;
    });

    res.json({ success: true, data: formattedData });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: error.message });
  }
});

// Tambah Data Keuangan (Hanya Developer & Admin)
router.post('/finance', verifyToken, allowDeveloperAndAdmin, async (req, res) => {
  try {
    const { id, tanggal, keterangan, tipe, jumlah } = req.body;
    const sheets = await getSheetClient();

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Keuangan!A:H',
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[id, tanggal, keterangan, tipe, jumlah]]
      }
    });

    res.json({ success: true, message: 'Data keuangan berhasil ditambahkan!' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: error.message });
  }
});

// Update Data Keuangan (Hanya Developer & Admin)
router.put('/finance', verifyToken, allowDeveloperAndAdmin, async (req, res) => {
  try {
    const { id, tanggal, keterangan, tipe, jumlah } = req.body;
    if (!id) return res.status(400).json({ success: false, message: 'ID wajib disertakan untuk update keuangan.' });

    const sheets = await getSheetClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Keuangan!A:H',
    });

    const rows = response.data.values;
    if (!rows || rows.length <= 1) return res.status(404).json({ success: false, message: 'Data keuangan kosong.' });

    const headers = rows[0].map(h => h.trim().toLowerCase());
    let rowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      const dbId = rows[i][0] || ''; 
      if (dbId.toString() === id.toString()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return res.status(404).json({ success: false, message: 'Data keuangan tidak ditemukan.' });

    const oldRow = rows[rowIndex - 1];
    const updatedRow = [
      id,
      tanggal !== undefined ? tanggal : oldRow[headers.indexOf('tanggal')],
      keterangan !== undefined ? keterangan : oldRow[headers.indexOf('keterangan')],
      tipe !== undefined ? tipe : oldRow[headers.indexOf('tipe')],
      jumlah !== undefined ? jumlah : oldRow[headers.indexOf('jumlah')]
    ];

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Keuangan!A${rowIndex}:E${rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [updatedRow] }
    });

    res.json({ success: true, message: 'Data keuangan berhasil diperbarui!' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: error.message });
  }
});

// Hapus Data Keuangan (Hanya Developer & Admin)
router.delete('/finance', verifyToken, allowDeveloperAndAdmin, async (req, res) => {
  try {
    const { id } = req.body;
    if (!id) return res.status(400).json({ success: false, message: 'ID wajib disertakan untuk menghapus data keuangan.' });

    const sheets = await getSheetClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Keuangan!A:H',
    });

    const rows = response.data.values;
    if (!rows || rows.length <= 1) return res.status(404).json({ success: false, message: 'Data keuangan kosong.' });

    let rowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      const dbId = rows[i][0] || ''; 
      if (dbId.toString() === id.toString()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return res.status(404).json({ success: false, message: 'Data keuangan tidak ditemukan.' });

    const sheetInfo = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const sheetId = sheetInfo.data.sheets.find(s => s.properties.title === 'Keuangan').properties.sheetId;

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [{
          deleteDimension: {
            range: {
              sheetId: sheetId,
              dimension: 'ROWS',
              startIndex: rowIndex - 1,
              endIndex: rowIndex
            }
          }
        }]
      }
    });

    res.json({ success: true, message: 'Data keuangan berhasil dihapus!' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: error.message });
  }
});

module.exports = router;