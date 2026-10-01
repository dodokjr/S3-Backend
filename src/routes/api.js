const express = require('express');
const router = express.Router();
const { getSheetClient, SPREADSHEET_ID } = require('../config/googleSheets');

// verifyToken, allowDeveloperAndAdmin, dan checkModuleAccess ada di satu file
// middleware yang sama: routes/middleware/auth.js
const { verifyToken, allowDeveloperAndAdmin, checkModuleAccess } = require('../routes/middleware/auth');

// ==========================================
// HELPER SINKRON STOCK <-> SALES
// Sheet Real_Stock: A:J
// A No_ID | B Nama_Barang | C Box | D PerPcs | E PerDus | F Harga | G Satuan | H Gambar | I Stok_Awal | J Stok_Akhir
// ==========================================
const STOCK_RANGE = 'Real_Stock!A:J';
const COL_NAMA = 1;   // B: Nama_Barang
const COL_AWAL = 8;   // I: Stok_Awal
const COL_AKHIR = 9;  // J: Stok_Akhir

const norm = (v) => (v || '').toString().trim().toLowerCase();
const toNum = (v) => {
  const n = Number(String(v ?? '').replace(',', '.'));
  return Number.isNaN(n) ? 0 : n;
};
// null = stok belum dilacak (Stok_Akhir kosong)
const getAkhir = (row) => {
  const v = row[COL_AKHIR];
  return v === undefined || v === '' ? null : toNum(v);
};
const httpError = (status, message) => Object.assign(new Error(message), { status });

async function loadStockRows(sheets) {
  const r = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: STOCK_RANGE });
  return r.data.values || [];
}

function findStockIdx(rows, nama) {
  for (let i = 1; i < rows.length; i++) {
    if (norm(rows[i][COL_NAMA]) === norm(nama)) return i;
  }
  return -1;
}

// Hitung perubahan stok TANPA menulis (divalidasi dulu sebelum data sales disimpan).
// oldItem/oldQty = barang yang dikembalikan, newItem/newQty = barang yang diambil.
async function planStockChange(sheets, { oldItem, oldQty = 0, newItem, newQty = 0 }) {
  const rows = await loadStockRows(sheets);
  const pending = new Map(); // index baris -> Stok_Akhir baru
  const current = (i) => (pending.has(i) ? pending.get(i) : getAkhir(rows[i]));

  if (oldItem && oldQty > 0) {
    const i = findStockIdx(rows, oldItem);
    if (i !== -1 && getAkhir(rows[i]) !== null) pending.set(i, current(i) + oldQty);
  }

  if (newItem && newQty > 0) {
    const i = findStockIdx(rows, newItem);
    if (i === -1) throw httpError(404, `Barang "${newItem}" tidak ditemukan di stock.`);
    const cur = current(i);
    if (cur === null) throw httpError(400, `Stok barang "${newItem}" belum diisi (Stok_Awal kosong).`);
    if (cur <= 0) throw httpError(400, `Stok barang "${newItem}" sudah habis.`);
    if (cur < newQty) {
      throw httpError(400, `Stok "${newItem}" tidak cukup. Tersedia: ${cur}, diminta: ${newQty}.`);
    }
    pending.set(i, cur - newQty);
  }
  return pending;
}

async function commitStockChange(sheets, pending) {
  if (!pending || pending.size === 0) return;
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: [...pending].map(([i, val]) => ({
        range: `Real_Stock!J${i + 1}`,
        values: [[val]],
      })),
    },
  });
}

// Ambil SEMUA data stock. Barang yang habis (Stok_Akhir <= 0) atau Box/PerPcs/PerDus = 0 tetap dikirim,
// tapi diberi flag hidden: true supaya frontend menyembunyikannya.
async function readStock() {
  const sheets = await getSheetClient();
  const rows = await loadStockRows(sheets);
  if (rows.length === 0) return [];
  const headers = rows[0].map((h) => h.trim());
  return rows.slice(1).map((row) => {
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = row[idx] || ''; });
    // Aturan hidden:
    // - Jika kolom Stok_Akhir ADA dan terisi: hidden = Stok_Akhir <= 0 (>= 1 tampil).
    // - Jika Stok_Akhir tidak ada/kosong: hidden hanya jika Box, PerPcs, dan PerDus semuanya <= 0.
    const hasAkhir = obj.Stok_Akhir !== undefined && String(obj.Stok_Akhir).trim() !== '';
    if (hasAkhir) {
      obj.hidden = toNum(obj.Stok_Akhir) <= 0;
    } else {
      obj.hidden = toNum(obj.Box) <= 0 && toNum(obj.PerPcs) <= 0 && toNum(obj.PerDus) <= 0;
    }
    return obj;
  });
}

// Total jumlah pengambilan di baris Sales (Pcs + Pack + Kilogram)
function salesRowQty(row, headers) {
  const idx = (names) => {
    for (const n of names) {
      const i = headers.indexOf(n);
      if (i !== -1) return i;
    }
    return -1;
  };
  const get = (names) => {
    const i = idx(names);
    return i === -1 ? 0 : toNum(row[i]);
  };
  return get(['pcs']) + get(['pack']) + get(['kilogram', 'kg']);
}

// ==========================================
// 1. ENDPOINT USERS
// ==========================================

// Membaca data Users (butuh login).
// Role 'sales', 'finance', dan 'gudang' tetap boleh GET /users, tapi hasilnya difilter
// supaya hanya melihat data profil dirinya sendiri (match by email).
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

    let formattedData = dataRows.map((row) => {
      let obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index] || '';
      });
      delete obj.password;
      return obj;
    });

    const requesterRole = (req.user?.role || '').toString().toLowerCase();
    if (requesterRole === 'sales' || requesterRole === 'finance' || requesterRole === 'gudang') {
      const requesterEmail = (req.user?.email || '').toString().toLowerCase();
      formattedData = formattedData.filter(
        (u) => (u.email || '').toString().toLowerCase() === requesterEmail
      );
    }

    res.json({ success: true, data: formattedData });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: error.message });
  }
});

// Tambah User Baru (Hanya Developer & Admin)
router.post('/users', verifyToken, allowDeveloperAndAdmin, async (req, res) => {
  try {
    const { name, email, password, role, is_login } = req.body;

    if (!name || !email) {
      return res.status(400).json({ success: false, message: 'Nama dan Email wajib diisi!' });
    }

    const sheets = await getSheetClient();

    const existingResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A:F',
    });

    const rows = existingResponse.data.values || [];
    const dataRows = rows.slice(1);

    let maxId = 0;
    dataRows.forEach((row) => {
      const idNumber = Number(row[0]);
      if (!Number.isNaN(idNumber) && idNumber > maxId) {
        maxId = idNumber;
      }
    });
    const newId = (maxId + 1).toString();

    const allowedRoles = ['admin', 'developer', 'karyawan', 'sales', 'finance', 'gudang'];
    const finalRole = allowedRoles.includes((role || '').toLowerCase()) ? role : 'karyawan';

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A:F',
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[newId, name, email, password || '', finalRole, is_login || false]]
      }
    });

    res.json({
      success: true,
      message: 'User berhasil ditambahkan!',
      data: { id: newId, name, email, role: finalRole, status: is_login ? 'TRUE' : 'FALSE' }
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
      range: 'Users!A:F',
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
    const allowedRoles = ['admin', 'developer', 'karyawan', 'sales', 'finance', 'gudang'];
    const finalRole = role !== undefined
      ? (allowedRoles.includes((role || '').toLowerCase()) ? role : oldRow[headers.indexOf('role')])
      : oldRow[headers.indexOf('role')];

    const updatedRow = [
      oldRow[headers.indexOf('id')],
      name !== undefined ? name : oldRow[headers.indexOf('name')],
      email,
      password !== undefined ? password : oldRow[headers.indexOf('password')],
      finalRole,
      is_login !== undefined ? is_login : oldRow[headers.indexOf('is_login')]
    ];

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Users!A${rowIndex}:F${rowIndex}`,
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
      range: 'Users!A:F',
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
// GET /stock: terbuka (tanpa token), mengirim SEMUA data; barang habis diberi hidden: true
// POST/PUT/DELETE: Gudang, Sales, Admin, Developer
// ==========================================

// GET /stock: semua data dari database. Barang habis ditandai hidden: true.
// Opsional: /stock?hidden=false -> hanya barang yang masih ada (yang tampil)
router.get('/stock', async (req, res) => {
  try {
    let data = await readStock();
    if (req.query.hidden === 'false') data = data.filter((item) => !item.hidden);
    res.json({ success: true, data });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: error.message });
  }
});

router.post('/stock', verifyToken, checkModuleAccess('gudang'), async (req, res) => {
  try {
    const { No_ID, Nama_Barang, Box, PerPcs, PerDus, Harga, Satuan, Gambar, Stok_Awal } = req.body;
    if (!No_ID || !Nama_Barang) {
      return res.status(400).json({ success: false, message: 'No_ID dan Nama_Barang wajib diisi!' });
    }
    const awal = toNum(Stok_Awal);
    const sheets = await getSheetClient();

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: STOCK_RANGE,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[No_ID, Nama_Barang, Box, PerPcs, PerDus, Harga, Satuan, Gambar, awal, awal]]
      }
    });

    res.json({ success: true, message: 'Stock barang berhasil ditambahkan!' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: error.message });
  }
});

router.put('/stock', verifyToken, checkModuleAccess('gudang'), async (req, res) => {
  try {
    const { No_ID, Nama_Barang, Box, PerPcs, PerDus, Harga, Satuan, Gambar, Stok_Awal } = req.body;
    if (!No_ID) return res.status(400).json({ success: false, message: 'No_ID wajib disertakan untuk update stock.' });

    const sheets = await getSheetClient();
    const rows = await loadStockRows(sheets);
    if (rows.length <= 1) return res.status(404).json({ success: false, message: 'Data stock kosong.' });

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

    // Jika Stok_Awal diubah (restock), Stok_Akhir ikut bergeser sebesar selisihnya,
    // jadi riwayat pengambilan sales tetap terhitung.
    let newAwal = toNum(oldRow[COL_AWAL]);
    let newAkhir = getAkhir(oldRow) ?? newAwal;
    if (Stok_Awal !== undefined) {
      const oldAwal = toNum(oldRow[COL_AWAL]);
      newAwal = toNum(Stok_Awal);
      newAkhir = (getAkhir(oldRow) ?? oldAwal) + (newAwal - oldAwal);
      if (newAkhir < 0) {
        return res.status(400).json({
          success: false,
          message: 'Stok_Awal baru lebih kecil dari jumlah yang sudah diambil sales.',
        });
      }
    }

    const updatedRow = [
      No_ID,
      Nama_Barang !== undefined ? Nama_Barang : oldRow[headers.indexOf('Nama_Barang')],
      Box !== undefined ? Box : oldRow[headers.indexOf('Box')],
      PerPcs !== undefined ? PerPcs : oldRow[headers.indexOf('PerPcs')],
      PerDus !== undefined ? PerDus : oldRow[headers.indexOf('PerDus')],
      Harga !== undefined ? Harga : oldRow[headers.indexOf('Harga')],
      Satuan !== undefined ? Satuan : oldRow[headers.indexOf('Satuan')],
      Gambar !== undefined ? Gambar : oldRow[headers.indexOf('Gambar')],
      newAwal,
      newAkhir,
    ];

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Real_Stock!A${rowIndex}:J${rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [updatedRow] }
    });

    res.json({ success: true, message: 'Stock barang berhasil diperbarui!' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: error.message });
  }
});

router.delete('/stock', verifyToken, checkModuleAccess('gudang'), async (req, res) => {
  try {
    const { No_ID } = req.body;
    if (!No_ID) return res.status(400).json({ success: false, message: 'No_ID wajib disertakan untuk menghapus stock.' });

    const sheets = await getSheetClient();
    const rows = await loadStockRows(sheets);
    if (rows.length <= 1) return res.status(404).json({ success: false, message: 'Data stock kosong.' });

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
// GET: terbuka (tanpa token)
// POST/PUT/DELETE: Finance, Developer, Admin
// ==========================================

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

router.post('/finance', verifyToken, checkModuleAccess('finance'), async (req, res) => {
  try {
    const { deskripsi, pemasukan, pengeluaran, tgl, bulan, tahun } = req.body;

    if (!deskripsi || !tgl || !bulan || !tahun) {
      return res.status(400).json({
        success: false,
        message: 'Deskripsi, tanggal, bulan, dan tahun wajib diisi!',
      });
    }

    const sheets = await getSheetClient();

    const existingResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Keuangan!A:H',
    });

    const rows = existingResponse.data.values || [];
    const dataRows = rows.slice(1);

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
      range: 'Keuangan!A:G',
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[
          newId,
          deskripsi,
          pengeluaran || '0',
          pemasukan || '0',
          tgl,
          bulan,
          tahun,
        ]],
      },
    });

    res.json({ success: true, message: 'Data keuangan berhasil ditambahkan!', data: { id: newId } });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: error.message });
  }
});

router.put('/finance', verifyToken, checkModuleAccess('finance'), async (req, res) => {
  try {
    const { id, deskripsi, pemasukan, pengeluaran, tgl, bulan, tahun } = req.body;
    if (!id) return res.status(400).json({ success: false, message: 'ID wajib disertakan untuk update keuangan.' });

    const sheets = await getSheetClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Keuangan!A:H',
    });

    const rows = response.data.values;
    if (!rows || rows.length <= 1) return res.status(404).json({ success: false, message: 'Data keuangan kosong.' });

    const headers = rows[0].map(h => h.trim().toLowerCase());

    // Kolom ID: coba 'no_id', lalu 'id', fallback ke kolom pertama (A).
    let idColIdx = headers.indexOf('no_id');
    if (idColIdx === -1) idColIdx = headers.indexOf('id');
    if (idColIdx === -1) idColIdx = 0;

    let rowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      const dbId = rows[i][idColIdx] || '';
      if (dbId.toString() === id.toString()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return res.status(404).json({ success: false, message: 'Data keuangan tidak ditemukan.' });

    const oldRow = rows[rowIndex - 1];
    const updatedRow = [
      id,
      deskripsi !== undefined ? deskripsi : oldRow[headers.indexOf('deskripsi')],
      pengeluaran !== undefined ? pengeluaran : oldRow[headers.indexOf('pengeluaran')],
      pemasukan !== undefined ? pemasukan : oldRow[headers.indexOf('pemasukan')],
      tgl !== undefined ? tgl : oldRow[headers.indexOf('tgl')],
      bulan !== undefined ? bulan : oldRow[headers.indexOf('bulan')],
      tahun !== undefined ? tahun : oldRow[headers.indexOf('tahun')],
    ];

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Keuangan!A${rowIndex}:G${rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [updatedRow] },
    });

    res.json({ success: true, message: 'Data keuangan berhasil diperbarui!' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: error.message });
  }
});

router.delete('/finance', verifyToken, checkModuleAccess('finance'), async (req, res) => {
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

// ==========================================
// 4. ENDPOINT SALES (SINKRON DENGAN STOCK)
// Akses GET/POST/PUT/DELETE: Sales, Developer, Admin
// - POST   : cek stok, simpan sales, kurangi Stok_Akhir
// - PUT    : kembalikan stok lama, potong stok baru
// - DELETE : kembalikan stok
// ==========================================

// GET Data Sales
router.get('/sales', verifyToken, checkModuleAccess('sales'), async (req, res) => {
  try {
    const sheets = await getSheetClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Sales!A:K',
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

// Tambah Data Sales (POST)
router.post('/sales', verifyToken, checkModuleAccess('sales'), async (req, res) => {
  try {
    const { Deskripsi, Costumer, tgl, status, satuan, jumlah, nama_seles } = req.body;
    const hargaJual = req.body['harga jual'];
    const hargaBeli = req.body['harga beli'];

    // Validasi field utama
    if (!Deskripsi || hargaJual === undefined || hargaBeli === undefined || !tgl || !nama_seles || !satuan || jumlah === undefined) {
      return res.status(400).json({
        success: false,
        message: 'Deskripsi, harga jual, harga beli, tgl, nama_seles, satuan, dan jumlah wajib diisi!',
      });
    }

    // Logika penentuan Pcs, Pack, dan Kilogram berdasarkan "satuan" dan "jumlah"
    let finalPcs = '0';
    let finalPack = '0';
    let finalKilogram = '0';

    const normalizedSatuan = satuan.trim().toLowerCase();

    if (normalizedSatuan === 'pcs') {
      finalPcs = jumlah;
    } else if (normalizedSatuan === 'pack') {
      finalPack = jumlah;
    } else if (normalizedSatuan === 'kilogram' || normalizedSatuan === 'kg') {
      finalKilogram = jumlah;
    } else {
      return res.status(400).json({
        success: false,
        message: 'Satuan tidak valid! Gunakan Pcs, Pack, atau Kilogram.',
      });
    }

    const qty = toNum(jumlah);
    if (qty <= 0) {
      return res.status(400).json({ success: false, message: 'Jumlah harus lebih dari 0.' });
    }

    const sheets = await getSheetClient();

    // 1) Validasi stok dulu (belum menulis apa pun)
    const stockPlan = await planStockChange(sheets, { newItem: Deskripsi, newQty: qty });

    const existingResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Sales!A:K',
    });

    const rows = existingResponse.data.values || [];
    const dataRows = rows.slice(1);

    let maxId = 0;
    dataRows.forEach((row) => {
      const idNumber = Number(row[0]);
      if (!Number.isNaN(idNumber) && idNumber > maxId) {
        maxId = idNumber;
      }
    });
    const newId = (maxId + 1).toString();
    const finalStatus = status || 'success';

    // Urutan kolom di Google Sheets:
    // id | Deskripsi | Costumer | harga jual | harga beli | tgl | status | nama_seles | Pcs | Pack | Kilogram
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Sales!A:K',
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[
          newId,
          Deskripsi,
          Costumer || '',
          hargaJual,
          hargaBeli,
          tgl,
          finalStatus,
          nama_seles,
          finalPcs,
          finalPack,
          finalKilogram
        ]]
      }
    });

    // 2) Setelah sales tersimpan, kurangi Stok_Akhir
    await commitStockChange(sheets, stockPlan);

    res.json({
      success: true,
      message: 'Data sales berhasil ditambahkan!',
      data: {
        id: newId,
        Deskripsi,
        Costumer: Costumer || '',
        'harga jual': hargaJual,
        'harga beli': hargaBeli,
        tgl,
        status: finalStatus,
        nama_seles,
        Pcs: finalPcs,
        Pack: finalPack,
        Kilogram: finalKilogram
      }
    });
  } catch (error) {
    console.error(error);
    res.status(error.status || 500).json({ success: false, message: error.message });
  }
});

// Update Data Sales (PUT)
router.put('/sales', verifyToken, checkModuleAccess('sales'), async (req, res) => {
  try {
    const { id, Deskripsi, Costumer, tgl, status, satuan, jumlah, nama_seles } = req.body;
    const hargaJual = req.body['harga jual'];
    const hargaBeli = req.body['harga beli'];

    if (!id) return res.status(400).json({ success: false, message: 'id wajib disertakan untuk update sales.' });

    const sheets = await getSheetClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Sales!A:K',
    });

    const rows = response.data.values;
    if (!rows || rows.length <= 1) return res.status(404).json({ success: false, message: 'Data sales kosong.' });

    const headers = rows[0].map(h => h.trim().toLowerCase());
    let rowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      const dbId = rows[i][headers.indexOf('id')] || '';
      if (dbId.toString() === id.toString()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return res.status(404).json({ success: false, message: 'Data sales dengan id tersebut tidak ditemukan.' });

    const oldRow = rows[rowIndex - 1];
    const getColIdx = (names) => {
      for (let name of names) {
        const idx = headers.indexOf(name.toLowerCase());
        if (idx !== -1) return idx;
      }
      return -1;
    };

    // Logika Update Satuan berdasarkan input satuan & jumlah baru
    let finalPcs, finalPack, finalKilogram;

    if (satuan !== undefined && jumlah !== undefined) {
      const normalizedSatuan = satuan.trim().toLowerCase();
      if (normalizedSatuan === 'pcs') {
        finalPcs = jumlah;
        finalPack = '0';
        finalKilogram = '0';
      } else if (normalizedSatuan === 'pack') {
        finalPcs = '0';
        finalPack = jumlah;
        finalKilogram = '0';
      } else if (normalizedSatuan === 'kilogram' || normalizedSatuan === 'kg') {
        finalPcs = '0';
        finalPack = '0';
        finalKilogram = jumlah;
      } else {
        finalPcs = oldRow[getColIdx(['pcs'])];
        finalPack = oldRow[getColIdx(['pack'])];
        finalKilogram = oldRow[getColIdx(['kilogram', 'kg'])];
      }
    } else {
      // Jika satuan/jumlah tidak dikirim saat update, gunakan nilai lama di database
      finalPcs = oldRow[getColIdx(['pcs'])];
      finalPack = oldRow[getColIdx(['pack'])];
      finalKilogram = oldRow[getColIdx(['kilogram', 'kg'])];
    }

    const updatedRow = [
      id,
      Deskripsi !== undefined ? Deskripsi : oldRow[getColIdx(['deskripsi'])],
      Costumer !== undefined ? Costumer : oldRow[getColIdx(['costumer', 'customer'])],
      hargaJual !== undefined ? hargaJual : oldRow[getColIdx(['harga jual', 'hargajual'])],
      hargaBeli !== undefined ? hargaBeli : oldRow[getColIdx(['harga beli', 'hargabeli'])],
      tgl !== undefined ? tgl : oldRow[getColIdx(['tgl'])],
      status !== undefined ? status : oldRow[getColIdx(['status'])],
      nama_seles !== undefined ? nama_seles : oldRow[getColIdx(['nama_seles', 'name_seles', 'namaseles'])],
      finalPcs,
      finalPack,
      finalKilogram,
    ];

    // Sinkron stok: kembalikan pengambilan lama, potong pengambilan baru
    const oldItem = oldRow[getColIdx(['deskripsi'])];
    const oldQty = salesRowQty(oldRow, headers);
    const newItem = updatedRow[1];
    const newQty = toNum(updatedRow[8]) + toNum(updatedRow[9]) + toNum(updatedRow[10]);

    // Validasi stok dulu sebelum menulis
    const stockPlan = await planStockChange(sheets, { oldItem, oldQty, newItem, newQty });

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Sales!A${rowIndex}:K${rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [updatedRow] }
    });

    await commitStockChange(sheets, stockPlan);

    res.json({ success: true, message: 'Data sales berhasil diperbarui!' });
  } catch (error) {
    console.error(error);
    res.status(error.status || 500).json({ success: false, message: error.message });
  }
});

// Hapus Data Sales (DELETE)
router.delete('/sales', verifyToken, checkModuleAccess('sales'), async (req, res) => {
  try {
    const { id } = req.body;
    if (!id) return res.status(400).json({ success: false, message: 'id wajib disertakan untuk menghapus data sales.' });

    const sheets = await getSheetClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Sales!A:K',
    });

    const rows = response.data.values;
    if (!rows || rows.length <= 1) return res.status(404).json({ success: false, message: 'Data sales kosong.' });

    const headers = rows[0].map(h => h.trim().toLowerCase());
    let rowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      const dbId = rows[i][headers.indexOf('id')] || '';
      if (dbId.toString() === id.toString()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return res.status(404).json({ success: false, message: 'Data sales tidak ditemukan.' });

    // Kembalikan stok barang yang pernah diambil
    const oldRow = rows[rowIndex - 1];
    const oldQty = salesRowQty(oldRow, headers);
    const oldItem = oldRow[headers.indexOf('deskripsi')];
    const stockPlan = await planStockChange(sheets, { oldItem, oldQty });

    const sheetInfo = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const sheetId = sheetInfo.data.sheets.find(s => s.properties.title === 'Sales').properties.sheetId;

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

    await commitStockChange(sheets, stockPlan);

    res.json({ success: true, message: 'Data sales berhasil dihapus!' });
  } catch (error) {
    console.error(error);
    res.status(error.status || 500).json({ success: false, message: error.message });
  }
});

module.exports = router;