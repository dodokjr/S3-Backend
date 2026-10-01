const express = require('express');
const router = express.Router();
const { getSheetClient, SPREADSHEET_ID } = require('../config/googleSheets');

// verifyToken, allowDeveloperAndAdmin, dan checkModuleAccess ada di satu file
// middleware yang sama: routes/middleware/auth.js
const { verifyToken, allowDeveloperAndAdmin, checkModuleAccess } = require('../routes/middleware/auth');

// Error HTTP terpusat (400, 404, 500, 503, dst) - satu file dipakai semua route
const { badRequest, notFound } = require('./Utils/Errors');

// ==========================================
// HELPER SINKRON STOCK <-> SALES
// Sheet Real_Stock: A:H
// A No_ID | B Nama_Barang | C Box | D PerPcs | E PerDus | F Harga | G Satuan | H Gambar
// Sisa stok = kolom PerPcs. Setiap sales mengambil barang, PerPcs dikurangi.
// ==========================================
const STOCK_RANGE = 'Real_Stock!A:H';

const norm = (v) => (v || '').toString().trim().toLowerCase();
const toNum = (v) => {
  const n = Number(String(v ?? '').replace(',', '.'));
  return Number.isNaN(n) ? 0 : n;
};

// Cari index kolom berdasarkan nama header (fallback ke posisi default)
const colIdx = (rows, name, fallback) => {
  const i = (rows[0] || []).map((h) => norm(h)).indexOf(name.toLowerCase());
  return i === -1 ? fallback : i;
};
const colLetter = (i) => String.fromCharCode(65 + i);

async function loadStockRows(sheets) {
  const r = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: STOCK_RANGE });
  return r.data.values || [];
}

function findStockIdx(rows, nama) {
  const c = colIdx(rows, 'nama_barang', 1);
  for (let i = 1; i < rows.length; i++) {
    if (norm(rows[i][c]) === norm(nama)) return i;
  }
  return -1;
}

// Hitung perubahan PerPcs TANPA menulis (divalidasi dulu sebelum data sales disimpan).
// oldItem/oldQty = barang yang dikembalikan, newItem/newQty = barang yang diambil.
async function planStockChange(sheets, { oldItem, oldQty = 0, newItem, newQty = 0 }) {
  const rows = await loadStockRows(sheets);
  const cPcs = colIdx(rows, 'perpcs', 3);
  const pending = new Map(); // index baris -> PerPcs baru
  const current = (i) => (pending.has(i) ? pending.get(i) : toNum(rows[i][cPcs]));

  if (oldItem && oldQty > 0) {
    const i = findStockIdx(rows, oldItem);
    if (i !== -1) pending.set(i, current(i) + oldQty);
  }

  if (newItem && newQty > 0) {
    const i = findStockIdx(rows, newItem);
    if (i === -1) throw notFound( `Barang "${newItem}" tidak ditemukan di stock.`);
    const cur = current(i);
    if (cur <= 0) throw badRequest( `Stok barang "${newItem}" sudah habis.`);
    if (cur < newQty) {
      throw badRequest( `Stok "${newItem}" tidak cukup. Tersedia: ${cur}, diminta: ${newQty}.`);
    }
    pending.set(i, cur - newQty);
  }
  return { pending, cPcs };
}

async function commitStockChange(sheets, plan) {
  if (!plan || plan.pending.size === 0) return;
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: [...plan.pending].map(([i, val]) => ({
        range: `Real_Stock!${colLetter(plan.cPcs)}${i + 1}`,
        values: [[val]],
      })),
    },
  });
}

// Ambil SEMUA data stock. hidden = true jika PerPcs <= 0, false jika PerPcs >= 1.
async function readStock() {
  const sheets = await getSheetClient();
  const rows = await loadStockRows(sheets);
  if (rows.length === 0) return [];
  const headers = rows[0].map((h) => h.trim());
  return rows.slice(1).map((row) => {
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = row[idx] || ''; });
    obj.hidden = toNum(obj.PerPcs) <= 0;
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
router.get('/users', verifyToken, async (req, res, next) => {
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
    next(error);
  }
});

// Tambah User Baru (Hanya Developer & Admin)
router.post('/users', verifyToken, allowDeveloperAndAdmin, async (req, res, next) => {
  try {
    const { name, email, password, role, is_login } = req.body;

    if (!name || !email) {
      throw badRequest('Nama dan Email wajib diisi!');
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
    next(error);
  }
});

// Update User (Hanya Developer & Admin)
router.put('/users', verifyToken, allowDeveloperAndAdmin, async (req, res, next) => {
  try {
    const { email, name, password, role, is_login } = req.body;
    if (!email) throw badRequest('Email wajib disertakan untuk update user.');

    const sheets = await getSheetClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A:F',
    });

    const rows = response.data.values;
    if (!rows || rows.length <= 1) throw notFound('Data user kosong.');

    const headers = rows[0].map(h => h.trim().toLowerCase());
    let rowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      const dbEmail = rows[i][headers.indexOf('email')] || '';
      if (dbEmail.toLowerCase() === email.toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) throw notFound('User tidak ditemukan.');

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
    next(error);
  }
});

// Hapus User (Hanya Developer & Admin)
router.delete('/users', verifyToken, allowDeveloperAndAdmin, async (req, res, next) => {
  try {
    const { email } = req.body;
    if (!email) throw badRequest('Email wajib disertakan untuk menghapus user.');

    const sheets = await getSheetClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A:F',
    });

    const rows = response.data.values;
    if (!rows || rows.length <= 1) throw notFound('Data user kosong.');

    const headers = rows[0].map(h => h.trim().toLowerCase());
    let rowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      const dbEmail = rows[i][headers.indexOf('email')] || '';
      if (dbEmail.toLowerCase() === email.toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) throw notFound('User tidak ditemukan.');

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
    next(error);
  }
});

// ==========================================
// 2. ENDPOINT REAL STOCK
// GET /stock: terbuka (tanpa token), mengirim SEMUA data; hidden = true jika PerPcs <= 0
// POST/PUT/DELETE: Gudang, Sales, Admin, Developer
// ==========================================

// GET /stock: semua data dari database. PerPcs >= 1 -> hidden false, PerPcs <= 0 -> hidden true.
// Opsional: /stock?hidden=false -> hanya barang yang masih ada (yang tampil)
router.get('/stock', async (req, res, next) => {
  try {
    let data = await readStock();
    if (req.query.hidden === 'false') data = data.filter((item) => !item.hidden);
    res.json({ success: true, data });
  } catch (error) {
    next(error);
  }
});

router.post('/stock', verifyToken, checkModuleAccess('gudang'), async (req, res, next) => {
  try {
    const { No_ID, Nama_Barang, Box, PerPcs, PerDus, Harga, Satuan, Gambar } = req.body;
    if (!No_ID || !Nama_Barang) {
      throw badRequest('No_ID dan Nama_Barang wajib diisi!');
    }
    const sheets = await getSheetClient();

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: STOCK_RANGE,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[No_ID, Nama_Barang, Box, PerPcs, PerDus, Harga, Satuan, Gambar]]
      }
    });

    res.json({ success: true, message: 'Stock barang berhasil ditambahkan!' });
  } catch (error) {
    next(error);
  }
});

router.put('/stock', verifyToken, checkModuleAccess('gudang'), async (req, res, next) => {
  try {
    const { No_ID, Nama_Barang, Box, PerPcs, PerDus, Harga, Satuan, Gambar } = req.body;
    if (!No_ID) throw badRequest('No_ID wajib disertakan untuk update stock.');

    const sheets = await getSheetClient();
    const rows = await loadStockRows(sheets);
    if (rows.length <= 1) throw notFound('Data stock kosong.');

    const headers = rows[0].map(h => h.trim());
    let rowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      const dbId = rows[i][headers.indexOf('No_ID')] || '';
      if (dbId.toString() === No_ID.toString()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) throw notFound('Barang dengan No_ID tersebut tidak ditemukan.');

    const oldRow = rows[rowIndex - 1];
    // Mengirim PerPcs = restock / koreksi stok (nilai baru menggantikan nilai lama)
    const updatedRow = [
      No_ID,
      Nama_Barang !== undefined ? Nama_Barang : oldRow[headers.indexOf('Nama_Barang')],
      Box !== undefined ? Box : oldRow[headers.indexOf('Box')],
      PerPcs !== undefined ? PerPcs : oldRow[headers.indexOf('PerPcs')],
      PerDus !== undefined ? PerDus : oldRow[headers.indexOf('PerDus')],
      Harga !== undefined ? Harga : oldRow[headers.indexOf('Harga')],
      Satuan !== undefined ? Satuan : oldRow[headers.indexOf('Satuan')],
      Gambar !== undefined ? Gambar : oldRow[7] // kolom H (Gambar / Full Stock)
    ];

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Real_Stock!A${rowIndex}:H${rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [updatedRow] }
    });

    res.json({ success: true, message: 'Stock barang berhasil diperbarui!' });
  } catch (error) {
    next(error);
  }
});

router.delete('/stock', verifyToken, checkModuleAccess('gudang'), async (req, res, next) => {
  try {
    const { No_ID } = req.body;
    if (!No_ID) throw badRequest('No_ID wajib disertakan untuk menghapus stock.');

    const sheets = await getSheetClient();
    const rows = await loadStockRows(sheets);
    if (rows.length <= 1) throw notFound('Data stock kosong.');

    const headers = rows[0].map(h => h.trim());
    let rowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      const dbId = rows[i][headers.indexOf('No_ID')] || '';
      if (dbId.toString() === No_ID.toString()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) throw notFound('Barang tidak ditemukan.');

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
    next(error);
  }
});

// ==========================================
// 3. ENDPOINT KEUANGAN (FINANCIAL)
// GET: terbuka (tanpa token)
// POST/PUT/DELETE: Finance, Developer, Admin
// ==========================================

router.get('/finance', async (req, res, next) => {
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
    next(error);
  }
});

router.post('/finance', verifyToken, checkModuleAccess('finance'), async (req, res, next) => {
  try {
    const { deskripsi, pemasukan, pengeluaran, tgl, bulan, tahun } = req.body;

    if (!deskripsi || !tgl || !bulan || !tahun) {
      throw badRequest('Deskripsi, tanggal, bulan, dan tahun wajib diisi!');
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
    next(error);
  }
});

router.put('/finance', verifyToken, checkModuleAccess('finance'), async (req, res, next) => {
  try {
    const { id, deskripsi, pemasukan, pengeluaran, tgl, bulan, tahun } = req.body;
    if (!id) throw badRequest('ID wajib disertakan untuk update keuangan.');

    const sheets = await getSheetClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Keuangan!A:H',
    });

    const rows = response.data.values;
    if (!rows || rows.length <= 1) throw notFound('Data keuangan kosong.');

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

    if (rowIndex === -1) throw notFound('Data keuangan tidak ditemukan.');

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
    next(error);
  }
});

router.delete('/finance', verifyToken, checkModuleAccess('finance'), async (req, res, next) => {
  try {
    const { id } = req.body;
    if (!id) throw badRequest('ID wajib disertakan untuk menghapus data keuangan.');

    const sheets = await getSheetClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Keuangan!A:H',
    });

    const rows = response.data.values;
    if (!rows || rows.length <= 1) throw notFound('Data keuangan kosong.');

    let rowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      const dbId = rows[i][0] || '';
      if (dbId.toString() === id.toString()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) throw notFound('Data keuangan tidak ditemukan.');

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
    next(error);
  }
});

// ==========================================
// 4. ENDPOINT SALES (SINKRON DENGAN STOCK)
// Akses GET/POST/PUT/DELETE: Sales, Developer, Admin
// - POST   : cek stok, simpan sales, kurangi PerPcs
// - PUT    : kembalikan stok lama, potong stok baru
// - DELETE : kembalikan stok
// ==========================================

// GET Data Sales
router.get('/sales', verifyToken, checkModuleAccess('sales'), async (req, res, next) => {
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
    next(error);
  }
});

// Tambah Data Sales (POST)
router.post('/sales', verifyToken, checkModuleAccess('sales'), async (req, res, next) => {
  try {
    const { Deskripsi, Costumer, tgl, status, satuan, jumlah, nama_seles } = req.body;
    const hargaJual = req.body['harga jual'];
    const hargaBeli = req.body['harga beli'];

    // Validasi field utama
    if (!Deskripsi || hargaJual === undefined || hargaBeli === undefined || !tgl || !nama_seles || !satuan || jumlah === undefined) {
      throw badRequest('Deskripsi, harga jual, harga beli, tgl, nama_seles, satuan, dan jumlah wajib diisi!');
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
      throw badRequest('Satuan tidak valid! Gunakan Pcs, Pack, atau Kilogram.');
    }

    const qty = toNum(jumlah);
    if (qty <= 0) {
      throw badRequest('Jumlah harus lebih dari 0.');
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

    // 2) Setelah sales tersimpan, kurangi PerPcs
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
    next(error);
  }
});

// Update Data Sales (PUT)
router.put('/sales', verifyToken, checkModuleAccess('sales'), async (req, res, next) => {
  try {
    const { id, Deskripsi, Costumer, tgl, status, satuan, jumlah, nama_seles } = req.body;
    const hargaJual = req.body['harga jual'];
    const hargaBeli = req.body['harga beli'];

    if (!id) throw badRequest('id wajib disertakan untuk update sales.');

    const sheets = await getSheetClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Sales!A:K',
    });

    const rows = response.data.values;
    if (!rows || rows.length <= 1) throw notFound('Data sales kosong.');

    const headers = rows[0].map(h => h.trim().toLowerCase());
    let rowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      const dbId = rows[i][headers.indexOf('id')] || '';
      if (dbId.toString() === id.toString()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) throw notFound('Data sales dengan id tersebut tidak ditemukan.');

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
    next(error);
  }
});

// Hapus Data Sales (DELETE)
router.delete('/sales', verifyToken, checkModuleAccess('sales'), async (req, res, next) => {
  try {
    const { id } = req.body;
    if (!id) throw badRequest('id wajib disertakan untuk menghapus data sales.');

    const sheets = await getSheetClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Sales!A:K',
    });

    const rows = response.data.values;
    if (!rows || rows.length <= 1) throw notFound('Data sales kosong.');

    const headers = rows[0].map(h => h.trim().toLowerCase());
    let rowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      const dbId = rows[i][headers.indexOf('id')] || '';
      if (dbId.toString() === id.toString()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) throw notFound('Data sales tidak ditemukan.');

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
    next(error);
  }
});

module.exports = router;