const { getSheetClient, SPREADSHEET_ID } = require('../../config/googleSheets');
const { badRequest, notFound } = require('../Utils/Errors');

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
    if (i === -1) throw notFound(`Barang "${newItem}" tidak ditemukan di stock.`);
    const cur = current(i);
    if (cur <= 0) throw badRequest(`Stok barang "${newItem}" sudah habis.`);
    if (cur < newQty) {
      throw badRequest(`Stok "${newItem}" tidak cukup. Tersedia: ${cur}, diminta: ${newQty}.`);
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

// Barang diketik manual: TANPA cek stok, langsung masuk ke Real_Stock.
// - Jika nama barang sudah ada di stok -> PerPcs ditambah qty.
// - Jika belum ada -> baris baru ditambahkan (PerPcs = qty).
async function addManualStock(sheets, { nama, qty, satuan = '', harga = '' }) {
  const item = (nama || '').toString().trim();
  const q = toNum(qty);
  if (!item || q <= 0) return;

  const rows = await loadStockRows(sheets);
  const cPcs = colIdx(rows, 'perpcs', 3);
  const i = findStockIdx(rows, item);

  if (i !== -1) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Real_Stock!${colLetter(cPcs)}${i + 1}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[toNum(rows[i][cPcs]) + q]] },
    });
    return;
  }

  // No_ID baru = max No_ID + 1
  const cId = colIdx(rows, 'no_id', 0);
  let maxId = 0;
  for (let r = 1; r < rows.length; r++) {
    const n = Number(rows[r][cId]);
    if (!Number.isNaN(n) && n > maxId) maxId = n;
  }

  // A No_ID | B Nama_Barang | C Box | D PerPcs | E PerDus | F Harga | G Satuan | H Gambar
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: STOCK_RANGE,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[(maxId + 1).toString(), item, '', q, '', harga, satuan, '']],
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

module.exports = {
  STOCK_RANGE,
  norm,
  toNum,
  loadStockRows,
  planStockChange,
  commitStockChange,
  addManualStock,
  readStock,
  salesRowQty,
};