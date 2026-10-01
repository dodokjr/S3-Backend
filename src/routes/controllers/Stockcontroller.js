const { getSheetClient, SPREADSHEET_ID } = require('../../config/googleSheets');
const { badRequest, notFound } = require('../Utils/Errors');
const { STOCK_RANGE, loadStockRows, readStock } = require('../helpers/Stockhelper');

// GET /stock: semua data dari database. PerPcs >= 1 -> hidden false, PerPcs <= 0 -> hidden true.
// Opsional: /stock?hidden=false -> hanya barang yang masih ada (yang tampil)
exports.getStock = async (req, res, next) => {
  try {
    let data = await readStock();
    if (req.query.hidden === 'false') data = data.filter((item) => !item.hidden);
    res.json({ success: true, data });
  } catch (error) {
    next(error);
  }
};

exports.createStock = async (req, res, next) => {
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
};

exports.updateStock = async (req, res, next) => {
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
};

exports.deleteStock = async (req, res, next) => {
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
};