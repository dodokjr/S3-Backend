const { getSheetClient, SPREADSHEET_ID } = require('../../config/googleSheets');
const { badRequest, notFound } = require('../Utils/Errors');
const {
  toNum,
  planStockChange,
  commitStockChange,
  salesRowQty,
} = require('../helpers/Stockhelper');

// GET Data Sales
exports.getSales = async (req, res, next) => {
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
};

// Tambah Data Sales (POST): cek stok, simpan sales, kurangi PerPcs
exports.createSales = async (req, res, next) => {
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
};

// Update Data Sales (PUT): kembalikan stok lama, potong stok baru
exports.updateSales = async (req, res, next) => {
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
};

// Hapus Data Sales (DELETE): kembalikan stok
exports.deleteSales = async (req, res, next) => {
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
};