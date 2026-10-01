const { getSheetClient, SPREADSHEET_ID } = require('../../config/googleSheets');
const { badRequest, notFound } = require('../Utils/Errors');

exports.getFinance = async (req, res, next) => {
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
};

exports.createFinance = async (req, res, next) => {
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
};

exports.updateFinance = async (req, res, next) => {
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
};

exports.deleteFinance = async (req, res, next) => {
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
};