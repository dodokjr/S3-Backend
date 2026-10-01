const { getSheetClient, SPREADSHEET_ID } = require('../../config/googleSheets');
const { badRequest, notFound } = require('../Utils/Errors');

const ALLOWED_ROLES = ['admin', 'developer', 'karyawan', 'sales', 'finance', 'gudang'];

// Membaca data Users (butuh login).
// Role 'sales', 'finance', dan 'gudang' tetap boleh GET /users, tapi hasilnya difilter
// supaya hanya melihat data profil dirinya sendiri (match by email).
exports.getUsers = async (req, res, next) => {
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
};

// Tambah User Baru (Hanya Developer & Admin)
exports.createUser = async (req, res, next) => {
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

    const finalRole = ALLOWED_ROLES.includes((role || '').toLowerCase()) ? role : 'karyawan';

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
};

// Update User (Hanya Developer & Admin)
exports.updateUser = async (req, res, next) => {
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
    const finalRole = role !== undefined
      ? (ALLOWED_ROLES.includes((role || '').toLowerCase()) ? role : oldRow[headers.indexOf('role')])
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
};

// Hapus User (Hanya Developer & Admin)
exports.deleteUser = async (req, res, next) => {
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
};