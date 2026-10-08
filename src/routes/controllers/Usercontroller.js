const { getSheetClient, SPREADSHEET_ID } = require('../../config/googleSheets');
const { badRequest, notFound } = require('../Utils/Errors');

const ALLOWED_ROLES = ['admin', 'developer', 'karyawan', 'sales', 'finance', 'gudang'];
const DEFAULT_PASSWORD = '123';
const DEFAULT_HEADERS = ['id', 'name', 'password', 'email', 'role', 'status'];

// Cari index kolom, mendukung beberapa nama header (status / is_login)
const colIndex = (headers, ...names) => {
  for (const n of names) {
    const i = headers.indexOf(n);
    if (i !== -1) return i;
  }
  return -1;
};

// Normalisasi status menjadi 'TRUE' / 'FALSE' (default FALSE)
const toStatus = (v) =>
  v === true || (v || '').toString().toLowerCase() === 'true' ? 'TRUE' : 'FALSE';

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
    const { name, email, password, role, status, is_login } = req.body;

    if (!name || !email) {
      throw badRequest('Nama dan Email wajib diisi!');
    }

    const sheets = await getSheetClient();
    const existingResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A:F',
    });

    const rows = existingResponse.data.values || [];
    const headers = rows.length > 0
      ? rows[0].map(h => h.trim().toLowerCase())
      : DEFAULT_HEADERS;
    const dataRows = rows.slice(1);

    const idxId = colIndex(headers, 'id');
    const idxName = colIndex(headers, 'name');
    const idxPass = colIndex(headers, 'password');
    const idxEmail = colIndex(headers, 'email');
    const idxRole = colIndex(headers, 'role');
    const idxStatus = colIndex(headers, 'status', 'is_login');

    // Cegah email duplikat
    const emailExists = dataRows.some(
      (r) => (r[idxEmail] || '').toLowerCase() === email.toLowerCase()
    );
    if (emailExists) throw badRequest('Email sudah terdaftar!');

    // ID baru
    let maxId = 0;
    dataRows.forEach((row) => {
      const n = Number(row[idxId]);
      if (!Number.isNaN(n) && n > maxId) maxId = n;
    });
    const newId = (maxId + 1).toString();

    const finalRole = ALLOWED_ROLES.includes((role || '').toLowerCase())
      ? role.toLowerCase()
      : 'karyawan';
    const finalPassword = password || DEFAULT_PASSWORD;
    const finalStatus = toStatus(status ?? is_login); // default FALSE

    // Susun baris sesuai urutan header sheet
    const newRow = new Array(headers.length).fill('');
    newRow[idxId] = newId;
    newRow[idxName] = name;
    newRow[idxPass] = finalPassword;
    newRow[idxEmail] = email;
    newRow[idxRole] = finalRole;
    newRow[idxStatus] = finalStatus;

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A:F',
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [newRow] },
    });

    res.json({
      success: true,
      message: 'User berhasil ditambahkan!',
      data: { id: newId, name, email, role: finalRole, status: finalStatus },
    });
  } catch (error) {
    next(error);
  }
};

// Update User (Hanya Developer & Admin)
exports.updateUser = async (req, res, next) => {
  try {
    const { email, name, password, role, status, is_login } = req.body;
    if (!email) throw badRequest('Email wajib disertakan untuk update user.');

    const sheets = await getSheetClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A:F',
    });

    const rows = response.data.values;
    if (!rows || rows.length <= 1) throw notFound('Data user kosong.');

    const headers = rows[0].map(h => h.trim().toLowerCase());
    const idxName = colIndex(headers, 'name');
    const idxPass = colIndex(headers, 'password');
    const idxEmail = colIndex(headers, 'email');
    const idxRole = colIndex(headers, 'role');
    const idxStatus = colIndex(headers, 'status', 'is_login');

    let rowIndex = -1;
    for (let i = 1; i < rows.length; i++) {
      const dbEmail = rows[i][idxEmail] || '';
      if (dbEmail.toLowerCase() === email.toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) throw notFound('User tidak ditemukan.');

    // Salin baris lama (isi sampai panjang header), lalu timpa field yang dikirim
    const updatedRow = Array.from(
      { length: headers.length },
      (_, i) => rows[rowIndex - 1][i] ?? ''
    );

    if (name !== undefined) updatedRow[idxName] = name;
    if (password !== undefined && password !== '') updatedRow[idxPass] = password;
    if (role !== undefined && ALLOWED_ROLES.includes((role || '').toLowerCase())) {
      updatedRow[idxRole] = role.toLowerCase();
    }
    const newStatus = status !== undefined ? status : is_login;
    if (newStatus !== undefined) updatedRow[idxStatus] = toStatus(newStatus);

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Users!A${rowIndex}:F${rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [updatedRow] },
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
    const idxEmail = colIndex(headers, 'email');

    let rowIndex = -1;
    for (let i = 1; i < rows.length; i++) {
      const dbEmail = rows[i][idxEmail] || '';
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