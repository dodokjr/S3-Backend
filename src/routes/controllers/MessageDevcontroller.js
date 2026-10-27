const { getSheetClient, SPREADSHEET_ID } = require('../../config/googleSheets');
const { badRequest, notFound, forbidden } = require('../Utils/Errors');

const SHEET_NAME = 'Messages';
const RANGE = `${SHEET_NAME}!A:F`;
const DEFAULT_HEADERS = ['id', 'nama', 'email', 'pesan', 'tag', 'status'];
const ALLOWED_TAGS = ['developer', 'semidev'];
const MAX_MESSAGE = 500;
const VIEW_ALL_ROLES = ['admin', 'developer'];

// ===== Status pesan =====
const ALLOWED_STATUS = ['unread', 'read', 'pending', 'success'];
const DEFAULT_STATUS = 'unread'; // default: belum dibaca
const STATUS_LABEL = {
  unread: 'Belum dibaca',
  read: 'Dibaca',
  pending: 'Sedang dikerjakan',
  success: 'Selesai',
};
// Alias supaya input fleksibel (Indonesia / Inggris)
const STATUS_ALIAS = {
  unread: 'unread', 'belum dibaca': 'unread',
  read: 'read', dibaca: 'read', 'di baca': 'read',
  pending: 'pending', pendding: 'pending', proses: 'pending',
  'sedang pengerjaan': 'pending', 'sedang dikerjakan': 'pending',
  success: 'success', succes: 'success', selesai: 'success', done: 'success',
};

const colIndex = (headers, ...names) => {
  for (const n of names) {
    const i = headers.indexOf(n);
    if (i !== -1) return i;
  }
  return -1;
};

const colLetter = (i) => String.fromCharCode(65 + i); // cukup untuk < 26 kolom

let sheetReady = false;
const ensureSheet = async (sheets) => {
  if (sheetReady) return;
  const info = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const exists = info.data.sheets.some((s) => s.properties.title === SHEET_NAME);
  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title: SHEET_NAME } } }] },
    });
  } else {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!1:1`,
    });
    const headerRow = (res.data.values?.[0] || []).map((h) => h.trim().toLowerCase());
    if (headerRow.length && !headerRow.includes('status')) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `${SHEET_NAME}!${colLetter(headerRow.length)}1`,
        valueInputOption: 'RAW',
        requestBody: { values: [['status']] },
      });
    }
  }
  sheetReady = true;
};

const normalizeTags = (input) => {
  const list = Array.isArray(input) ? input : (input || '').toString().split(',');
  const tags = [
    ...new Set(
      list
        .map((t) => (t || '').toString().trim().replace(/^@+/, '').toLowerCase())
        .filter(Boolean)
    ),
  ];
  const invalid = tags.filter((t) => !ALLOWED_TAGS.includes(t));
  if (invalid.length) {
    throw badRequest(
      `Tag tidak valid: ${invalid.map((t) => `@${t}`).join(', ')}. Pilihan: ${ALLOWED_TAGS.map((t) => `@${t}`).join(', ')}.`
    );
  }
  return tags;
};

const normalizeStatus = (input) => {
  const key = (input || '').toString().trim().toLowerCase();
  const status = STATUS_ALIAS[key];
  if (!status) {
    throw badRequest(`Status tidak valid. Pilihan: ${ALLOWED_STATUS.join(', ')}.`);
  }
  return status;
};

const findUserByEmail = async (sheets, email) => {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'Users!A:F',
  });
  const rows = response.data.values || [];
  if (rows.length <= 1) throw notFound('Data user kosong.');

  const headers = rows[0].map((h) => h.trim().toLowerCase());
  const idxName = colIndex(headers, 'name', 'nama');
  const idxEmail = colIndex(headers, 'email');

  const row = rows
    .slice(1)
    .find((r) => (r[idxEmail] || '').toLowerCase() === email.toLowerCase());
  if (!row) throw notFound('User tidak ditemukan.');

  return { nama: row[idxName] || '', email: row[idxEmail] || email };
};

exports.createMessage = async (req, res, next) => {
  try {
    const pesan = (req.body.pesan ?? req.body.message ?? '').toString().trim();
    if (!pesan) throw badRequest('Pesan wajib diisi!');
    if (pesan.length > MAX_MESSAGE) {
      throw badRequest(`Pesan maksimal ${MAX_MESSAGE} karakter.`);
    }
    const tags = normalizeTags(req.body.tags ?? req.body.tag);

    const requesterEmail = (req.user?.email || '').toString();
    if (!requesterEmail) throw badRequest('Email user tidak ditemukan pada token.');

    const sheets = await getSheetClient();
    await ensureSheet(sheets);
    const user = await findUserByEmail(sheets, requesterEmail);

    const existing = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: RANGE,
    });
    const rows = existing.data.values || [];
    const needsHeader = rows.length === 0;
    const headers = needsHeader ? DEFAULT_HEADERS : rows[0].map((h) => h.trim().toLowerCase());

    const idxId = colIndex(headers, 'id');
    const idxNama = colIndex(headers, 'nama');
    const idxEmail = colIndex(headers, 'email');
    const idxPesan = colIndex(headers, 'pesan');
    const idxTag = colIndex(headers, 'tag');
    const idxStatus = colIndex(headers, 'status');
    if ([idxId, idxNama, idxEmail, idxPesan, idxTag, idxStatus].includes(-1)) {
      throw badRequest(`Header sheet ${SHEET_NAME} harus: ${DEFAULT_HEADERS.join(', ')}`);
    }

    // ID baru (auto increment)
    let maxId = 0;
    rows.slice(1).forEach((row) => {
      const n = Number(row[idxId]);
      if (!Number.isNaN(n) && n > maxId) maxId = n;
    });
    const newId = (maxId + 1).toString();

    const newRow = new Array(headers.length).fill('');
    newRow[idxId] = newId;
    newRow[idxNama] = user.nama;
    newRow[idxEmail] = user.email;
    newRow[idxPesan] = pesan;
    newRow[idxTag] = tags.join(',');
    newRow[idxStatus] = DEFAULT_STATUS;

    // RAW (bukan USER_ENTERED) supaya pesan yang diawali "=" tidak dieksekusi sebagai formula sheet
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: RANGE,
      valueInputOption: 'RAW',
      requestBody: { values: needsHeader ? [DEFAULT_HEADERS, newRow] : [newRow] },
    });

    res.status(201).json({
      success: true,
      message: 'Pesan berhasil dikirim!',
      data: {
        id: newId,
        nama: user.nama,
        email: user.email,
        pesan,
        tag: tags,
        status: DEFAULT_STATUS,
        statusLabel: STATUS_LABEL[DEFAULT_STATUS],
      },
    });
  } catch (error) {
    // Log detail asli (termasuk balasan error dari Google API) supaya terlihat di log server/Vercel
    console.error('[createMessage]', error?.response?.data || error);
    next(error);
  }
};


exports.getMessages = async (req, res, next) => {
  try {
    const sheets = await getSheetClient();
    await ensureSheet(sheets);

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: RANGE,
    });

    const rows = response.data.values || [];
    if (rows.length <= 1) return res.json({ success: true, data: [] });

    const headers = rows[0].map((h) => h.trim().toLowerCase());

    let data = rows
      .slice(1)
      .filter((row) => row.some((cell) => (cell || '').toString().trim() !== ''))
      .map((row) => {
        const obj = {};
        headers.forEach((header, i) => {
          obj[header] = row[i] || '';
        });
        obj.tag = obj.tag ? obj.tag.split(',').map((t) => t.trim()).filter(Boolean) : [];
        // Data lama yang kolom statusnya kosong/tidak dikenal dianggap belum dibaca
        obj.status = STATUS_ALIAS[(obj.status || '').toString().trim().toLowerCase()] || DEFAULT_STATUS;
        obj.statusLabel = STATUS_LABEL[obj.status];
        return obj;
      });

    const requesterRole = (req.user?.role || '').toString().toLowerCase();
    if (!VIEW_ALL_ROLES.includes(requesterRole)) {
      const requesterEmail = (req.user?.email || '').toString().toLowerCase();
      data = data.filter((m) => (m.email || '').toLowerCase() === requesterEmail);
    }

    if (req.query.tag) {
      const wanted = req.query.tag.toString().replace(/^@+/, '').toLowerCase();
      data = data.filter((m) => m.tag.includes(wanted));
    }

    if (req.query.status) {
      const wanted = normalizeStatus(req.query.status);
      data = data.filter((m) => m.status === wanted);
    }

    // Terbaru di atas
    data.sort((a, b) => Number(b.id) - Number(a.id));

    res.json({ success: true, data });
  } catch (error) {
    console.error('[getMessages]', error?.response?.data || error);
    next(error);
  }
};

exports.updateMessageStatus = async (req, res, next) => {
  try {
    const role = (req.user?.role || '').toString().toLowerCase();
    if (!VIEW_ALL_ROLES.includes(role)) {
      throw forbidden('Anda tidak berhak mengubah status pesan.');
    }

    const id = (req.params.id || '').toString().trim();
    const status = normalizeStatus(req.body.status);

    const sheets = await getSheetClient();
    await ensureSheet(sheets);

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: RANGE,
    });
    const rows = response.data.values || [];
    if (rows.length <= 1) throw notFound('Pesan tidak ditemukan.');

    const headers = rows[0].map((h) => h.trim().toLowerCase());
    const idxId = colIndex(headers, 'id');
    const idxStatus = colIndex(headers, 'status');
    if (idxId === -1 || idxStatus === -1) {
      throw badRequest(`Header sheet ${SHEET_NAME} harus: ${DEFAULT_HEADERS.join(', ')}`);
    }

    const rowIndex = rows.findIndex((r, i) => i > 0 && (r[idxId] || '').toString() === id);
    if (rowIndex === -1) throw notFound('Pesan tidak ditemukan.');

    // rowIndex 0-based -> nomor baris sheet = rowIndex + 1
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!${colLetter(idxStatus)}${rowIndex + 1}`,
      valueInputOption: 'RAW',
      requestBody: { values: [[status]] },
    });

    res.json({
      success: true,
      message: 'Status pesan berhasil diperbarui!',
      data: { id, status, statusLabel: STATUS_LABEL[status] },
    });
  } catch (error) {
    console.error('[updateMessageStatus]', error?.response?.data || error);
    next(error);
  }
};