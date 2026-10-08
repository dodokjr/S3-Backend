const { getSheetClient, SPREADSHEET_ID } = require('../../config/googleSheets');
const { badRequest, notFound } = require('../Utils/Errors');

const SHEET_NAME = 'Developer_message';
const RANGE = `${SHEET_NAME}!A:E`;
const DEFAULT_HEADERS = ['id', 'nama', 'email', 'pesan', 'tag'];
const ALLOWED_TAGS = ['developer', 'semidev'];
const MAX_MESSAGE = 500;
// Role yang boleh melihat semua pesan. Role lain hanya melihat pesan miliknya sendiri.
const VIEW_ALL_ROLES = ['admin', 'developer'];

const colIndex = (headers, ...names) => {
  for (const n of names) {
    const i = headers.indexOf(n);
    if (i !== -1) return i;
  }
  return -1;
};

// Pastikan sheet "Messages" ada (dibuat otomatis kalau belum ada). Dicek sekali per proses.
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
  }
  sheetReady = true;
};

// Terima array ["developer"] / ["@semidev"] atau string "developer,semidev".
// Hasilnya selalu array id tanpa "@", tanpa duplikat, dan hanya user yang diizinkan.
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

// Cari nama & email dari sheet Users berdasarkan email di token
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

// Kirim pesan baru (semua user yang sudah punya token / login).
// Body: { pesan: "...", tags: ["developer", "semidev"] }  (alias: message, tag)
// Nama & email TIDAK dikirim dari client, diambil dari sheet Users berdasarkan token.
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
    if ([idxId, idxNama, idxEmail, idxPesan, idxTag].includes(-1)) {
      throw new Error(`Header sheet ${SHEET_NAME} harus: ${DEFAULT_HEADERS.join(', ')}`);
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
      data: { id: newId, nama: user.nama, email: user.email, pesan, tag: tags },
    });
  } catch (error) {
    next(error);
  }
};

// Ambil pesan (semua user yang punya token).
// admin & developer melihat semua pesan; role lain hanya pesan miliknya (match email).
// Query opsional: ?tag=developer
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

    // Terbaru di atas
    data.sort((a, b) => Number(b.id) - Number(a.id));

    res.json({ success: true, data });
  } catch (error) {
    next(error);
  }
};