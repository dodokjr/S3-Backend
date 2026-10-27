const { getSheetClient, SPREADSHEET_ID } = require('../../config/googleSheets');
const { badRequest, notFound } = require('../Utils/Errors');
const { signCustomerToken } = require('../../routes/middleware/auth');

const SHEET_NAME = 'PopChat';
const RANGE = `${SHEET_NAME}!A:I`;
const KEY_PREFIX = 'S3';
const DEFAULT_CS = 'cs.s3';

const norm = (v) => (v ?? '').toString().trim().toLowerCase();
const str = (v) => (v ?? '').toString().trim();

const nowString = () =>
  new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta', hour12: false });

const makeKey = (id) => `${KEY_PREFIX}${id}`;
const padId = (n) => String(n).padStart(2, '0');

// Baris sheet -> object
const rowToObj = (row = []) => ({
  no: str(row[0]),
  id: str(row[1]),
  key: str(row[2]),
  nama: str(row[3]),
  email: str(row[4]),
  pesan: str(row[5]),
  pengirim: norm(row[6]),
  tgl: str(row[7]),
  reply_to: str(row[8]),
});

async function loadChats(sheets) {
  const r = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: RANGE });
  const rows = r.data.values || [];
  return rows
    .slice(1)
    .map((row, i) => ({ ...rowToObj(row), sheetRow: i + 2 }))
    .filter((o) => o.no !== '');
}

const nextNo = (chats) =>
  String(chats.reduce((max, c) => (Number(c.no) > max ? Number(c.no) : max), 0) + 1);

const toReply = (c) => ({
  no: c.no,
  id: c.id,
  key: c.key,
  nama: c.nama,
  email: c.email,
  pesan: c.pesan,
  pengirim: c.pengirim, // "cs" atau "customer"
  tgl: c.tgl,
  reply_to: c.reply_to,
});

const buildThreads = (chats) => {
  const byNo = new Map(chats.map((c) => [c.no, c]));

  // telusuri rantai reply_to sampai ketemu pesan utama
  const rootNoOf = (msg) => {
    let cur = msg;
    const seen = new Set();
    while (cur.reply_to && byNo.has(cur.reply_to) && !seen.has(cur.no)) {
      seen.add(cur.no);
      cur = byNo.get(cur.reply_to);
    }
    return cur.no;
  };

  const roots = chats.filter((c) => c.pengirim === 'customer' && !c.reply_to);
  const repliesByRoot = new Map();
  chats
    .filter((c) => c.reply_to)
    .forEach((c) => {
      const rootNo = rootNoOf(c);
      if (!repliesByRoot.has(rootNo)) repliesByRoot.set(rootNo, []);
      repliesByRoot.get(rootNo).push(c);
    });

  return roots
    .sort((a, b) => Number(a.no) - Number(b.no))
    .map((c) => ({
      no: c.no,
      id: c.id,
      key: c.key,
      nama: c.nama,
      email: c.email,
      pesan: c.pesan,
      tgl: c.tgl,
      balasan: (repliesByRoot.get(c.no) || [])
        .sort((a, b) => Number(a.no) - Number(b.no))
        .map(toReply),
    }));
};

// Ambil pesan + semua turunan balasannya (rekursif lewat reply_to)
const collectWithReplies = (chats, startNos) => {
  const result = new Set(startNos);
  let grew = true;
  while (grew) {
    grew = false;
    for (const c of chats) {
      if (c.reply_to && result.has(c.reply_to) && !result.has(c.no)) {
        result.add(c.no);
        grew = true;
      }
    }
  }
  return chats.filter((c) => result.has(c.no));
};

const appendRow = (sheets, row) =>
  sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: RANGE,
    // RAW: id "01" tidak berubah jadi angka 1, dan pesan berawalan "=" tidak dieksekusi sebagai formula
    valueInputOption: 'RAW',
    requestBody: { values: [row] },
  });

exports.createCustomerToken = async (req, res, next) => {
  try {
    const nama = str(req.body.nama);
    const email = str(req.body.email);

    if (!nama || !email) throw badRequest('nama dan email wajib diisi!');
    if (!/^\S+@\S+$/.test(email)) throw badRequest('Format email tidak valid.');

    const token = signCustomerToken({ nama, email });

    res.json({
      success: true,
      message: 'Token customer berhasil dibuat!',
      data: { token, nama, email: email.toLowerCase() },
    });
  } catch (error) {
    next(error);
  }
};

exports.getChats = async (req, res, next) => {
  try {
    const sheets = await getSheetClient();
    let chats = await loadChats(sheets);

    if (norm(req.user?.role) === 'customer') {
      // customer hanya boleh melihat chat dengan emailnya sendiri
      const myKeys = new Set(
        chats.filter((c) => norm(c.email) === norm(req.user.email)).map((c) => c.key)
      );
      chats = chats.filter((c) => myKeys.has(c.key));
    } else {
      const key = str(req.query.key);
      if (key) chats = chats.filter((c) => norm(c.key) === norm(key));
    }

    res.json({ success: true, data: buildThreads(chats) });
  } catch (error) {
    next(error);
  }
};

exports.createChat = async (req, res, next) => {
  try {
    // nama & email dari JWT (tidak bisa dipalsukan client)
    const nama = str(req.user.nama);
    const email = str(req.user.email);
    const pesan = str(req.body.pesan);

    if (!pesan) throw badRequest('pesan wajib diisi!');

    const sheets = await getSheetClient();
    const chats = await loadChats(sheets);

    const existing = chats.find((c) => c.pengirim === 'customer' && norm(c.email) === norm(email));
    let id;
    if (existing) {
      id = existing.id;
    } else {
      const maxId = chats.reduce((max, c) => (Number(c.id) > max ? Number(c.id) : max), 0);
      id = padId(maxId + 1);
    }

    const key = makeKey(id);
    const no = nextNo(chats);
    const tgl = nowString();

    await appendRow(sheets, [no, id, key, nama, email, pesan, 'customer', tgl, '']);

    res.json({
      success: true,
      message: 'Pesan berhasil dikirim!',
      data: { no, id, key, nama, email, pesan, tgl, balasan: [] },
    });
  } catch (error) {
    next(error);
  }
};

exports.customerReply = async (req, res, next) => {
  try {
    const nama = str(req.user.nama);
    const email = str(req.user.email);
    const pesan = str(req.body.pesan);
    const replyTo = str(req.body.reply_to);

    if (!pesan) throw badRequest('pesan wajib diisi!');

    const sheets = await getSheetClient();
    const chats = await loadChats(sheets);

    // room chat milik customer ini (berdasarkan email di token)
    const mine = chats.find((c) => c.pengirim === 'customer' && norm(c.email) === norm(email));
    if (!mine) throw notFound('Kamu belum punya chat. Kirim pesan terlebih dahulu.');

    const csMsgs = chats
      .filter((c) => c.pengirim === 'cs' && c.key === mine.key)
      .sort((a, b) => Number(a.no) - Number(b.no));

    if (csMsgs.length === 0) throw notFound('Belum ada balasan dari CS yang bisa dibalas.');

    let parent;
    if (replyTo) {
      // hanya boleh membalas pesan CS di room miliknya sendiri
      parent = csMsgs.find((c) => c.no === replyTo);
      if (!parent) throw notFound('Pesan CS yang dibalas tidak ditemukan.');
    } else {
      parent = csMsgs[csMsgs.length - 1];
    }

    const no = nextNo(chats);
    const tgl = nowString();

    await appendRow(sheets, [no, mine.id, mine.key, nama, email, pesan, 'customer', tgl, parent.no]);

    res.json({
      success: true,
      message: 'Balasan berhasil dikirim!',
      data: { no, id: mine.id, key: mine.key, nama, email, pesan, pengirim: 'customer', tgl, reply_to: parent.no },
    });
  } catch (error) {
    next(error);
  }
};

exports.replyChat = async (req, res, next) => {
  try {
    const key = str(req.body.key);
    const pesan = str(req.body.pesan);
    const replyTo = str(req.body.reply_to);
    const nama = str(req.body.nama) || DEFAULT_CS;
    const email = str(req.body.email) || DEFAULT_CS;

    if (!key || !pesan) throw badRequest('key dan pesan wajib diisi!');

    const sheets = await getSheetClient();
    const chats = await loadChats(sheets);

    const customerMsgs = chats
      .filter((c) => c.pengirim === 'customer' && norm(c.key) === norm(key))
      .sort((a, b) => Number(a.no) - Number(b.no));

    if (customerMsgs.length === 0) throw notFound('Chat dengan key tersebut tidak ditemukan.');

    let parent;
    if (replyTo) {
      parent = customerMsgs.find((c) => c.no === replyTo);
      if (!parent) throw notFound('Pesan customer yang dibalas tidak ditemukan.');
    } else {
      parent = customerMsgs[customerMsgs.length - 1];
    }

    const no = nextNo(chats);
    const tgl = nowString();

    await appendRow(sheets, [no, parent.id, parent.key, nama, email, pesan, 'cs', tgl, parent.no]);

    res.json({
      success: true,
      message: 'Balasan berhasil dikirim!',
      data: { no, id: parent.id, key: parent.key, nama, email, pesan, pengirim: 'cs', tgl, reply_to: parent.no },
    });
  } catch (error) {
    next(error);
  }
};

exports.deleteChat = async (req, res, next) => {
  try {
    const no = str(req.body.no);
    const key = str(req.body.key);
    if (!no && !key) throw badRequest('no atau key wajib disertakan untuk menghapus chat.');

    const sheets = await getSheetClient();
    const chats = await loadChats(sheets);

    const targets = key
      ? chats.filter((c) => norm(c.key) === norm(key))
      : collectWithReplies(chats, [no]); // pesan tsb + seluruh balasan di bawahnya

    if (targets.length === 0) throw notFound('Data chat tidak ditemukan.');

    const sheetInfo = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const sheet = sheetInfo.data.sheets.find((s) => s.properties.title === SHEET_NAME);
    if (!sheet) throw notFound(`Sheet "${SHEET_NAME}" tidak ditemukan.`);

    // Hapus dari baris paling bawah supaya index baris di atasnya tidak bergeser
    const requests = targets
      .map((t) => t.sheetRow)
      .sort((a, b) => b - a)
      .map((rowNum) => ({
        deleteDimension: {
          range: {
            sheetId: sheet.properties.sheetId,
            dimension: 'ROWS',
            startIndex: rowNum - 1,
            endIndex: rowNum,
          },
        },
      }));

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests },
    });

    res.json({ success: true, message: `${targets.length} pesan berhasil dihapus!` });
  } catch (error) {
    next(error);
  }
};