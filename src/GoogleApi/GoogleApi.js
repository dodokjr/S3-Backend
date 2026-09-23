const { google } = require('googleapis');

async function bacaSpreadsheet() {
  try {
    // 1. Autentikasi menggunakan file kredensial service account
    const auth = new google.auth.GoogleAuth({
      keyFile: 'credentials.json',
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: 'v4', auth: client });

    // ID Spreadsheet (bisa diambil dari URL spreadsheet Anda)
    const spreadsheetId = 'https://docs.google.com/spreadsheets/d/1RkRKEt8AW0CNArv-FqEM9JJYUox1fGHZkV9JLuQgqXw/edit?gid=0#gid=0';

    // 2. Membaca data dari range tertentu (contoh: Sheet1!A1:B10)
    const response = await googleSheets.spreadsheets.values.get({
      spreadsheetId,
      range: 'Sheet1!A1:C4',
    });

    const rows = response.data.values;
    if (rows.length) {
      console.log('Data ditemukan:');
      rows.map((row) => {
        console.log(`${row[0]} - ${row[1]}`);
      });
    } else {
      console.log('Tidak ada data ditemukan.');
    }
  } catch (error) {
    console.error('Terjadi kesalahan:', error);
  }
}

bacaSpreadsheet();