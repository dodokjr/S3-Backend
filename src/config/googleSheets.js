const { google } = require('googleapis');

const SPREADSHEET_ID = '1RkRKEt8AW0CNArv-FqEM9JJYUox1fGHZkV9JLuQgqXw';
const OFFICIAL_DEV_EMAIL = 's3-600@appss-3c587.iam.gserviceaccount.com';

async function getSheetClient() {
  const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS);
  
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  
  const client = await auth.getClient();
  return google.sheets({ version: 'v4', auth: client });
}

module.exports = {
  getSheetClient,
  SPREADSHEET_ID,
  OFFICIAL_DEV_EMAIL
};