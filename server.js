const express = require('express');
const bodyParser = require('body-parser');
const axios = require('axios');
const { google } = require('googleapis');
const multer = require('multer');
const cloudinary = require('cloudinary').v2;
const XLSX = require('xlsx');
const fs = require('fs');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ extended: true, limit: '50mb' }));

const upload = multer({ dest: '/tmp/' });

const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
const ADMIN_CHAT_ID = process.env.ADMIN_CHAT_ID;
const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID;

// Debug Endpoint to check exact environment variables status
app.get('/api/debug-auth', (req, res) => {
  const email = process.env.GOOGLE_CLIENT_EMAIL;
  const privKey = process.env.GOOGLE_PRIVATE_KEY;
  const sheetId = process.env.GOOGLE_SHEET_ID;

  res.json({
    hasEmail: !!email,
    emailValue: email || 'MISSING',
    hasPrivateKey: !!privKey,
    privateKeyLength: privKey ? privKey.length : 0,
    hasSheetId: !!sheetId,
    sheetIdValue: sheetId || 'MISSING'
  });
});

// Google Sheets Authentication using Separated Environment Variables
let auth;
try {
  let privateKey = process.env.GOOGLE_PRIVATE_KEY || '';
  if (privateKey.includes('\\n')) {
    privateKey = privateKey.replace(/\\n/g, '\n');
  }

  if (process.env.GOOGLE_CLIENT_EMAIL && privateKey) {
    auth = new google.auth.GoogleAuth({
      credentials: {
        client_email: process.env.GOOGLE_CLIENT_EMAIL,
        private_key: privateKey,
      },
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
  }
} catch (e) {
  console.error("Auth Error:", e.message);
}

const sheets = google.sheets({ version: 'v4', auth });

app.get('/api/items', async (req, res) => {
  try {
    if (!auth) {
      return res.status(500).json({ 
        status: 'error', 
        message: 'Auth is null. GOOGLE_CLIENT_EMAIL or GOOGLE_PRIVATE_KEY is missing in Vercel Environment Variables.' 
      });
    }

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: 'Sheet1!A:L',
    });

    res.json({ status: 'success', rowsCount: response.data.values ? response.data.values.length : 0 });
  } catch (err) {
    res.status(500).json({ status: 'error', message: err.toString() });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
