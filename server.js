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

app.get('/', (req, res) => {
  res.send('Server is running!');
});

// Items API with full error catching so you can see the exact cause
app.get('/api/items', async (req, res) => {
  try {
    const clientEmail = process.env.GOOGLE_CLIENT_EMAIL;
    let privateKey = process.env.GOOGLE_PRIVATE_KEY || '';
    const sheetId = process.env.GOOGLE_SHEET_ID;

    // Check if any env variable is missing
    if (!clientEmail || !privateKey || !sheetId) {
      return res.status(500).json({
        success: false,
        error_type: "MISSING_ENV_VARIABLES",
        hasEmail: !!clientEmail,
        hasPrivateKey: !!privateKey,
        hasSheetId: !!sheetId
      });
    }

    // Handle newlines in private key
    if (privateKey.includes('\\n')) {
      privateKey = privateKey.replace(/\\n/g, '\n');
    }

    const auth = new google.auth.GoogleAuth({
      credentials: {
        client_email: clientEmail,
        private_key: privateKey,
      },
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const sheets = google.sheets({ version: 'v4', auth });

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: 'Sheet1!A:L',
    });

    const rows = response.data.values;
    res.json({ 
      success: true, 
      totalRows: rows ? rows.length : 0,
      firstRow: rows ? rows[0] : null
    });

  } catch (err) {
    // This will catch the EXACT Google API or crypto error and show it in browser
    res.status(500).json({
      success: false,
      error_message: err.message,
      error_stack: err.stack
    });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
