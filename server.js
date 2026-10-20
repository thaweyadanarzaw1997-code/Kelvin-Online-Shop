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

cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
  api_key: process.env.CLOUDINARY_API_KEY,
  api_secret: process.env.CLOUDINARY_API_SECRET
});

// Google Sheets Authentication using GOOGLE_SERVICE_ACCOUNT JSON env variable
let auth;
try {
  const serviceAccountKey = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
  auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: serviceAccountKey.client_email,
      private_key: serviceAccountKey.private_key,
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
} catch (e) {
  console.error("Google Auth Setup Error:", e.message);
}

const sheets = google.sheets({ version: 'v4', auth });
const adminStates = {};

// 1. Get Items API (Sheet1 မှ Product များကို Odoo သို့ ပို့ပေးမည်)
app.get('/api/items', async (req, res) => {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: 'Sheet1!A:L',
    });
    const rows = response.data.values;
    if (!rows || rows.length < 2) return res.json({ items: [] });
    
    let items = [];
    for (let i = 1; i < rows.length; i++) {
      let r = rows[i];
      items.push({
        productId: r[0], productName: r[1], imageUrl: r[2],
        price3ml: Number(r[3]) || 0, price5ml: Number(r[4]) || 0,
        price10ml: Number(r[5]) || 0, price30ml: Number(r[6]) || 0,
        price100ml: Number(r[7]) || 0, category: r[8] || 'All',
        type: r[9] || '', description: r[10] || '', status: r[11] || 'Active'
      });
    }
    res.json({ status: 'success', items, sysDate: new Date().toISOString().split('T')[0] });
  } catch (err) {
    res.status(500).json({ status: 'error', message: err.toString() });
  }
});

// 2. Customer Order API
app.post('/api/order', async (req, res) => {
  try {
    const data = req.body;
    let photoUrl = "";
    
    if (data.photoBase64) {
      const uploadResponse = await cloudinary.uploader.upload(`data:image/jpeg;base64,${data.photoBase64}`, {
        folder: 'fragrance_vouchers'
      });
      photoUrl = uploadResponse.secure_url;
    }

    const todayStr = new Date().toISOString().split('T')[0].replace(/-/g, '');
    const items = typeof data.items === 'string' ? JSON.parse(data.items) : data.items;
    
    const sheetData = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: 'Sheet2!A:M',
    });
    const nextRow = (sheetData.data.values ? sheetData.data.values.length : 1) + 1;
    const orderId = `ORD-${todayStr}-${String(nextRow).padStart(4, '0')}`;

    let rowValues = [
      data.date, data.customer, data.phone, data.viber || '', data.address,
      items.map(i => `${i.name} (x${i.quantity})`).join(', '),
      items.reduce((sum, i) => sum + (i.price * i.quantity), 0),
      "Pending", photoUrl, data.paymentMethod, orderId, "", "Pending"
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: 'Sheet2!A:M',
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [rowValues] }
    });

    const caption = `🛒 *New Order Received (Odoo)*\n\n` +
                    `🆔 Order ID: \`${orderId}\`\n` +
                    `👤 Name: ${data.customer}\n` +
                    `📞 Phone: ${data.phone}\n` +
                    `📍 Address: ${data.address}\n` +
                    `💰 Payment: ${data.paymentMethod}\n` +
                    `💵 Total: ${rowValues[6].toLocaleString()} Ks`;

    const inlineKeyboard = {
      inline_keyboard: [
        [
          { text: "➕ Add WayBill", callback_data: `add_wb_${orderId}` },
          { text: "❌ Cancel", callback_data: `cancel_${orderId}` }
        ]
      ]
    };

    if (photoUrl) {
      await axios.post(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendPhoto`, {
        chat_id: ADMIN_CHAT_ID, photo: photoUrl, caption: caption, parse_mode: 'Markdown', reply_markup: inlineKeyboard
      });
    } else {
      await axios.post(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
        chat_id: ADMIN_CHAT_ID, text: caption, parse_mode: 'Markdown', reply_markup: inlineKeyboard
      });
    }

    res.json({ result: 'success', orderId });
  } catch (err) {
    res.status(500).json({ result: 'error', message: err.toString() });
  }
});

// 3. Telegram Webhook
app.post('/api/telegram-webhook', upload.single('document'), async (req, res) => {
  try {
    const update = req.body;
    
    if (update.callback_query) {
      const cb = update.callback_query;
      const chatId = cb.message.chat.id;
      const data = cb.data;

      if (data.startsWith('add_wb_')) {
        const orderId = data.replace('add_wb_', '');
        adminStates[chatId] = { action: 'waiting_waybill', orderId };
        await axios.post(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
          chat_id: chatId, text: `✍️ ကျေးဇူးပြု၍ Order ID (${orderId}) အတွက် WayBill နံပါတ်ကို ရိုက်ထည့်ပေးပါ:`
        });
      } else if (data.startsWith('cancel_')) {
        const orderId = data.replace('cancel_', '');
        await updateOrderStatus(orderId, 'Cancelled');
        await axios.post(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
          chat_id: chatId, text: `❌ Order ID (${orderId}) အား Cancel လုပ်လိုက်ပါပြီ။`
        });
      }
      await axios.post(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/answerCallbackQuery`, { callback_query_id: cb.id });
    }

    if (update.message) {
      const msg = update.message;
      const chatId = msg.chat.id;

      if (msg.text === '/pending' || msg.text === '/inprogress' || msg.text === '/finish') {
        const statusMap = { '/pending': 'Pending', '/inprogress': 'In Progress', '/finish': 'Finish' };
        await sendReportExcel(chatId, statusMap[msg.text]);
      }
      else if (adminStates[chatId] && adminStates[chatId].action === 'waiting_waybill') {
        const waybillNo = msg.text.trim();
        const orderId = adminStates[chatId].orderId;
        delete adminStates[chatId];

        await updateOrderWaybillAndStatus(orderId, waybillNo, 'In Progress');
        await axios.post(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
          chat_id: chatId, text: `✅ Order ID (${orderId}) တွင် WayBill (${waybillNo}) ထည့်သွင်းပြီး Status ကို 'In Progress' သို့ ပြောင်းလိုက်ပါပြီ။`
        });
      }
      else if (msg.document) {
        const fileId = msg.document.file_id;
        const fileRes = await axios.get(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/getFile?file_id=${fileId}`);
        const filePath = fileRes.data.result.file_path;
        const downloadUrl = `https://api.telegram.org/file/bot${TELEGRAM_BOT_TOKEN}/${filePath}`;

        const excelResponse = await axios({ method: 'GET', url: downloadUrl, responseType: 'arraybuffer' });
        const tmpFilePath = `/tmp/${Date.now()}.xlsx`;
        fs.writeFileSync(tmpFilePath, excelResponse.data);

        const workbook = XLSX.readFile(tmpFilePath);
        const sheetName = workbook.SheetNames[0];
        const excelRows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

        await processExcelUpload(excelRows, chatId);
        fs.unlinkSync(tmpFilePath);
      }
    }
    res.sendStatus(200);
  } catch (err) {
    console.error(err);
    res.sendStatus(500);
  }
});

async function updateOrderStatus(orderId, status) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: GOOGLE_SHEET_ID, range: 'Sheet2!A:M' });
  const rows = res.data.values;
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][10] === orderId) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: GOOGLE_SHEET_ID, range: `Sheet2!H${i+1}`,
        valueInputOption: 'USER_ENTERED', requestBody: { values: [[status]] }
      });
      break;
    }
  }
}

async function updateOrderWaybillAndStatus(orderId, waybill, status) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: GOOGLE_SHEET_ID, range: 'Sheet2!A:M' });
  const rows = res.data.values;
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][10] === orderId) {
      const rowIndex = i + 1;
      await sheets.spreadsheets.values.update({
        spreadsheetId: GOOGLE_SHEET_ID, range: `Sheet2!H${rowIndex}:M${rowIndex}`,
        valueInputOption: 'USER_ENTERED', requestBody: { values: [[status, rows[i][8], rows[i][9], orderId, waybill, status]] }
      });
      break;
    }
  }
}

async function processExcelUpload(excelRows, chatId) {
  const sheet3Res = await sheets.spreadsheets.values.get({ spreadsheetId: GOOGLE_SHEET_ID, range: 'Sheet3!A:Z' });
  const sheet3Values = sheet3Res.data.values || [];
  const sheet2Res = await sheets.spreadsheets.values.get({ spreadsheetId: GOOGLE_SHEET_ID, range: 'Sheet2!A:M' });
  const sheet2Rows = sheet2Res.data.values || [];

  const existingWaybillsSheet3 = new Set(sheet3Values.slice(1).map(r => r[4]?.toString().trim()));
  let newRowsCount = 0;
  let updatedOrdersCount = 0;

  for (let i = 1; i < excelRows.length; i++) {
    const row = excelRows[i];
    const waybill = row[4]?.toString().trim();
    if (!waybill) continue;

    if (!existingWaybillsSheet3.has(waybill)) {
      await sheets.spreadsheets.values.append({
        spreadsheetId: GOOGLE_SHEET_ID, range: 'Sheet3!A:O',
        valueInputOption: 'USER_ENTERED', requestBody: { values: [row] }
      });
      existingWaybillsSheet3.add(waybill);
      newRowsCount++;
    }

    for (let j = 0; j < sheet2Rows.length; j++) {
      const s2Row = sheet2Rows[j];
      const s2Waybill = s2Row[12]?.toString().trim();
      if (s2Waybill === waybill) {
        await sheets.spreadsheets.values.update({
          spreadsheetId: GOOGLE_SHEET_ID, range: `Sheet2!H${j+1}`,
          valueInputOption: 'USER_ENTERED', requestBody: { values: [["Finish"]] }
        });
        updatedOrdersCount++;
      }
    }
  }

  await axios.post(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
    chat_id: chatId,
    text: `📊 Excel Upload အောင်မြင်ပါသည်။\n- Sheet 3 သို့ အသစ်ထည့်သွင်းမှု: ${newRowsCount} ခု\n- Finish သို့ ပြောင်းလဲပြီးစီးမှု: ${updatedOrdersCount} ခု`
  });
}

async function sendReportExcel(chatId, statusFilter) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: GOOGLE_SHEET_ID, range: 'Sheet2!A:M' });
  const rows = res.data.values || [];
  
  let filtered = rows.filter(r => r[7] === statusFilter);
  let wbData = [["Date", "Customer", "Phone", "Address", "Items", "Total", "Status", "Payment", "OrderID", "Waybill"]];
  filtered.forEach(r => {
    wbData.push([r[0], r[1], r[2], r[4], r[5], r[6], r[7], r[9], r[10], r[12]]);
  });

  const ws = XLSX.utils.aoa_to_sheet(wbData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, `${statusFilter} Orders`);
  
  const filePath = `/tmp/${statusFilter}_Report.xlsx`;
  XLSX.writeFile(wb, filePath);

  const formData = new (require('form-data'))();
  formData.append('chat_id', chatId);
  formData.append('document', fs.createReadStream(filePath));
  formData.append('caption', `📋 Report for Status: *${statusFilter}*`);

  await axios.post(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendDocument`, formData, {
    headers: formData.getHeaders()
  });

  fs.unlinkSync(filePath);
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));

// Google Sheets Authentication using GOOGLE_SERVICE_ACCOUNT JSON env variable
let auth;
try {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT) {
    throw new Error("GOOGLE_SERVICE_ACCOUNT environment variable is missing!");
  }
  const serviceAccountKey = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
  auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: serviceAccountKey.client_email,
      private_key: serviceAccountKey.private_key,
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  console.log("Google Auth initialized successfully for:", serviceAccountKey.client_email);
} catch (e) {
  console.error("Google Auth Setup Error:", e.message);
}
