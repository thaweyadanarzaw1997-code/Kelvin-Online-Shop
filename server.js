const express = require('express');
const bodyParser = require('body-parser');
const axios = require('axios');
const { google } = require('googleapis');
const multer = require('multer');
const cloudinary = require('cloudinary').v2;
const XLSX = require('xlsx');
const fs = require('fs');
const cors = require('cors');
const Jimp = require('jimp');
const FormData = require('form-data');

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

// Google Sheets Authentication
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
    console.log("Google Auth initialized successfully for:", process.env.GOOGLE_CLIENT_EMAIL);
  }
} catch (e) {
  console.error("Google Auth Initialization Error:", e.message);
}

const sheets = google.sheets({ version: 'v4', auth });
const adminStates = {};

app.get('/', (req, res) => {
  res.send('Kelvin Online Shop Backend API is running!');
});

// 1. Get Items API
app.get('/api/items', async (req, res) => {
  try {
    if (!auth) {
      return res.status(500).json({ status: 'error', message: 'Google Auth is not initialized.' });
    }

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: 'Sheet1!A:L',
    });

    const rows = response.data.values;
    if (!rows || rows.length < 2) {
      return res.json({ status: 'success', items: [], sysDate: new Date().toISOString().split('T')[0] });
    }

    let items = [];
    for (let i = 1; i < rows.length; i++) {
      let r = rows[i];
      items.push({
        productId: r[0] || '', 
        productName: r[1] || '', 
        imageUrl: r[2] || '',
        price3ml: Number(r[3]) || 0, 
        price5ml: Number(r[4]) || 0,
        price10ml: Number(r[5]) || 0, 
        price30ml: Number(r[6]) || 0,
        price100ml: Number(r[7]) || 0, 
        category: r[8] || 'All',
        type: r[9] || '', 
        description: r[10] || '', 
        status: r[11] || 'Active'
      });
    }

    res.json({ status: 'success', items, sysDate: new Date().toISOString().split('T')[0] });
  } catch (err) {
    console.error("Error fetching items:", err.message);
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

    const caption = `✨ *FRAGRANCE BY KELVIN* ✨\n\n` +
                    `🛒 *New Order Received*\n\n` +
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
    console.error("Order Error:", err.message);
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
      const messageId = cb.message.message_id;
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
        try {
          await axios.post(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/editMessageReplyMarkup`, {
            chat_id: chatId, message_id: messageId, reply_markup: { inline_keyboard: [] }
          });
        } catch (e) {}
      }
      await axios.post(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/answerCallbackQuery`, { callback_query_id: cb.id });
    }

    if (update.message) {
      const msg = update.message;
      const chatId = msg.chat.id;
      const textMsg = (msg.text || '').trim();

      if (textMsg === '/pending') {
        await sendOrdersListByStatus(chatId, 'Pending', true);
      }
      else if (textMsg === '/inprogress') {
        await sendOrdersListByStatus(chatId, 'In Progress', false);
      }
      else if (textMsg === '/finish') {
        await sendReportExcel(chatId, 'Finish');
      }
      else if (textMsg.startsWith('/ORD-') || textMsg.startsWith('/ord-')) {
        const lookupOrderId = textMsg.replace('/', '');
        await sendVoucherPngByOrderId(chatId, lookupOrderId);
      }
      else if (adminStates[chatId] && adminStates[chatId].action === 'waiting_waybill') {
        const waybillNo = textMsg;
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
    console.error("Webhook Error:", err);
    res.sendStatus(500);
  }
});

async function updateOrderStatus(orderId, status) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: GOOGLE_SHEET_ID, range: 'Sheet2!A:M' });
  const rows = res.data.values;
  if (!rows) return;
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
  if (!rows) return;
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

// Generate PNG Voucher Image using Jimp and send as Photo
async function sendVoucherPngByOrderId(chatId, orderId) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: GOOGLE_SHEET_ID, range: 'Sheet2!A:M' });
  const rows = res.data.values || [];
  
  let foundRow = null;
  for (let r of rows) {
    if (r[10] && r[10].toString().trim().toLowerCase() === orderId.toLowerCase()) {
      foundRow = r;
      break;
    }
  }

  if (!foundRow) {
    await axios.post(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
      chat_id: chatId, text: `❌ Order ID (${orderId}) ကို ရှာမတွေ့ပါ။ နံပါတ်မှန်ကန်မှု ရှိမရှိ စစ်ဆေးပါ။`
    });
    return;
  }

  const date = foundRow[0] || 'N/A';
  const customer = foundRow[1] || 'N/A';
  const phone = foundRow[2] || 'N/A';
  const address = foundRow[4] || 'N/A';
  const items = foundRow[5] || 'N/A';
  const total = Number(foundRow[6] || 0).toLocaleString() + ' Ks';
  const status = foundRow[7] || 'N/A';
  const payment = foundRow[9] || 'N/A';
  const waybill = foundRow[12] || foundRow[11] || 'None';

  // Create Voucher image using Jimp
  const image = await Jimp.create(650, 850, 0xFFFAF9F6);
  const fontTitle = await Jimp.loadFont(Jimp.FONT_SANS_32_BLACK);
  const fontSub = await Jimp.loadFont(Jimp.FONT_SANS_14_BLACK);
  const fontText = await Jimp.loadFont(Jimp.FONT_SANS_16_BLACK);

  image.print(fontTitle, 0, 40, { text: 'FRAGRANCE', alignmentX: Jimp.HORIZONTAL_ALIGN_CENTER }, 650);
  image.print(fontSub, 0, 80, { text: 'BY KELVIN', alignmentX: Jimp.HORIZONTAL_ALIGN_CENTER }, 650);

  image.print(fontText, 50, 130, `Order ID: ${orderId}`);
  image.print(fontText, 50, 170, `Date: ${date}`);
  image.print(fontText, 50, 210, `Customer: ${customer}`);
  image.print(fontText, 50, 250, `Phone: ${phone}`);
  image.print(fontText, 50, 290, `Address: ${address}`);
  image.print(fontText, 50, 350, `Items: ${items}`);
  image.print(fontText, 50, 530, `Payment: ${payment}`);
  image.print(fontText, 50, 570, `Status: ${status}`);
  image.print(fontText, 50, 610, `Waybill: ${waybill}`);
  image.print(fontText, 50, 680, `Total Amount: ${total}`);

  const pngBuffer = await image.getBufferAsync(Jimp.MIME_PNG);

  let keyboardRows = [];
  if (status === 'Pending') {
    keyboardRows.push([
      { text: "➕ Add WayBill", callback_data: `add_wb_${orderId}` },
      { text: "❌ Cancel", callback_data: `cancel_${orderId}` }
    ]);
  } else if (status === 'In Progress') {
    keyboardRows.push([
      { text: "❌ Cancel", callback_data: `cancel_${orderId}` }
    ]);
  }

  const replyMarkup = keyboardRows.length > 0 ? { inline_keyboard: keyboardRows } : undefined;

  const formData = new FormData();
  formData.append('chat_id', chatId);
  formData.append('photo', pngBuffer, { filename: `${orderId}.png`, contentType: 'image/png' });
  formData.append('caption', `✨ Voucher for Order: \`${orderId}\``);
  formData.append('parse_mode', 'Markdown');
  if (replyMarkup) {
    formData.append('reply_markup', JSON.stringify(replyMarkup));
  }

  await axios.post(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendPhoto`, formData, {
    headers: formData.getHeaders()
  });
}

async function sendOrdersListByStatus(chatId, statusFilter, showAddWbButton) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: GOOGLE_SHEET_ID, range: 'Sheet2!A:M' });
  const rows = res.data.values || [];
  
  let filtered = rows.filter(r => (r[7] === statusFilter || r[12] === statusFilter));

  if (filtered.length === 0) {
    await axios.post(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
      chat_id: chatId, text: `📭 ${statusFilter} Orders များ မရှိသေးပါ။`
    });
    return;
  }

  for (let r of filtered) {
    const orderId = r[10] || 'N/A';
    const customer = r[1] || 'N/A';
    const phone = r[2] || 'N/A';
    const address = r[4] || 'N/A';
    const items = r[5] || 'N/A';
    const total = r[6] || 0;
    const payment = r[9] || 'N/A';

    const text = `✨ *FRAGRANCE BY KELVIN* ✨\n\n` +
                 `📦 *${statusFilter} Order*\n\n` +
                 `🆔 Order ID: \`${orderId}\`\n` +
                 `👤 Name: ${customer}\n` +
                 `📞 Phone: ${phone}\n` +
                 `📍 Address: ${address}\n` +
                 `🛍 Items: ${items}\n` +
                 `💰 Payment: ${payment}\n` +
                 `💵 Total: ${Number(total).toLocaleString()} Ks`;

    let keyboardRows = [];
    if (showAddWbButton) {
      keyboardRows.push([
        { text: "➕ Add WayBill", callback_data: `add_wb_${orderId}` },
        { text: "❌ Cancel", callback_data: `cancel_${orderId}` }
      ]);
    } else {
      keyboardRows.push([
        { text: "❌ Cancel", callback_data: `cancel_${orderId}` }
      ]);
    }

    await axios.post(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
      chat_id: chatId,
      text: text,
      parse_mode: 'Markdown',
      reply_markup: { inline_keyboard: keyboardRows }
    });
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
    const waybill = (row[4] || row[3] || row[2] || '').toString().trim();
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
      const s2Waybill = (s2Row[11] || s2Row[12] || '').toString().trim();
      
      if (s2Waybill && s2Waybill === waybill) {
        const rowIndex = j + 1;
        await sheets.spreadsheets.values.update({
          spreadsheetId: GOOGLE_SHEET_ID, range: `Sheet2!H${rowIndex}`,
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

  let filtered = rows.filter(r => (r[7] === statusFilter || r[12] === statusFilter));
  let wbData = [["Date", "Customer", "Phone", "Address", "Items", "Total", "Status", "Payment", "OrderID", "Waybill"]];
  filtered.forEach(r => {
    wbData.push([r[0], r[1], r[2], r[4], r[5], r[6], r[7], r[9], r[10], r[12]]);
  });

  const ws = XLSX.utils.aoa_to_sheet(wbData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, `${statusFilter} Orders`);

  const filePath = `/tmp/${statusFilter}_Report.xlsx`;
  XLSX.writeFile(wb, filePath);

  const formData = new FormData();
  formData.append('chat_id', chatId);
  formData.append('document', fs.createReadStream(filePath), { filename: `${statusFilter}_Report.xlsx` });
  formData.append('caption', `📋 Report for Status: *${statusFilter}*`);

  await axios.post(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendDocument`, formData, {
    headers: formData.getHeaders()
  });

  fs.unlinkSync(filePath);
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
