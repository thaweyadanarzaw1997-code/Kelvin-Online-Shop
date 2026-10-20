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
    
    // DEBUG: Sheet ထဲက Data တွေ အမှန်တကယ် ဝင်မဝင် Browser မှာ တိုက်ရိုက်ကြည့်ရန်
    if (!rows || rows.length === 0) {
      return res.json({ status: 'empty', message: 'Sheet1 is completely empty or name is incorrect!' });
    }

    let items = [];
    // Skip header row (i = 1)
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

    res.json({ status: 'success', rawRowsCount: rows.length, items });
  } catch (err) {
    console.error("Error fetching items:", err.message);
    res.status(500).json({ status: 'error', message: err.toString() });
  }
});
