const express = require('express');
const path = require('path');
const XLSX = require('xlsx');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

// Excel file path
const EXCEL_PATH = path.join(__dirname, '摩点订单.xlsx');

let rows = [];

function loadExcel() {
  try {
    const workbook = XLSX.readFile(EXCEL_PATH);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    rows = [];
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (!r) continue;
      const orderId = (r[0] || '').toString().trim();
      const rewardTitle = (r[1] || '').toString().trim();
      const userId = (r[7] || '').toString().trim();
      const serialNumber = (r[11] || '').toString().trim();
      if (orderId || userId) {
        rows.push({ orderId, userId, rewardTitle, serialNumber, rawRowIndex: i + 1 });
      }
    }
    console.log(`Loaded ${rows.length} rows from Excel.`);
  } catch (err) {
    console.error('Failed to load Excel:', err.message);
  }
}

loadExcel();

app.get('/api/lookup', (req, res) => {
  const order = (req.query.order || '').trim().toLowerCase();
  const userid = (req.query.userid || '').trim().toLowerCase();
  const found = rows.find(r => r.orderId.toLowerCase() === order && r.userId.toLowerCase() === userid);
  if (!found) return res.json({ ok: false, message: '没有找到匹配的订单。' });

  res.json({ ok: true, serialNumber: found.serialNumber, rewardTitle: found.rewardTitle, rawRowIndex: found.rawRowIndex });
});

app.use(express.static(path.join(__dirname, 'public')));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
