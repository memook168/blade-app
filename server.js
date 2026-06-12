const express = require('express');
const ExcelJS = require('exceljs');
const cors = require('cors');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// ── เปลี่ยน path นี้ให้ตรงกับที่เก็บไฟล์ Excel จริงๆ ──
const EXCEL_PATH = path.join(
  process.env.USERPROFILE || process.env.HOME,
  'OneDrive - Mahidol University',
  'blade_template_final.xlsx'
);

app.post('/api/log', async (req, res) => {
  try {
    const { timestamp, productCode, bladeType, action,
            statusTo, machine, note, operator } = req.body;

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_PATH);

    const sheet = workbook.getWorksheet('Transaction_Log');

    sheet.addRow([
      timestamp, '', productCode, bladeType,
      action, statusTo, machine, note, operator
    ]);

    await workbook.xlsx.writeFile(EXCEL_PATH);

    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/products', async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_PATH);

    const sheet = workbook.getWorksheet('Master_Blade');
    const products = [];

    sheet.eachRow((row, i) => {
      if (i <= 2) return;
      const code = row.getCell(2).value;
      const name = row.getCell(3).value;
      if (code && !products.find(p => p.code === code)) {
        products.push({ code, name });
      }
    });

    res.json(products);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));