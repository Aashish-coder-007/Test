const express = require('express');
const fs = require('fs');
const XLSX = require('xlsx');
const cors = require('cors');
const app = express();

app.use(cors());
app.use(express.json()); // For parsing JSON request bodies

const EXCEL_FILE = './data.xlsx';

// Helper: Read Excel file and convert to JSON
function readExcel() {
  if (!fs.existsSync(EXCEL_FILE)) {
    // Create empty workbook with a sheet if missing
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([["id", "name", "age"]]);
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, EXCEL_FILE);
  }
  const workbook = XLSX.readFile(EXCEL_FILE);
  const sheetName = workbook.SheetNames[0];
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
  return data;
}

// Helper: Write JSON data to Excel file
function writeExcel(data) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  XLSX.writeFile(wb, EXCEL_FILE);
}

// GET /data - send all Excel data as JSON
app.get('/data', (req, res) => {
  const data = readExcel();
  res.json(data);
});

// POST /data - update Excel data with posted JSON array
app.post('/data', (req, res) => {
  const newData = req.body;
  if (!Array.isArray(newData)) {
    return res.status(400).json({ error: 'Expected an array of records' });
  }
  writeExcel(newData);
  res.json({ message: 'Excel updated successfully' });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server listening on port ${PORT}`));
