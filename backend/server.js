const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

const upload = multer({ storage: multer.memoryStorage() });

// POST /parse — reads workbook from memory buffer and returns JSON
app.post('/api/parse', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    const wb = XLSX.read(req.file.buffer, { type: 'buffer' });

    const result = {};
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      const data = XLSX.utils.sheet_to_json(ws, { defval: '' });
      const headers = data.length > 0 ? Object.keys(data[0]) : [];
      result[name] = { headers, rows: data };
    }

    res.json({ sheetNames: wb.SheetNames, sheets: result });
  } catch (err) {
    console.error('Parse error:', err);
    res.status(500).json({ error: err.message });
  }
});

// POST /export — receives modified data and returns xlsx
app.post('/api/export', (req, res) => {
  try {
    const { sheetNames, sheets, analysisSheet } = req.body;

    const wb = XLSX.utils.book_new();

    // Write each sheet
    for (const name of sheetNames) {
      const sheetData = sheets[name];
      if (!sheetData || !sheetData.rows) continue;
      const ws = XLSX.utils.json_to_sheet(sheetData.rows);
      // Set column widths
      if (sheetData.headers) {
        ws['!cols'] = sheetData.headers.map(h => ({
          wch: Math.max(h.length + 2, 12)
        }));
      }
      XLSX.utils.book_append_sheet(wb, ws, name);
    }

    // Add analysis sheet if provided
    if (analysisSheet) {
      const aoa = [];
      aoa.push(['ANALYSIS OF TARGETS SECTION WISE']);
      aoa.push(analysisSheet.headers);
      for (const row of analysisSheet.rows) {
        aoa.push(analysisSheet.headers.map(h => row[h] !== undefined ? row[h] : ''));
      }
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws['!cols'] = analysisSheet.headers.map((_, i) => ({ wch: i === 0 ? 14 : 12 }));
      XLSX.utils.book_append_sheet(wb, ws, 'GENERATED ANALYSIS');
    }

    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Disposition', 'attachment; filename="Report.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.document');
    res.send(Buffer.from(buf));
  } catch (err) {
    console.error('Export error:', err);
    res.status(500).json({ error: err.message });
  }
});

// For local development
if (process.env.NODE_ENV !== 'production') {
  const PORT = process.env.PORT || 5000;
  app.listen(PORT, () => {
    console.log(`Backend running on http://localhost:${PORT}`);
  });
}

// Export the app for Vercel Serverless
module.exports = app;
