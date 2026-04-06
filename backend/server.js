const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

// Healthcheck
const healthCheck = (req, res) => res.status(200).json({ ok: true, message: 'Backend is running on Vercel!' });
app.get('/health', healthCheck);
app.get('/api/health', healthCheck);

const upload = multer({ storage: multer.memoryStorage() });

// POST /parse — reads workbook from memory buffer and returns JSON
const handleParse = (req, res) => {
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
};
app.post('/parse', upload.single('file'), handleParse);
app.post('/api/parse', upload.single('file'), handleParse);

// POST /export — receives modified data and returns xlsx
const handleExport = (req, res) => {
  try {
    const { sheetNames, sheets, analysisSheet, perSheetAnalysis } = req.body;

    const wb = XLSX.utils.book_new();

    // Write each data sheet
    for (const name of sheetNames) {
      const sheetData = sheets[name];
      if (!sheetData || !sheetData.rows) continue;
      const ws = XLSX.utils.json_to_sheet(sheetData.rows);
      if (sheetData.headers) {
        ws['!cols'] = sheetData.headers.map(h => ({
          wch: Math.max(h.length + 2, 12)
        }));
      }
      XLSX.utils.book_append_sheet(wb, ws, name);
    }

    // Add cross-section analysis sheet
    if (analysisSheet) {
      const aoa = [];
      aoa.push(['ANALYSIS OF TARGETS — SECTION WISE']);
      aoa.push([]);
      aoa.push(analysisSheet.headers);
      for (const row of analysisSheet.rows) {
        aoa.push(analysisSheet.headers.map(h => row[h] !== undefined ? row[h] : ''));
      }
      aoa.push([]);
      aoa.push(['Generated on: ' + new Date().toLocaleDateString('en-IN', { 
        year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit' 
      })]);
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws['!cols'] = analysisSheet.headers.map((_, i) => ({ wch: i === 0 ? 16 : 14 }));
      // Merge the title row across all columns
      ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: analysisSheet.headers.length - 1 } }];
      XLSX.utils.book_append_sheet(wb, ws, 'SECTION ANALYSIS');
    }

    // Add per-sheet analysis sheets
    if (perSheetAnalysis) {
      for (const name of sheetNames) {
        const sa = perSheetAnalysis[name];
        if (!sa || !sa.rows) continue;

        const aoa = [];
        aoa.push([`SCORE DISTRIBUTION — ${name}`]);
        aoa.push([]);
        aoa.push(['Score Range', 'No. of Students']);
        for (const row of sa.rows) {
          aoa.push([row.Range, row.Count]);
        }
        aoa.push([]);
        aoa.push([`Total Students: ${sa.totalStudents}`]);
        aoa.push([`Target Column: ${sa.targetCol}`]);

        // Trim sheet name to fit Excel's 31 char limit
        const analysisName = `${name} Analysis`.substring(0, 31);
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        ws['!cols'] = [{ wch: 16 }, { wch: 16 }];
        ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
        XLSX.utils.book_append_sheet(wb, ws, analysisName);
      }
    }

    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Disposition', 'attachment; filename="Report.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.document');
    res.send(Buffer.from(buf));
  } catch (err) {
    console.error('Export error:', err);
    res.status(500).json({ error: err.message });
  }
};
app.post('/export', handleExport);
app.post('/api/export', handleExport);

// For local development
if (process.env.NODE_ENV !== 'production') {
  const PORT = process.env.PORT || 5000;
  app.listen(PORT, () => {
    console.log(`Backend running on http://localhost:${PORT}`);
  });
}

// Export the app for Vercel Serverless
module.exports = app;
