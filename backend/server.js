const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const cors = require('cors');
let ChartJSNodeCanvas;
let canvasAvailable = false;
try {
  ChartJSNodeCanvas = require('chartjs-node-canvas').ChartJSNodeCanvas;
  canvasAvailable = true;
} catch (e) {
  console.warn('chartjs-node-canvas not available (likely serverless environment). Charts will be skipped in exports.');
}

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

// Healthcheck
const healthCheck = (req, res) => res.status(200).json({ ok: true, message: 'Backend is running!' });
app.get('/health', healthCheck);
app.get('/api/health', healthCheck);

const upload = multer({ storage: multer.memoryStorage() });

// ========== Chart Generation Config ==========
const CHART_WIDTH = 800;
const CHART_HEIGHT = 400;
let chartCanvas = null;
if (canvasAvailable) {
  try {
    chartCanvas = new ChartJSNodeCanvas({ 
      width: CHART_WIDTH, 
      height: CHART_HEIGHT,
      backgroundColour: 'white',
    });
  } catch (e) {
    console.warn('Failed to create chart canvas:', e.message);
  }
}

const CHART_COLORS = ['#6366f1', '#ef4444', '#22c55e', '#3b82f6', '#f59e0b', '#8b5cf6', '#ec4899', '#14b8a6'];
const RANGE_LABELS = ['95-100', '90-94', '80-89', '60-79', '50-59', 'below 50'];
const RANGE_BAR_COLORS = ['#22c55e', '#3b82f6', '#8b5cf6', '#f59e0b', '#f97316', '#ef4444'];
const RANGES = [
  { label: '95-100', min: 95, max: 100 },
  { label: '90-94', min: 90, max: 94.999 },
  { label: '80-89', min: 80, max: 89.999 },
  { label: '60-79', min: 60, max: 79.999 },
  { label: '50-59', min: 50, max: 59.999 },
  { label: 'below 50', min: 0, max: 49.999 },
];

// ========== Helper Functions ==========
function extractCellValue(cell) {
  let val = cell.value;
  if (val === null || val === undefined) return '';
  if (typeof val === 'object') {
    if (val.result !== undefined && val.result !== null) return val.result;
    if (val.formula) return val.result !== undefined ? val.result : '';
    if (val.sharedFormula) return val.result !== undefined ? val.result : '';
    if (val.richText && Array.isArray(val.richText)) return val.richText.map(t => t.text || '').join('');
    if (val instanceof Date) return val.toLocaleDateString('en-IN');
    if (val.text) return val.text;
    if (val.hyperlink) return val.text || val.hyperlink;
    return String(val);
  }
  return val;
}

function isNotOpted(val) {
  if (val === null || val === undefined || val === '') return true;
  const s = String(val).trim();
  return s === '-' || s === '—' || s === '–' || s === 'N/A' || s === 'NA' || s === '';
}

function findTargetColumn(headers) {
  const priorities = ['% in IX+30', 'Grand Total', '% in IX'];
  for (const col of priorities) {
    const found = headers.find(h => h && h.toString().trim().toLowerCase() === col.toLowerCase());
    if (found) return found;
  }
  return null;
}

function getSubjectColumns(headers) {
  return headers.filter(h => {
    const lowerH = h.toLowerCase();
    return !lowerH.includes('s.no') && !lowerH.includes('sr.') &&
      !lowerH.includes('name') && !lowerH.includes('%') &&
      !lowerH.includes('admn') && !lowerH.includes('admin') &&
      !lowerH.includes('roll') && !lowerH.includes('rank') &&
      !lowerH.includes('dob') && !lowerH.includes('date') &&
      !lowerH.includes('father') && !lowerH.includes('mother') &&
      !lowerH.includes('gender') && !lowerH.includes('enrollment') &&
      !lowerH.includes('grand total') && !lowerH.includes('total') &&
      !lowerH.includes('column') &&
      !lowerH.includes('+30') && !lowerH.includes('+ 30') &&
      !lowerH.includes('ix 100') && !lowerH.includes('x target');
  });
}

// Compute score distribution for a sheet
function computeDistribution(rows, targetCol) {
  const dist = [];
  let totalStudents = 0;
  for (const range of RANGES) {
    let count = 0;
    for (const student of rows) {
      const v = parseFloat(student[targetCol]);
      if (!isNaN(v) && v >= range.min && v <= range.max) count++;
    }
    dist.push({ range: range.label, count });
    totalStudents += count;
  }
  return { dist, totalStudents };
}

// ========== Server-Side Chart Generators ==========

// Generate a bar chart for score distribution of a single section
async function generateDistributionChart(sheetName, dist) {
  if (!chartCanvas) return null;
  const config = {
    type: 'bar',
    data: {
      labels: dist.map(d => d.range),
      datasets: [{
        label: sheetName,
        data: dist.map(d => d.count),
        backgroundColor: RANGE_BAR_COLORS,
        borderColor: RANGE_BAR_COLORS.map(c => c),
        borderWidth: 1,
        borderRadius: 6,
      }]
    },
    options: {
      responsive: false,
      plugins: {
        title: { display: true, text: `${sheetName} — Score Distribution`, font: { size: 16, weight: 'bold' }, color: '#1e293b' },
        legend: { display: true, position: 'bottom' },
      },
      scales: {
        y: { beginAtZero: true, ticks: { stepSize: 1, font: { size: 12 } }, title: { display: true, text: 'No. of Students' } },
        x: { ticks: { font: { size: 12 } }, title: { display: true, text: 'Score Range' } },
      },
    },
  };
  return await chartCanvas.renderToBuffer(config);
}

// Generate a grouped bar chart comparing all sections
async function generateSectionComparisonChart(sectionData) {
  if (!chartCanvas) return null;
  const datasets = sectionData.map((sec, i) => ({
    label: sec.name,
    data: sec.dist.map(d => d.count),
    backgroundColor: CHART_COLORS[i % CHART_COLORS.length] + 'CC',
    borderColor: CHART_COLORS[i % CHART_COLORS.length],
    borderWidth: 1,
    borderRadius: 4,
  }));

  const config = {
    type: 'bar',
    data: {
      labels: RANGE_LABELS,
      datasets,
    },
    options: {
      responsive: false,
      plugins: {
        title: { display: true, text: 'Section-wise Comparison — Score Distribution', font: { size: 16, weight: 'bold' }, color: '#1e293b' },
        legend: { display: true, position: 'bottom' },
      },
      scales: {
        y: { beginAtZero: true, ticks: { stepSize: 1, font: { size: 11 } }, title: { display: true, text: 'No. of Students' } },
        x: { ticks: { font: { size: 11 } }, title: { display: true, text: 'Score Range' } },
      },
    },
  };
  return await chartCanvas.renderToBuffer(config);
}

// Generate a line chart for trend comparison across sections
async function generateTrendChart(sectionData) {
  if (!chartCanvas) return null;
  const datasets = sectionData.map((sec, i) => ({
    label: sec.name,
    data: sec.dist.map(d => d.count),
    borderColor: CHART_COLORS[i % CHART_COLORS.length],
    backgroundColor: CHART_COLORS[i % CHART_COLORS.length] + '33',
    borderWidth: 3,
    pointRadius: 5,
    pointBackgroundColor: CHART_COLORS[i % CHART_COLORS.length],
    fill: false,
    tension: 0.3,
  }));

  const config = {
    type: 'line',
    data: {
      labels: RANGE_LABELS,
      datasets,
    },
    options: {
      responsive: false,
      plugins: {
        title: { display: true, text: 'Section-wise Trend Analysis', font: { size: 16, weight: 'bold' }, color: '#1e293b' },
        legend: { display: true, position: 'bottom' },
      },
      scales: {
        y: { beginAtZero: true, ticks: { stepSize: 1, font: { size: 11 } }, title: { display: true, text: 'No. of Students' } },
        x: { ticks: { font: { size: 11 } } },
      },
    },
  };
  return await chartCanvas.renderToBuffer(config);
}

// Generate a pie/doughnut chart for overall distribution
async function generatePieChart(sectionData) {
  if (!canvasAvailable) return null;
  // Aggregate all sections
  const totals = RANGES.map((_, i) => {
    let sum = 0;
    sectionData.forEach(sec => { sum += sec.dist[i].count; });
    return sum;
  });

  const config = {
    type: 'doughnut',
    data: {
      labels: RANGE_LABELS,
      datasets: [{
        data: totals,
        backgroundColor: RANGE_BAR_COLORS.map(c => c + 'CC'),
        borderColor: RANGE_BAR_COLORS,
        borderWidth: 2,
      }]
    },
    options: {
      responsive: false,
      plugins: {
        title: { display: true, text: 'Overall Score Distribution — All Sections', font: { size: 16, weight: 'bold' }, color: '#1e293b' },
        legend: { display: true, position: 'bottom', labels: { font: { size: 12 } } },
      },
    },
  };
  // Use a square canvas for pie chart
  const pieCanvas = new ChartJSNodeCanvas({ width: 500, height: 400, backgroundColour: 'white' });
  return await pieCanvas.renderToBuffer(config);
}

// Generate subject-wise average comparison across sections
async function generateSubjectComparisonChart(sheetsData, sheetNames) {
  if (!canvasAvailable) return null;
  // Find common subject columns from the first sheet
  const firstSheet = sheetsData[sheetNames[0]];
  if (!firstSheet) return null;
  const subjectCols = getSubjectColumns(firstSheet.headers);
  if (subjectCols.length === 0) return null;

  const labels = subjectCols.map(h => h.replace(/\s*\+\s*30/g, '').replace(/\s+\d+$/, '').trim());
  const datasets = sheetNames.map((name, i) => {
    const sheet = sheetsData[name];
    if (!sheet) return null;
    const avgs = subjectCols.map(subject => {
      let sum = 0, count = 0;
      for (const student of sheet.rows) {
        const val = student[subject];
        if (isNotOpted(val)) continue;
        const v = parseFloat(val);
        if (!isNaN(v)) { sum += v; count++; }
      }
      return count > 0 ? parseFloat((sum / count).toFixed(1)) : 0;
    });
    return {
      label: name,
      data: avgs,
      backgroundColor: CHART_COLORS[i % CHART_COLORS.length] + 'CC',
      borderColor: CHART_COLORS[i % CHART_COLORS.length],
      borderWidth: 1,
      borderRadius: 4,
    };
  }).filter(Boolean);

  if (datasets.length === 0) return null;

  const config = {
    type: 'bar',
    data: { labels, datasets },
    options: {
      responsive: false,
      plugins: {
        title: { display: true, text: 'Subject-wise Average Comparison Across Sections', font: { size: 15, weight: 'bold' }, color: '#1e293b' },
        legend: { display: true, position: 'bottom' },
      },
      scales: {
        y: { beginAtZero: true, ticks: { font: { size: 10 } }, title: { display: true, text: 'Average Score' } },
        x: { ticks: { font: { size: 10 }, maxRotation: 45 } },
      },
    },
  };

  const wideCanvas = new ChartJSNodeCanvas({ width: 900, height: 450, backgroundColour: 'white' });
  return await wideCanvas.renderToBuffer(config);
}

// ========== Color mapping for Excel cells ==========
const RANGE_COLORS_EXCEL = {
  '95-100': { bg: 'FFD4EDDA', font: 'FF155724' },
  '90-94': { bg: 'FFDBEAFE', font: 'FF1E3A8A' },
  '80-89': { bg: 'FFEDE9FE', font: 'FF5B21B6' },
  '60-79': { bg: 'FFFFF3CD', font: 'FF856404' },
  '50-59': { bg: 'FFFED7AA', font: 'FF9A3412' },
  'below 50': { bg: 'FFF8D7DA', font: 'FF721C24' },
  'Total': { bg: 'FFE0F2FE', font: 'FF0C4A6E' },
};

// ========== PARSE ==========
const handleParse = async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(req.file.buffer);

    const sheetNames = [];
    const result = {};

    workbook.eachSheet((worksheet) => {
      const name = worksheet.name;
      sheetNames.push(name);
      const headers = [];
      const rows = [];

      const headerRow = worksheet.getRow(1);
      let maxCol = 0;
      headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        if (colNumber > maxCol) maxCol = colNumber;
      });

      for (let col = 1; col <= maxCol; col++) {
        const cell = headerRow.getCell(col);
        let val = extractCellValue(cell);
        if (typeof val === 'object' || val === '' || val === null || val === undefined) {
          if (col === 1) val = 'S.No';
          else val = `Column${col}`;
        }
        headers.push(String(val).trim());
      }

      for (let i = 2; i <= worksheet.rowCount; i++) {
        const row = worksheet.getRow(i);
        const rowData = {};
        let hasData = false;

        headers.forEach((h, idx) => {
          const cell = row.getCell(idx + 1);
          let val = extractCellValue(cell);
          if (typeof val === 'string' && val.trim() !== '' && !isNaN(Number(val.trim()))) {
            val = Number(val.trim());
          }
          if (val !== '' && val !== null && val !== undefined) hasData = true;
          rowData[h] = val !== null && val !== undefined ? val : '';
        });

        if (hasData) rows.push(rowData);
      }

      result[name] = { headers, rows };
    });

    res.json({ sheetNames, sheets: result });
  } catch (err) {
    console.error('Parse error:', err);
    res.status(500).json({ error: err.message });
  }
};
app.post('/parse', upload.single('file'), handleParse);
app.post('/api/parse', upload.single('file'), handleParse);

// ========== EXPORT ==========
const handleExport = async (req, res) => {
  try {
    const { sheetNames, sheets } = req.body;

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Target Analysis Report Generator';
    workbook.created = new Date();

    // ====== 1. Write each data sheet ======
    for (const name of sheetNames) {
      const sheetData = sheets[name];
      if (!sheetData || !sheetData.rows) continue;

      const ws = workbook.addWorksheet(name);
      const headers = sheetData.headers || Object.keys(sheetData.rows[0] || {});

      const headerRow = ws.addRow(headers);
      headerRow.eachCell((cell) => {
        cell.font = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F6EF7' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFD1D5DB' } },
          bottom: { style: 'thin', color: { argb: 'FFD1D5DB' } },
          left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
          right: { style: 'thin', color: { argb: 'FFD1D5DB' } },
        };
      });
      headerRow.height = 24;

      for (const row of sheetData.rows) {
        const dataRow = ws.addRow(headers.map(h => row[h] !== undefined ? row[h] : ''));
        dataRow.eachCell((cell, colNumber) => {
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFE5E7EB' } },
            bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } },
            left: { style: 'thin', color: { argb: 'FFE5E7EB' } },
            right: { style: 'thin', color: { argb: 'FFE5E7EB' } },
          };
          cell.alignment = { vertical: 'middle' };
          const headerName = headers[colNumber - 1] || '';
          if (headerName.toLowerCase().includes('%') || headerName.toLowerCase().includes('total')) {
            const v = parseFloat(cell.value);
            if (!isNaN(v)) {
              if (v >= 95) cell.font = { color: { argb: 'FF15803D' }, bold: true };
              else if (v >= 90) cell.font = { color: { argb: 'FF1D4ED8' }, bold: true };
              else if (v >= 80) cell.font = { color: { argb: 'FF7C3AED' }, bold: true };
              else if (v >= 60) cell.font = { color: { argb: 'FFD97706' }, bold: true };
              else if (v >= 50) cell.font = { color: { argb: 'FFEA580C' }, bold: true };
              else cell.font = { color: { argb: 'FFDC2626' }, bold: true };
            }
          }
        });
      }

      ws.columns.forEach((col, idx) => { col.width = Math.max((headers[idx] || '').length + 4, 12); });
    }

    // ====== 2. AUTO-GENERATE per-sheet analysis + charts ======
    const allSectionData = []; // for cumulative report

    for (const name of sheetNames) {
      const sheetData = sheets[name];
      if (!sheetData || !sheetData.rows || sheetData.rows.length === 0) continue;

      const targetCol = findTargetColumn(sheetData.headers);
      if (!targetCol) continue;

      const { dist, totalStudents } = computeDistribution(sheetData.rows, targetCol);
      allSectionData.push({ name, dist, totalStudents });

      // Create analysis sheet
      const analysisName = `${name} Analysis`.substring(0, 31);
      const ws = workbook.addWorksheet(analysisName);

      // Title
      const titleRow = ws.addRow([`SCORE DISTRIBUTION — ${name}`]);
      titleRow.getCell(1).font = { bold: true, size: 13, color: { argb: 'FF1E293B' } };
      titleRow.getCell(1).alignment = { horizontal: 'center' };
      titleRow.height = 28;
      ws.mergeCells(1, 1, 1, 2);

      ws.addRow([]);

      // Table headers
      const hRow = ws.addRow(['Score Range', 'No. of Students']);
      hRow.eachCell((cell) => {
        cell.font = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F6EF7' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
      });
      hRow.height = 24;

      // Data rows
      for (const d of dist) {
        const dataRow = ws.addRow([d.range, d.count]);
        dataRow.eachCell((cell, colNumber) => {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFD1D5DB' } },
            bottom: { style: 'thin', color: { argb: 'FFD1D5DB' } },
            left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
            right: { style: 'thin', color: { argb: 'FFD1D5DB' } },
          };
          if (colNumber === 1 && RANGE_COLORS_EXCEL[d.range]) {
            const rc = RANGE_COLORS_EXCEL[d.range];
            cell.font = { bold: true, color: { argb: rc.font } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rc.bg } };
          }
        });
      }

      // Total row
      const totalRowData = ws.addRow(['Total', totalStudents]);
      totalRowData.eachCell((cell) => {
        cell.font = { bold: true, size: 11 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0F2FE' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
      });

      ws.addRow([]);
      ws.addRow([`Target Column: ${targetCol}`]).getCell(1).font = { italic: true, size: 10, color: { argb: 'FF64748B' } };

      ws.getColumn(1).width = 18;
      ws.getColumn(2).width = 18;

      // Generate and embed the chart image
      try {
        const chartBuffer = await generateDistributionChart(name, dist);
        if (chartBuffer) {
          const imageId = workbook.addImage({ buffer: chartBuffer, extension: 'png' });
          const startRow = dist.length + 8;
          ws.addImage(imageId, {
            tl: { col: 0, row: startRow },
            ext: { width: CHART_WIDTH, height: CHART_HEIGHT },
          });
        }
      } catch (chartErr) {
        console.warn(`Chart generation failed for ${name}:`, chartErr.message);
      }
    }

    // ====== 3. CUMULATIVE SECTION COMPARISON SHEET ======
    if (allSectionData.length > 0) {
      const ws = workbook.addWorksheet('SECTION ANALYSIS');

      // Title
      const titleRow = ws.addRow(['CUMULATIVE ANALYSIS — ALL SECTIONS COMPARISON']);
      const colCount = allSectionData.length + 3; // Range + sections + Total Students + %
      titleRow.getCell(1).font = { bold: true, size: 14, color: { argb: 'FF1E293B' } };
      titleRow.getCell(1).alignment = { horizontal: 'center' };
      titleRow.height = 32;
      if (colCount > 1) ws.mergeCells(1, 1, 1, colCount);

      ws.addRow([]);

      // Headers: Range | Section1 | Section2 | ... | Total Students | %
      const tableHeaders = ['Range', ...allSectionData.map(s => s.name), 'Total Students', '%'];
      const hRow = ws.addRow(tableHeaders);
      hRow.eachCell((cell) => {
        cell.font = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F6EF7' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
      });
      hRow.height = 24;

      // Grand total for percentage
      const grandTotal = allSectionData.reduce((sum, s) => sum + s.totalStudents, 0);

      // Data rows
      for (let ri = 0; ri < RANGES.length; ri++) {
        const rangeLabel = RANGES[ri].label;
        const rowValues = [rangeLabel];
        let rowTotal = 0;
        allSectionData.forEach(sec => {
          rowValues.push(sec.dist[ri].count);
          rowTotal += sec.dist[ri].count;
        });
        rowValues.push(rowTotal);
        rowValues.push(grandTotal > 0 ? `${((rowTotal / grandTotal) * 100).toFixed(1)}%` : '0%');

        const dataRow = ws.addRow(rowValues);
        dataRow.eachCell((cell, colNumber) => {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFD1D5DB' } },
            bottom: { style: 'thin', color: { argb: 'FFD1D5DB' } },
            left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
            right: { style: 'thin', color: { argb: 'FFD1D5DB' } },
          };
          if (colNumber === 1 && RANGE_COLORS_EXCEL[rangeLabel]) {
            const rc = RANGE_COLORS_EXCEL[rangeLabel];
            cell.font = { bold: true, color: { argb: rc.font } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rc.bg } };
          }
        });
      }

      // Total row
      const totalRowValues = ['Total'];
      allSectionData.forEach(sec => totalRowValues.push(sec.totalStudents));
      totalRowValues.push(grandTotal);
      totalRowValues.push('100%');
      const totalRow = ws.addRow(totalRowValues);
      totalRow.eachCell((cell) => {
        cell.font = { bold: true, size: 11 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0F2FE' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
      });

      // Timestamp
      ws.addRow([]);
      const tsRow = ws.addRow(['Generated on: ' + new Date().toLocaleDateString('en-IN', {
        year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit'
      })]);
      tsRow.getCell(1).font = { italic: true, size: 9, color: { argb: 'FF94A3B8' } };

      ws.columns.forEach((col, idx) => { col.width = idx === 0 ? 16 : 14; });

      // ====== Generate ALL comparison charts and embed them ======
      let chartRow = RANGES.length + 9;

      // Chart 1: Grouped bar comparison
      try {
        const barBuf = await generateSectionComparisonChart(allSectionData);
        if (barBuf) {
          const imgId = workbook.addImage({ buffer: barBuf, extension: 'png' });
          ws.addImage(imgId, { tl: { col: 0, row: chartRow }, ext: { width: CHART_WIDTH, height: CHART_HEIGHT } });
          chartRow += 22;
        }
      } catch (e) { console.warn('Bar comparison chart failed:', e.message); }

      // Chart 2: Trend line
      try {
        const lineBuf = await generateTrendChart(allSectionData);
        if (lineBuf) {
          const imgId = workbook.addImage({ buffer: lineBuf, extension: 'png' });
          ws.addImage(imgId, { tl: { col: 0, row: chartRow }, ext: { width: CHART_WIDTH, height: CHART_HEIGHT } });
          chartRow += 22;
        }
      } catch (e) { console.warn('Trend chart failed:', e.message); }

      // Chart 3: Pie/doughnut
      try {
        const pieBuf = await generatePieChart(allSectionData);
        if (pieBuf) {
          const imgId = workbook.addImage({ buffer: pieBuf, extension: 'png' });
          ws.addImage(imgId, { tl: { col: 0, row: chartRow }, ext: { width: 500, height: 400 } });
          chartRow += 22;
        }
      } catch (e) { console.warn('Pie chart failed:', e.message); }

      // Chart 4: Subject-wise comparison
      try {
        const subjectBuf = await generateSubjectComparisonChart(sheets, sheetNames);
        if (subjectBuf) {
          const imgId = workbook.addImage({ buffer: subjectBuf, extension: 'png' });
          ws.addImage(imgId, { tl: { col: 0, row: chartRow }, ext: { width: 900, height: 450 } });
        }
      } catch (e) { console.warn('Subject comparison chart failed:', e.message); }
    }

    // Write to buffer and send
    const buf = await workbook.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="Report.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
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

module.exports = app;
