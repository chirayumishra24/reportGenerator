const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const cors = require('cors');
const { prisma, dbEnabled } = require('../config/database');
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

function findNameColumn(headers) {
  return headers.find(h => {
    const lower = String(h || '').toLowerCase();
    return lower.includes('name') && !lower.includes('father') && !lower.includes('mother');
  }) || null;
}

function normalizeHeaderKey(value) {
  return String(value ?? '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function normalizeIdentifier(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return '';
  const collapsed = raw.replace(/\s+/g, '');
  if (/^\d+(\.0+)?$/.test(collapsed)) {
    return String(parseInt(collapsed, 10));
  }
  return collapsed.toLowerCase().replace(/[^a-z0-9]/g, '');
}

function findAdmissionColumn(headers) {
  return headers.find(h => {
    const lower = normalizeHeaderKey(h);
    return lower.includes('admn') || lower.includes('admission') || lower.includes('adm no') ||
      lower.includes('admission no') || lower.includes('enroll no') ||
      lower.includes('enrollment') || lower.includes('enrolment') ||
      lower.includes('reg no') || lower.includes('registration no') || lower.includes('roll no') ||
      lower.includes('scholar no') || lower.includes('sch no') || lower.includes('student id');
  }) || null;
}

function findClass9Column(headers) {
  const priorities = [
    '% in IX', 'IX %', 'Class IX %', 'Class 9 %', 'IX Percentage', '9th %', 'Class 9th %',
    'IX 100', 'IX (100)', 'Percentage in IX', 'IX Percent', '9th Percentage'
  ];
  
  for (const p of priorities) {
    const found = headers.find(h => {
      const lower = String(h || '').toLowerCase().replace(/\s+/g, ' ').trim();
      return lower === p.toLowerCase();
    });
    if (found) return found;
  }

  return headers.find(h => {
    const lower = String(h || '').toLowerCase().replace(/\s+/g, ' ').trim();
    const isBaseline = (
      lower.includes('% in ix') || 
      lower.includes('ix %') || 
      lower.includes('class 9') || 
      lower.includes('class ix') ||
      lower.includes('9th class') || 
      lower.includes('ix percent') || 
      lower.includes('ix marks') || 
      lower.includes('baseline') ||
      lower.includes('class ix %') ||
      lower.includes('class 9 %') ||
      lower.includes('class 9th %') ||
      lower.includes('ix percentage') ||
      lower.includes('ix 100') ||
      lower.includes('ix 80') ||
      lower.includes('ix (100)') ||
      lower.includes('ix (80)') ||
      lower.includes('percentage in ix') ||
      lower.includes('percent in ix') ||
      (lower.includes('9th') && (lower.includes('%') || lower.includes('percentage') || lower.includes('marks'))) ||
      (lower.includes('ix') && (lower.includes('%') || lower.includes('percentage') || lower.includes('score') || lower.includes('marks')))
    );
    const isTarget = lower.includes('+30') || lower.includes('target') || lower.includes('+ 30') || lower.includes('projected') || lower.includes('improvement');
    return isBaseline && !isTarget;
  }) || null;
}

function findTarget100Column(headers) {
  const priorities = [
    '% in IX+30', 'X Target', 'Target', 'Class X Target', 'Target %', 
    'Target Percentage', 'IX+30', 'IX + 30', 'IX +30', 'IX+ 30',
    'X TARGET %', 'X TARGET PERCENTAGE', 'Target 100', 'X 100',
    'Projected %', 'Projected Target'
  ];
  for (const col of priorities) {
    const found = headers.find(h => {
      if (!h) return false;
      const normalized = String(h).trim().toLowerCase().replace(/\s+/g, ' ');
      return normalized === col.toLowerCase();
    });
    if (found) return found;
  }

  return headers.find(h => {
    const lower = String(h || '').toLowerCase().replace(/\s+/g, ' ');
    const hasTarget = lower.includes('target') || lower.includes('+30') || lower.includes('+ 30') || lower.includes('projected');
    const isIxBaselineOnly = (lower.includes('ix') || lower.includes('9th')) && !lower.includes('+30') && !lower.includes('+ 30') && !lower.includes('target') && !lower.includes('projected');
    return hasTarget && !isIxBaselineOnly;
  }) || null;
}

function findExamPercentColumn(headers) {
  return headers.find((h) => {
    const lower = String(h || '').toLowerCase().trim();
    return lower === '%' || lower === 'percentage' || (
      lower.includes('%') &&
      !lower.includes('ix') &&
      !lower.includes('target') &&
      !lower.includes('+30')
    );
  }) || null;
}

function getStudentKey(row, headers) {
  const admissionCol = findAdmissionColumn(headers);
  const nameCol = findNameColumn(headers);
  const admission = admissionCol ? normalizeIdentifier(row[admissionCol]) : '';
  const name = nameCol ? String(row[nameCol] ?? '').trim() : '';
  if (admission) return `adm:${admission.toLowerCase()}`;
  if (name) return `name:${name.toLowerCase()}`;
  return null;
}

function toNumber(value) {
  const num = parseFloat(value);
  return Number.isFinite(num) ? num : null;
}

function normalizeText(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function removeFileExtension(name = '') {
  return String(name).replace(/\.[^.]+$/, '');
}

function isSummaryRow(row) {
  const values = Object.values(row || {}).map(v => normalizeText(v));
  return values.some(v => RANGE_LABELS.includes(v));
}

function normalizeSheetName(name) {
  return normalizeText(name).replace(/[^a-z0-9]+/g, '-');
}

function parseClassSection(value) {
  const text = String(value || '').toUpperCase().replace(/[_]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!text) return { className: null, sectionName: null };
  const classTokens = ['XII', 'XI', 'IX', 'X', 'VIII', 'VII', 'VI', '12', '11', '10', '9', '8', '7', '6'];

  for (const token of text.split(' ')) {
    const compactMatch = token.match(/^(XII|XI|IX|X|VIII|VII|VI|12|11|10|9|8|7|6)([A-Z])$/);
    if (!compactMatch || classTokens.includes(token)) continue;
    return {
      className: compactMatch[1] || null,
      sectionName: compactMatch[2] || null,
    };
  }

  const patterns = [
    /\b(CLASS\s*)?(XII|XI|IX|X|VIII|VII|VI|12|11|10|9|8|7|6)\s*[-/ ]\s*([A-Z])\b/,
    /\b(CLASS\s*)?(XII|XI|IX|X|VIII|VII|VI|12|11|10|9|8|7|6)\b.*?\bSECTION\s*([A-Z])\b/,
    /\bSECTION\s*([A-Z])\b/,
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (!match) continue;
    if (pattern === patterns[2]) {
      return { className: null, sectionName: match[1] };
    }
    return {
      className: match[2] || null,
      sectionName: match[3] || null,
    };
  }

  const classOnly = text.match(/\b(XII|XI|IX|X|VIII|VII|VI|12|11|10|9|8|7|6)\b/);
  return {
    className: classOnly ? classOnly[1] : null,
    sectionName: null,
  };
}

function detectSheetMeta(sheetName, headers = [], rows = [], sourceFileName = '') {
  const classSectionHeader = headers.find(h => normalizeText(h).includes('class section'));
  let className = null;
  let sectionName = null;

  if (classSectionHeader) {
    const values = rows.map(r => r[classSectionHeader]).filter(Boolean);
    for (const val of values) {
      const parsed = parseClassSection(val);
      if (parsed.className || parsed.sectionName) {
        className = parsed.className || className;
        sectionName = parsed.sectionName || sectionName;
        break;
      }
    }
  }

  if (!className || !sectionName) {
    const fromSheetName = parseClassSection(sheetName);
    className = className || fromSheetName.className;
    sectionName = sectionName || fromSheetName.sectionName;
  }

  if (!className || !sectionName) {
    const fromFileName = parseClassSection(sourceFileName);
    className = className || fromFileName.className;
    sectionName = sectionName || fromFileName.sectionName;
  }

  return {
    examName: removeFileExtension(sourceFileName) || sheetName,
    className,
    sectionName,
  };
}

function detectExamStage(sheetName, headers = []) {
  const combined = `${sheetName} ${headers.join(' ')}`.toLowerCase();
  const normalizedCombined = combined.replace(/[^a-z0-9]+/g, ' ').trim();
  
  // 1. Explicit Frontend Prefixes (Absolute Priority)
  if (normalizedCombined.includes('baseline class10')) return 'BASELINE';
  if (normalizedCombined.includes('hy class10')) return 'HY';
  if (normalizedCombined.includes('pb1 class10')) return 'PB1';
  if (normalizedCombined.includes('pb2 class10')) return 'PB2';
  if (normalizedCombined.includes('board class10')) return 'BOARD';

  // 2. Specific Exam Stages (HY, then PB1, then PB2)
  if (normalizedCombined.includes('half yearly') || normalizedCombined.includes('halfyearly') || /\bhy\b/.test(normalizedCombined)) return 'HY';
  
  if (normalizedCombined.includes('preboard 1') || normalizedCombined.includes('pre board 1') || normalizedCombined.includes('preboard i') || normalizedCombined.includes('pre board i') || /\bpb1\b/.test(normalizedCombined)) return 'PB1';
  if (normalizedCombined.includes('preboard 2') || normalizedCombined.includes('pre board 2') || normalizedCombined.includes('preboard ii') || normalizedCombined.includes('pre board ii') || /\bpb2\b/.test(normalizedCombined)) return 'PB2';

  // 3. Board Results (Specific keywords)
  const isBoardResult = normalizedCombined.includes('cbse result') || 
                        (normalizedCombined.includes('cbse') && normalizedCombined.includes('result')) || 
                        normalizedCombined.includes('all subject wise report') ||
                        normalizedCombined.includes('board result') ||
                        normalizedCombined.includes('final result') ||
                        normalizedCombined.includes('annual result');
  
  // 4. Broad Board check, but excluding metadata columns that often mention "Board"
  const isMetadataSheet = normalizedCombined.includes('registration') || normalizedCombined.includes('roll number') || normalizedCombined.includes('roll list');
  if ((isBoardResult || /\bboard\b/.test(normalizedCombined)) && !isMetadataSheet) {
    return 'BOARD';
  }

  // 5. Baseline
  const class9Col = findClass9Column(headers);
  const targetCol = findTarget100Column(headers);
  if (class9Col && targetCol) return 'BASELINE';
  if (normalizedCombined.includes('target sheet') || normalizedCombined.includes('baseline') || normalizedCombined.includes('class 9 target')) return 'BASELINE';
  
  return 'UNKNOWN';
}

function countMarksLikeHeaders(row = []) {
  return row.filter((value) => /\(\s*\d+\s*\)|\b\d+\s*$/.test(String(value || '').trim())).length;
}

function detectHeaderRowIndex(matrix = []) {
  const scanLimit = Math.min(matrix.length, 8);
  let bestIndex = 0;
  let bestScore = -1;

  for (let rowIndex = 0; rowIndex < scanLimit; rowIndex += 1) {
    const row = (matrix[rowIndex] || []).map((cell) => String(cell ?? '').trim());
    const normalized = row.map(normalizeHeaderKey);
    const hasName = normalized.some((cell) => cell.includes('name') && !cell.includes('father') && !cell.includes('mother'));
    const hasEnroll = normalized.some((cell) =>
      cell.includes('enroll no') || cell.includes('enrollment') || cell.includes('enrolment') ||
      cell.includes('admn') || cell.includes('admission') || cell.includes('roll no') || cell.includes('reg no')
    );
    const hasPercent = normalized.some((cell) => cell === '%' || cell === 'percentage');
    const hasGrandTotal = normalized.some((cell) => cell.includes('grand total') || cell === 'total');
    const marksHeaders = countMarksLikeHeaders(row);

    let score = 0;
    if (hasName) score += 4;
    if (hasEnroll) score += 4;
    if (hasPercent) score += 2;
    if (hasGrandTotal) score += 2;
    score += Math.min(marksHeaders, 6);

    if (score > bestScore) {
      bestScore = score;
      bestIndex = rowIndex;
    }
  }

  return bestIndex;
}

function buildHeadersFromRow(rawHeaders = []) {
  return rawHeaders.map((value, index) => {
    const cellValue = String(value ?? '').trim();
    if (cellValue) return cellValue;
    return index === 0 ? 'S.No' : `Column${index + 1}`;
  });
}

function validateParsedSheet(sheetName, headers, rows, meta = {}) {
  const issues = [];
  const examStage = detectExamStage(meta.examName || sheetName, headers);
  const admissionCol = findAdmissionColumn(headers);
  const headerRowIndex = meta.headerRowIndex ?? null;

  if (headerRowIndex === null || headerRowIndex < 0) {
    issues.push(`Could not detect the header row for ${sheetName}.`);
  }
  if (examStage === 'UNKNOWN') {
    issues.push(`Could not detect exam stage for ${sheetName}.`);
  }
  if (!admissionCol) {
    issues.push(`Missing Enroll No. column in ${sheetName}.`);
  }
  if (examStage !== 'BASELINE' && !findExamPercentColumn(headers) && rows.length > 0) {
    const anyDerived = rows.some((row) => {
      const metrics = extractStudentMetrics(row, headers, meta.examName || sheetName);
      return metrics.examPercent !== null;
    });
    if (!anyDerived) {
      issues.push(`Could not determine exam percentage data in ${sheetName}.`);
    }
  }

  const duplicates = new Set();
  if (admissionCol) {
    const seen = new Set();
    rows.forEach((row) => {
      const normalized = normalizeIdentifier(row[admissionCol]);
      if (!normalized) return;
      if (seen.has(normalized)) duplicates.add(String(row[admissionCol]).trim());
      seen.add(normalized);
    });
  }
  if (duplicates.size > 0) {
    issues.push(`Duplicate Enroll No. values in ${sheetName}: ${Array.from(duplicates).slice(0, 5).join(', ')}`);
  }

  return {
    ok: issues.length === 0,
    issues,
    examStage,
    sectionName: meta.sectionName || '',
    headerRowIndex,
  };
}

function buildComparisonSummary(sheetNames, sheets) {
  const students = new Map();

  for (const sheetName of sheetNames) {
    const sheet = sheets[sheetName];
    if (!sheet?.rows?.length) continue;

    const headers = sheet.headers || [];
    const nameCol = findNameColumn(headers);
    const class9Col = findClass9Column(headers);
    const targetCol = findTarget100Column(headers);
    const examCol = findExamPercentColumn(headers);

    for (const row of sheet.rows) {
      const key = getStudentKey(row, headers);
      if (!key) continue;

      const existing = students.get(key) || {
        name: nameCol ? row[nameCol] : 'Student',
        class9: toNumber(class9Col ? row[class9Col] : null),
        target: toNumber(targetCol ? row[targetCol] : null),
        exams: [],
      };

      if (!existing.name && nameCol) existing.name = row[nameCol];
      if (existing.class9 === null && class9Col) existing.class9 = toNumber(row[class9Col]);
      if (existing.target === null && targetCol) existing.target = toNumber(row[targetCol]);

      const metrics = extractStudentMetrics(row, headers, sheetName);
      const examScore = metrics.examPercent;
      if (examScore !== null) {
        existing.exams.push({ sheetName, score: examScore });
      }

      students.set(key, existing);
    }
  }

  const studentRows = [];
  const summary = {
    totalStudents: 0,
    metTarget: 0,
    improvedTowardTarget: 0,
    behindTarget: 0,
  };

  students.forEach((student) => {
    const validExams = student.exams.filter(exam => exam.score !== null);
    if (validExams.length === 0 && student.class9 === null && student.target === null) return;

    summary.totalStudents += 1;
    const latestExam = validExams[validExams.length - 1] || null;
    const previousExam = validExams.length > 1 ? validExams[validExams.length - 2] : null;
    const latestScore = latestExam?.score ?? null;
    const target = student.target;
    const baseline = student.class9;

    let status = 'No Target';
    if (latestScore !== null && target !== null && latestScore >= target) {
      status = 'Achieved Target';
      summary.metTarget += 1;
    } else if (latestScore !== null && target !== null) {
      const reference = previousExam?.score ?? baseline;
      if (reference !== null && latestScore > reference) {
        status = 'Improving Toward Target';
        summary.improvedTowardTarget += 1;
      } else {
        status = 'Below Target';
        summary.behindTarget += 1;
      }
    } else if (latestScore !== null && baseline !== null && latestScore > baseline) {
      status = 'Improved from Class 9';
      summary.improvedTowardTarget += 1;
    } else {
      status = 'Needs Review';
      summary.behindTarget += 1;
    }

    studentRows.push({
      name: student.name || 'Student',
      class9: baseline ?? '',
      target: target ?? '',
      latestExam: latestExam?.sheetName || '',
      latestScore: latestScore ?? '',
      status,
    });
  });

  return { summary, studentRows };
}

function getSubjectColumns(headers) {
  return headers.filter(h => {
    const header = String(h || '').trim();
    const lowerH = header.toLowerCase();
    const hasTrailingMaxMarks = /\d+\s*$/.test(header);
    return !lowerH.includes('s.no') && !lowerH.includes('sr.') &&
      !lowerH.includes('name') && !lowerH.includes('%') &&
      !lowerH.includes('admn') && !lowerH.includes('admin') &&
      !lowerH.includes('roll') && !lowerH.includes('rank') &&
      !lowerH.includes('dob') && !lowerH.includes('date') &&
      !lowerH.includes('father') && !lowerH.includes('mother') &&
      !lowerH.includes('gender') && !lowerH.includes('enrollment') &&
      !lowerH.includes('source_') && !lowerH.includes('source file') &&
      !lowerH.includes('source sheet') && !lowerH.includes('unnamed') &&
      !lowerH.includes('class section') &&
      !lowerH.includes('grand total') && !lowerH.includes('total') &&
      !lowerH.includes('column') &&
      !lowerH.includes('+30') && !lowerH.includes('+ 30') &&
      !lowerH.includes('ix 100') && !lowerH.includes('eng 100 ix') &&
      !lowerH.includes('x target') && !lowerH.includes('analysis') &&
      !lowerH.includes('target') && !lowerH.includes(' ix') &&
      hasTrailingMaxMarks;
  });
}

function getMaxMarksFromHeader(header) {
  const match = String(header || '').match(/(\d+)\s*$/);
  return match ? parseInt(match[1], 10) : null;
}

function findSubjectEntriesMatchingTotal(entries, targetTotal) {
  if (!Number.isFinite(targetTotal) || entries.length === 0 || entries.length > 15) return null;

  let bestMatch = null;
  const maxMask = 1 << entries.length;
  for (let mask = 1; mask < maxMask; mask++) {
    let sum = 0;
    const picked = [];
    for (let i = 0; i < entries.length; i++) {
      if (mask & (1 << i)) {
        sum += entries[i].score;
        picked.push(entries[i]);
      }
    }
    if (Math.abs(sum - targetTotal) < 0.001) {
      if (!bestMatch || picked.length > bestMatch.length) {
        bestMatch = picked;
      }
    }
  }
  return bestMatch;
}

function getContributingSubjectEntries(row, headers) {
  const subjectCols = getSubjectColumns(headers);
  const totalCol = headers.find(h => h.toLowerCase().includes('grand total') || h.toLowerCase() === 'total');
  const reportedTotal = totalCol ? toNumber(row[totalCol]) : null;

  const entries = subjectCols.map((header) => {
    const value = row[header];
    if (isNotOpted(value)) return null;
    const score = toNumber(value);
    if (score === null) return null;
    return {
      header,
      score,
      maxScore: getMaxMarksFromHeader(header) || 80,
    };
  }).filter(Boolean);

  if (reportedTotal === null) return entries;

  const exactMatch = findSubjectEntriesMatchingTotal(entries, reportedTotal);
  return exactMatch || entries;
}

function recalcGrandTotal(row, headers) {
  const totalCol = headers.find(h => h.toLowerCase().includes('grand total') || h.toLowerCase() === 'total');
  if (!totalCol) return row;

  const entries = getContributingSubjectEntries(row, headers);
  const sum = entries.reduce((acc, entry) => acc + entry.score, 0);
  const maxMarks = entries.reduce((acc, entry) => acc + entry.maxScore, 0);
  const hasAny = entries.length > 0;

  row[totalCol] = hasAny ? sum : '';

  const pctCol = headers.find(h => h.toLowerCase().includes('% in ix') && !h.toLowerCase().includes('+30'));
  if (pctCol && hasAny && maxMarks > 0) {
    row[pctCol] = parseFloat(((sum / maxMarks) * 100).toFixed(2));
  }

  return row;
}

function extractStudentMetrics(row, headers, fallbackExamName = '') {
  const nameCol = findNameColumn(headers);
  const admissionCol = findAdmissionColumn(headers);
  const class9Col = findClass9Column(headers);
  const targetCol = findTarget100Column(headers);
  const totalCol = headers.find(h => h.toLowerCase().includes('grand total') || h.toLowerCase() === 'total');
  const examStage = detectExamStage(fallbackExamName, headers);
  const examCol = examStage === 'BASELINE' ? null : findExamPercentColumn(headers);

  const subjectBreakdown = getContributingSubjectEntries(row, headers).map((entry) => ({
    header: entry.header,
    subject: String(entry.header).replace(/\s+\d+$/, '').trim(),
    score: entry.score,
    maxScore: entry.maxScore,
  }));
  const obtainedMarks = subjectBreakdown.reduce((acc, entry) => acc + entry.score, 0);
  const maxMarks = subjectBreakdown.reduce((acc, entry) => acc + entry.maxScore, 0);

  const class9Percent = toNumber(class9Col ? row[class9Col] : null);
  const targetPercent = toNumber(targetCol ? row[targetCol] : null);
  const explicitExamPercent = toNumber(examCol ? row[examCol] : null);
  const derivedExamPercent = maxMarks > 0 ? parseFloat(((obtainedMarks / maxMarks) * 100).toFixed(2)) : null;
  const examPercent = examStage === 'BASELINE' ? null : (explicitExamPercent ?? derivedExamPercent);

  return {
    studentKey: getStudentKey(row, headers),
    name: nameCol ? row[nameCol] : '',
    admissionNo: admissionCol ? row[admissionCol] : '',
    class9Percent,
    targetPercent,
    totalValue: totalCol ? row[totalCol] : obtainedMarks,
    examPercent,
    obtainedMarks,
    maxMarks,
    subjectBreakdown,
    fallbackExamName,
  };
}

function parseWorkbookBuffer(buffer, sourceFileName = 'Upload') {
  const workbook = XLSX.read(buffer, {
    type: 'buffer',
    cellDates: true,
    raw: false,
  });

  const sheetNames = [];
  const result = {};
  const parsedSheets = [];

  workbook.SheetNames.forEach((name) => {
    const worksheet = workbook.Sheets[name];
    const matrix = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: '',
      blankrows: false,
      raw: false,
    });

    if (!matrix.length) {
      sheetNames.push(name);
      result[name] = { headers: [], rows: [] };
      parsedSheets.push({ sheetName: name, headers: [], rows: [], meta: detectSheetMeta(name, [], [], sourceFileName) });
      return;
    }

    const headerRowIndex = detectHeaderRowIndex(matrix);
    const rawHeaders = matrix[headerRowIndex] || [];
    const headers = buildHeadersFromRow(rawHeaders);

    const rows = matrix.slice(headerRowIndex + 1).map((rawRow) => {
      const rowData = {};
      let hasData = false;
      headers.forEach((header, index) => {
        let value = rawRow[index];
        if (typeof value === 'string') {
          const trimmed = value.trim();
          value = trimmed !== '' && !Number.isNaN(Number(trimmed)) ? Number(trimmed) : trimmed;
        }
        if (value !== '' && value !== null && value !== undefined) hasData = true;
        rowData[header] = value ?? '';
      });
      if (!hasData || isSummaryRow(rowData)) return null;
      recalcGrandTotal(rowData, headers);
      return rowData;
    }).filter(Boolean);

    const meta = {
      ...detectSheetMeta(name, headers, rows, sourceFileName),
      headerRowIndex,
      titleRow: headerRowIndex > 0 ? String((matrix[0] || [])[0] ?? '').trim() : '',
    };
    const validation = validateParsedSheet(name, headers, rows, meta);
    sheetNames.push(name);
    result[name] = { headers, rows, meta, validation };
    parsedSheets.push({ sheetName: name, headers, rows, meta, validation });
  });

  return { sheetNames, sheets: result, parsedSheets };
}

async function upsertStudentRecord(tx, metrics, meta) {
  const normalizedName = normalizeText(metrics.name || 'student');
  const admissionNo = String(metrics.admissionNo || '').trim() || null;
  const className = meta.className || null;
  const sectionName = meta.sectionName || null;

  if (admissionNo) {
    const existing = await tx.student.findFirst({ where: { admissionNo } });
    if (existing) {
      return tx.student.update({
        where: { id: existing.id },
        data: {
          name: metrics.name || existing.name,
          normalizedName,
          className,
          sectionName,
        },
      });
    }
  }

  return tx.student.upsert({
    where: {
      student_identity: {
        normalizedName,
        className,
        sectionName,
      },
    },
    update: {
      name: metrics.name || undefined,
      admissionNo: admissionNo || undefined,
    },
    create: {
      name: metrics.name || 'Student',
      normalizedName,
      admissionNo,
      className,
      sectionName,
    },
  });
}

function buildPerformanceStatus(performances) {
  if (!performances.length) return 'Needs Review';
  const latest = performances[performances.length - 1];
  const previous = performances.length > 1 ? performances[performances.length - 2] : null;
  const latestScore = latest.examPercent;
  const previousScore = previous?.examPercent ?? latest.class9Percent;
  const target = latest.targetPercent ?? null;

  if (latestScore !== null && target !== null && latestScore >= target) return 'Achieved Target';
  if (latestScore !== null && target !== null && previousScore !== null && latestScore > previousScore) return 'Improving Toward Target';
  if (latestScore !== null && previousScore !== null && latestScore > previousScore) return 'Improved';
  if (latestScore !== null && target !== null) return 'Below Target';
  return 'Needs Review';
}

async function persistParsedFiles(files) {
  if (!dbEnabled || !prisma) {
    throw new Error('Database is not configured. Add DATABASE_URL to enable persistent imports.');
  }

  const importSummary = [];

  for (const file of files) {
    const parsed = parseWorkbookBuffer(file.buffer, file.originalname);
    const validationErrors = parsed.parsedSheets.flatMap((sheet) =>
      (sheet.validation?.issues || []).map((issue) => `${file.originalname} / ${sheet.sheetName}: ${issue}`)
    );
    if (validationErrors.length > 0) {
      throw new Error(validationErrors.join(' | '));
    }
    const uploadBatch = await prisma.uploadBatch.create({
      data: { fileName: file.originalname },
    });

    for (const parsedSheet of parsed.parsedSheets) {
      const examSheet = await prisma.examSheet.create({
        data: {
          name: parsedSheet.sheetName,
          normalizedName: normalizeSheetName(parsedSheet.sheetName),
          examName: parsedSheet.meta.examName,
          className: parsedSheet.meta.className,
          sectionName: parsedSheet.meta.sectionName,
          headers: parsedSheet.headers,
          uploadBatchId: uploadBatch.id,
        },
      });

      for (const row of parsedSheet.rows) {
        const metrics = extractStudentMetrics(row, parsedSheet.headers, parsedSheet.meta.examName);
        if (!metrics.studentKey) continue;

        await prisma.$transaction(async (tx) => {
          const student = await upsertStudentRecord(tx, metrics, parsedSheet.meta);
          await tx.studentPerformance.upsert({
            where: {
              student_exam_unique: {
                studentId: student.id,
                examSheetId: examSheet.id,
              },
            },
            update: {
              class9Percent: metrics.class9Percent,
              targetPercent: metrics.targetPercent,
              examPercent: metrics.examPercent,
              obtainedMarks: metrics.obtainedMarks,
              maxMarks: metrics.maxMarks,
              totalValue: toNumber(metrics.totalValue),
              subjects: metrics.subjectBreakdown,
              rowData: row,
            },
            create: {
              studentId: student.id,
              examSheetId: examSheet.id,
              rowKey: metrics.studentKey,
              class9Percent: metrics.class9Percent,
              targetPercent: metrics.targetPercent,
              examPercent: metrics.examPercent,
              obtainedMarks: metrics.obtainedMarks,
              maxMarks: metrics.maxMarks,
              totalValue: toNumber(metrics.totalValue),
              subjects: metrics.subjectBreakdown,
              rowData: row,
            },
          });
        });
      }
    }

    importSummary.push({
      fileName: file.originalname,
      sheets: parsed.sheetNames.length,
      students: parsed.parsedSheets.reduce((acc, sheet) => acc + sheet.rows.length, 0),
    });
  }

  return importSummary;
}

function average(values) {
  const valid = values.filter(v => Number.isFinite(v));
  if (!valid.length) return null;
  return parseFloat((valid.reduce((sum, v) => sum + v, 0) / valid.length).toFixed(2));
}

function buildMasterCumulativeRows(students = []) {
  return students.map((student) => {
    const row = {
      studentId: student.id,
      enrollmentNo: student.admissionNo || '',
      name: student.name,
      section: student.sectionName || '',
      class9Percent: '',
      targetPercent: '',
      hyPercent: '',
      pb1Percent: '',
      pb2Percent: '',
      boardPercent: '',
      targetGap: '',
      improvement: '',
      status: 'Needs Review',
    };

    student.performances.forEach((perf) => {
      const stage = detectExamStage(perf.examSheet.examName || perf.examSheet.name, []);
      if (stage === 'BASELINE') {
        if (perf.class9Percent !== null) row.class9Percent = perf.class9Percent;
        if (perf.targetPercent !== null) row.targetPercent = perf.targetPercent;
      } else if (stage === 'HY' && perf.examPercent !== null) {
        row.hyPercent = perf.examPercent;
      } else if (stage === 'PB1' && perf.examPercent !== null) {
        row.pb1Percent = perf.examPercent;
      } else if (stage === 'PB2' && perf.examPercent !== null) {
        row.pb2Percent = perf.examPercent;
      } else if (stage === 'BOARD' && perf.examPercent !== null) {
        row.boardPercent = perf.examPercent;
      }
    });

    const latest = [row.boardPercent, row.pb2Percent, row.pb1Percent, row.hyPercent]
      .find(v => Number.isFinite(v)) ?? null;
    if (latest !== null && Number.isFinite(row.targetPercent)) {
      row.targetGap = parseFloat((latest - row.targetPercent).toFixed(2));
    }
    if (latest !== null && Number.isFinite(row.class9Percent)) {
      row.improvement = parseFloat((latest - row.class9Percent).toFixed(2));
    }
    if (latest !== null && Number.isFinite(row.targetPercent) && latest >= row.targetPercent) row.status = 'Achieved Target';
    else if (latest !== null && Number.isFinite(row.targetPercent) && Number.isFinite(row.class9Percent) && latest > row.class9Percent) row.status = 'Improving Toward Target';
    else if (latest !== null && Number.isFinite(row.class9Percent) && latest > row.class9Percent) row.status = 'Improved';
    else if (latest !== null && Number.isFinite(row.targetPercent)) row.status = 'Below Target';

    return row;
  }).sort((a, b) => {
    const av = [a.boardPercent, a.pb2Percent, a.pb1Percent, a.hyPercent].find(v => Number.isFinite(v)) ?? -1;
    const bv = [b.boardPercent, b.pb2Percent, b.pb1Percent, b.hyPercent].find(v => Number.isFinite(v)) ?? -1;
    return bv - av;
  });
}

async function buildCumulativeReport() {
  if (!dbEnabled || !prisma) {
    return {
      databaseEnabled: false,
      summary: { uploads: 0, sheets: 0, students: 0, performances: 0 },
      studentComparison: [],
      classComparison: [],
      sectionComparison: [],
      examTimeline: [],
    };
  }

  const [uploads, sheets, students] = await Promise.all([
    prisma.uploadBatch.findMany({ orderBy: { createdAt: 'desc' } }),
    prisma.examSheet.findMany({
      orderBy: { createdAt: 'asc' },
      include: { performances: { include: { student: true } } },
    }),
    prisma.student.findMany({
      include: {
        performances: {
          include: { examSheet: true },
          orderBy: { createdAt: 'asc' },
        },
      },
    }),
  ]);

  const studentComparison = students.map((student) => {
    const performances = student.performances
      .filter(perf => perf.examPercent !== null || perf.class9Percent !== null || perf.targetPercent !== null)
      .map((perf) => ({
        examName: perf.examSheet.examName || perf.examSheet.name,
        sheetName: perf.examSheet.name,
        examPercent: perf.examPercent,
        class9Percent: perf.class9Percent,
        targetPercent: perf.targetPercent,
      }));
    const latest = performances[performances.length - 1] || null;
    return {
      studentId: student.id,
      name: student.name,
      admissionNo: student.admissionNo,
      className: student.className,
      sectionName: student.sectionName,
      class9Percent: latest?.class9Percent ?? null,
      targetPercent: latest?.targetPercent ?? null,
      latestExam: latest?.examName ?? '',
      latestScore: latest?.examPercent ?? null,
      status: buildPerformanceStatus(performances),
      exams: performances,
    };
  }).sort((a, b) => (b.latestScore ?? -1) - (a.latestScore ?? -1));

  const classMap = new Map();
  const sectionMap = new Map();
  const timeline = [];

  sheets.forEach((sheet) => {
    const scores = sheet.performances.map(p => p.examPercent).filter(v => Number.isFinite(v));
    timeline.push({
      examName: sheet.examName || sheet.name,
      sheetName: sheet.name,
      className: sheet.className || 'Unassigned',
      sectionName: sheet.sectionName || 'N/A',
      avgPercent: average(scores),
      students: scores.length,
      createdAt: sheet.createdAt,
    });

    const classKey = sheet.className || 'Unassigned';
    if (!classMap.has(classKey)) {
      classMap.set(classKey, { className: classKey, scores: [], students: new Set(), sections: new Set() });
    }
    const classEntry = classMap.get(classKey);
    scores.forEach(score => classEntry.scores.push(score));
    sheet.performances.forEach(perf => classEntry.students.add(perf.studentId));
    if (sheet.sectionName) classEntry.sections.add(sheet.sectionName);

    const sectionKey = `${classKey}__${sheet.sectionName || 'N/A'}`;
    if (!sectionMap.has(sectionKey)) {
      sectionMap.set(sectionKey, { className: classKey, sectionName: sheet.sectionName || 'N/A', scores: [], students: new Set(), exams: [] });
    }
    const sectionEntry = sectionMap.get(sectionKey);
    scores.forEach(score => sectionEntry.scores.push(score));
    sheet.performances.forEach(perf => sectionEntry.students.add(perf.studentId));
    sectionEntry.exams.push({
      examName: sheet.examName || sheet.name,
      avgPercent: average(scores),
      students: scores.length,
    });
  });

  return {
    databaseEnabled: true,
    summary: {
      uploads: uploads.length,
      sheets: sheets.length,
      students: students.length,
      performances: sheets.reduce((sum, sheet) => sum + sheet.performances.length, 0),
    },
    studentComparison,
    classComparison: Array.from(classMap.values()).map((entry) => ({
      className: entry.className,
      sections: entry.sections.size,
      students: entry.students.size,
      avgPercent: average(entry.scores),
    })).sort((a, b) => (b.avgPercent ?? -1) - (a.avgPercent ?? -1)),
    sectionComparison: Array.from(sectionMap.values()).map((entry) => ({
      className: entry.className,
      sectionName: entry.sectionName,
      students: entry.students.size,
      avgPercent: average(entry.scores),
      exams: entry.exams,
    })).sort((a, b) => (b.avgPercent ?? -1) - (a.avgPercent ?? -1)),
    examTimeline: timeline.sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt)),
    masterCumulativeSheet: {
      headers: ['Enrollment No', 'Student Name', 'Section', 'Class 9 %', 'Target %', 'HY %', 'PB1 %', 'PB2 %', 'Board %', 'Target Gap', 'Improvement', 'Status'],
      rows: buildMasterCumulativeRows(students).map((row) => ({
        'Enrollment No': row.enrollmentNo,
        'Student Name': row.name,
        'Section': row.section,
        'Class 9 %': row.class9Percent,
        'Target %': row.targetPercent,
        'HY %': row.hyPercent,
        'PB1 %': row.pb1Percent,
        'PB2 %': row.pb2Percent,
        'Board %': row.boardPercent,
        'Target Gap': row.targetGap,
        'Improvement': row.improvement,
        'Status': row.status,
      })),
    },
  };
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
    const parsed = parseWorkbookBuffer(req.file.buffer, req.file.originalname);
    const issues = parsed.parsedSheets.flatMap((sheet) =>
      (sheet.validation?.issues || []).map((issue) => ({
        sheetName: sheet.sheetName,
        message: issue,
      }))
    );
    res.json({
      sheetNames: parsed.sheetNames,
      sheets: parsed.sheets,
      parsedSheets: parsed.parsedSheets,
      validationPassed: issues.length === 0,
      issues,
    });
  } catch (err) {
    console.error('Parse error:', err);
    res.status(500).json({ error: err.message });
  }
};
app.post('/parse', upload.single('file'), handleParse);
app.post('/api/parse', upload.single('file'), handleParse);

app.get('/db-status', async (req, res) => {
  if (!dbEnabled || !prisma) {
    return res.json({ enabled: false, message: 'DATABASE_URL not configured' });
  }
  try {
    await prisma.$queryRaw`SELECT 1`;
    return res.json({ enabled: true });
  } catch (err) {
    console.error('DB status check failed:', err.message);
    return res.status(500).json({ enabled: false, error: err.message });
  }
});
app.get('/api/db-status', async (req, res) => {
  if (!dbEnabled || !prisma) {
    return res.json({ enabled: false, message: 'DATABASE_URL not configured' });
  }
  try {
    await prisma.$queryRaw`SELECT 1`;
    return res.json({ enabled: true });
  } catch (err) {
    console.error('DB status check failed:', err.message);
    return res.status(500).json({ enabled: false, error: err.message });
  }
});

const handlePersistentImport = async (req, res) => {
  try {
    if (!dbEnabled || !prisma) {
      return res.status(400).json({ error: 'Database is not configured. Set DATABASE_URL and run Prisma migration first.' });
    }
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }

    const summary = await persistParsedFiles(req.files);
    const cumulativeReport = await buildCumulativeReport();
    return res.json({
      message: 'Files imported into cumulative database successfully.',
      imported: summary,
      cumulativeReport,
    });
  } catch (err) {
    console.error('Persistent import failed:', err);
    return res.status(500).json({ error: err.message });
  }
};
app.post('/import-persistent', upload.array('files'), handlePersistentImport);
app.post('/api/import-persistent', upload.array('files'), handlePersistentImport);

const handleCumulativeReport = async (req, res) => {
  try {
    const cumulativeReport = await buildCumulativeReport();
    return res.json(cumulativeReport);
  } catch (err) {
    console.error('Cumulative report failed:', err);
    return res.status(500).json({ error: err.message });
  }
};
app.get('/cumulative-report', handleCumulativeReport);
app.get('/api/cumulative-report', handleCumulativeReport);

app.get('/student-history/:studentId', async (req, res) => {
  try {
    if (!dbEnabled || !prisma) {
      return res.status(400).json({ error: 'Database is not configured.' });
    }
    const student = await prisma.student.findUnique({
      where: { id: req.params.studentId },
      include: {
        performances: {
          include: { examSheet: true },
          orderBy: { createdAt: 'asc' },
        },
      },
    });
    if (!student) return res.status(404).json({ error: 'Student not found' });
    return res.json(student);
  } catch (err) {
    console.error('Student history failed:', err);
    return res.status(500).json({ error: err.message });
  }
});
app.get('/api/student-history/:studentId', async (req, res) => {
  try {
    if (!dbEnabled || !prisma) {
      return res.status(400).json({ error: 'Database is not configured.' });
    }
    const student = await prisma.student.findUnique({
      where: { id: req.params.studentId },
      include: {
        performances: {
          include: { examSheet: true },
          orderBy: { createdAt: 'asc' },
        },
      },
    });
    if (!student) return res.status(404).json({ error: 'Student not found' });
    return res.json(student);
  } catch (err) {
    console.error('Student history failed:', err);
    return res.status(500).json({ error: err.message });
  }
});

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

    const comparison = buildComparisonSummary(sheetNames, sheets);
    if (comparison.studentRows.length > 0) {
      const summarySheet = workbook.addWorksheet('STUDENT COMPARISON');

      const titleRow = summarySheet.addRow(['Student Target Progress Summary']);
      titleRow.getCell(1).font = { bold: true, size: 14, color: { argb: 'FF1E293B' } };
      summarySheet.mergeCells(1, 1, 1, 6);

      summarySheet.addRow([]);
      summarySheet.addRow(['Total Students', comparison.summary.totalStudents]);
      summarySheet.addRow(['Achieved Target', comparison.summary.metTarget]);
      summarySheet.addRow(['Improving Toward Target', comparison.summary.improvedTowardTarget]);
      summarySheet.addRow(['Below Target / Needs Review', comparison.summary.behindTarget]);
      summarySheet.addRow([]);

      const headerRow = summarySheet.addRow([
        'Student Name',
        'Class 9 Marks % (Out of 80 based result)',
        'Class 10 Target % (Out of 100)',
        'Latest Exam',
        'Latest Score %',
        'Status',
      ]);
      headerRow.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F6EF7' } };
      });

      comparison.studentRows.forEach((row) => {
        summarySheet.addRow([
          row.name,
          row.class9,
          row.target,
          row.latestExam,
          row.latestScore,
          row.status,
        ]);
      });

      summarySheet.columns = [
        { width: 28 },
        { width: 24 },
        { width: 24 },
        { width: 24 },
        { width: 18 },
        { width: 26 },
      ];
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
if (require.main === module && process.env.NODE_ENV !== 'production') {
  const PORT = process.env.PORT || 5000;
  app.listen(PORT, () => {
    console.log(`Backend running on http://localhost:${PORT}`);
  });
}

module.exports = {
  legacyApp: app,
  parseWorkbookBuffer,
  persistParsedFiles,
  buildCumulativeReport,
};
