const XLSX = require('xlsx');

const RANGE_LABELS = ['95-100', '90-94', '80-89', '60-79', '50-59', 'below 50'];

function isNotOpted(val) {
  if (val === null || val === undefined || val === '') return true;
  const normalized = String(val).trim();
  return normalized === '-' || normalized === '—' || normalized === '–' || normalized === 'N/A' || normalized === 'NA' || normalized === '';
}

function findTargetColumn(headers) {
  const priorities = ['% in IX+30', 'Grand Total', '% in IX'];
  for (const column of priorities) {
    const found = headers.find((header) => header && header.toString().trim().toLowerCase() === column.toLowerCase());
    if (found) return found;
  }
  return null;
}

function isGeneratedColumnName(header) {
  return /^column\d+$/i.test(String(header || '').trim());
}

function findRollColumn(headers) {
  return headers.find((header) => {
    const lower = normalizeHeaderKey(header);
    return lower.includes('roll no') || lower === 'roll no' || lower === 'roll';
  }) || null;
}

function findEnrollmentColumn(headers) {
  return headers.find((header) => {
    const lower = normalizeHeaderKey(header);
    return lower.includes('admn') || lower.includes('admission') || lower.includes('adm no')
      || lower.includes('admission no') || lower.includes('enroll no')
      || lower.includes('enrollment') || lower.includes('enrolment')
      || lower.includes('reg no') || lower.includes('registration no');
  }) || null;
}

function findNameColumn(headers) {
  const explicitNameColumn = headers.find((header) => {
    const lower = String(header || '').toLowerCase();
    return lower.includes('name') && !lower.includes('father') && !lower.includes('mother');
  }) || null;
  if (explicitNameColumn) return explicitNameColumn;

  const admissionColumn = findAdmissionColumn(headers);
  const admissionIndex = admissionColumn ? headers.indexOf(admissionColumn) : -1;
  const genderIndex = headers.findIndex((header) => normalizeHeaderKey(header).includes('gender'));
  const fallbackColumn = admissionIndex >= 0 ? headers[admissionIndex + 1] : null;

  if (
    fallbackColumn
    && genderIndex === admissionIndex + 2
    && isGeneratedColumnName(fallbackColumn)
  ) {
    return fallbackColumn;
  }

  return null;
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
  return findEnrollmentColumn(headers) || findRollColumn(headers);
}

function findClass9Column(headers) {
  return headers.find((header) => {
    const lower = String(header || '').toLowerCase();
    return (lower.includes('% in ix') || lower.includes('ix %') || lower.includes('class 9') || lower.includes('ix percent') || lower.includes('ix marks') || lower.includes('9th %'))
      && !lower.includes('+30') && !lower.includes('target');
  }) || null;
}

function findTarget100Column(headers) {
  const priorities = ['% in IX+30', 'X Target', 'Target', 'Class X Target', 'Target %', 'Target Percentage'];
  for (const column of priorities) {
    const found = headers.find((header) => {
      if (!header) return false;
      const normalized = header.toString().trim().toLowerCase().replace(/\s+/g, ' ');
      return normalized === column.toLowerCase();
    });
    if (found) return found;
  }

  return headers.find((header) => {
    const lower = String(header || '').toLowerCase();
    const hasTarget = lower.includes('target') || lower.includes('+30');
    // Don't exclude 'ix' if it's paired with '+30' (common in target headers)
    const isIxBaselineOnly = lower.includes('ix') && !lower.includes('+30');
    return hasTarget && !isIxBaselineOnly;
  }) || null;
}

function findExamPercentColumn(headers) {
  // Prioritize exact '%' match as requested by the user
  const exactPercent = headers.find((h) => String(h || '').trim() === '%');
  if (exactPercent) return exactPercent;

  return headers.find((header) => {
    const lower = String(header || '').toLowerCase().trim();
    return lower === 'percentage' || (
      lower.includes('%')
      && !lower.includes('ix')
      && !lower.includes('target')
      && !lower.includes('+30')
    );
  }) || null;
}

function getStudentKey(row, headers) {
  const admissionCol = findAdmissionColumn(headers);
  const nameCol = findNameColumn(headers);
  const admission = admissionCol ? normalizeIdentifier(row[admissionCol]) : '';
  const name = nameCol ? String(row[nameCol] ?? '').trim() : '';
  const normalizedName = normalizeStudentName(name);

  if (admission) return `adm:${admission}`;
  if (normalizedName) return `name:${normalizedName}`;
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

function cleanStudentName(value) {
  return String(value ?? '')
    .replace(/\s+/g, ' ')
    .replace(/\s+\d{3}(?:\s+\d{3})+\s+(PASS|COMP|FAIL)\s*$/i, '')
    .trim();
}

function normalizeStudentName(value) {
  return cleanStudentName(value)
    .trim()
    .toLowerCase()
    .replace(/[.'`]+/g, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function removeFileExtension(name = '') {
  return String(name).replace(/\.[^.]+$/, '');
}

function isSummaryRow(row) {
  const values = Object.values(row || {}).map((value) => normalizeText(value));
  return values.some((value) => RANGE_LABELS.includes(value));
}

function getFirstMeaningfulCellValue(row = {}) {
  const values = Object.values(row);
  for (const value of values) {
    const normalized = normalizeText(value);
    if (normalized) return normalized;
  }
  return '';
}

function isRepeatedHeaderRow(row = {}, headers = []) {
  const normalizedHeaders = headers.map((header) => normalizeHeaderKey(header)).filter(Boolean);
  if (!normalizedHeaders.length) return false;

  const normalizedValues = Object.values(row).map((value) => normalizeHeaderKey(value)).filter(Boolean);
  if (!normalizedValues.length) return false;

  const overlap = normalizedValues.filter((value) => normalizedHeaders.includes(value)).length;
  return overlap >= Math.min(3, normalizedHeaders.length);
}

function isAggregateRow(row = {}, headers = []) {
  if (isSummaryRow(row) || isRepeatedHeaderRow(row, headers)) return true;

  const firstCell = getFirstMeaningfulCellValue(row);
  const aggregateLabels = [
    'average',
    'avg',
    'class teacher',
    'class incharge',
    'teacher',
    'coordinator',
    'principal',
    'signature',
    'remarks',
    'summary',
    'result',
    'rankwise',
    'rank wise',
    'topper',
    'toppers',
    'subject topper',
    'pass',
    'passed',
    'fail',
    'failed',
    'pass percentage',
    'overall',
    'grand total',
    'total',
  ];

  return aggregateLabels.some((label) => firstCell === label || firstCell.startsWith(`${label} `));
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
  const classSectionHeader = headers.find((header) => normalizeText(header).includes('class section'));
  let className = null;
  let sectionName = null;

  if (classSectionHeader) {
    const values = rows.map((row) => row[classSectionHeader]).filter(Boolean);
    for (const value of values) {
      const parsed = parseClassSection(value);
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
  const class9Col = findClass9Column(headers);
  const targetCol = findTarget100Column(headers);

  if (class9Col && targetCol) return 'BASELINE';
  if (normalizedCombined.includes('target sheet') || normalizedCombined.includes('baseline') || normalizedCombined.includes('class 9 target')) return 'BASELINE';
  if (normalizedCombined.includes('half yearly') || normalizedCombined.includes('halfyearly') || /\bhy\b/.test(normalizedCombined)) return 'HY';
  if (normalizedCombined.includes('preboard 2') || normalizedCombined.includes('pre board 2') || normalizedCombined.includes('preboard ii') || normalizedCombined.includes('pre board ii') || /\bpb2\b/.test(normalizedCombined)) return 'PB2';
  if (normalizedCombined.includes('preboard 1') || normalizedCombined.includes('pre board 1') || normalizedCombined.includes('preboard i') || normalizedCombined.includes('pre board i') || /\bpb1\b/.test(normalizedCombined)) return 'PB1';
  if (normalizedCombined.includes('cbse result') || (normalizedCombined.includes('cbse') && normalizedCombined.includes('result')) || normalizedCombined.includes('all subject wise report')) return 'BOARD';
  if (/\bboard\b/.test(normalizedCombined)) return 'BOARD';
  return 'UNKNOWN';
}

function countMarksLikeHeaders(row = []) {
  return row.filter((value) => /\(\s*\d+\s*\)|\b\d+\s*$/.test(String(value || '').trim())).length;
}

function detectHeaderRowIndex(matrix = []) {
  const scanLimit = Math.min(matrix.length, 12);
  let bestIndex = 0;
  let bestScore = -1;

  for (let rowIndex = 0; rowIndex < scanLimit; rowIndex += 1) {
    const row = (matrix[rowIndex] || []).map((cell) => String(cell ?? '').trim());
    const normalized = row.map(normalizeHeaderKey);
    const hasName = normalized.some((cell) => cell.includes('name') && !cell.includes('father') && !cell.includes('mother'));
    const hasEnroll = normalized.some((cell) =>
      cell.includes('enroll no') || cell.includes('enrollment') || cell.includes('enrolment')
      || cell.includes('admn') || cell.includes('admission') || cell.includes('roll no') || cell.includes('reg no')
    );
    const hasPercent = normalized.some((cell) => cell === '%' || cell === 'percentage' || cell.includes('%') || cell.includes('percent'));
    const hasGrandTotal = normalized.some((cell) => cell.includes('grand total') || cell === 'total');
    const hasBaseline = normalized.some((cell) => (cell.includes('ix') || cell.includes('9th') || cell.includes('class 9')) && !cell.includes('target'));
    const hasTarget = normalized.some((cell) => cell.includes('target') || cell.includes('+30'));
    const marksHeaders = countMarksLikeHeaders(row);

    let score = 0;
    if (hasName) score += 5;
    if (hasEnroll) score += 5;
    if (hasPercent) score += 3;
    if (hasBaseline) score += 4;
    if (hasTarget) score += 4;
    if (hasGrandTotal) score += 2;
    score += Math.min(marksHeaders, 8);

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

function getSubjectColumns(headers) {
  return headers.filter((header) => {
    const normalized = String(header || '').trim();
    const lower = normalized.toLowerCase();
    const hasTrailingMaxMarks = /(\(\s*\d+\s*\)|\d+\s*)$/.test(normalized);

    return !lower.includes('s.no') && !lower.includes('sr.')
      && !lower.includes('name') && !lower.includes('%')
      && !lower.includes('admn') && !lower.includes('admin')
      && !lower.includes('roll') && !lower.includes('rank')
      && !lower.includes('dob') && !lower.includes('date')
      && !lower.includes('father') && !lower.includes('mother')
      && !lower.includes('gender') && !lower.includes('enrollment')
      && !lower.includes('source_') && !lower.includes('source file')
      && !lower.includes('source sheet') && !lower.includes('unnamed')
      && !lower.includes('class section')
      && !lower.includes('grand total') && !lower.includes('total')
      && !lower.includes('column')
      && !lower.includes('+30') && !lower.includes('+ 30')
      && !lower.includes('ix 100') && !lower.includes('eng 100 ix')
      && !lower.includes('x target') && !lower.includes('analysis')
      && !lower.includes('target') && !lower.includes(' ix')
      && hasTrailingMaxMarks;
  });
}

function getMaxMarksFromHeader(header) {
  const match = String(header || '').match(/(\d+)(?:\s*\)|\s*)$/);
  return match ? parseInt(match[1], 10) : null;
}

function findSubjectEntriesMatchingTotal(entries, targetTotal) {
  if (!Number.isFinite(targetTotal) || entries.length === 0 || entries.length > 15) return null;

  let bestMatch = null;
  const maxMask = 1 << entries.length;

  for (let mask = 1; mask < maxMask; mask += 1) {
    let sum = 0;
    const picked = [];

    for (let index = 0; index < entries.length; index += 1) {
      if (mask & (1 << index)) {
        sum += entries[index].score;
        picked.push(entries[index]);
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
  const totalCol = headers.find((header) => String(header || '').toLowerCase().includes('grand total') || String(header || '').toLowerCase() === 'total');
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
  const totalCol = headers.find((header) => String(header || '').toLowerCase().includes('grand total') || String(header || '').toLowerCase() === 'total');
  if (!totalCol) return row;

  const entries = getContributingSubjectEntries(row, headers);
  const sum = entries.reduce((acc, entry) => acc + entry.score, 0);
  const maxMarks = entries.reduce((acc, entry) => acc + entry.maxScore, 0);
  const hasAny = entries.length > 0;

  row[totalCol] = hasAny ? sum : '';

  // Do NOT overwrite baseline columns during recalculation
  return row;
}

function extractStudentMetrics(row, headers, fallbackExamName = '') {
  const nameCol = findNameColumn(headers);
  const admissionCol = findAdmissionColumn(headers);
  const class9Col = findClass9Column(headers);
  const targetCol = findTarget100Column(headers);
  const totalCol = headers.find((header) => String(header || '').toLowerCase().includes('grand total') || String(header || '').toLowerCase() === 'total');
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

  const rawClass9 = toNumber(class9Col ? row[class9Col] : null);
  const rawTarget = toNumber(targetCol ? row[targetCol] : null);

  // Scale baseline metrics if they are in decimal format (e.g., 0.59 -> 59.0)
  let class9Percent = rawClass9;
  if (class9Percent !== null && class9Percent > 0 && class9Percent < 2) {
    class9Percent = parseFloat((class9Percent * 100).toFixed(2));
  }

  let targetPercent = rawTarget;
  if (targetPercent !== null && targetPercent > 0 && targetPercent < 2) {
    targetPercent = parseFloat((targetPercent * 100).toFixed(2));
  }

  // Cap target at 100 as requested
  if (targetPercent !== null && targetPercent > 100) {
    targetPercent = 100;
  }

  const explicitExamPercent = toNumber(examCol ? row[examCol] : null);
  const derivedExamPercent = maxMarks > 0 ? parseFloat(((obtainedMarks / maxMarks) * 100).toFixed(2)) : null;
  const examPercent = examStage === 'BASELINE' ? null : (explicitExamPercent ?? derivedExamPercent);

  return {
    studentKey: getStudentKey(row, headers),
    name: nameCol ? cleanStudentName(row[nameCol]) : '',
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

function isLikelyStudentDataRow(row = {}, headers = []) {
  const admissionCol = findAdmissionColumn(headers);
  const nameCol = findNameColumn(headers);
  const totalCol = headers.find((header) => String(header || '').toLowerCase().includes('grand total') || String(header || '').toLowerCase() === 'total');
  const examCol = findExamPercentColumn(headers);
  const subjectEntries = getContributingSubjectEntries(row, headers);

  const admissionNo = admissionCol ? normalizeIdentifier(row[admissionCol]) : '';
  const studentName = nameCol ? String(row[nameCol] ?? '').trim() : '';
  const totalValue = totalCol ? toNumber(row[totalCol]) : null;
  const examPercent = examCol ? toNumber(row[examCol]) : null;

  if (admissionNo) return true;
  if (studentName && (subjectEntries.length > 0 || totalValue !== null || examPercent !== null)) return true;
  return false;
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

function parseWorkbookBuffer(buffer, sourceFileName = 'Upload') {
  const workbook = XLSX.read(buffer, {
    type: 'buffer',
    cellDates: true,
    raw: false,
  });

  const sheetNames = [];
  const sheets = {};
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
      const emptyMeta = detectSheetMeta(name, [], [], sourceFileName);
      sheetNames.push(name);
      sheets[name] = { headers: [], rows: [], meta: emptyMeta };
      parsedSheets.push({ sheetName: name, headers: [], rows: [], meta: emptyMeta });
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

      if (!hasData) return null;
      recalcGrandTotal(rowData, headers);
      if (isAggregateRow(rowData, headers)) return null;
      if (!isLikelyStudentDataRow(rowData, headers)) return null;
      return rowData;
    }).filter(Boolean);

    const meta = {
      ...detectSheetMeta(name, headers, rows, sourceFileName),
      headerRowIndex,
      titleRow: headerRowIndex > 0 ? String((matrix[0] || [])[0] ?? '').trim() : '',
    };
    if (meta.titleRow && detectExamStage(meta.examName, headers) === 'UNKNOWN') {
      meta.examName = `${meta.examName} ${meta.titleRow}`.trim();
    }
    const validation = validateParsedSheet(name, headers, rows, meta);

    sheetNames.push(name);
    sheets[name] = { headers, rows, meta, validation };
    parsedSheets.push({ sheetName: name, headers, rows, meta, validation });
  });

  return { sheetNames, sheets, parsedSheets };
}

module.exports = {
  RANGE_LABELS,
  findTargetColumn,
  findNameColumn,
  normalizeHeaderKey,
  findRollColumn,
  findEnrollmentColumn,
  normalizeIdentifier,
  findAdmissionColumn,
  findClass9Column,
  findTarget100Column,
  findExamPercentColumn,
  getStudentKey,
  toNumber,
  normalizeText,
  cleanStudentName,
  normalizeStudentName,
  removeFileExtension,
  isSummaryRow,
  isAggregateRow,
  isLikelyStudentDataRow,
  normalizeSheetName,
  parseClassSection,
  detectSheetMeta,
  detectExamStage,
  countMarksLikeHeaders,
  detectHeaderRowIndex,
  buildHeadersFromRow,
  validateParsedSheet,
  getSubjectColumns,
  getMaxMarksFromHeader,
  findSubjectEntriesMatchingTotal,
  getContributingSubjectEntries,
  recalcGrandTotal,
  extractStudentMetrics,
  parseWorkbookBuffer,
};
