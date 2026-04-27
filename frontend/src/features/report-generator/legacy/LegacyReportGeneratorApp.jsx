import React, { useState, useRef, useMemo, useCallback, useEffect } from 'react';
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell, RadarChart, Radar, PolarGrid, PolarAngleAxis, PolarRadiusAxis
} from 'recharts';
import {
  Upload, FileSpreadsheet, CheckCircle, Download, BarChart3, ArrowLeft, X,
  Loader2, Info, Plus, Trash2, UserCheck, FileDown, RotateCcw, Eye, Edit3, Save,
  Printer, Search, TrendingUp, TrendingDown, Award, Users, Undo2, Redo2, ArrowUpDown, FileText,
  ChevronUp, ChevronDown, AlertCircle, CheckCircle2, History, Network, Calendar, Database
} from 'lucide-react';
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';
import {
  API,
  SESSION_STORAGE_KEY,
  SESSION_VERSION,
  CUMULATIVE_SHEET_NAME,
  CLASS10_SECTIONS,
  PHASE_ONE_CUMULATIVE_ONLY,
  RANGES,
  CHART_COLORS,
} from '../reportGenerator.config.js';

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

function toSafeNumber(value) {
  const num = parseFloat(value);
  return Number.isFinite(num) ? num : null;
}

function findClass9Column(headers) {
  const priorities = [
    'Class 9th %', 'Class IX %', 'IX %', '% in IX', 'Class 9 %', 'IX Percentage', '9th %', '9 %', 'Class 9th Percentage',
    'IX(100)', 'IX (100)', 'IX_PERCENT', 'CLASS_IX_PERC'
  ];
  
  for (const p of priorities) {
    const found = headers.find(h => {
      const normalized = String(h || '').trim().toLowerCase().replace(/\s+/g, ' ');
      return normalized === p.toLowerCase();
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
    'Target %', 'Target Percentage', 'X Target', 'Class 10 Target', 'Target 100', 'Target (100)',
    'Target % (X)', 'X_TARGET', 'TARGET_PERCENT'
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
  const priorities = [
    'Board %', 'Board Percentage', 'Final %', 'Final Percentage', 'CBSE %', 'Annual %',
    '%', 'Percentage', 'Exam %', 'Grand Total %', 'Agg %', 'Agg. %', 'Total %'
  ];
  const exactPercent = headers.find((h) => String(h || '').trim() === '%');
  if (exactPercent) return exactPercent;

  return headers.find((h) => {
    const lower = String(h || '').toLowerCase().trim();
    return lower === 'percentage' || (
      lower.includes('%') &&
      !lower.includes('ix') &&
      !lower.includes('target') &&
      !lower.includes('+30')
    );
  }) || null;
}

function normalizeText(value) {
  return String(value ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function normalizeStudentName(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/[.'`]+/g, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function parseClassSectionText(value) {
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
      return {
        className: null,
        sectionName: match[1] || null,
      };
    }

    return {
      className: match[2] || null,
      sectionName: match[3] || null,
    };
  }

  return { className: null, sectionName: null };
}

function simplifySheetDisplayName(sheetName) {
  const raw = String(sheetName || '').trim();
  if (!raw) return 'Sheet';
  if (raw === CUMULATIVE_SHEET_NAME) return 'Cumulative Class 10';

  const examStage = detectExamStage(raw, []);
  const { sectionName } = parseClassSectionText(raw);
  const normalized = raw
    .replace(/\.[^.]+$/, '')
    .replace(/\bSheet\s*\d+\b/gi, '')
    .replace(/\bRankwise\b/gi, '')
    .replace(/\bRANKWISE\b/gi, '')
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (examStage === 'BASELINE') return sectionName ? `Baseline - ${sectionName}` : 'Baseline';
  if (examStage === 'HY') return sectionName ? `HY - ${sectionName}` : 'HY';
  if (examStage === 'PB1') return sectionName ? `PB1 - ${sectionName}` : 'PB1';
  if (examStage === 'PB2') return sectionName ? `PB2 - ${sectionName}` : 'PB2';
  if (examStage === 'BOARD') {
    if (/best five/i.test(raw)) return 'Board - Best Five';
    if (/rank/i.test(raw)) return 'Board - Rankwise';
    return 'Board';
  }

  return normalized || raw;
}

function countMarksLikeHeaders(row = []) {
  return row.filter((value) => /\(\s*\d+\s*\)|\b\d+\s*$/.test(String(value || '').trim())).length;
}

function detectHeaderRowIndex(matrix = []) {
  const scanLimit = Math.min(matrix.length, 8);
  let bestIndex = 0;
  let bestScore = -1;

  for (const rowIndex = 0; rowIndex < scanLimit; rowIndex += 1) {
    const row = (matrix[rowIndex] || []).map((cell) => String(cell ?? '').trim());
    const normalized = row.map(normalizeHeaderKey);
    const hasName = normalized.some((cell) => cell.includes('name') && !cell.includes('father') && !cell.includes('mother'));
    const hasEnroll = normalized.some((cell) =>
      cell.includes('enroll no') || cell.includes('enrollment') || cell.includes('enrolment') ||
      cell.includes('admn') || cell.includes('admission') || cell.includes('roll no') || cell.includes('reg no')
    );
    const hasPercent = normalized.some((cell) => cell === '%' || cell === 'percentage' || cell.includes('%') || cell.includes('percent'));
    const hasGrandTotal = normalized.some((cell) => cell.includes('grand total') || cell === 'total');
    const hasBaseline = normalized.some((cell) => (cell.includes('ix') || cell.includes('9th') || cell.includes('class 9')) && !cell.includes('target') && !cell.includes('+30'));
    const hasTarget = normalized.some((cell) => (cell.includes('target') || cell.includes('+30')) && !((cell.includes('ix') || cell.includes('9th')) && !cell.includes('+30')));
    const marksHeaders = countMarksLikeHeaders(row);

    let score = 0;
    if (hasName) score += 4;
    if (hasEnroll) score += 4;
    if (hasPercent) score += 2;
    if (hasBaseline) score += 2;
    if (hasTarget) score += 2;
    if (hasGrandTotal) score += 2;
    score += Math.min(marksHeaders, 6);

    if (score > bestScore) {
      bestScore = score;
      bestIndex = rowIndex;
    }
  }

  return bestIndex;
}

function buildHeadersFromRow(rawHeaders = [], prevRow = []) {
  return rawHeaders.map((value, index) => {
    const cellValue = String(value ?? '').trim();
    const contextValue = String(prevRow[index] ?? '').trim();
    
    const isGeneric = ['%', 'marks', 'total', 'grand total', 'target', '100', '80', 'percent', 'percentage'].includes(cellValue.toLowerCase());
    if (isGeneric && contextValue && !['s.no', 'name', 'enrollment', 'admission'].includes(contextValue.toLowerCase())) {
      return `${contextValue} ${cellValue}`;
    }
    
    if (cellValue) return cellValue;
    return index === 0 ? 'S.No' : `Column${index + 1}`;
  });
}

function detectExamStage(sheetName, headers = []) {
  const combined = `${sheetName} ${headers.join(' ')}`.toLowerCase();
  const normalizedCombined = combined.replace(/[^a-z0-9]+/g, ' ').trim();
  
  if (normalizedCombined.includes('baseline class10')) return 'BASELINE';
  if (normalizedCombined.includes('hy class10')) return 'HY';
  if (normalizedCombined.includes('pb1 class10')) return 'PB1';
  if (normalizedCombined.includes('pb2 class10')) return 'PB2';
  if (normalizedCombined.includes('board class10')) return 'BOARD';

  if (normalizedCombined.includes('half yearly') || normalizedCombined.includes('halfyearly') || /\bhy\b/.test(normalizedCombined)) return 'HY';
  
  if (normalizedCombined.includes('preboard 1') || normalizedCombined.includes('pre board 1') || normalizedCombined.includes('preboard i') || normalizedCombined.includes('pre board i') || /\bpb1\b/.test(normalizedCombined)) return 'PB1';
  if (normalizedCombined.includes('preboard 2') || normalizedCombined.includes('pre board 2') || normalizedCombined.includes('preboard ii') || normalizedCombined.includes('pre board ii') || /\bpb2\b/.test(normalizedCombined)) return 'PB2';

  const isBoardResult = normalizedCombined.includes('cbse result') || 
                        (normalizedCombined.includes('cbse') && normalizedCombined.includes('result')) || 
                        normalizedCombined.includes('all subject wise report') ||
                        normalizedCombined.includes('board result') ||
                        normalizedCombined.includes('final result') ||
                        normalizedCombined.includes('annual result');
  
  const isMetadataSheet = normalizedCombined.includes('registration') || normalizedCombined.includes('roll number') || normalizedCombined.includes('roll list');
  if ((isBoardResult || /\bboard\b/.test(normalizedCombined)) && !isMetadataSheet) {
    return 'BOARD';
  }

  const class9Col = findClass9Column(headers);
  const targetCol = findTarget100Column(headers);
  if (class9Col && targetCol) return 'BASELINE';
  if (normalizedCombined.includes('target sheet') || normalizedCombined.includes('baseline') || normalizedCombined.includes('class 9 target')) return 'BASELINE';
  
  return 'UNKNOWN';
}

function detectSection(row, headers, sheetName) {
  const classSectionHeader = headers.find(h => normalizeText(h).includes('class section') || normalizeText(h) === 'section');
  const fromRow = classSectionHeader ? parseClassSectionText(row[classSectionHeader]) : { className: null, sectionName: null };
  if (fromRow.sectionName) return fromRow.sectionName;
  const fromSheet = parseClassSectionText(sheetName);
  return fromSheet.sectionName || '';
}

function validateParsedSheet(sheetName, sheet) {
  const headers = sheet?.headers || [];
  const rows = sheet?.rows || [];
  const meta = sheet?.meta || {};
  const issues = [];
  const examStage = detectExamStage(meta.examName || sheetName, headers);
  const admissionCol = findAdmissionColumn(headers);

  if (meta.headerRowIndex === undefined || meta.headerRowIndex === null || meta.headerRowIndex < 0) {
    issues.push(`Could not detect the header row for ${sheetName}.`);
  }
  if (examStage === 'UNKNOWN') {
    issues.push(`Could not detect exam stage for ${sheetName}.`);
  }
  if (!admissionCol) {
    issues.push(`Missing Enroll No. column in ${sheetName}.`);
  }
  if (examStage !== 'BASELINE' && !findExamPercentColumn(headers)) {
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
    headerRowIndex: meta.headerRowIndex ?? null,
  };
}

function computeStudentIdentity(row, headers, sheetName) {
  const admissionCol = findAdmissionColumn(headers);
  const nameCol = findNameColumn(headers);
  const enrollmentNo = admissionCol ? String(row[admissionCol] ?? '').trim() : '';
  const normalizedEnrollmentNo = normalizeIdentifier(enrollmentNo);
  const name = nameCol ? String(row[nameCol] ?? '').trim() : '';
  const normalizedName = normalizeStudentName(name);
  const section = detectSection(row, headers, sheetName);
  const normalizedSection = section.toLowerCase();
  const enrollmentKey = normalizedEnrollmentNo ? `enrollment:${normalizedEnrollmentNo}` : '';
  const nameSectionKey = normalizedName ? `${normalizedName}|${normalizedSection}` : '';
  const nameOnlyKey = normalizedName || '';
  const key = enrollmentKey || nameSectionKey || nameOnlyKey;
  return { key, enrollmentNo, normalizedEnrollmentNo, name, normalizedName, section, normalizedSection, enrollmentKey, nameSectionKey, nameOnlyKey };
}

function choosePreferredText(currentValue = '', nextValue = '') {
  const current = String(currentValue || '').trim();
  const next = String(nextValue || '').trim();

  if (!current) return next;
  if (!next) return current;
  return next.length > current.length ? next : current;
}

function setAliasMatch(aliasMap, aliasKey, studentKey) {
  if (!aliasKey || !studentKey) return;
  if (!aliasMap.has(aliasKey)) {
    aliasMap.set(aliasKey, new Set());
  }
  aliasMap.get(aliasKey).add(studentKey);
}

function resolveUniqueAlias(aliasMap, aliasKey) {
  if (!aliasKey || !aliasMap.has(aliasKey)) return null;
  const matches = aliasMap.get(aliasKey);
  if (!matches || matches.size !== 1) return null;
  return Array.from(matches)[0];
}

function getStageScore(entry, stage) {
  if (stage === 'HY') return entry['HY %'];
  if (stage === 'PB1') return entry['PB1 %'];
  if (stage === 'PB2') return entry['PB2 %'];
  if (stage === 'BOARD') return entry['Board %'];
  return '';
}

function setStageScore(entry, stage, value) {
  if (stage === 'HY') entry['HY %'] = value;
  if (stage === 'PB1') entry['PB1 %'] = value;
  if (stage === 'PB2') entry['PB2 %'] = value;
  if (stage === 'BOARD') entry['Board %'] = value;
}

function buildClass10CumulativeSheet(sheetNames, sheets) {
  const students = new Map();
  const nameSectionAliases = new Map();
  const nameOnlyAliases = new Map();

  const resolveStudentKey = (identity) => {
    if (identity.enrollmentKey) {
      return identity.enrollmentKey;
    }

    const exactSectionMatch = resolveUniqueAlias(nameSectionAliases, identity.nameSectionKey);
    if (exactSectionMatch) return exactSectionMatch;

    return resolveUniqueAlias(nameOnlyAliases, identity.nameOnlyKey);
  };

  const getOrCreateStudent = (studentKey, identity) => {
    const existing = students.get(studentKey) || {
      'Enrollment No': identity.enrollmentNo,
      'Student Name': identity.name || '',
      Section: identity.section,
      'Class 9 %': '',
      'Target %': '',
      'HY %': '',
      'PB1 %': '',
      'PB2 %': '',
      'Board %': '',
    };

    if (!existing['Enrollment No'] && identity.enrollmentNo) existing['Enrollment No'] = identity.enrollmentNo;
    existing['Student Name'] = choosePreferredText(existing['Student Name'], identity.name);
    existing.Section = choosePreferredText(existing.Section, identity.section);

    students.set(studentKey, existing);
    return existing;
  };

  const registerIdentityAliases = (studentKey, identity) => {
    if (!identity.nameOnlyKey) return;
    setAliasMatch(nameOnlyAliases, identity.nameOnlyKey, studentKey);
    if (identity.nameSectionKey) {
      setAliasMatch(nameSectionAliases, identity.nameSectionKey, studentKey);
    }
  };

  sheetNames.forEach((sheetName) => {
    const sheet = sheets[sheetName];
    if (!sheet?.rows?.length) return;

    const headers = sheet.headers || [];
    sheet.rows.forEach((row) => {
      const identity = computeStudentIdentity(row, headers, sheetName);
      if (!identity.enrollmentKey) return;
      getOrCreateStudent(identity.enrollmentKey, identity);
      registerIdentityAliases(identity.enrollmentKey, identity);
    });
  });

  sheetNames.forEach((sheetName) => {
    const sheet = sheets[sheetName];
    if (!sheet?.rows?.length) return;
    const headers = sheet.headers || [];
    const examStage = detectExamStage(sheet.meta?.examName || sheetName, headers);
    sheet.rows.forEach((row) => {
      const identity = computeStudentIdentity(row, headers, sheetName);
      if (!identity.key) return;
      const metrics = extractStudentMetrics(row, headers, sheet.meta?.examName || sheetName);
      const studentKey = resolveStudentKey(identity);
      if (!studentKey) return;

      const existing = getOrCreateStudent(studentKey, identity);
      if (identity.enrollmentKey) {
        registerIdentityAliases(studentKey, identity);
      }

      const class9 = metrics.class9Percent;
      const target = metrics.targetPercent;
      if (class9 !== null && (toSafeNumber(existing['Class 9 %']) === null || existing['Class 9 %'] === '')) {
        existing['Class 9 %'] = class9;
      }
      if (target !== null && (toSafeNumber(existing['Target %']) === null || existing['Target %'] === '')) {
        existing['Target %'] = target;
      }

      if (examStage !== 'BASELINE' && examStage !== 'UNKNOWN') {
        const examPercent = metrics.examPercent;
        if (examPercent !== null && (toSafeNumber(getStageScore(existing, examStage)) === null || getStageScore(existing, examStage) === '')) {
          setStageScore(existing, examStage, examPercent);
        }
      }

      students.set(studentKey, existing);
    });
  });

  const headers = [
    'Enrollment No',
    'Student Name',
    'Section',
    'Class 9 %',
    'Target %',
    'HY %',
    'PB1 %',
    'PB2 %',
    'Board %',
    'Target Gap',
    'Improvement',
    'Status',
  ];

  const rows = Array.from(students.values()).map((entry) => {
    const board = toSafeNumber(entry['Board %']);
    const pb2 = toSafeNumber(entry['PB2 %']);
    const pb1 = toSafeNumber(entry['PB1 %']);
    const hy = toSafeNumber(entry['HY %']);
    const latest = board ?? pb2 ?? pb1 ?? hy ?? null;
    const target = toSafeNumber(entry['Target %']);
    const class9 = toSafeNumber(entry['Class 9 %']);
    const targetGap = latest !== null && target !== null ? parseFloat((latest - target).toFixed(2)) : '';
    const improvement = latest !== null && class9 !== null ? parseFloat((latest - class9).toFixed(2)) : '';
    let status = 'Needs Review';
    if (latest !== null && target !== null && latest >= target) status = 'Achieved Target';
    else if (latest !== null && target !== null && class9 !== null && latest > class9) status = 'Improving Toward Target';
    else if (latest !== null && class9 !== null && latest > class9) status = 'Improved';
    else if (latest !== null && target !== null) status = 'Below Target';

    return {
      ...entry,
      'Target Gap': targetGap,
      Improvement: improvement,
      Status: status,
    };
  }).sort((a, b) => {
    const sectionA = String(a.Section || '').toUpperCase();
    const sectionB = String(b.Section || '').toUpperCase();
    if (sectionA !== sectionB) return sectionA.localeCompare(sectionB);

    const enrollA = String(a['Enrollment No'] || '');
    const enrollB = String(b['Enrollment No'] || '');
    return enrollA.localeCompare(enrollB, undefined, { numeric: true, sensitivity: 'base' });
  });

  if (rows.length === 0) return null;
  return { headers, rows };
}

function isMissingProjectedValue(value) {
  return value === null || value === undefined || value === '' || value === 'null' || value === 'undefined';
}

function buildCumulativeRowLookupKey(row = {}) {
  const enrollmentNo = normalizeIdentifier(row['Enrollment No']);
  if (enrollmentNo) return `enrollment:${enrollmentNo}`;

  const normalizedName = normalizeStudentName(row['Student Name']);
  const normalizedSection = String(row.Section || '').trim().toLowerCase();
  if (normalizedName && normalizedSection) return `name-section:${normalizedName}|${normalizedSection}`;
  if (normalizedName) return `name:${normalizedName}`;
  return '';
}

function mergeImportedCumulativeSheet(importedSheet, fallbackSheet) {
  if (!importedSheet) return fallbackSheet;
  if (!fallbackSheet?.rows?.length) return importedSheet;

  const fallbackRowsByKey = new Map();
  fallbackSheet.rows.forEach((row) => {
    const key = buildCumulativeRowLookupKey(row);
    if (key) fallbackRowsByKey.set(key, row);
  });

  const rows = (importedSheet.rows || []).map((row) => {
    const key = buildCumulativeRowLookupKey(row);
    const fallbackRow = key ? fallbackRowsByKey.get(key) : null;
    if (!fallbackRow) return row;

    const mergedRow = { ...row };
    if (isMissingProjectedValue(mergedRow['Class 9 %']) && !isMissingProjectedValue(fallbackRow['Class 9 %'])) {
      mergedRow['Class 9 %'] = fallbackRow['Class 9 %'];
    }
    if (isMissingProjectedValue(mergedRow['Target %']) && !isMissingProjectedValue(fallbackRow['Target %'])) {
      mergedRow['Target %'] = fallbackRow['Target %'];
    }
    if (isMissingProjectedValue(mergedRow.Section) && !isMissingProjectedValue(fallbackRow.Section)) {
      mergedRow.Section = fallbackRow.Section;
    }
    return mergedRow;
  });

  return {
    ...importedSheet,
    headers: importedSheet.headers?.length ? importedSheet.headers : fallbackSheet.headers,
    rows,
  };
}

function recalculateCumulativeDerivedFields(row = {}) {
  const board = toSafeNumber(row['Board %']);
  const pb2 = toSafeNumber(row['PB2 %']);
  const pb1 = toSafeNumber(row['PB1 %']);
  const hy = toSafeNumber(row['HY %']);
  const latest = board !== null ? board : pb2 !== null ? pb2 : pb1 !== null ? pb1 : hy !== null ? hy : null;
  const target = toSafeNumber(row['Target %']);
  const class9 = toSafeNumber(row['Class 9 %']);
  const targetGap = latest !== null && target !== null ? parseFloat((latest - target).toFixed(2)) : '';
  const improvement = latest !== null && class9 !== null ? parseFloat((latest - class9).toFixed(2)) : '';
  let status = 'Needs Review';
  if (latest !== null && target !== null && latest >= target) status = 'Achieved Target';
  else if (latest !== null && target !== null && class9 !== null && latest > class9) status = 'Improving Toward Target';
  else if (latest !== null && class9 !== null && latest > class9) status = 'Improved';
  else if (latest !== null && target !== null) status = 'Below Target';

  return {
    ...row,
    'Target Gap': targetGap,
    Improvement: improvement,
    Status: status,
  };
}

function buildCumulativeRowOptionKey(row = {}, index = 0) {
  return buildCumulativeRowLookupKey(row) || `row:${index}`;
}

function buildBaselineReportRowKey(row = {}, index = 0) {
  const enrollment = normalizeIdentifier(row.baselineEnrollmentNo);
  const name = normalizeStudentName(row.baselineStudentName);
  const section = String(row.section || '').trim().toLowerCase();
  return `${enrollment || 'no-enrollment'}|${name || 'no-name'}|${section || 'no-section'}|${index}`;
}

function buildEmptySectionFileMap() {
  return Object.fromEntries(CLASS10_SECTIONS.map((section) => [section, null]));
}

function createEmptySectionWiseStructuredUpload() {
  return {
    baseline: null,
    hy: buildEmptySectionFileMap(),
    pb1: buildEmptySectionFileMap(),
    pb2: buildEmptySectionFileMap(),
    board: null,
  };
}

function createEmptyMergedStructuredUpload() {
  return {
    baseline: null,
    hy: null,
    pb1: null,
    pb2: null,
    board: null,
  };
}

function createStructuredUploadState() {
  return {
    sectionWise: createEmptySectionWiseStructuredUpload(),
    merged: createEmptyMergedStructuredUpload(),
  };
}


function normalizeStudentKey(row, headers) {
  const admissionCol = findAdmissionColumn(headers);
  const nameCol = findNameColumn(headers);
  const admission = admissionCol ? normalizeIdentifier(row[admissionCol]) : '';
  const name = nameCol ? String(row[nameCol] ?? '').trim() : '';
  const normalizedName = normalizeStudentName(name);
  if (admission) return `adm:${admission}`;
  if (normalizedName) return `name:${normalizedName}`;
  return null;
}

function getScoreColor(val) {
  const v = parseFloat(val);
  if (isNaN(v)) return '';
  if (v >= 95) return 'score-excellent';
  if (v >= 90) return 'score-great';
  if (v >= 80) return 'score-good';
  if (v >= 60) return 'score-average';
  if (v >= 50) return 'score-low';
  return 'score-poor';
}

function isNotOpted(val) {
  if (val === null || val === undefined || val === '') return true;
  const s = String(val).trim();
  return s === '-' || s === '—' || s === '–' || s === 'N/A' || s === 'NA' || s === '';
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
  const match = header.match(/(\d+)\s*$/);
  return match ? parseInt(match[1]) : null;
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
  const reportedTotal = totalCol ? toSafeNumber(row[totalCol]) : null;

  const entries = subjectCols.map((header) => {
    const score = toSafeNumber(row[header]);
    if (score === null || isNotOpted(row[header])) return null;
    return {
      header,
      subject: header.replace(/\s*\+\s*30/g, '').replace(/\s+\d+$/, '').trim(),
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
    ...entry,
    target: null,
  }));
  const obtainedMarks = subjectBreakdown.reduce((acc, entry) => acc + entry.score, 0);
  const maxMarks = subjectBreakdown.reduce((acc, entry) => acc + entry.maxScore, 0);

  const rawClass9 = toSafeNumber(class9Col ? row[class9Col] : null);
  const rawTarget = toSafeNumber(targetCol ? row[targetCol] : null);

  let class9Percent = rawClass9;
  if (class9Percent !== null && class9Percent > 0 && class9Percent < 2) {
    class9Percent = parseFloat((class9Percent * 100).toFixed(2));
  }

  let targetPercent = rawTarget;
  if (targetPercent !== null && targetPercent > 0 && targetPercent < 2) {
    targetPercent = parseFloat((targetPercent * 100).toFixed(2));
  }

  if (targetPercent !== null && targetPercent > 100) {
    targetPercent = 100;
  }

  const explicitExamPercent = toSafeNumber(examCol ? row[examCol] : null);
  const derivedExamPercent = maxMarks > 0 ? parseFloat(((obtainedMarks / maxMarks) * 100).toFixed(2)) : null;
  const examPercent = examStage === 'BASELINE' ? null : (explicitExamPercent ?? derivedExamPercent);

  return {
    studentKey: normalizeStudentKey(row, headers),
    name: nameCol ? row[nameCol] : '',
    admissionNo: admissionCol ? row[admissionCol] : '',
    class9Col,
    class9Percent,
    targetCol,
    targetPercent,
    totalCol,
    totalValue: totalCol ? row[totalCol] : obtainedMarks,
    examCol,
    examPercent,
    obtainedMarks,
    maxMarks,
    examName: fallbackExamName,
    subjectBreakdown,
  };
}

function buildWorkbookExamTimeline(sheetNames, sheets) {
  const STAGE_ORDER = { 'BASELINE': 0, 'HY': 1, 'PB1': 2, 'PB2': 3, 'BOARD': 4, 'UNKNOWN': 5 };
  
  return sheetNames
    .filter(name => sheets[name]?.rows?.length)
    .map((name, index) => {
      const headers = sheets[name].headers || [];
      const stage = detectExamStage(sheets[name].meta?.examName || name, headers);
      return {
        id: `sheet-${index}-${name}`,
        name,
        stage,
        date: '',
        maxMarks: 100,
        sheets: { [name]: sheets[name] },
      };
    })
    .sort((a, b) => (STAGE_ORDER[a.stage] ?? 99) - (STAGE_ORDER[b.stage] ?? 99));
}

function buildStudentComparisonData(studentRow, activeHeaders, activeSheetName, sheetNames, sheets) {
  const activeMetrics = extractStudentMetrics(studentRow, activeHeaders, activeSheetName);
  const studentKey = activeMetrics.studentKey;
  if (!studentKey) return null;

  const exams = [];
  sheetNames.forEach((sheetName) => {
    const sheet = sheets[sheetName];
    if (!sheet?.rows?.length) return;
    const match = sheet.rows.find(row => normalizeStudentKey(row, sheet.headers) === studentKey);
    if (!match) return;
    const metrics = extractStudentMetrics(match, sheet.headers, sheetName);
    exams.push({
      examName: sheetName,
      class9Percent: metrics.class9Percent,
      targetPercent: metrics.targetPercent,
      examPercent: metrics.examPercent,
      obtainedMarks: metrics.obtainedMarks,
      maxMarks: metrics.maxMarks,
      totalValue: metrics.totalValue,
      subjectBreakdown: metrics.subjectBreakdown,
    });
  });

  const class9Percent = activeMetrics.class9Percent ?? exams.find(exam => exam.class9Percent !== null)?.class9Percent ?? null;
  const targetPercent = activeMetrics.targetPercent ?? exams.find(exam => exam.targetPercent !== null)?.targetPercent ?? null;
  const latestExam = exams[exams.length - 1] || null;
  const previousExam = exams.length > 1 ? exams[exams.length - 2] : null;

  let status = 'Needs Review';
  if (latestExam?.examPercent !== null && targetPercent !== null && latestExam.examPercent >= targetPercent) {
    status = 'Achieved Target';
  } else if (latestExam?.examPercent !== null) {
    const baseline = previousExam?.examPercent ?? class9Percent;
    if (baseline !== null && latestExam.examPercent > baseline) {
      status = 'Improving Toward Target';
    } else {
      status = 'Below Target';
    }
  }

  return {
    ...activeMetrics,
    class9Percent,
    targetPercent,
    exams,
    latestExam,
    previousExam,
    status,
  };
}

function computeSheetAnalysis(sheet) {
  if (!sheet || !sheet.rows || sheet.rows.length === 0) return null;
  const targetCol = findTargetColumn(sheet.headers);
  if (!targetCol) return null;

  const rows = [];
  let totalStudents = 0;
  for (const range of RANGES) {
    let count = 0;
    for (const student of sheet.rows) {
      const v = parseFloat(student[targetCol]);
      if (!isNaN(v) && v >= range.min && v <= range.max) count++;
    }
    rows.push({ Range: range.label, Count: count, color: range.color });
    totalStudents += count;
  }

  rows.push({ Range: 'Total', Count: totalStudents, color: '#1e293b' });
  return { rows, totalStudents, targetCol };
}

function normalizeSheetForApp(sheet) {
  if (!sheet?.headers || !Array.isArray(sheet.rows)) return sheet;
  const validRows = [];
  for (const rawRow of sheet.rows) {
    const row = { ...rawRow };
    const values = Object.values(row).map(v => String(v).trim().toLowerCase());
    const isSummary = values.some(v =>
      v === '95-100' || v === '90-94' || v === '80-89' ||
      v === '60-79' || v === '50-59' || v === 'below 50'
    );
    const isEmpty = values.every(v => v === '' || v === 'null' || v === 'undefined');
    if (!isSummary && !isEmpty) {
      recalcGrandTotal(row, sheet.headers);
      validRows.push(row);
    }
  }
  return { headers: sheet.headers, rows: validRows, meta: sheet.meta, validation: sheet.validation };
}

function parseCSV(text) {
  const lines = text.split(/\r?\n/).filter(l => l.trim());
  if (lines.length === 0) return null;
  const parseLine = (line) => {
    const result = [];
    let current = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (inQuotes && line[i + 1] === '"') { current += '"'; i++; }
        else inQuotes = !inQuotes;
      } else if (ch === ',' && !inQuotes) {
        result.push(current.trim());
        current = '';
      } else {
        current += ch;
      }
    }
    result.push(current.trim());
    return result;
  };
  const matrix = lines.map(parseLine);
  const headerRowIndex = detectHeaderRowIndex(matrix);
  const prevRow = headerRowIndex > 0 ? matrix[headerRowIndex - 1] : [];
  const headers = buildHeadersFromRow(matrix[headerRowIndex] || [], prevRow);
  const rows = [];
  for (let i = headerRowIndex + 1; i < matrix.length; i++) {
    const values = matrix[i];
    if (values.every(v => v === '')) continue;
    const row = {};
    headers.forEach((h, idx) => {
      let val = values[idx] || '';
      const num = parseFloat(val);
      if (val !== '' && !isNaN(num) && String(num) === val) val = num;
      row[h] = val;
    });
    rows.push(row);
  }
  const meta = {
    headerRowIndex,
    titleRow: headerRowIndex > 0 ? String((matrix[0] || [])[0] ?? '').trim() : '',
  };
  const validation = validateParsedSheet('CSV Import', { headers, rows, meta });
  return { headers, rows, meta, validation };
}

export default function App() {
  const [screen, setScreen] = useState('upload');
  const [fileName, setFileName] = useState('');
  const [sheetNames, setSheetNames] = useState([]);
  const [sheets, setSheets] = useState({});
  const [importedCumulativeSheet, setImportedCumulativeSheet] = useState(null);
  const [activeSheet, setActiveSheet] = useState('');
  const [loading, setLoading] = useState(false);
  const [showAnalysis, setShowAnalysis] = useState(false);
  const [analysisSheets, setAnalysisSheets] = useState([]);
  const [studentReport, setStudentReport] = useState(null);
  const [editingCell, setEditingCell] = useState(null);
  const [editValue, setEditValue] = useState('');
  const [showCreateModal, setShowCreateModal] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [pdfExporting, setPdfExporting] = useState(false);
  const [toasts, setToasts] = useState([]);
  const [sortConfig, setSortConfig] = useState({ key: null, direction: null });
  const [editHistory, setEditHistory] = useState([]);
  const [historyIndex, setHistoryIndex] = useState(-1);
  const [examTimeline, setExamTimeline] = useState([]);
  const [showExamManager, setShowExamManager] = useState(false);
  const [showProgressTracker, setShowProgressTracker] = useState(false);
  const [dbEnabled, setDbEnabled] = useState(false);
  const [cumulativeData, setCumulativeData] = useState(null);
  const [cumulativeLoading, setCumulativeLoading] = useState(false);
  const [importIssues, setImportIssues] = useState([]);
  const [baselineMatchReport, setBaselineMatchReport] = useState(null);
  const [structuredWorkflowMode, setStructuredWorkflowMode] = useState('merged');
  const [structuredUpload, setStructuredUpload] = useState(createStructuredUploadState);
  const fileRef = useRef(null);
  const importMoreRef = useRef(null);
  const pdfReportRef = useRef(null);

  const removeToast = useCallback((id) => {
    setToasts(prev => prev.filter(t => t.id !== id));
  }, []);

  const addToast = useCallback((message, type = 'info', duration = 4000) => {
    const id = Date.now() + Math.random();
    setToasts(prev => [...prev, { id, message, type, duration }]);
    if (duration > 0) setTimeout(() => removeToast(id), duration);
  }, [removeToast]);

  useEffect(() => {
    try {
      const saved = localStorage.getItem(SESSION_STORAGE_KEY);
      if (saved) {
        const session = JSON.parse(saved);
        if (session.version === SESSION_VERSION && session.sheetNames && session.sheetNames.length > 0) {
          const normalizedSheets = Object.fromEntries(
            Object.entries(session.sheets || {}).map(([name, sheet]) => [name, normalizeSheetForApp(sheet)])
          );
          const rebuiltCumulativeSheet = buildClass10CumulativeSheet(session.sheetNames, normalizedSheets);
          setFileName(session.fileName || 'Restored Session');
          setSheetNames(session.sheetNames);
          setSheets(normalizedSheets);
          setActiveSheet(session.activeSheet || session.sheetNames[0]);
          setAnalysisSheets(session.analysisSheets || session.sheetNames.slice());
          if (session.examTimeline) setExamTimeline(session.examTimeline);
          if (session.importedCumulativeSheet) {
            setImportedCumulativeSheet(mergeImportedCumulativeSheet(session.importedCumulativeSheet, rebuiltCumulativeSheet));
          }
          if (session.baselineMatchReport) {
            setBaselineMatchReport(session.baselineMatchReport);
          }
          setScreen('dashboard');
          setTimeout(() => addToast('✅ Previous session restored!', 'success'), 500);
        } else if (session.version !== SESSION_VERSION) {
          localStorage.removeItem(SESSION_STORAGE_KEY);
        }
      }
    } catch (e) {
      console.warn('Failed to restore session:', e.message);
    }
  }, []);

  useEffect(() => {
    if (screen !== 'dashboard' && screen !== 'progress' || sheetNames.length === 0) return;
    const timer = setTimeout(() => {
      try {
        const session = {
          version: SESSION_VERSION,
          fileName,
          sheetNames,
          sheets,
          activeSheet,
          analysisSheets,
          examTimeline,
          importedCumulativeSheet,
          baselineMatchReport,
        };
        localStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(session));
      } catch (e) {
        console.warn('Session save failed:', e.message);
      }
    }, 1000);
    return () => clearTimeout(timer);
  }, [sheets, sheetNames, activeSheet, fileName, analysisSheets, screen, examTimeline, importedCumulativeSheet, baselineMatchReport]);

  const fetchCumulativeData = useCallback(async () => {
    setCumulativeLoading(true);
    try {
      const res = await fetch(`${API}/cumulative-report`);
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Failed to load cumulative report');
      setCumulativeData(data);
    } catch (err) {
      console.warn('Cumulative report unavailable:', err.message);
    } finally {
      setCumulativeLoading(false);
    }
  }, []);

  useEffect(() => {
    if (PHASE_ONE_CUMULATIVE_ONLY) return undefined;
    const initDb = async () => {
      try {
        const res = await fetch(`${API}/db-status`);
        const data = await res.json();
        setDbEnabled(Boolean(data.enabled));
        if (data.enabled) {
          fetchCumulativeData();
        }
      } catch (err) {
        console.warn('Database not reachable:', err.message);
      }
    };
    initDb();
  }, [API, fetchCumulativeData]);

  const buildUniqueSheetName = useCallback((preferredName, usedNames) => {
    if (!usedNames.has(preferredName)) {
      usedNames.add(preferredName);
      return preferredName;
    }
    let count = 2;
    let nextName = `${preferredName} (${count})`;
    while (usedNames.has(nextName)) {
      count += 1;
      nextName = `${preferredName} (${count})`;
    }
    usedNames.add(nextName);
    return nextName;
  }, []);

  const normalizeImportedSheet = useCallback((sheetData) => normalizeSheetForApp(sheetData), []);

  const setStructuredFile = useCallback((mode, group, section, file) => {
    setStructuredUpload(prev => {
      const currentModeState = prev[mode];
      if (!currentModeState) return prev;
      if (mode === 'merged' || group === 'baseline' || group === 'board') {
        return {
          ...prev,
          [mode]: {
            ...currentModeState,
            [group]: file || null,
          },
        };
      }
      return {
        ...prev,
        [mode]: {
          ...currentModeState,
          [group]: {
            ...currentModeState[group],
            [section]: file || null,
          },
        },
      };
    });
  }, []);

  const buildStructuredFiles = useCallback((mode = structuredWorkflowMode) => {
    const upload = structuredUpload[mode];
    const prepared = [];
    if (!upload) return prepared;

    if (upload.baseline) {
      const file = upload.baseline;
      prepared.push(new File([file], `BASELINE_CLASS10_${file.name}`, { type: file.type }));
    }

    if (mode === 'merged') {
      [
        ['hy', 'HY'],
        ['pb1', 'PB1'],
        ['pb2', 'PB2'],
      ].forEach(([key, label]) => {
        const file = upload[key];
        if (!file) return;
        prepared.push(new File([file], `${label}_CLASS10_MERGED_${file.name}`, { type: file.type }));
      });
    } else {
      [
        ['hy', 'HY'],
        ['pb1', 'PB1'],
        ['pb2', 'PB2'],
      ].forEach(([key, label]) => {
        CLASS10_SECTIONS.forEach((section) => {
          const file = upload[key][section];
          if (!file) return;
          prepared.push(new File([file], `${label}_10${section}_${file.name}`, { type: file.type }));
        });
      });
    }

    if (upload.board) {
      const file = upload.board;
      prepared.push(new File([file], `BOARD_CLASS10_${file.name}`, { type: file.type }));
    }
    return prepared;
  }, [structuredUpload, structuredWorkflowMode]);

  const persistFilesToDatabase = useCallback(async (files) => {
    if (!dbEnabled || !files?.length) return;
    try {
      const formData = new FormData();
      files.forEach(file => formData.append('files', file));
      const res = await fetch(`${API}/import-persistent`, { method: 'POST', body: formData });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Persistent import failed');
      setCumulativeData(data.cumulativeReport || null);
      addToast('Files imported into cumulative database successfully.', 'success');
    } catch (err) {
      console.error('Persistent import error:', err);
      addToast(`Persistent import skipped: ${err.message}`, 'warning');
    }
  }, [API, addToast, dbEnabled]);

  const handleFiles = useCallback(async (fileList) => {
    const files = Array.from(fileList || []);
    if (files.length === 0) return;

    const invalidFiles = files.filter(file => {
      const ext = file.name.split('.').pop().toLowerCase();
      return ext !== 'xlsx' && ext !== 'xls' && ext !== 'csv';
    });

    if (invalidFiles.length > 0) {
      addToast('Please upload only .xls, .xlsx, or .csv files', 'error');
      return;
    }

    setLoading(true);
    setImportIssues([]);
    setImportedCumulativeSheet(null);
    setBaselineMatchReport(null);

    try {
      const usedNames = new Set();
      const mergedSheetNames = [];
      const mergedSheets = {};
      const validationIssues = [];

      for (const file of files) {
        const ext = file.name.split('.').pop().toLowerCase();
        const baseName = file.name.replace(/\.[^/.]+$/, '');

        if (ext === 'csv') {
          const text = await file.text();
          const parsed = parseCSV(text);
          if (!parsed || parsed.rows.length === 0) continue;
          const parsedValidation = validateParsedSheet(baseName, parsed);
          if (!parsedValidation.ok) {
            parsedValidation.issues.forEach((issue) => validationIssues.push(`${file.name}: ${issue}`));
            continue;
          }

          parsed.rows.forEach(r => recalcGrandTotal(r, parsed.headers));
          const uniqueSheetName = buildUniqueSheetName(baseName, usedNames);
          mergedSheetNames.push(uniqueSheetName);
          mergedSheets[uniqueSheetName] = parsed;
          continue;
        }

        const formData = new FormData();
        formData.append('file', file);
        const res = await fetch(`${API}/v1/reports/parse`, { method: 'POST', body: formData });
        const data = await res.json();
        if (!res.ok) {
          throw new Error(data.error || `Failed to parse ${file.name}`);
        }

        if (Array.isArray(data.issues) && data.issues.length > 0) {
          data.issues.forEach((issue) => validationIssues.push(`${file.name} / ${issue.sheetName}: ${issue.message}`));
        }

        for (const sheetName of data.sheetNames || []) {
          const importedSheet = data.sheets?.[sheetName];
          if (!importedSheet) continue;

          const preferredName = files.length === 1 ? sheetName : `${baseName} - ${sheetName}`;
          const uniqueSheetName = buildUniqueSheetName(preferredName, usedNames);
          mergedSheetNames.push(uniqueSheetName);
          mergedSheets[uniqueSheetName] = normalizeImportedSheet(importedSheet);
        }
      }

      if (validationIssues.length > 0) {
        setImportIssues(validationIssues);
        addToast('Import blocked. Please fix the sheet structure issues shown above.', 'error', 6000);
        return;
      }

      if (mergedSheetNames.length === 0) {
        addToast('No valid data found in the selected files', 'error');
        return;
      }

      setFileName(files.length === 1 ? files[0].name : `${files.length} files loaded`);
      setSheetNames(mergedSheetNames);
      setSheets(mergedSheets);
      setActiveSheet(CUMULATIVE_SHEET_NAME);
      setAnalysisSheets(mergedSheetNames.slice());
      setScreen('dashboard');
      addToast(`${files.length} file${files.length > 1 ? 's' : ''} loaded with ${mergedSheetNames.length} sheet${mergedSheetNames.length > 1 ? 's' : ''}`, 'success');
    } catch (err) {
      console.error('Upload error:', err);
      addToast(err.message || 'Cannot connect to backend. Check the server is running.', 'error');
    } finally {
      setLoading(false);
    }
  }, [API, addToast, buildUniqueSheetName, normalizeImportedSheet]);

  const handleStructuredSubmit = useCallback(async () => {
    const currentUpload = structuredUpload[structuredWorkflowMode];
    const files = buildStructuredFiles(structuredWorkflowMode);
    const hasAllRequiredFiles = Boolean(
      currentUpload?.baseline
      && currentUpload?.hy
      && currentUpload?.pb1
      && currentUpload?.pb2
      && currentUpload?.board
    );

    if (!hasAllRequiredFiles || files.length !== 5) {
      addToast('Please add exactly 5 files before importing: Baseline, HY, PB1, PB2, and Board.', 'error');
      return;
    }

    setLoading(true);
    setImportIssues([]);
    setBaselineMatchReport(null);

    try {
      const formData = new FormData();
      files.forEach((file) => formData.append('files', file));
      const res = await fetch(`${API}/v1/reports/structured-import`, { method: 'POST', body: formData });
      const data = await res.json();

      if (!res.ok) {
        const issues = Array.isArray(data.issues)
          ? data.issues.map((issue) => `${issue.fileName} / ${issue.sheetName}: ${issue.message}`)
          : [];
        if (issues.length > 0) setImportIssues(issues);
        throw new Error(data.error || 'Structured import failed');
      }

      const validationIssues = Array.isArray(data.issues)
        ? data.issues.map((issue) => `${issue.fileName} / ${issue.sheetName}: ${issue.message}`)
        : [];
      if (validationIssues.length > 0) {
        setImportIssues(validationIssues);
        addToast('Structured import blocked. Please fix the sheet structure issues shown above.', 'error', 6000);
        return;
      }

      const normalizedSheets = Object.fromEntries(
        Object.entries(data.sheets || {}).map(([name, sheet]) => [name, normalizeImportedSheet(sheet)])
      );
      const rebuiltCumulativeSheet = buildClass10CumulativeSheet(data.sheetNames || [], normalizedSheets);
      const mergedCumulativeSheet = mergeImportedCumulativeSheet(data.masterCumulativeSheet || null, rebuiltCumulativeSheet);
      const nextBaselineMatchReport = data.baselineMatchReport || null;

      setFileName(`${files.length} structured file${files.length > 1 ? 's' : ''} loaded`);
      setSheetNames(data.sheetNames || []);
      setSheets(normalizedSheets);
      setImportedCumulativeSheet(mergedCumulativeSheet);
      setBaselineMatchReport(nextBaselineMatchReport);
      setActiveSheet(CUMULATIVE_SHEET_NAME);
      setAnalysisSheets((data.sheetNames || []).slice());
      setScreen('dashboard');
      setStructuredUpload(createStructuredUploadState());
      if ((nextBaselineMatchReport?.unmatchedCount || 0) > 0) {
        addToast(`Structured import completed with ${nextBaselineMatchReport.unmatchedCount} unmatched baseline row${nextBaselineMatchReport.unmatchedCount > 1 ? 's' : ''}.`, 'warning', 6000);
      } else {
        addToast('Structured import completed without database dependency.', 'success');
      }
    } catch (err) {
      console.error('Structured import error:', err);
      addToast(err.message || 'Structured import failed', 'error');
    } finally {
      setLoading(false);
    }
  }, [API, addToast, buildStructuredFiles, structuredUpload, structuredWorkflowMode]);

  const startEdit = (sheetName, rowIdx, col) => {
    setEditingCell({ sheetName, rowIdx, col });
    setEditValue(sheets[sheetName].rows[rowIdx][col] ?? '');
  };

  const saveEdit = () => {
    if (!editingCell) return;
    const { sheetName, rowIdx, col } = editingCell;
    const oldValue = sheets[sheetName].rows[rowIdx][col];
    const updated = { ...sheets };
    const val = editValue;
    const num = parseFloat(val);
    const finalVal = isNaN(num) || val === '' ? val : num;

    if (oldValue !== finalVal) {
      const newHistory = editHistory.slice(0, historyIndex + 1);
      newHistory.push({ sheetName, rowIdx, col, oldValue, newValue: finalVal });
      if (newHistory.length > 20) newHistory.shift();
      setEditHistory(newHistory);
      setHistoryIndex(newHistory.length - 1);
    }

    updated[sheetName].rows[rowIdx][col] = finalVal;

    const row = updated[sheetName].rows[rowIdx];
    const headers = updated[sheetName].headers;
    recalcGrandTotal(row, headers);

    setSheets(updated);
    setImportedCumulativeSheet(null);
    setEditingCell(null);
  };

  const cancelEdit = () => setEditingCell(null);

  const handleCreateTemplate = (config) => {
    const { className, numSections, subjects } = config;
    const newSheetNames = [];
    const newSheets = {};
    const sectionLetters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

    const headers = ['S.No', 'Name', ...subjects, 'Grand Total', '% in IX+30'];

    for (let i = 0; i < numSections; i++) {
      const sec = sectionLetters[i] || `${i + 1}`;
      const sheetName = `${className} ${sec}`;
      newSheetNames.push(sheetName);
      newSheets[sheetName] = { headers, rows: [] };
    }

    setFileName(`NEW_${className}_DATA.xlsx`);
    setSheetNames(newSheetNames);
    setSheets(newSheets);
    setImportedCumulativeSheet(null);
    setActiveSheet(CUMULATIVE_SHEET_NAME);
    setAnalysisSheets(newSheetNames.slice());
    setShowCreateModal(false);
    setScreen('dashboard');
  };

  const handleSaveExam = (examDetails) => {
    const newExam = {
      id: Date.now().toString(),
      name: examDetails.name,
      date: examDetails.date,
      maxMarks: examDetails.maxMarks,
      sheets: JSON.parse(JSON.stringify(sheets))
    };
    
    setExamTimeline(prev => [...prev, newExam].sort((a, b) => new Date(a.date) - new Date(b.date)));
    setShowExamManager(false);
    addToast('Exam saved to timeline!', 'success');
  };

  const addStudent = () => {
    if (!activeSheet || !sheets[activeSheet]) return;
    const updated = { ...sheets };
    const headers = updated[activeSheet].headers;
    const newRow = {};
    headers.forEach(h => { newRow[h] = ''; });
    const snCol = headers.find(h => h.toLowerCase().includes('s.no') || h.toLowerCase() === 'sr. no.');
    if (snCol) newRow[snCol] = updated[activeSheet].rows.length + 1;
    updated[activeSheet].rows.push(newRow);
    setSheets(updated);
  };

  const deleteStudent = (rowIdx) => {
    if (!confirm('Delete this student?')) return;
    const updated = { ...sheets };
    updated[activeSheet].rows.splice(rowIdx, 1);
    setSheets(updated);
  };

  const addSheet = () => {
    const name = prompt('Enter new sheet name:');
    if (!name || sheetNames.includes(name)) return;
    const headers = ['S.No', 'Name', 'Score'];
    setSheetNames([...sheetNames, name]);
    setSheets({
      ...sheets,
      [name]: { headers, rows: [] }
    });
    setActiveSheet(name);
  };

  const analysisData = useMemo(() => {
    const selected = analysisSheets.filter(s => sheets[s]);
    const sums = {};
    selected.forEach(s => sums[s] = 0);
    const rows = [];

    for (const range of RANGES) {
      const row = { Range: range.label };
      let total = 0;
      for (const sn of selected) {
        const sheet = sheets[sn];
        const targetCol = findTargetColumn(sheet.headers);
        if (!targetCol) { row[sn] = 0; continue; }
        let count = 0;
        for (const student of sheet.rows) {
          const v = parseFloat(student[targetCol]);
          if (!isNaN(v) && v >= range.min && v <= range.max) count++;
        }
        row[sn] = count;
        total += count;
        sums[sn] += count;
      }
      row.students = total;
      rows.push(row);
    }

    const grandTotal = Object.values(sums).reduce((a, b) => a + b, 0);
    rows.forEach(r => {
      r['per%'] = grandTotal > 0 ? parseFloat(((r.students / grandTotal) * 100).toFixed(2)) : 0;
    });

    const totalRow = { Range: 'Total' };
    selected.forEach(s => totalRow[s] = sums[s]);
    totalRow.students = grandTotal;
    totalRow['per%'] = 100;
    rows.push(totalRow);

    return { sections: selected, rows, headers: ['Range', ...selected, 'students', 'per%'] };
  }, [analysisSheets, sheets]);

  const subjectComparison = useMemo(() => {
    const selected = analysisSheets.filter(s => sheets[s]);
    if (selected.length === 0) return null;

    const firstSheet = sheets[selected[0]];
    if (!firstSheet) return null;
    const subjectCols = getSubjectColumns(firstSheet.headers);
    if (subjectCols.length === 0) return null;

    const data = subjectCols.map(subject => {
      const entry = { subject: subject.replace(/\s*\+\s*30/g, '').replace(/\s+\d+$/, '').trim() };
      for (const sn of selected) {
        const sheet = sheets[sn];
        let sum = 0, count = 0;
        for (const student of sheet.rows) {
          const val = student[subject];
          if (isNotOpted(val)) continue;
          const v = parseFloat(val);
          if (!isNaN(v)) { sum += v; count++; }
        }
        entry[sn] = count > 0 ? parseFloat((sum / count).toFixed(1)) : 0;
      }
      return entry;
    });

    return { data, sections: selected, subjects: subjectCols };
  }, [analysisSheets, sheets]);

  const openStudentReport = (row) => {
    setStudentReport(row);
  };

  const exportExcel = async () => {
    setLoading(true);
    try {
      const exportSheetNames = cumulativeSheet ? [...sheetNames, CUMULATIVE_SHEET_NAME] : sheetNames;
      const exportSheets = cumulativeSheet ? { ...sheets, [CUMULATIVE_SHEET_NAME]: cumulativeSheet } : sheets;
      const res = await fetch(`${API}/export`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sheetNames: exportSheetNames, sheets: exportSheets })
      });

      if (!res.ok) {
        const errData = await res.json().catch(() => ({}));
        throw new Error(errData.error || 'Export failed');
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Report_${fileName}`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Export error:', err);
      alert('Export failed: ' + err.message);
    }
    setLoading(false);
  };

  const exportCSV = () => {
    if (!effectiveCurrentSheet) return;
    const headers = effectiveCurrentSheet.headers;
    const escape = (val) => {
      const s = String(val ?? '');
      return s.includes(',') || s.includes('"') || s.includes('\n') ? `"${s.replace(/"/g, '""')}"` : s;
    };
    const lines = [headers.map(escape).join(',')];
    effectiveCurrentSheet.rows.forEach(row => {
      lines.push(headers.map(h => escape(row[h])).join(','));
    });
    const blob = new Blob([lines.join('\n')], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${activeSheet || 'sheet'}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    addToast('CSV exported successfully!', 'success');
  };

  const exportPDF = async () => {
    setPdfExporting(true);
    await new Promise(r => setTimeout(r, 1500));

    try {
      const el = pdfReportRef.current;
      if (!el) throw new Error('PDF content not ready');

      const blocks = el.querySelectorAll('.pdf-page-block');
      if (!blocks || blocks.length === 0) throw new Error('No PDF blocks found');

      const pdf = new jsPDF('p', 'mm', 'a4');
      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();
      const margin = 10;
      const usableWidth = pageWidth - (margin * 2);

      let currentY = margin;
      let isFirstPage = true;

      for (let i = 0; i < blocks.length; i++) {
        const block = blocks[i];
        const canvas = await html2canvas(block, {
          backgroundColor: '#ffffff',
          scale: 2,
          useCORS: true,
          logging: false,
          windowWidth: 1050,
        });

        const imgWidth = canvas.width;
        const imgHeight = canvas.height;
        const ratio = usableWidth / imgWidth;
        const scaledHeight = imgHeight * ratio;

        if (!isFirstPage && currentY + scaledHeight > pageHeight - margin) {
          pdf.addPage();
          currentY = margin;
        }

        const imgData = canvas.toDataURL('image/png');
        if (scaledHeight > pageHeight - (margin * 2)) {
          let remainingHeight = imgHeight;
          let sourceY = 0;
          const pageContentHeight = (pageHeight - (margin * 2)) / ratio;

          while (remainingHeight > 0) {
            const sliceHeight = Math.min(pageContentHeight, remainingHeight);
            const tempCanvas = document.createElement('canvas');
            tempCanvas.width = imgWidth;
            tempCanvas.height = sliceHeight;
            const ctx = tempCanvas.getContext('2d');
            ctx.drawImage(canvas, 0, sourceY, imgWidth, sliceHeight, 0, 0, imgWidth, sliceHeight);

            const sliceData = tempCanvas.toDataURL('image/png');
            const scaledSliceHeight = sliceHeight * ratio;

            if (sourceY > 0) pdf.addPage();
            const drawY = sourceY === 0 ? currentY : margin;
            pdf.addImage(sliceData, 'PNG', margin, drawY, usableWidth, scaledSliceHeight);

            sourceY += sliceHeight;
            remainingHeight -= sliceHeight;
            if (sourceY === sliceHeight) {
              currentY = drawY + scaledSliceHeight;
            }
          }
          currentY = margin;
        } else {
          pdf.addImage(imgData, 'PNG', margin, currentY, usableWidth, scaledHeight);
          currentY += scaledHeight + 10;
        }
        isFirstPage = false;
      }

      pdf.save(`Analysis_Report_${fileName.replace('.xlsx', '')}.pdf`);
    } catch (err) {
      console.error('PDF export failed:', err);
      addToast('PDF export failed: ' + err.message, 'error');
    }
    setPdfExporting(false);
  };

  const resetAll = () => {
    if (!confirm('Reset everything? All changes will be lost.')) return;
    setScreen('upload');
    setSheets({});
    setSheetNames([]);
    setActiveSheet('');
    setShowAnalysis(false);
    setImportIssues([]);
    setStudentReport(null);
    setEditHistory([]);
    setHistoryIndex(-1);
    setSortConfig({ key: null, direction: null });
    setImportedCumulativeSheet(null);
    setBaselineMatchReport(null);
    setStructuredUpload(createStructuredUploadState());
    setStructuredWorkflowMode('merged');
    try { localStorage.removeItem('report-generator-session'); } catch (e) { /* ignore */ }
    addToast('All data cleared', 'info', 2000);
  };

  const currentSheet = sheets[activeSheet];
  const rebuiltCumulativeSheet = useMemo(
    () => buildClass10CumulativeSheet(sheetNames, sheets),
    [sheetNames, sheets]
  );
  const cumulativeSheet = useMemo(
    () => mergeImportedCumulativeSheet(importedCumulativeSheet, rebuiltCumulativeSheet),
    [importedCumulativeSheet, rebuiltCumulativeSheet]
  );
  const cumulativeStudentOptions = useMemo(
    () => (cumulativeSheet?.rows || []).map((row, index) => ({
      key: buildCumulativeRowOptionKey(row, index),
      enrollmentNo: row['Enrollment No'] || '',
      studentName: row['Student Name'] || '',
      section: row.Section || '',
      hasBaselineValues: !isMissingProjectedValue(row['Class 9 %']) || !isMissingProjectedValue(row['Target %']),
    })),
    [cumulativeSheet]
  );
  const applyBaselineMapping = useCallback((reportRow, targetKey, mode = 'manual') => {
    if (!cumulativeSheet?.rows?.length || !targetKey) {
      addToast('Select a cumulative student before applying the baseline values.', 'error');
      return;
    }

    const baselineClass9 = reportRow.baselineClass9Percent;
    const baselineTarget = reportRow.baselineTargetPercent;
    if (isMissingProjectedValue(baselineClass9) && isMissingProjectedValue(baselineTarget)) {
      addToast('This baseline row has no Class IX or Target value to apply.', 'error');
      return;
    }

    const targetIndex = cumulativeSheet.rows.findIndex((row, index) => buildCumulativeRowOptionKey(row, index) === targetKey);
    if (targetIndex < 0) {
      addToast('Selected cumulative student was not found.', 'error');
      return;
    }

    const targetRow = cumulativeSheet.rows[targetIndex];
    if (!isMissingProjectedValue(targetRow['Class 9 %']) || !isMissingProjectedValue(targetRow['Target %'])) {
      addToast('Selected student already has Class IX / Target values. Exact matches are not overwritten.', 'warning', 6000);
      return;
    }

    const nextRows = cumulativeSheet.rows.map((row, index) => {
      if (index !== targetIndex) return row;
      return recalculateCumulativeDerivedFields({
        ...row,
        'Class 9 %': isMissingProjectedValue(baselineClass9) ? row['Class 9 %'] : baselineClass9,
        'Target %': isMissingProjectedValue(baselineTarget) ? row['Target %'] : baselineTarget,
      });
    });
    const nextSheet = { ...cumulativeSheet, rows: nextRows };
    const nextTargetRow = nextRows[targetIndex];

    setImportedCumulativeSheet(nextSheet);
    setBaselineMatchReport((currentReport) => {
      if (!currentReport?.rows?.length) return currentReport;
      let changed = false;
      const rows = currentReport.rows.map((row, index) => {
        if (buildBaselineReportRowKey(row, index) !== buildBaselineReportRowKey(reportRow, reportRow._reportIndex ?? index)) return row;
        changed = true;
        return {
          ...row,
          confidence: 'exact',
          reason: mode === 'suggested' ? 'Approved suggested baseline mapping' : 'Approved manual baseline mapping',
          suggestedStudentName: nextTargetRow['Student Name'] || '',
          suggestedEnrollmentNo: nextTargetRow['Enrollment No'] || '',
          suggestedSection: nextTargetRow.Section || '',
          suggestedStudentKey: buildCumulativeRowOptionKey(nextTargetRow, targetIndex),
          suggestionScore: row.suggestionScore || '',
        };
      });
      if (!changed) return currentReport;
      return {
        ...currentReport,
        matchedCount: (currentReport.matchedCount || 0) + 1,
        unmatchedCount: Math.max(0, (currentReport.unmatchedCount || 0) - 1),
        rows,
      };
    });
    setActiveSheet(CUMULATIVE_SHEET_NAME);
    addToast('Baseline Class IX / Target values applied to the cumulative sheet.', 'success');
  }, [addToast, cumulativeSheet]);
  const displaySheetNames = useMemo(
    () => {
      if (PHASE_ONE_CUMULATIVE_ONLY) {
        return cumulativeSheet ? [CUMULATIVE_SHEET_NAME] : [];
      }
      return cumulativeSheet ? [...sheetNames, CUMULATIVE_SHEET_NAME] : sheetNames;
    },
    [sheetNames, cumulativeSheet]
  );
  const effectiveCurrentSheet = PHASE_ONE_CUMULATIVE_ONLY
    ? cumulativeSheet
    : (activeSheet === CUMULATIVE_SHEET_NAME ? cumulativeSheet : currentSheet);
  const isCumulativeView = PHASE_ONE_CUMULATIVE_ONLY ? true : activeSheet === CUMULATIVE_SHEET_NAME;
  const totalStudents = effectiveCurrentSheet ? effectiveCurrentSheet.rows.length : 0;

  const filteredRows = useMemo(() => {
    if (!effectiveCurrentSheet) return [];
    if (!searchQuery.trim()) return effectiveCurrentSheet.rows;
    const q = searchQuery.toLowerCase();
    const nameCol = effectiveCurrentSheet.headers.find(h => h.toLowerCase().includes('name') && !h.toLowerCase().includes('father') && !h.toLowerCase().includes('mother'));
    return effectiveCurrentSheet.rows.filter(row => {
      if (nameCol && String(row[nameCol] || '').toLowerCase().includes(q)) return true;
      return Object.values(row).some(v => String(v).toLowerCase().includes(q));
    });
  }, [effectiveCurrentSheet, searchQuery]);

  const handleSort = (key) => {
    setSortConfig(prev => {
      if (prev.key === key) {
        if (prev.direction === 'asc') return { key, direction: 'desc' };
        if (prev.direction === 'desc') return { key: null, direction: null };
      }
      return { key, direction: 'asc' };
    });
  };

  const sortedRows = useMemo(() => {
    let rows = filteredRows;
    if (sortConfig.key) {
      rows = [...rows].sort((a, b) => {
        const aVal = a[sortConfig.key];
        const bVal = b[sortConfig.key];
        const aNum = parseFloat(aVal);
        const bNum = parseFloat(bVal);
        let comparison = 0;
        if (!isNaN(aNum) && !isNaN(bNum)) {
          comparison = aNum - bNum;
        } else {
          comparison = String(aVal || '').localeCompare(String(bVal || ''));
        }
        return sortConfig.direction === 'desc' ? -comparison : comparison;
      });
    }
    return rows;
  }, [filteredRows, sortConfig]);

  const undo = useCallback(() => {
    if (historyIndex < 0 || editHistory.length === 0) return;
    const entry = editHistory[historyIndex];
    const updated = { ...sheets };
    if (updated[entry.sheetName]?.rows[entry.rowIdx]) {
      updated[entry.sheetName].rows[entry.rowIdx][entry.col] = entry.oldValue;
      recalcGrandTotal(updated[entry.sheetName].rows[entry.rowIdx], updated[entry.sheetName].headers);
      setSheets(updated);
      setImportedCumulativeSheet(null);
      setHistoryIndex(prev => prev - 1);
      addToast('Undo: Reverted cell edit', 'info', 2000);
    }
  }, [historyIndex, editHistory, sheets, addToast]);

  const redo = useCallback(() => {
    if (historyIndex >= editHistory.length - 1) return;
    const entry = editHistory[historyIndex + 1];
    const updated = { ...sheets };
    if (updated[entry.sheetName]?.rows[entry.rowIdx]) {
      updated[entry.sheetName].rows[entry.rowIdx][entry.col] = entry.newValue;
      recalcGrandTotal(updated[entry.sheetName].rows[entry.rowIdx], updated[entry.sheetName].headers);
      setSheets(updated);
      setImportedCumulativeSheet(null);
      setHistoryIndex(prev => prev + 1);
      addToast('Redo: Restored cell edit', 'info', 2000);
    }
  }, [historyIndex, editHistory, sheets, addToast]);

  useEffect(() => {
    const handleKeyDown = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) {
        e.preventDefault();
        undo();
      }
      if ((e.ctrlKey || e.metaKey) && (e.key === 'y' || (e.key === 'z' && e.shiftKey))) {
        e.preventDefault();
        redo();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [undo, redo]);

  const classStats = useMemo(() => {
    if (!effectiveCurrentSheet || effectiveCurrentSheet.rows.length === 0) return null;
    const totalCol = effectiveCurrentSheet.headers.find(h => h.toLowerCase().includes('grand total') || h.toLowerCase() === 'total');
    const cumulativeScoreCol = isCumulativeView
      ? effectiveCurrentSheet.headers.find(h => ['Board %', 'PB2 %', 'PB1 %', 'HY %'].includes(h))
      : null;
    const percentCol = cumulativeScoreCol || findTargetColumn(effectiveCurrentSheet.headers);
    const nameCol = effectiveCurrentSheet.headers.find(h => h.toLowerCase().includes('name') && !h.toLowerCase().includes('father') && !h.toLowerCase().includes('mother'));

    const avgCol = percentCol || totalCol;
    if (!avgCol) return null;

    const rankCol = totalCol || avgCol;

    const scored = effectiveCurrentSheet.rows.filter(r => !isNaN(parseFloat(r[rankCol])) && parseFloat(r[rankCol]) > 0);
    if (scored.length === 0) return null;

    const avgScores = scored.filter(r => !isNaN(parseFloat(r[avgCol]))).map(r => parseFloat(r[avgCol]));
    const avg = avgScores.length > 0 ? avgScores.reduce((a, b) => a + b, 0) / avgScores.length : 0;

    const rankScores = scored.map(r => parseFloat(r[rankCol]));
    const maxScore = Math.max(...rankScores);
    const minScore = Math.min(...rankScores);
    const topStudent = scored.find(r => parseFloat(r[rankCol]) === maxScore);
    const bottomStudent = scored.find(r => parseFloat(r[rankCol]) === minScore);

    const ranked = [...scored].sort((a, b) => parseFloat(b[rankCol]) - parseFloat(a[rankCol]));
    const rankMap = new Map();
    ranked.forEach((r, i) => rankMap.set(r, i + 1));

    return {
      avg: avg.toFixed(1),
      topName: nameCol ? (topStudent?.[nameCol] || '—') : '—',
      topScore: maxScore,
      bottomName: nameCol ? (bottomStudent?.[nameCol] || '—') : '—',
      bottomScore: minScore,
      totalScored: scored.length,
      rankMap,
    };
  }, [effectiveCurrentSheet, isCumulativeView]);

  return (
    <div className="app-wrapper">
      <div className="app-bg" />

      {screen === 'upload' && (
        <div className="main-card fade-in upload-card">
          <div className="card-icon"><FileSpreadsheet size={48} strokeWidth={1.5} /></div>
          <h1>Target Analysis Report Generator</h1>
          <p className="subtitle">Import exactly 5 files only: Class 9 + Target, HY, PB1, PB2, and Board. Each exam file can be a workbook with section tabs or one combined sheet for all sections.</p>
          {importIssues.length > 0 && (
            <div style={{ marginBottom: '1rem', border: '1px solid rgba(239,68,68,0.18)', background: 'rgba(254,242,242,0.95)', borderRadius: '18px', padding: '1rem 1.1rem', textAlign: 'left' }}>
              <div style={{ fontWeight: 700, color: '#991b1b', marginBottom: '0.5rem' }}>Import blocked due to sheet validation errors</div>
              <div style={{ display: 'grid', gap: '0.35rem', color: '#7f1d1d', fontSize: '0.92rem' }}>
                {importIssues.slice(0, 8).map((issue, index) => <div key={`${issue}-${index}`}>• {issue}</div>)}
                {importIssues.length > 8 && <div>• {importIssues.length - 8} more issues not shown</div>}
              </div>
            </div>
          )}
          <div className="upload-shell" style={{ gridTemplateColumns: '1fr', maxWidth: '980px', margin: '0 auto' }}>
            <StructuredUploadPanel
              loading={loading}
              workflowMode={structuredWorkflowMode}
              setWorkflowMode={setStructuredWorkflowMode}
              structuredUpload={structuredUpload}
              setStructuredFile={setStructuredFile}
              onSubmit={handleStructuredSubmit}
            />
          </div>

        </div>
      )}

      {showCreateModal && (
        <CreateTemplateModal
          onClose={() => setShowCreateModal(false)}
          onCreate={handleCreateTemplate}
        />
      )}

      {showExamManager && (
        <ExamManagerModal
          onClose={() => setShowExamManager(false)}
          onSave={handleSaveExam}
        />
      )}

      {screen === 'progress' && (
        <ProgressTracker
          examTimeline={examTimeline}
          onBack={() => setScreen('dashboard')}
          currentSheets={sheets}
          sheetNames={sheetNames}
        />
      )}

      {screen === 'dashboard' && (
        <div className="main-card fade-in dashboard-card">
          <div className="dash-header">
            <div className="dash-header-left">
              <FileSpreadsheet size={28} className="header-icon" />
              <div>
                <h2>Student Data Manager</h2>
                <p className="header-sub">View, edit, and analyze student performance data</p>
              </div>
            </div>
          </div>

          <div className="file-info-bar">
            <div className="info-chip"><strong>File:</strong> {fileName}</div>
            <div className="info-chip"><strong>Sheets:</strong> {displaySheetNames.length}</div>
            <div className="info-chip"><strong>Active:</strong> {simplifySheetDisplayName(activeSheet)}</div>
            <div className="info-chip highlight"><strong>Students:</strong> {totalStudents}</div>
          </div>

          <div className="dashboard-layout">
            <aside className="sheet-sidebar">
              <div className="sidebar-title">
                <FileSpreadsheet size={18} />
                <span>Workbook Sheets</span>
              </div>
              <div className="sheet-list">
                {displaySheetNames.map(name => (
                  <button
                    key={name}
                    className={`sheet-list-item ${activeSheet === name ? 'active' : ''}`}
                    title={name}
                    onClick={() => {
                      setActiveSheet(name);
                      setShowAnalysis(false);
                      setSortConfig({ key: null, direction: null });
                    }}
                  >
                    <FileSpreadsheet size={16} className="sheet-icon" />
                    <span className="sheet-name">{simplifySheetDisplayName(name)}</span>
                  </button>
                ))}
              </div>
            </aside>

            <div className="sheet-content">
              <div className="toolbar">
                <input
                  ref={importMoreRef}
                  type="file"
                  multiple
                  accept=".xls,.xlsx,.csv"
                  style={{ display: 'none' }}
                  onChange={e => handleFiles(e.target.files)}
                />
                {dbEnabled && (
                  <button className="tool-btn outline" onClick={() => importMoreRef.current?.click()}>
                    <Database size={16} /> Import More Files
                  </button>
                )}
                <button className="tool-btn primary" onClick={addStudent} disabled={isCumulativeView}><Plus size={16} /> Add Student</button>
                <button className="tool-btn success" onClick={() => { setShowAnalysis(true); }} disabled={isCumulativeView}><BarChart3 size={16} /> Class Analysis</button>
                <button className="tool-btn outline" onClick={exportExcel} disabled={loading}>
                  {loading ? <Loader2 size={16} className="spin" /> : <FileDown size={16} />} Export Excel
                </button>
                <button className="tool-btn outline" onClick={exportCSV}><FileText size={16} /> Export CSV</button>
                <button className="tool-btn outline" onClick={exportPDF} disabled={pdfExporting}>
                  {pdfExporting ? <Loader2 size={16} className="spin" /> : <Printer size={16} />} Export PDF
                </button>
                <button className="tool-btn outline" onClick={undo} disabled={historyIndex < 0} title="Undo (Ctrl+Z)"><Undo2 size={16} /> Undo</button>
                <button className="tool-btn outline" onClick={redo} disabled={historyIndex >= editHistory.length - 1} title="Redo (Ctrl+Y)"><Redo2 size={16} /> Redo</button>
                <button className="tool-btn outline" onClick={addSheet} disabled={isCumulativeView}><Plus size={16} /> New Sheet</button>
                <button className="tool-btn outline" onClick={() => setShowExamManager(true)}><Save size={16} /> Save Exam</button>
                <button className="tool-btn primary" onClick={() => { setScreen('progress'); setShowProgressTracker(true); }}><Network size={16} /> Progress Tracker</button>
                <button className="tool-btn danger-outline" onClick={resetAll}><RotateCcw size={16} /> Reset All</button>
                <div className="search-box">
                  <Search size={15} className="search-icon" />
                  <input className="search-input" placeholder="Search students..." value={searchQuery} onChange={e => setSearchQuery(e.target.value)} />
                  {searchQuery && <button className="search-clear" onClick={() => setSearchQuery('')}><X size={14} /></button>}
                </div>
              </div>

              {cumulativeSheet && (
                <div style={{ marginBottom: '1rem', padding: '0.9rem 1rem', borderRadius: '16px', background: 'rgba(16,185,129,0.08)', border: '1px solid rgba(16,185,129,0.14)' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', gap: '1rem', flexWrap: 'wrap', alignItems: 'center' }}>
                    <div>
                      <h3 style={{ margin: 0, fontSize: '1rem', color: 'var(--text-primary)' }}>Class 10 cumulative sheet prepared</h3>
                      <p style={{ margin: '0.25rem 0 0', color: 'var(--text-secondary)', fontSize: '0.9rem' }}>
                        Enrollment-based merge of <strong>Class 9 + Target</strong>, <strong>HY</strong>, <strong>PB1</strong>, <strong>PB2</strong> and <strong>Board</strong>.
                      </p>
                    </div>
                    <button className="tool-btn outline" onClick={() => setActiveSheet(CUMULATIVE_SHEET_NAME)}>
                      <Database size={16} /> Open cumulative sheet
                    </button>
                  </div>
                </div>
              )}

              {(baselineMatchReport?.unmatchedCount || 0) > 0 && (
                <BaselineMatchWarningPanel
                  report={baselineMatchReport}
                  studentOptions={cumulativeStudentOptions}
                  onApplyMapping={applyBaselineMapping}
                />
              )}

              {classStats && (
                <div className="stats-bar fade-in">
                  <div className="stat-card">
                    <div className="stat-icon"><Users size={18} /></div>
                    <div><span className="stat-label">Class Average</span><span className="stat-value">{classStats.avg}</span></div>
                  </div>
                  <div className="stat-card stat-top">
                    <div className="stat-icon"><TrendingUp size={18} /></div>
                    <div><span className="stat-label">Top Scorer</span><span className="stat-value">{classStats.topName} <small>({classStats.topScore})</small></span></div>
                  </div>
                  <div className="stat-card stat-bottom">
                    <div className="stat-icon"><TrendingDown size={18} /></div>
                    <div><span className="stat-label">Needs Attention</span><span className="stat-value">{classStats.bottomName} <small>({classStats.bottomScore})</small></span></div>
                  </div>
                  <div className="stat-card">
                    <div className="stat-icon"><Award size={18} /></div>
                    <div><span className="stat-label">Scored Students</span><span className="stat-value">{classStats.totalScored} / {totalStudents}</span></div>
                  </div>
                </div>
              )}

              {showAnalysis ? (
                <AnalysisPanel
                  data={analysisData}
                  subjectComparison={subjectComparison}
                  sheets={sheetNames}
                  selected={analysisSheets}
                  setSelected={setAnalysisSheets}
                  onClose={() => setShowAnalysis(false)}
                />
              ) : effectiveCurrentSheet ? (
                <>
                  <div className="table-container">
                    <table className="data-table">
                      <thead>
                        <tr>
                          <th className="row-num">#</th>
                          {classStats && <th className="rank-col">Rank</th>}
                          {effectiveCurrentSheet.headers.map(h => (
                            <th key={h} className="sortable-th" onClick={() => handleSort(h)}>
                              <span className="th-content">
                                {h}
                                <span className="sort-icon">
                                  {sortConfig.key === h ? (
                                    sortConfig.direction === 'asc' ? <ChevronUp size={14} /> : <ChevronDown size={14} />
                                  ) : (
                                    <ArrowUpDown size={12} className="sort-idle" />
                                  )}
                                </span>
                              </span>
                            </th>
                          ))}
                          <th className="actions-col">Actions</th>
                        </tr>
                      </thead>
                      <tbody>
                        {sortedRows.length === 0 ? (
                          <tr><td colSpan={effectiveCurrentSheet.headers.length + (classStats ? 3 : 2)} className="empty-msg">
                            {searchQuery ? `No students matching "${searchQuery}"` : (isCumulativeView ? 'No cumulative records available yet.' : 'No students added yet. Click "Add Student" to begin.')}
                          </td></tr>
                        ) : sortedRows.map((row, fi) => {
                          const ri = effectiveCurrentSheet.rows.indexOf(row);
                          const rank = classStats?.rankMap?.get(row);
                          return (
                            <tr key={ri}>
                              <td className="row-num">{ri + 1}</td>
                              {classStats && (
                                <td className="rank-col">
                                  {rank ? (
                                    <span className={`rank-badge ${rank <= 3 ? 'rank-top' : rank <= 10 ? 'rank-mid' : ''}`}>
                                      {rank <= 3 ? ['🥇', '🥈', '🥉'][rank - 1] : `#${rank}`}
                                    </span>
                                  ) : '—'}
                                </td>
                              )}
                              {effectiveCurrentSheet.headers.map(h => {
                                const isEditing = editingCell && editingCell.sheetName === activeSheet && editingCell.rowIdx === ri && editingCell.col === h;
                                const val = row[h];
                                const isSubjectCol = getSubjectColumns(effectiveCurrentSheet.headers).includes(h);
                                const notOpted = isSubjectCol && isNotOpted(val);
                                const scoreClass = (h.toLowerCase().includes('%') || h.toLowerCase().includes('total')) ? getScoreColor(val) : '';
                                const isBlankValue = val === '' || val === null || val === undefined;
                                const displayValue = notOpted
                                  ? '—'
                                  : isBlankValue
                                    ? (isCumulativeView ? '' : '—')
                                    : String(val);
                                return (
                                  <td key={h} className={`data-cell ${scoreClass} ${notOpted ? 'not-opted-cell' : ''}`} onDoubleClick={() => !isCumulativeView && startEdit(activeSheet, ri, h)}>
                                    {isEditing ? (
                                      <input
                                        className="cell-input"
                                        value={editValue}
                                        onChange={e => setEditValue(e.target.value)}
                                        onBlur={saveEdit}
                                        onKeyDown={e => { if (e.key === 'Enter') saveEdit(); if (e.key === 'Escape') cancelEdit(); }}
                                        autoFocus
                                      />
                                    ) : (
                                      <span>{displayValue}</span>
                                    )}
                                  </td>
                                );
                              })}
                              <td className="actions-col">
                                <button className="icon-btn view" title="View Report" onClick={() => openStudentReport(row)}>
                                  <Eye size={15} />
                                </button>
                                {!isCumulativeView && (
                                  <button className="icon-btn delete" title="Delete" onClick={() => deleteStudent(ri)}>
                                    <Trash2 size={15} />
                                  </button>
                                )}
                              </td>
                            </tr>
                          );
                        })}
                        {searchQuery && sortedRows.length > 0 && (
                          <tr><td colSpan={effectiveCurrentSheet.headers.length + (classStats ? 3 : 2)} className="search-info">Showing {sortedRows.length} of {effectiveCurrentSheet.rows.length} students</td></tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                  {!isCumulativeView && <SheetAnalysis sheetName={activeSheet} sheet={effectiveCurrentSheet} />}
                </>
              ) : null}
            </div>
          </div>


          <div className="footer-info">
            Target Analysis Report Generator — Double-click any cell to edit • Ctrl+Z/Y for undo/redo • Subjects marked with "-" are excluded from totals & graphs
          </div>
        </div>
      )}

      {!PHASE_ONE_CUMULATIVE_ONLY && studentReport && (
        <StudentReportModal
          student={studentReport}
          headers={effectiveCurrentSheet?.headers || []}
          sheetName={activeSheet}
          sheetNames={sheetNames}
          sheets={sheets}
          onClose={() => setStudentReport(null)}
        />
      )}

      {!PHASE_ONE_CUMULATIVE_ONLY && pdfExporting && (
        <div style={{ position: 'fixed', left: '-9999px', top: 0, zIndex: -1 }}>
          <PDFReportContent
            ref={pdfReportRef}
            fileName={fileName}
            sheetNames={sheetNames}
            sheets={sheets}
            analysisData={analysisData}
            subjectComparison={subjectComparison}
          />
        </div>
      )}

      <div className="toast-container">
        {toasts.map(toast => (
          <div key={toast.id} className={`toast toast-${toast.type}`}>
            <div className="toast-icon">
              {toast.type === 'success' && <CheckCircle2 size={18} />}
              {toast.type === 'error' && <AlertCircle size={18} />}
              {toast.type === 'warning' && <AlertCircle size={18} />}
              {toast.type === 'info' && <Info size={18} />}
            </div>
            <span className="toast-message">{toast.message}</span>
            <button className="toast-close" onClick={() => removeToast(toast.id)}><X size={14} /></button>
            {toast.duration > 0 && <div className="toast-progress" style={{ animationDuration: `${toast.duration}ms` }} />}
          </div>
        ))}
      </div>
    </div>
  );
}

function BaselineMatchWarningPanel({ report, studentOptions = [], onApplyMapping }) {
  const reviewRows = (report?.rows || [])
    .map((row, index) => ({ ...row, _reportIndex: index }))
    .filter((row) => row.confidence !== 'exact');
  const suggestedRows = reviewRows.filter((row) => row.confidence === 'fuzzy');
  const [manualTargets, setManualTargets] = useState({});

  const getSuggestedTargetKey = (row) => {
    if (row.suggestedStudentKey && studentOptions.some((student) => student.key === row.suggestedStudentKey)) {
      return row.suggestedStudentKey;
    }
    const enrollment = normalizeIdentifier(row.suggestedEnrollmentNo);
    if (enrollment) return `enrollment:${enrollment}`;
    const normalizedName = normalizeStudentName(row.suggestedStudentName);
    const normalizedSection = String(row.suggestedSection || '').trim().toLowerCase();
    if (normalizedName && normalizedSection) return `name-section:${normalizedName}|${normalizedSection}`;
    if (normalizedName) return `name:${normalizedName}`;
    return '';
  };

  const formatBaselineValue = (value) => isMissingProjectedValue(value) ? '—' : value;

  const exportCsv = () => {
    if (!reviewRows.length) return;

    const toCsvCell = (value) => `"${String(value ?? '').replace(/"/g, '""')}"`;
    const lines = [
      ['Baseline Student Name', 'Baseline Enrollment No', 'Section', 'Class 9 %', 'Target %', 'Reason', 'Suggested Student Name', 'Suggested Enrollment No', 'Suggested Section', 'Suggestion Score', 'Confidence'],
      ...reviewRows.map((row) => [
        row.baselineStudentName,
        row.baselineEnrollmentNo,
        row.section,
        row.baselineClass9Percent,
        row.baselineTargetPercent,
        row.reason,
        row.suggestedStudentName,
        row.suggestedEnrollmentNo,
        row.suggestedSection,
        row.suggestionScore,
        row.confidence,
      ]),
    ];
    const csv = lines.map((line) => line.map(toCsvCell).join(',')).join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'baseline-mismatch-report.csv';
    link.click();
    URL.revokeObjectURL(link.href);
  };

  if (!reviewRows.length) return null;

  return (
    <div className="baseline-review-panel">
      <div className="baseline-review-head">
        <div>
          <h3>Baseline rows need review</h3>
          <p>
            Matched {report?.matchedCount || 0} baseline rows automatically. {reviewRows.length} row{reviewRows.length > 1 ? 's' : ''} still need approval or manual mapping.
          </p>
          <div className="baseline-review-stats">
            <span>{suggestedRows.length} suggested</span>
            <span>{reviewRows.length - suggestedRows.length} without safe suggestion</span>
          </div>
        </div>
        <button className="tool-btn outline" onClick={exportCsv}>
          <Download size={16} /> Export review CSV
        </button>
      </div>

      <div className="baseline-review-table-wrap">
        <table className="data-table">
          <thead>
            <tr>
              <th>Baseline Student</th>
              <th>Class IX / Target</th>
              <th>Section</th>
              <th>Suggested Match</th>
              <th>Manual Match</th>
              <th>Action</th>
            </tr>
          </thead>
          <tbody>
            {reviewRows.slice(0, 25).map((row, index) => {
              const rowKey = buildBaselineReportRowKey(row, row._reportIndex);
              const suggestedTargetKey = getSuggestedTargetKey(row);
              const suggestedOption = studentOptions.find((student) => student.key === suggestedTargetKey);
              const selectedTargetKey = manualTargets[rowKey] || '';
              const canApplySuggestion = row.confidence === 'fuzzy' && suggestedTargetKey && !suggestedOption?.hasBaselineValues;

              return (
                <tr key={rowKey}>
                  <td>
                    <strong>{row.baselineStudentName || '—'}</strong>
                    <span className="baseline-review-sub">{row.baselineEnrollmentNo || 'No enrollment'}</span>
                  </td>
                  <td>
                    <strong>{formatBaselineValue(row.baselineClass9Percent)} / {formatBaselineValue(row.baselineTargetPercent)}</strong>
                    <span className="baseline-review-sub">Class IX / Target</span>
                  </td>
                <td>{row.section || '—'}</td>
                  <td>
                    {row.suggestedStudentName ? (
                      <>
                        <strong>{row.suggestedStudentName}</strong>
                        <span className="baseline-review-sub">
                          {[row.suggestedEnrollmentNo, row.suggestedSection ? `Sec ${row.suggestedSection}` : '', row.suggestionScore ? `${Math.round(Number(row.suggestionScore) * 100)}%` : '']
                            .filter(Boolean)
                            .join(' • ')}
                        </span>
                        {suggestedOption?.hasBaselineValues && <span className="baseline-review-sub">Already has exact baseline values</span>}
                      </>
                    ) : (
                      <span className="baseline-review-sub">No safe candidate</span>
                    )}
                    <span className={`baseline-review-chip ${row.confidence === 'fuzzy' ? 'suggested' : ''}`}>{row.reason || row.confidence}</span>
                  </td>
                  <td>
                    <select
                      className="baseline-review-select"
                      value={selectedTargetKey}
                      onChange={(event) => setManualTargets((prev) => ({ ...prev, [rowKey]: event.target.value }))}
                    >
                      <option value="">Choose student...</option>
                      {studentOptions.map((student) => (
                        <option key={student.key} value={student.key} disabled={student.hasBaselineValues}>
                          {student.studentName || 'Unnamed'}{student.section ? ` - ${student.section}` : ''}{student.enrollmentNo ? ` (${student.enrollmentNo})` : ''}{student.hasBaselineValues ? ' - already filled' : ''}
                        </option>
                      ))}
                    </select>
                  </td>
                  <td>
                    <div className="baseline-review-actions">
                      <button
                        className="tool-btn success"
                        disabled={!canApplySuggestion}
                        onClick={() => onApplyMapping(row, suggestedTargetKey, 'suggested')}
                      >
                        <CheckCircle2 size={15} /> Approve
                      </button>
                      <button
                        className="tool-btn outline"
                        disabled={!selectedTargetKey}
                        onClick={() => onApplyMapping(row, selectedTargetKey, 'manual')}
                      >
                        Apply manual
                      </button>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {reviewRows.length > 25 && (
        <div className="baseline-review-foot">
          Showing 25 of {reviewRows.length} baseline review rows. Use the CSV export for the full list.
        </div>
      )}
    </div>
  );
}

function StructuredUploadPanel({ loading, workflowMode, setWorkflowMode, structuredUpload, setStructuredFile, onSubmit }) {
  const activeUpload = structuredUpload[workflowMode] || {};

  const mergedExamInput = (groupKey, label) => {
    const file = activeUpload[groupKey] || null;
    return (
      <label key={groupKey} className="structured-base-input">
        <span className="structured-slot-title">{label} — Class 10 (All Sections)</span>
        <input type="file" accept=".xls,.xlsx,.csv" onChange={(e) => setStructuredFile(workflowMode, groupKey, null, e.target.files?.[0] || null)} />
        <span className="structured-slot-file">{file ? file.name : 'No file selected'}</span>
      </label>
    );
  };

  const boardFile = activeUpload.board;

  return (
    <div className="structured-panel">
      <div className="structured-panel-head">
        <div className="structured-panel-copy">
          <span className="structured-panel-eyebrow">5-file workflow</span>
          <h3>Class 10 five-file import</h3>
          <p>
            Use one file for each stage: Baseline, HY, PB1, PB2, and Board. HY / PB1 / PB2 can be multi-sheet workbooks with section tabs, or a single combined sheet that includes `Section` or `Class Section` data.
          </p>
        </div>
        <button className="tool-btn primary structured-submit-btn" onClick={onSubmit} disabled={loading}>
          {loading ? <Loader2 size={16} className="spin" /> : <Upload size={16} />}
          Import 5 files
        </button>
      </div>

      <div className="structured-panel-body">
        <div className="structured-base-card">
          <div className="structured-section-head">
            <h4>Baseline sheet</h4>
            <span>Use the Class 9 + Target sheet here</span>
          </div>
          <label className="structured-base-input">
            <span className="structured-slot-title">Class 9 + Target</span>
            <input type="file" accept=".xls,.xlsx,.csv" onChange={(e) => setStructuredFile(workflowMode, 'baseline', null, e.target.files?.[0] || null)} />
            <span className="structured-slot-file">{activeUpload.baseline ? activeUpload.baseline.name : 'No file selected'}</span>
          </label>
        </div>

        {[
          ['hy', 'Half Yearly'],
          ['pb1', 'Preboard 1'],
          ['pb2', 'Preboard 2'],
        ].map(([groupKey, label]) => (
          <div key={groupKey} className="structured-exam-card">
            <div className="structured-section-head">
              <h4>{label}</h4>
              <span>Upload one file for all sections. Multi-sheet workbooks are supported.</span>
            </div>
            {mergedExamInput(groupKey, label)}
          </div>
        ))}

        <div className="structured-exam-card">
          <div className="structured-section-head">
            <h4>Board Result</h4>
            <span>Upload one combined board-result file for all sections</span>
          </div>
          <label className="structured-base-input">
            <span className="structured-slot-title">Board Result — Class 10 (All Sections)</span>
            <input type="file" accept=".xls,.xlsx,.csv" onChange={(e) => setStructuredFile(workflowMode, 'board', null, e.target.files?.[0] || null)} />
            <span className="structured-slot-file">{boardFile ? boardFile.name : 'No file selected'}</span>
          </label>
        </div>
      </div>
    </div>
  );
}

function AnalysisPanel({ data, subjectComparison, sheets, selected, setSelected, onClose }) {
  const toggleSheet = (name) => {
    setSelected(prev => prev.includes(name) ? prev.filter(s => s !== name) : [...prev, name]);
  };

  if (!data) return null;

  const pieData = data.rows.slice(0, -1).map(r => ({
    name: r.Range,
    value: r.students,
  }));
  const PIE_COLORS = ['#22c55e', '#3b82f6', '#8b5cf6', '#f59e0b', '#f97316', '#ef4444'];

  return (
    <div className="analysis-panel fade-in">
      <div className="analysis-header">
        <h3><BarChart3 size={20} /> Section-wise Target Analysis</h3>
        <button className="icon-btn" onClick={onClose}><X size={18} /></button>
      </div>

      <div className="analysis-sheets">
        <span className="label">Include sheets:</span>
        {sheets.map(s => (
          <button key={s} className={`sheet-pill ${selected.includes(s) ? 'active' : ''}`} onClick={() => toggleSheet(s)}>
            {selected.includes(s) && <CheckCircle size={14} />} {s}
          </button>
        ))}
      </div>

      <div className="chart-box" id="section-analysis-charts" style={{ padding: '2rem 1.5rem', background: '#fff' }}>
        <h4 style={{ marginBottom: '1.5rem', color: 'var(--text)', borderBottom: '1px solid var(--border)', paddingBottom: '0.75rem', fontSize: '1rem' }}>
          {data.sections.map(s => s.replace(/^X\s*/, 'X-')).join(', ')} — Performance Visuals
        </h4>

        <div style={{ display: 'grid', gridTemplateColumns: 'minmax(0, 1fr) minmax(0, 1fr)', gap: '2rem', marginBottom: '2rem' }}>
          <div className="chart-inner-box">
            <h5 className="chart-subtitle">Section-wise Bar Analysis</h5>
            <ResponsiveContainer width="100%" height={300}>
              <BarChart data={data.rows.slice(0, -1)} margin={{ top: 10, right: 10, left: 0, bottom: 5 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.05)" />
                <XAxis dataKey="Range" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip contentStyle={{ borderRadius: 10, border: 'none', boxShadow: '0 4px 20px rgba(0,0,0,0.1)' }} />
                <Legend />
                {data.sections.map((s, i) => (
                  <Bar key={s} dataKey={s} fill={CHART_COLORS[i % CHART_COLORS.length]} radius={[4, 4, 0, 0]} />
                ))}
              </BarChart>
            </ResponsiveContainer>
          </div>

          <div className="chart-inner-box">
            <h5 className="chart-subtitle">Trend Line Analysis</h5>
            <ResponsiveContainer width="100%" height={300}>
              <LineChart data={data.rows.slice(0, -1)} margin={{ top: 10, right: 10, left: 0, bottom: 5 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.05)" />
                <XAxis dataKey="Range" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip contentStyle={{ borderRadius: 10, border: 'none', boxShadow: '0 4px 20px rgba(0,0,0,0.1)' }} />
                <Legend />
                {data.sections.map((s, i) => (
                  <Line key={s} type="monotone" dataKey={s} stroke={CHART_COLORS[i % CHART_COLORS.length]} strokeWidth={3} activeDot={{ r: 6 }} />
                ))}
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: 'minmax(0, 1fr) minmax(0, 1fr)', gap: '2rem' }}>
          <div className="chart-inner-box">
            <h5 className="chart-subtitle">Overall Distribution</h5>
            <ResponsiveContainer width="100%" height={300}>
              <PieChart>
                <Pie
                  data={pieData.filter(d => d.value > 0)}
                  cx="50%" cy="50%"
                  innerRadius={60} outerRadius={110}
                  paddingAngle={3}
                  dataKey="value"
                  label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`}
                >
                  {pieData.map((entry, i) => (
                    <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />
                  ))}
                </Pie>
                <Tooltip />
                <Legend />
              </PieChart>
            </ResponsiveContainer>
          </div>

          {subjectComparison && subjectComparison.data.length > 0 && (
            <div className="chart-inner-box">
              <h5 className="chart-subtitle">Subject-wise Average (Section Comparison)</h5>
              <ResponsiveContainer width="100%" height={300}>
                <BarChart data={subjectComparison.data} margin={{ top: 10, right: 10, left: 0, bottom: 5 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.05)" />
                  <XAxis dataKey="subject" tick={{ fontSize: 10 }} interval={0} angle={-25} textAnchor="end" height={60} />
                  <YAxis tick={{ fontSize: 11 }} />
                  <Tooltip contentStyle={{ borderRadius: 10, border: 'none', boxShadow: '0 4px 20px rgba(0,0,0,0.1)' }} />
                  <Legend />
                  {subjectComparison.sections.map((s, i) => (
                    <Bar key={s} dataKey={s} fill={CHART_COLORS[i % CHART_COLORS.length]} radius={[4, 4, 0, 0]} />
                  ))}
                </BarChart>
              </ResponsiveContainer>
            </div>
          )}
        </div>
      </div>

      <div className="table-container">
        <table className="data-table analysis-tbl">
          <thead>
            <tr>
              {data.headers.map(h => <th key={h}>{h}</th>)}
            </tr>
          </thead>
          <tbody>
            {data.rows.map((row, i) => (
              <tr key={i} className={i === data.rows.length - 1 ? 'total-row' : ''}>
                {data.headers.map(h => (
                  <td key={h} className={h === 'Range' ? `range-cell range-${row[h]?.replace(/[^a-zA-Z0-9]/g, '')}` : ''}>
                    {h === 'per%' ? `${row[h]}%` : row[h]}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function SheetAnalysis({ sheetName, sheet }) {
  const analysis = useMemo(() => computeSheetAnalysis(sheet), [sheet]);
  if (!analysis) return null;

  const chartData = analysis.rows.slice(0, -1);
  const chartId = `sheet-chart-${sheetName.replace(/\s+/g, '_')}`;

  return (
    <div className="sheet-analysis fade-in">
      <h3 className="sheet-analysis-title">
        <BarChart3 size={18} /> {sheetName} — Score Distribution
      </h3>
      <div className="sheet-analysis-body">
        <div className="sheet-analysis-table">
          <table className="data-table analysis-tbl mini-tbl">
            <thead>
              <tr>
                <th>Range</th>
                <th>Students</th>
              </tr>
            </thead>
            <tbody>
              {analysis.rows.map((row, i) => (
                <tr key={i} className={i === analysis.rows.length - 1 ? 'total-row' : ''}>
                  <td className={`range-cell range-${row.Range?.replace(/[^a-zA-Z0-9]/g, '')}`}>
                    {row.Range}
                  </td>
                  <td style={{ textAlign: 'center', fontWeight: 600 }}>{row.Count}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className="sheet-analysis-chart" id={chartId}>
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={chartData} margin={{ top: 10, right: 10, left: 0, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.05)" />
              <XAxis dataKey="Range" tick={{ fontSize: 11 }} />
              <YAxis tick={{ fontSize: 11 }} />
              <Tooltip />
              <Legend />
              <Bar dataKey="Count" name={sheetName} radius={[4, 4, 0, 0]}>
                {chartData.map((entry, i) => (
                  <Cell key={i} fill={entry.color} />
                ))}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
}

const PDFReportContent = React.forwardRef(function PDFReportContent({ fileName, sheetNames, sheets, analysisData, subjectComparison }, ref) {
  const PIE_COLORS = ['#22c55e', '#3b82f6', '#8b5cf6', '#f59e0b', '#f97316', '#ef4444'];

  const sheetAnalyses = sheetNames.map(name => {
    const sheet = sheets[name];
    if (!sheet) return null;
    const analysis = computeSheetAnalysis(sheet);
    return analysis ? { name, ...analysis } : null;
  }).filter(Boolean);

  const pieData = analysisData ? analysisData.rows.slice(0, -1).map(r => ({
    name: r.Range,
    value: r.students,
  })) : [];

  return (
    <div ref={ref} style={{ width: '1050px', padding: '40px', background: '#fff', fontFamily: 'Inter, system-ui, sans-serif', color: '#1e293b' }}>
      <div className="pdf-page-block" style={{ textAlign: 'center', marginBottom: '15px', borderBottom: '3px solid #6366f1', paddingBottom: '20px' }}>
        <h1 style={{ fontSize: '24px', fontWeight: 700, color: '#1e293b', margin: 0 }}>📊 Target Analysis Report</h1>
        <p style={{ fontSize: '14px', color: '#64748b', margin: '8px 0 0' }}>{fileName} • Generated on {new Date().toLocaleDateString('en-IN', { year: 'numeric', month: 'long', day: 'numeric' })}</p>
      </div>

      {sheetAnalyses.map((sa) => (
        <div key={sa.name} className="pdf-page-block" style={{ marginBottom: '15px', padding: '10px 0' }}>
          <h2 style={{ fontSize: '16px', fontWeight: 700, color: '#4f46e5', marginBottom: '12px', borderLeft: '4px solid #6366f1', paddingLeft: '10px' }}>
            {sa.name} — Score Distribution (Target: {sa.targetCol})
          </h2>
          <div style={{ display: 'flex', gap: '20px', alignItems: 'flex-start' }}>
            <table style={{ borderCollapse: 'collapse', fontSize: '13px', minWidth: '220px' }}>
              <thead>
                <tr>
                  <th style={{ padding: '8px 14px', background: '#4f46e5', color: '#fff', textAlign: 'center', borderRadius: '6px 0 0 0' }}>Range</th>
                  <th style={{ padding: '8px 14px', background: '#4f46e5', color: '#fff', textAlign: 'center', borderRadius: '0 6px 0 0' }}>Students</th>
                </tr>
              </thead>
              <tbody>
                {sa.rows.map((row, i) => (
                  <tr key={i} style={{ background: i === sa.rows.length - 1 ? '#e0f2fe' : i % 2 === 0 ? '#f8fafc' : '#fff' }}>
                    <td style={{ padding: '6px 14px', borderBottom: '1px solid #e2e8f0', fontWeight: i === sa.rows.length - 1 ? 700 : 500 }}>{row.Range}</td>
                    <td style={{ padding: '6px 14px', borderBottom: '1px solid #e2e8f0', textAlign: 'center', fontWeight: 600 }}>{row.Count}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div style={{ flex: 1 }}>
              <BarChart width={730} height={240} data={sa.rows.slice(0, -1)} margin={{ top: 5, right: 10, left: 0, bottom: 5 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.06)" />
                <XAxis dataKey="Range" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip />
                <Bar dataKey="Count" name={sa.name} radius={[4, 4, 0, 0]} isAnimationActive={false}>
                  {sa.rows.slice(0, -1).map((entry, i) => (
                    <Cell key={i} fill={entry.color} />
                  ))}
                </Bar>
              </BarChart>
            </div>
          </div>
        </div>
      ))}

      {analysisData && analysisData.sections.length > 0 && (
        <div style={{ marginTop: '10px' }}>
          <div className="pdf-page-block" style={{ marginBottom: '15px', padding: '10px 0' }}>
            <h2 style={{ fontSize: '18px', fontWeight: 700, color: '#1e293b', marginBottom: '16px', textAlign: 'center', borderBottom: '2px solid #e2e8f0', paddingBottom: '10px' }}>
              📈 Cumulative Section-wise Analysis
            </h2>

            <table style={{ borderCollapse: 'collapse', fontSize: '13px', width: '100%', marginBottom: '25px' }}>
              <thead>
                <tr>
                  {analysisData.headers.map(h => (
                    <th key={h} style={{ padding: '8px 12px', background: '#4f46e5', color: '#fff', textAlign: 'center' }}>{h === 'per%' ? '%' : h === 'students' ? 'Total' : h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {analysisData.rows.map((row, i) => (
                  <tr key={i} style={{ background: i === analysisData.rows.length - 1 ? '#e0f2fe' : i % 2 === 0 ? '#f8fafc' : '#fff' }}>
                    {analysisData.headers.map(h => (
                      <td key={h} style={{ padding: '6px 12px', borderBottom: '1px solid #e2e8f0', textAlign: 'center', fontWeight: i === analysisData.rows.length - 1 || h === 'Range' ? 700 : 400 }}>
                        {h === 'per%' ? `${row[h]}%` : row[h]}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="pdf-page-block" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '25px', marginBottom: '15px', padding: '10px 0' }}>
            <div>
              <h3 style={{ fontSize: '13px', fontWeight: 600, color: '#475569', textAlign: 'center', marginBottom: '8px' }}>Section-wise Bar Comparison</h3>
              <BarChart width={460} height={280} data={analysisData.rows.slice(0, -1)} margin={{ top: 10, right: 10, left: 0, bottom: 5 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.05)" />
                <XAxis dataKey="Range" tick={{ fontSize: 10 }} />
                <YAxis tick={{ fontSize: 10 }} />
                <Tooltip />
                <Legend wrapperStyle={{ fontSize: '11px' }} />
                {analysisData.sections.map((s, i) => (
                  <Bar key={s} dataKey={s} fill={CHART_COLORS[i % CHART_COLORS.length]} radius={[3, 3, 0, 0]} isAnimationActive={false} />
                ))}
              </BarChart>
            </div>

            <div>
              <h3 style={{ fontSize: '13px', fontWeight: 600, color: '#475569', textAlign: 'center', marginBottom: '8px' }}>Trend Line Analysis</h3>
              <LineChart width={460} height={280} data={analysisData.rows.slice(0, -1)} margin={{ top: 10, right: 10, left: 0, bottom: 5 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.05)" />
                <XAxis dataKey="Range" tick={{ fontSize: 10 }} />
                <YAxis tick={{ fontSize: 10 }} />
                <Tooltip />
                <Legend wrapperStyle={{ fontSize: '11px' }} />
                {analysisData.sections.map((s, i) => (
                  <Line key={s} type="monotone" dataKey={s} stroke={CHART_COLORS[i % CHART_COLORS.length]} strokeWidth={2.5} dot={{ r: 4 }} isAnimationActive={false} />
                ))}
              </LineChart>
            </div>
          </div>

          <div className="pdf-page-block" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '25px', marginBottom: '15px', padding: '10px 0' }}>
            <div>
              <h3 style={{ fontSize: '13px', fontWeight: 600, color: '#475569', textAlign: 'center', marginBottom: '8px' }}>Overall Distribution</h3>
              <PieChart width={460} height={280}>
                <Pie data={pieData.filter(d => d.value > 0)} cx="50%" cy="50%" innerRadius={55} outerRadius={100} paddingAngle={3} dataKey="value"
                  label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`} isAnimationActive={false}
                >
                  {pieData.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                </Pie>
                <Tooltip />
                <Legend wrapperStyle={{ fontSize: '11px' }} />
              </PieChart>
            </div>

            {subjectComparison && subjectComparison.data.length > 0 && (
              <div>
                <h3 style={{ fontSize: '13px', fontWeight: 600, color: '#475569', textAlign: 'center', marginBottom: '8px' }}>Subject-wise Average</h3>
                <BarChart width={460} height={280} data={subjectComparison.data} margin={{ top: 10, right: 10, left: 0, bottom: 5 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.05)" />
                  <XAxis dataKey="subject" tick={{ fontSize: 9 }} interval={0} angle={-25} textAnchor="end" height={55} />
                  <YAxis tick={{ fontSize: 10 }} />
                  <Tooltip />
                  <Legend wrapperStyle={{ fontSize: '11px' }} />
                  {subjectComparison.sections.map((s, i) => (
                    <Bar key={s} dataKey={s} fill={CHART_COLORS[i % CHART_COLORS.length]} radius={[3, 3, 0, 0]} isAnimationActive={false} />
                  ))}
                </BarChart>
              </div>
            )}
          </div>
        </div>
      )}

      <div className="pdf-page-block" style={{ textAlign: 'center', fontSize: '11px', color: '#94a3b8', borderTop: '1px solid #e2e8f0', paddingTop: '15px', marginTop: '15px' }}>
        <p>Auto-generated by Target Analysis Report Generator • Subjects marked with "-" are excluded from totals & charts</p>
      </div>
    </div>
  );
});

function StudentReportModal({ student, headers, sheetName, sheetNames, sheets, onClose }) {
  const [downloading, setDownloading] = useState(false);
  const reportRef = useRef(null);
  const comparisonData = useMemo(
    () => buildStudentComparisonData(student, headers, sheetName, sheetNames, sheets),
    [student, headers, sheetName, sheetNames, sheets]
  );

  const subjectCols = getSubjectColumns(headers);
  const optedSubjects = subjectCols.filter(h => !isNotOpted(student[h]));
  const notOptedSubjects = subjectCols.filter(h => isNotOpted(student[h]));
  const nameCol = findNameColumn(headers);
  const name = nameCol ? student[nameCol] : 'Student';

  const obtainedMarks = comparisonData?.obtainedMarks ?? 0;
  const maxMarks = comparisonData?.maxMarks ?? 0;
  const calculatedPercent = comparisonData?.examPercent ?? (maxMarks > 0 ? parseFloat(((obtainedMarks / maxMarks) * 100).toFixed(2)) : null);
  const class9Percent = comparisonData?.class9Percent ?? null;
  const targetPercent = comparisonData?.targetPercent ?? null;
  const examMatrixData = comparisonData?.exams?.map((exam) => ({
    examName: exam.examName,
    examPercent: exam.examPercent,
    targetPercent: exam.targetPercent ?? targetPercent,
    class9Percent,
  })) || [];

  const barData = (comparisonData?.subjectBreakdown || []).map((subject, i) => ({
    name: subject.subject,
    value: subject.score || 0,
    maxMarks: subject.maxScore || 80,
    color: CHART_COLORS[i % CHART_COLORS.length]
  }));

  const radarData = barData.map((subject) => ({
    subject: subject.name,
    score: subject.value,
    percentage: parseFloat(((subject.value / (subject.maxMarks || 80)) * 100).toFixed(1)),
  }));

  const metaFields = headers.filter(h => {
    return !subjectCols.includes(h) &&
      !h.toLowerCase().includes('grand total') &&
      !h.toLowerCase().includes('total') &&
      !h.toLowerCase().includes('%') &&
      !h.toLowerCase().includes('target');
  });

  const downloadPDF = async () => {
    if (!reportRef.current) return;
    setDownloading(true);

    try {
      await new Promise(r => setTimeout(r, 500));

      const canvas = await html2canvas(reportRef.current, {
        backgroundColor: '#ffffff',
        scale: 2,
        useCORS: true,
        logging: false,
        windowWidth: 800,
      });

      const imgData = canvas.toDataURL('image/png');
      const pdf = new jsPDF('p', 'mm', 'a4');

      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();
      const margin = 10;
      const usableWidth = pageWidth - (margin * 2);

      const imgWidth = canvas.width;
      const imgHeight = canvas.height;
      const ratio = usableWidth / imgWidth;
      const scaledHeight = imgHeight * ratio;

      if (scaledHeight <= pageHeight - (margin * 2)) {
        pdf.addImage(imgData, 'PNG', margin, margin, usableWidth, scaledHeight);
      } else {
        let remainingHeight = imgHeight;
        let sourceY = 0;
        const pageContentHeight = (pageHeight - (margin * 2)) / ratio;

        while (remainingHeight > 0) {
          const sliceHeight = Math.min(pageContentHeight, remainingHeight);
          const tempCanvas = document.createElement('canvas');
          tempCanvas.width = imgWidth;
          tempCanvas.height = sliceHeight;
          const ctx = tempCanvas.getContext('2d');
          ctx.drawImage(canvas, 0, sourceY, imgWidth, sliceHeight, 0, 0, imgWidth, sliceHeight);

          const sliceData = tempCanvas.toDataURL('image/png');
          const scaledSliceHeight = sliceHeight * ratio;

          if (sourceY > 0) pdf.addPage();
          pdf.addImage(sliceData, 'PNG', margin, margin, usableWidth, scaledSliceHeight);

          sourceY += sliceHeight;
          remainingHeight -= sliceHeight;
        }
      }

      pdf.save(`${name}_Report.pdf`);
    } catch (err) {
      console.error('PDF generation failed:', err);
      alert('PDF download failed. Try again.');
    }
    setDownloading(false);
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-card fade-in" onClick={e => e.stopPropagation()} style={{ maxWidth: '850px' }}>
        <div className="modal-header">
          <h3><UserCheck size={20} /> Student Report</h3>
          <div style={{ display: 'flex', gap: '0.5rem' }}>
            <button className="tool-btn primary" onClick={downloadPDF} disabled={downloading}>
              {downloading ? <Loader2 size={14} className="spin" /> : <Download size={14} />}
              {downloading ? ' Generating...' : ' Download PDF'}
            </button>
            <button className="icon-btn" onClick={onClose}><X size={18} /></button>
          </div>
        </div>

        <div ref={reportRef} className="pdf-content">
          <div className="pdf-header">
            <h2 className="pdf-title">Student Performance Report</h2>
            <p className="pdf-subtitle">{sheetName} • Generated on {new Date().toLocaleDateString('en-IN', { year: 'numeric', month: 'long', day: 'numeric' })}</p>
          </div>

          <div className="report-name-bar">
            <div className="student-avatar">{name?.charAt(0)}</div>
            <div>
              <h2>{name}</h2>
              {calculatedPercent !== null && (
                <span className={`score-badge ${getScoreColor(calculatedPercent)}`}>
                  {calculatedPercent}% ({calculatedPercent >= 90 ? 'Excellent' : calculatedPercent >= 80 ? 'Very Good' : calculatedPercent >= 60 ? 'Good' : calculatedPercent >= 50 ? 'Average' : 'Needs Improvement'})
                </span>
              )}
              {comparisonData?.status && (
                <p style={{ marginTop: '0.45rem', color: 'var(--text-secondary)', fontSize: '0.9rem' }}>
                  Status: <strong>{comparisonData.status}</strong>
                </p>
              )}
            </div>
            <div className="total-badge-group">
              {class9Percent !== null && (
                <div className="total-badge">
                  <span className="total-label">Class 9 %</span>
                  <span className="total-value">{class9Percent}</span>
                </div>
              )}
              {targetPercent !== null && (
                <div className="total-badge">
                  <span className="total-label">Class 10 Target %</span>
                  <span className="total-value total-max">{targetPercent}</span>
                </div>
              )}
              {maxMarks > 0 && (
                <div className="total-badge">
                  <span className="total-label">Current Exam Marks</span>
                  <span className="total-value total-max">{obtainedMarks} / {maxMarks}</span>
                </div>
              )}
            </div>
          </div>

          {(class9Percent !== null || targetPercent !== null || calculatedPercent !== null) && (
            <div className="report-section">
              <h4 className="report-section-title">Target Matrix</h4>
              <div className="report-grid">
                {class9Percent !== null && (
                  <div className="report-field highlight-field">
                    <label>Class 9 Marks %</label>
                    <span>{class9Percent}%</span>
                  </div>
                )}
                {targetPercent !== null && (
                  <div className="report-field highlight-field">
                    <label>Class 10 Target %</label>
                    <span>{targetPercent}%</span>
                  </div>
                )}
                {calculatedPercent !== null && (
                  <div className="report-field highlight-field">
                    <label>Current Exam %</label>
                    <span className={getScoreColor(calculatedPercent)}>{calculatedPercent}%</span>
                  </div>
                )}
              </div>
              <p style={{ marginTop: '0.75rem', color: 'var(--text-secondary)', fontSize: '0.84rem' }}>
                Subject marks remain shown in actual marks scored in class exams, usually out of 80, while target is shown separately as the Class 10 target out of 100.
              </p>
            </div>
          )}

          {metaFields.length > 0 && (
            <div className="report-section">
              <h4 className="report-section-title">📋 Student Details</h4>
              <div className="report-grid">
                {metaFields.map(h => {
                  const val = student[h];
                  return (
                    <div key={h} className="report-field">
                      <label>{h}</label>
                      <span>{val !== '' && val !== null && val !== undefined ? String(val) : '—'}</span>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {optedSubjects.length > 0 && (
            <div className="report-section">
              <h4 className="report-section-title">📊 Subject Scores (Opted Subjects Only)</h4>
              <div className="report-grid">
                {optedSubjects.map(h => {
                  const val = student[h];
                  const maxM = getMaxMarksFromHeader(h);
                  return (
                    <div key={h} className="report-field">
                      <label>{h}</label>
                      <span className={typeof val === 'number' ? getScoreColor((val / (maxM || 100)) * 100) : ''}>
                        {String(val)} {maxM ? `/ ${maxM}` : ''}
                      </span>
                    </div>
                  );
                })}
                {calculatedPercent !== null && (
                  <div className="report-field highlight-field">
                    <label>Exam Percentage</label>
                    <span className={getScoreColor(calculatedPercent)}>{calculatedPercent}%</span>
                  </div>
                )}
              </div>
            </div>
          )}

          {examMatrixData.length > 0 && (
            <div className="report-section">
              <h4 className="report-section-title">Exam-wise Comparison</h4>
              <div className="chart-box report-chart">
                <ResponsiveContainer width="100%" height={300}>
                  <LineChart data={examMatrixData} margin={{ top: 10, right: 20, left: 0, bottom: 5 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.05)" />
                    <XAxis dataKey="examName" tick={{ fontSize: 11 }} />
                    <YAxis tick={{ fontSize: 11 }} domain={[0, 100]} />
                    <Tooltip formatter={(value) => [`${value}%`, '']} />
                    <Legend />
                    <Line type="monotone" dataKey="examPercent" name="Exam %" stroke="#2563eb" strokeWidth={3} />
                    {targetPercent !== null && (
                      <Line type="monotone" dataKey="targetPercent" name="Target %" stroke="#dc2626" strokeDasharray="5 5" dot={false} />
                    )}
                    {class9Percent !== null && (
                      <Line type="monotone" dataKey="class9Percent" name="Class 9 %" stroke="#16a34a" strokeDasharray="4 4" connectNulls />
                    )}
                  </LineChart>
                </ResponsiveContainer>
              </div>

              <table className="data-table" style={{ marginTop: '1rem' }}>
                <thead>
                  <tr>
                    <th>Exam</th>
                    <th>Marks</th>
                    <th>Exam %</th>
                    <th>Target %</th>
                    <th>Direction</th>
                  </tr>
                </thead>
                <tbody>
                  {comparisonData.exams.map((exam, index) => {
                    const previous = index > 0 ? comparisonData.exams[index - 1] : null;
                    const baseline = previous?.examPercent ?? class9Percent;
                    const direction = exam.examPercent === null ? 'No Data'
                      : targetPercent !== null && exam.examPercent >= targetPercent ? 'Achieved Target'
                      : baseline !== null && exam.examPercent > baseline ? 'Improving'
                      : 'Below Target';
                    return (
                      <tr key={exam.examName}>
                        <td>{exam.examName}</td>
                        <td>{exam.obtainedMarks} / {exam.maxMarks}</td>
                        <td className={getScoreColor(exam.examPercent)}>{exam.examPercent ?? '—'}{exam.examPercent !== null ? '%' : ''}</td>
                        <td>{exam.targetPercent ?? targetPercent ?? '—'}{(exam.targetPercent ?? targetPercent) !== null ? '%' : ''}</td>
                        <td>{direction}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}

          {notOptedSubjects.length > 0 && (
            <div className="report-section">
              <h4 className="report-section-title" style={{ color: 'var(--text-secondary)', fontSize: '0.85rem' }}>
                🚫 Not Opted Subjects
              </h4>
              <div className="not-opted-list">
                {notOptedSubjects.map(h => (
                  <span key={h} className="not-opted-tag">{h.replace(/\s+\d+$/, '').trim()}</span>
                ))}
              </div>
            </div>
          )}

          {barData.length > 0 && (
            <div className="report-section">
              <h4 className="report-section-title">📈 Subject-wise Score Chart</h4>
              <div className="chart-box report-chart">
                <ResponsiveContainer width="100%" height={280}>
                  <BarChart data={barData} margin={{ top: 10, right: 20, left: 0, bottom: 5 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.05)" />
                    <XAxis dataKey="name" tick={{ fontSize: 10 }} interval={0} angle={-30} textAnchor="end" height={70} />
                    <YAxis tick={{ fontSize: 12 }} />
                    <Tooltip />
                    <Bar dataKey="value" name="Score" radius={[4, 4, 0, 0]}>
                      {barData.map((entry, i) => <Cell key={i} fill={entry.color} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          )}

          {radarData.length >= 3 && (
            <div className="report-section">
              <h4 className="report-section-title">🎯 Performance Radar</h4>
              <div className="chart-box report-chart">
                <ResponsiveContainer width="100%" height={280}>
                  <RadarChart data={radarData} cx="50%" cy="50%" outerRadius="75%">
                    <PolarGrid stroke="rgba(0,0,0,0.08)" />
                    <PolarAngleAxis dataKey="subject" tick={{ fontSize: 10 }} />
                    <PolarRadiusAxis tick={{ fontSize: 10 }} domain={[0, 100]} />
                    <Radar name="Score %" dataKey="percentage" stroke="#6366f1" fill="#6366f1" fillOpacity={0.25} strokeWidth={2} />
                    <Tooltip />
                  </RadarChart>
                </ResponsiveContainer>
              </div>
            </div>
          )}

          <div className="pdf-footer">
            <p>This report was auto-generated by the Target Analysis Report Generator.</p>
            <p>Class 9 marks, Class 10 target, and exam-wise progress are shown separately for clearer comparison.</p>
          </div>
        </div>
      </div>
    </div>
  );
}

function CreateTemplateModal({ onClose, onCreate }) {
  const [className, setClassName] = useState('Class X');
  const [numSections, setNumSections] = useState(5);
  const [subjectInput, setSubjectInput] = useState('');
  const [subjects, setSubjects] = useState(['English', 'Hindi', 'Maths', 'Science', 'Social Science', 'IT / AI']);

  const addSubject = (e) => {
    e.preventDefault();
    if (subjectInput.trim() && !subjects.includes(subjectInput.trim())) {
      setSubjects([...subjects, subjectInput.trim()]);
      setSubjectInput('');
    }
  };

  const removeSubject = (subj) => {
    setSubjects(subjects.filter(s => s !== subj));
  };

  const handleCreate = () => {
    if (!className.trim()) { alert('Class name is required'); return; }
    if (numSections < 1) { alert('At least 1 section is required'); return; }
    if (subjects.length === 0) { alert('Add at least 1 subject'); return; }
    const newHeaders = ['S.No', 'Admn. No.', 'Name', ...subjects, 'Grand Total', '% in IX'];
    onCreate({ className: className.trim(), numSections, subjects: newHeaders.slice(3, -2) });
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-card fade-in" style={{ maxWidth: '500px' }} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h3><Plus size={20} /> Create Blank Class Template</h3>
          <button className="icon-btn" onClick={onClose}><X size={18} /></button>
        </div>

        <div style={{ marginBottom: '1.5rem' }}>
          <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, color: 'var(--text-secondary)', marginBottom: '0.4rem' }}>Class Level / Name</label>
          <input className="cell-input" value={className} onChange={e => setClassName(e.target.value)} placeholder="e.g. Class X" style={{ padding: '0.6rem', fontSize: '0.9rem' }} />
        </div>

        <div style={{ marginBottom: '1.5rem' }}>
          <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, color: 'var(--text-secondary)', marginBottom: '0.4rem' }}>Number of Sections (A, B, C...)</label>
          <input type="number" min="1" max="20" className="cell-input" value={numSections} onChange={e => setNumSections(parseInt(e.target.value) || 1)} style={{ padding: '0.6rem', fontSize: '0.9rem' }} />
        </div>

        <div style={{ marginBottom: '1.5rem' }}>
          <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, color: 'var(--text-secondary)', marginBottom: '0.4rem' }}>Subjects</label>
          <form onSubmit={addSubject} style={{ display: 'flex', gap: '0.5rem', marginBottom: '0.75rem' }}>
            <input className="cell-input" value={subjectInput} onChange={e => setSubjectInput(e.target.value)} placeholder="Type new subject" style={{ padding: '0.6rem', fontSize: '0.9rem' }} />
            <button type="submit" className="tool-btn outline" style={{ minWidth: "60px" }}>Add</button>
          </form>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.4rem' }}>
            {subjects.map(s => (
              <span key={s} style={{ display: 'inline-flex', alignItems: 'center', gap: '0.3rem', padding: '0.3rem 0.6rem', background: '#f1f5f9', borderRadius: '6px', fontSize: '0.8rem', fontWeight: 500, border: '1px solid #e2e8f0' }}>
                {s} <X size={14} style={{ cursor: 'pointer', color: '#64748b' }} onClick={() => removeSubject(s)} />
              </span>
            ))}
            {subjects.length === 0 && <span style={{ fontSize: '0.8rem', color: '#94a3b8' }}>No subjects added</span>}
          </div>
        </div>

        <div style={{ marginTop: '2rem', display: 'flex', gap: '0.75rem', justifyContent: 'flex-end', paddingTop: '1rem', borderTop: '1px solid var(--border)' }}>
          <button className="tool-btn outline" onClick={onClose}>Cancel</button>
          <button className="tool-btn primary" onClick={handleCreate}>Create Template</button>
        </div>
      </div>
    </div>
  );
}

function ExamManagerModal({ onClose, onSave }) {
  const [examName, setExamName] = useState('Unit Test 1');
  const [examDate, setExamDate] = useState(new Date().toISOString().slice(0, 10));
  const [examMaxMarks, setExamMaxMarks] = useState(100);

  const handleSave = () => {
    if (!examName.trim()) { alert('Exam name is required'); return; }
    if (examMaxMarks < 1) { alert('Max marks must be > 0'); return; }
    
    onSave({
      name: examName.trim(),
      date: examDate,
      maxMarks: examMaxMarks
    });
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-card fade-in" style={{ maxWidth: '450px' }} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h3><Database size={20} /> Save Exam to Timeline</h3>
          <button className="icon-btn" onClick={onClose}><X size={18} /></button>
        </div>
        
        <div style={{ marginBottom: '1.25rem' }}>
          <p style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginBottom: '1rem' }}>
            Take a snapshot of the current sheets and save it as an exam milestone for progress tracking.
          </p>
        </div>

        <div style={{ marginBottom: '1.25rem' }}>
          <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, color: 'var(--text-secondary)', marginBottom: '0.4rem' }}>Exam Name</label>
          <input className="cell-input" value={examName} onChange={e => setExamName(e.target.value)} placeholder="e.g. Pre-Board" style={{ padding: '0.6rem', fontSize: '0.9rem' }} />
        </div>

        <div style={{ marginBottom: '1.25rem' }}>
          <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, color: 'var(--text-secondary)', marginBottom: '0.4rem' }}>Date</label>
          <input type="date" className="cell-input" value={examDate} onChange={e => setExamDate(e.target.value)} style={{ padding: '0.6rem', fontSize: '0.9rem' }} />
        </div>

        <div style={{ marginBottom: '1.5rem' }}>
          <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, color: 'var(--text-secondary)', marginBottom: '0.4rem' }}>Global Max Marks (Optional)</label>
          <input type="number" min="1" className="cell-input" value={examMaxMarks} onChange={e => setExamMaxMarks(parseInt(e.target.value) || 100)} placeholder="100" style={{ padding: '0.6rem', fontSize: '0.9rem' }} />
          <span style={{ fontSize: '0.75rem', color: '#94a3b8' }}>Used if column headers don't specify max marks.</span>
        </div>

        <div style={{ marginTop: '2rem', display: 'flex', gap: '0.75rem', justifyContent: 'flex-end', paddingTop: '1rem', borderTop: '1px solid var(--border)' }}>
          <button className="tool-btn outline" onClick={onClose}>Cancel</button>
          <button className="tool-btn success" onClick={handleSave}><Save size={16} /> Save</button>
        </div>
      </div>
    </div>
  );
}

function ProgressTracker({ examTimeline, onBack, currentSheets, sheetNames }) {
  const [selectedStudent, setSelectedStudent] = useState(null);
  const [searchQuery, setSearchQuery] = useState('');
  const effectiveTimeline = useMemo(() => {
    if (examTimeline.length > 0) return examTimeline;
    return buildWorkbookExamTimeline(sheetNames, currentSheets);
  }, [examTimeline, sheetNames, currentSheets]);

  const allStudents = useMemo(() => {
    const map = new Map();
    Object.values(currentSheets).forEach(sheet => {
      sheet.rows.forEach(row => {
        const nameCol = sheet.headers.find(h => h.toLowerCase().includes('name') && !h.toLowerCase().includes('father') && !h.toLowerCase().includes('mother'));
        const name = nameCol ? row[nameCol] : null;
        if (name && !map.has(name)) {
          map.set(name, {
            name,
            latestGrandTotal: row[sheet.headers.find(h => h.toLowerCase().includes('grand total') || h.toLowerCase() === 'total')]
          });
        }
      });
    });
    return Array.from(map.values()).sort((a, b) => a.name.localeCompare(b.name));
  }, [currentSheets]);

  const filteredStudents = useMemo(() => {
    if (!searchQuery.trim()) return allStudents;
    const q = searchQuery.toLowerCase();
    return allStudents.filter(s => s.name.toLowerCase().includes(q));
  }, [allStudents, searchQuery]);

  return (
    <div className="main-card fade-in dashboard-card">
      <div className="dash-header">
        <div className="dash-header-left">
          <button className="icon-btn" onClick={onBack} style={{ marginRight: '10px' }}><ArrowLeft size={24} /></button>
          <Network size={28} className="header-icon" />
          <div>
            <h2>Progress Tracker</h2>
            <p className="header-sub">Track student journey across multiple exams</p>
          </div>
        </div>
      </div>

      {effectiveTimeline.length === 0 ? (
        <div className="empty-msg" style={{ padding: '4rem 1rem' }}>
          <History size={48} strokeWidth={1} style={{ margin: '0 auto 1rem', display: 'block', color: '#cbd5e1' }} />
          <h3>No Exam Sheets Found</h3>
          <p style={{ marginTop: '0.5rem' }}>Upload multiple exam sheets or save exam snapshots from the dashboard to start tracking progress.</p>
          <button className="btn-primary" onClick={onBack} style={{ margin: '1.5rem auto 0', display: 'flex' }}>Go Back Dashboard</button>
        </div>
      ) : (
        <div className="progress-tracker-container fade-in">
          <div className="progress-tracker-sidebar">
            <h4 className="sidebar-title">Students</h4>
            <div className="search-box" style={{ width: '100%', marginBottom: '1rem' }}>
              <Search size={15} className="search-icon" />
              <input className="search-input" style={{ width: '100%' }} placeholder="Search name..." value={searchQuery} onChange={e => setSearchQuery(e.target.value)} />
            </div>
            
            <div style={{ maxHeight: '600px', overflowY: 'auto' }} className="student-list">
              {filteredStudents.length === 0 ? (
                <div style={{ padding: '1rem', textAlign: 'center', color: 'var(--text-secondary)', fontSize: '0.85rem' }}>No students found</div>
              ) : (
                filteredStudents.map(s => (
                  <button 
                    key={s.name} 
                    className={`student-list-item ${selectedStudent?.name === s.name ? 'active' : ''}`}
                    onClick={() => setSelectedStudent(s)}
                  >
                    <div className="avatar">{s.name.charAt(0)}</div>
                    <div className="info">
                      <span className="name">{s.name}</span>
                      <span className="sub">Latest Total: {s.latestGrandTotal ?? 'N/A'}</span>
                    </div>
                  </button>
                ))
              )}
            </div>
          </div>
          
          <div className="progress-tracker-content">
            {selectedStudent ? (
              <StudentProgressView student={selectedStudent} timeline={effectiveTimeline} currentSheets={currentSheets} />
            ) : (
              <div style={{ height: '100%', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', color: 'var(--text-secondary)' }}>
                <TrendingUp size={64} strokeWidth={1} style={{ color: '#cbd5e1', marginBottom: '1rem' }} />
                <h3>Select a student to view progress</h3>
                <p>View their journey across {effectiveTimeline.length} recorded exams.</p>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

// ========== STUDENT PROGRESS VIEW ==========
function StudentProgressView({ student, timeline, currentSheets }) {
  // Extract data for the student across all timeline points
  const progressData = useMemo(() => {
    if (!student || timeline.length === 0) return [];
    
    return timeline.map(exam => {
      // Find the student in this exam's sheets
      let studentRecord = null;
      let headers = [];
      
      for (const sheetName in exam.sheets) {
        const sheet = exam.sheets[sheetName];
        headers = sheet.headers;
        const nameCol = headers.find(h => h.toLowerCase().includes('name') && !h.toLowerCase().includes('father') && !h.toLowerCase().includes('mother'));
        if (!nameCol) continue;
        
        const found = sheet.rows.find(r => r[nameCol] === student.name);
        if (found) {
          studentRecord = found;
          break;
        }
      }
      
      if (!studentRecord) return null; // Student wasn't in this exam
      
      const metrics = extractStudentMetrics(studentRecord, headers, exam.name);
      const subjects = {};
      metrics.subjectBreakdown.forEach((subject) => {
        if (subject.score === null) return;
        subjects[subject.subject] = {
          score: subject.score,
          maxScore: subject.maxScore,
          percent: parseFloat(((subject.score / subject.maxScore) * 100).toFixed(1)),
          target: metrics.targetPercent,
        };
      });
      
      return {
        examName: exam.name,
        date: exam.date,
        overallPercent: metrics.examPercent ?? 0,
        class9Percent: metrics.class9Percent,
        targetPercent: metrics.targetPercent,
        obtained: metrics.obtainedMarks,
        max: metrics.maxMarks,
        subjects
      };
    }).filter(Boolean); // Remove nulls (student absent/not in exam)
  }, [student, timeline]);

  if (progressData.length === 0) {
     return <div className="empty-msg">No progress data found for {student.name}. They might have been absent.</div>
  }
  
  // Extract unique subjects across all participated exams
  const allParticipatedSubjects = new Set();
  progressData.forEach(pd => {
      Object.keys(pd.subjects).forEach(s => allParticipatedSubjects.add(s));
  });
  const subjectList = Array.from(allParticipatedSubjects);
  const overallChartData = progressData.map((entry) => ({
    ...entry,
    baselineClass9: entry.class9Percent,
  }));

  return (
    <div className="fade-in">
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.5rem', paddingBottom: '1rem', borderBottom: '1px solid var(--border)' }}>
        <div>
           <h3 style={{ fontSize: '1.4rem' }}>{student.name}'s Progress</h3>
           <p style={{ color: 'var(--text-secondary)' }}>Comparing {progressData.length} exams</p>
        </div>
      </div>
      
      <div className="chart-grid">
        <div className="chart-inner-box">
          <h5 className="chart-subtitle">Overall Progress vs Target</h5>
          <ResponsiveContainer width="100%" height={250}>
            <BarChart data={overallChartData} margin={{ top: 10, right: 10, left: -20, bottom: 5 }} barGap={6}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.05)" />
              <XAxis dataKey="examName" tick={{ fontSize: 11 }} />
              <YAxis tick={{ fontSize: 11 }} domain={[0, 100]} />
              <Tooltip formatter={(value, name) => [`${value}%`, name]} />
              <Legend />
              <Bar dataKey="overallPercent" name="Actual %" fill="#4f6ef7" radius={[4, 4, 0, 0]} maxBarSize={26} />
              {overallChartData.some(d => d.targetPercent !== null) && (
                <Bar dataKey="targetPercent" name="Target %" fill="#f87171" radius={[4, 4, 0, 0]} maxBarSize={26} />
              )}
              {overallChartData.some(d => d.baselineClass9 !== null) && (
                <Bar dataKey="baselineClass9" name="Class 9 %" fill="#22c55e" radius={[4, 4, 0, 0]} maxBarSize={26} />
              )}
            </BarChart>
          </ResponsiveContainer>
        </div>
        
        <div className="chart-inner-box" style={{ overflowY: 'auto', maxHeight: '315px' }}>
          <h5 className="chart-subtitle">Exams Summary</h5>
          <div className="table-responsive">
            <table className="data-table">
              <thead>
                 <tr>
                   <th>Exam</th>
                   <th>Score</th>
                   <th>%</th>
                   <th>Target</th>
                   <th>Status</th>
                 </tr>
              </thead>
              <tbody>
                {progressData.map((d, i) => (
                  <tr key={i}>
                    <td>{d.examName}</td>
                    <td>{d.obtained} / {d.max}</td>
                    <td className={getScoreColor(d.overallPercent)}>{d.overallPercent}%</td>
                    <td>{d.targetPercent ?? '—'}{d.targetPercent !== null ? '%' : ''}</td>
                    <td>
                      {d.targetPercent !== null && d.overallPercent >= d.targetPercent
                        ? 'Achieved Target'
                        : d.class9Percent !== null && d.overallPercent > d.class9Percent
                          ? 'Improving'
                          : 'Below Target'}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
      
      <h4 style={{ marginBottom: '1rem', paddingBottom: '0.5rem', borderBottom: '2px solid var(--primary-light)' }}>Subject-wise Progress</h4>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: '1.5rem' }}>
         {subjectList.map((subject, idx) => {
            // Prepare standard chart data for this subject
            const subjData = progressData.map(pd => {
                const sData = pd.subjects[subject];
                return {
                    exam: pd.examName,
                    percent: sData ? sData.percent : null,
                    target: sData && sData.target ? sData.target : null
                };
            });
            
            return (
               <div key={subject} className="chart-inner-box" style={{ padding: '1rem' }}>
                  <h6 style={{ fontSize: '0.9rem', marginBottom: '0.5rem', color: 'var(--text)' }}>{subject}</h6>
                  <ResponsiveContainer width="100%" height={150}>
                    <BarChart data={subjData} margin={{ top: 5, right: 5, left: -25, bottom: 0 }} barGap={6}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="rgba(0,0,0,0.05)" />
                      <XAxis dataKey="exam" tick={{ fontSize: 9 }} />
                      <YAxis tick={{ fontSize: 9 }} domain={[0, 100]} />
                      <Tooltip formatter={(value, name) => [`${value}%`, name]} labelStyle={{fontSize: 11}} itemStyle={{fontSize: 11}} />
                      <Bar dataKey="percent" name="Score" fill={CHART_COLORS[idx % CHART_COLORS.length]} radius={[4, 4, 0, 0]} maxBarSize={22} />
                      {subjData.some(d => d.target !== null) && (
                         <Bar dataKey="target" name="Target" fill="#fca5a5" radius={[4, 4, 0, 0]} maxBarSize={22} />
                      )}
                    </BarChart>
                  </ResponsiveContainer>
               </div>
            )
         })}
      </div>
    </div>
  );
}
