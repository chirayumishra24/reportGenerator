const {
  detectExamStage,
  extractStudentMetrics,
  cleanStudentName,
  findEnrollmentColumn,
  findNameColumn,
  findRollColumn,
  normalizeIdentifier,
  normalizeStudentName,
} = require('./parser.service');

function buildEmptyCumulativeReport() {
  return {
    databaseEnabled: false,
    summary: { uploads: 0, sheets: 0, students: 0, performances: 0 },
    studentComparison: [],
    classComparison: [],
    sectionComparison: [],
    examTimeline: [],
  };
}

function average(values) {
  const valid = values.filter((value) => Number.isFinite(value));
  if (!valid.length) return null;
  return parseFloat((valid.reduce((sum, value) => sum + value, 0) / valid.length).toFixed(2));
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

function getLatestKnownScore(entry) {
  const stages = ['Board %', 'PB2 %', 'PB1 %', 'HY %'];
  for (const stage of stages) {
    const value = parseFloat(entry[stage]);
    if (Number.isFinite(value)) return value;
  }
  return null;
}

function detectSubjectSignature(row = {}, headers = []) {
  const normalizedHeaders = headers.map((header) => String(header || '').toLowerCase());
  const getValue = (pattern) => {
    const index = normalizedHeaders.findIndex((header) => pattern.test(header));
    if (index < 0) return '';
    return String(row[headers[index]] ?? '').trim();
  };

  const secondLanguage = getValue(/\bhindi\b/)
    ? 'hindi'
    : getValue(/\bsanskrit\b/)
      ? 'sanskrit'
      : getValue(/\bfrench\b/)
        ? 'french'
        : '';

  const optionalSubject = getValue(/\bai\b/)
    ? 'ai'
    : getValue(/\bit\b/)
      ? 'it'
      : '';

  const mathType = getValue(/\bbasic maths\b/)
    ? 'basic-maths'
    : getValue(/\bmathematics\b|\bmaths\b/)
      ? 'maths'
      : '';

  return {
    secondLanguage,
    optionalSubject,
    mathType,
  };
}

function buildStudentEntry(identity) {
  return {
    'Enrollment No': identity.enrollmentNo,
    'Student Name': identity.name || '',
    Section: identity.section,
    'Class 9 %': '',
    'Target %': '',
    'HY %': '',
    'PB1 %': '',
    'PB2 %': '',
    'Board %': '',
    _subjectSignature: identity.subjectSignature || { secondLanguage: '', optionalSubject: '', mathType: '' },
  };
}

function updatePreferredIdentityFields(entry, identity) {
  if (!entry['Enrollment No'] && identity.enrollmentNo) entry['Enrollment No'] = identity.enrollmentNo;
  entry['Student Name'] = choosePreferredText(entry['Student Name'], identity.name);
  entry.Section = choosePreferredText(entry.Section, identity.section);
  if (!entry._subjectSignature?.secondLanguage && identity.subjectSignature?.secondLanguage) {
    entry._subjectSignature.secondLanguage = identity.subjectSignature.secondLanguage;
  }
  if (!entry._subjectSignature?.optionalSubject && identity.subjectSignature?.optionalSubject) {
    entry._subjectSignature.optionalSubject = identity.subjectSignature.optionalSubject;
  }
  if (!entry._subjectSignature?.mathType && identity.subjectSignature?.mathType) {
    entry._subjectSignature.mathType = identity.subjectSignature.mathType;
  }
}

function applyStageMetrics(entry, examStage, metrics) {
  // Always extract baseline metrics if they are available in the current sheet
  if (metrics.class9Percent !== null && !Number.isFinite(parseFloat(entry['Class 9 %']))) {
    entry['Class 9 %'] = metrics.class9Percent;
  }
  if (metrics.targetPercent !== null && !Number.isFinite(parseFloat(entry['Target %']))) {
    entry['Target %'] = metrics.targetPercent;
  }

  // Then handle stage-specific marks
  if (examStage === 'BASELINE' || metrics.examPercent === null) return;

  const existingStageScore = parseFloat(getStageScore(entry, examStage));
  if (!Number.isFinite(existingStageScore)) {
    setStageScore(entry, examStage, metrics.examPercent);
  }
}

function getBoardSheetPriority(sheetName = '') {
  const normalized = String(sheetName || '').trim().toLowerCase();
  if (normalized === 'rank' || normalized.includes('rank')) return 1;
  if (normalized.includes('all subjects report')) return 2;
  if (normalized.includes('best five')) return 3;
  return 4;
}

function getProcessingSheetNames(sheetNames = [], sheets = {}) {
  const preferredBoardSheets = new Map();
  const selected = [];

  sheetNames.forEach((sheetName) => {
    const sheet = sheets[sheetName];
    if (!sheet?.rows?.length) return;
    const headers = sheet.headers || [];
    const examStage = detectExamStage(sheet.meta?.examName || sheetName, headers);

    if (examStage !== 'BOARD') {
      selected.push(sheetName);
      return;
    }

    const boardGroupKey = String(sheet.meta?.examName || sheetName).trim().toLowerCase();
    const candidate = {
      sheetName,
      priority: getBoardSheetPriority(sheetName),
    };
    const current = preferredBoardSheets.get(boardGroupKey);
    if (!current || candidate.priority < current.priority) {
      preferredBoardSheets.set(boardGroupKey, candidate);
    }
  });

  preferredBoardSheets.forEach(({ sheetName }) => selected.push(sheetName));
  return selected;
}

function permute(items = []) {
  if (items.length <= 1) return [items.slice()];
  const output = [];
  items.forEach((item, index) => {
    const rest = items.slice(0, index).concat(items.slice(index + 1));
    permute(rest).forEach((tail) => output.push([item, ...tail]));
  });
  return output;
}

function buildBoardMatchCost(candidate, boardRow) {
  let cost = 0;
  const candidateSignature = candidate._subjectSignature || {};
  const boardSignature = boardRow.subjectSignature || {};

  if (
    candidateSignature.secondLanguage
    && boardSignature.secondLanguage
    && candidateSignature.secondLanguage !== boardSignature.secondLanguage
  ) {
    cost += 5000;
  }

  if (
    candidateSignature.optionalSubject
    && boardSignature.optionalSubject
    && candidateSignature.optionalSubject !== boardSignature.optionalSubject
  ) {
    cost += 5000;
  }

  if (
    candidateSignature.mathType
    && boardSignature.mathType
    && candidateSignature.mathType !== 'maths'
    && boardSignature.mathType !== candidateSignature.mathType
  ) {
    cost += 2500;
  }

  const latestScore = getLatestKnownScore(candidate);
  if (Number.isFinite(latestScore) && Number.isFinite(boardRow.examPercent)) {
    const diff = boardRow.examPercent - latestScore;
    cost += diff * diff;
  }

  return cost;
}

function assignBoardRowsToCandidates(boardRows = [], candidateEntries = []) {
  if (!boardRows.length || !candidateEntries.length) return [];

  const boardPermutations = permute(boardRows);
  let bestAssignment = [];
  let bestCost = Number.POSITIVE_INFINITY;

  boardPermutations.forEach((permutation) => {
    const cost = permutation.reduce(
      (sum, row, index) => sum + buildBoardMatchCost(candidateEntries[index], row),
      0,
    );

    if (cost < bestCost) {
      bestCost = cost;
      bestAssignment = permutation.map((row, index) => ({
        studentKey: candidateEntries[index].studentKey,
        row,
      }));
    }
  });

  return bestAssignment;
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

    student.performances.forEach((performance) => {
      const stage = detectExamStage(performance.examSheet.examName || performance.examSheet.name, []);
      if (stage === 'BASELINE') {
        if (performance.class9Percent !== null) row.class9Percent = performance.class9Percent;
        if (performance.targetPercent !== null) row.targetPercent = performance.targetPercent;
      } else if (stage === 'HY' && performance.examPercent !== null) {
        row.hyPercent = performance.examPercent;
      } else if (stage === 'PB1' && performance.examPercent !== null) {
        row.pb1Percent = performance.examPercent;
      } else if (stage === 'PB2' && performance.examPercent !== null) {
        row.pb2Percent = performance.examPercent;
      } else if (stage === 'BOARD' && performance.examPercent !== null) {
        row.boardPercent = performance.examPercent;
      }
    });

    const latest = [row.boardPercent, row.pb2Percent, row.pb1Percent, row.hyPercent].find((value) => Number.isFinite(value)) ?? null;
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
    const enrollA = String(a.enrollmentNo || '');
    const enrollB = String(b.enrollmentNo || '');
    return enrollA.localeCompare(enrollB, undefined, { numeric: true, sensitivity: 'base' });
  });
}

function buildWorkbookCumulativeSheet(sheetNames = [], sheets = {}) {
  const students = new Map();
  const studentsByName = new Map();
  const selectedSheetNames = getProcessingSheetNames(sheetNames, sheets);

  function computeStudentIdentity(row, sheet = {}, sheetName = '', examStage = 'UNKNOWN') {
    const headers = sheet.headers || [];
    const enrollmentCol = findEnrollmentColumn(headers);
    const rollCol = findRollColumn(headers);
    const nameCol = findNameColumn(headers);
    const enrollmentNo = enrollmentCol ? String(row[enrollmentCol] ?? '').trim() : '';
    const rollNo = rollCol ? String(row[rollCol] ?? '').trim() : '';
    const normalizedEnrollmentNo = normalizeIdentifier(enrollmentNo);
    const normalizedRollNo = normalizeIdentifier(rollNo);
    const name = nameCol ? cleanStudentName(row[nameCol]) : '';
    const normalizedName = normalizeStudentName(name);
    const section = String(sheet.meta?.sectionName || '').trim();

    return {
      enrollmentNo,
      rollNo,
      name,
      normalizedName,
      section,
      examStage,
      enrollmentKey: normalizedEnrollmentNo ? `enrollment:${normalizedEnrollmentNo}` : '',
      rollKey: normalizedRollNo ? `roll:${normalizedRollNo}` : '',
      subjectSignature: detectSubjectSignature(row, headers),
      sheetName,
    };
  }

  function getOrCreateStudent(studentKey, identity) {
    const existing = students.get(studentKey) || buildStudentEntry(identity);
    updatePreferredIdentityFields(existing, identity);
    students.set(studentKey, existing);
    return existing;
  }

  function registerStudentName(studentKey, identity) {
    if (!identity.normalizedName) return;
    if (!studentsByName.has(identity.normalizedName)) {
      studentsByName.set(identity.normalizedName, new Set());
    }
    studentsByName.get(identity.normalizedName).add(studentKey);
  }

  function resolveUniqueStudentKeyByName(normalizedName) {
    if (!normalizedName || !studentsByName.has(normalizedName)) return null;
    const matches = studentsByName.get(normalizedName);
    if (!matches || matches.size !== 1) return null;
    return Array.from(matches)[0];
  }

  const anchorSheetNames = selectedSheetNames.filter((sheetName) => {
    const sheet = sheets[sheetName];
    if (!sheet?.rows?.length) return false;
    const headers = sheet.headers || [];
    const examStage = detectExamStage(sheet.meta?.examName || sheetName, headers);
    return examStage !== 'BASELINE' && examStage !== 'BOARD' && Boolean(findEnrollmentColumn(headers));
  });

  const fallbackAnchorSheetNames = anchorSheetNames.length
    ? anchorSheetNames
    : selectedSheetNames.filter((sheetName) => {
      const sheet = sheets[sheetName];
      if (!sheet?.rows?.length) return false;
      const headers = sheet.headers || [];
      const examStage = detectExamStage(sheet.meta?.examName || sheetName, headers);
      return examStage !== 'BOARD' && Boolean(findEnrollmentColumn(headers));
    });

  fallbackAnchorSheetNames.forEach((sheetName) => {
    const sheet = sheets[sheetName];
    if (!sheet?.rows?.length) return;

    const headers = sheet.headers || [];
    const examStage = detectExamStage(sheet.meta?.examName || sheetName, headers);
    sheet.rows.forEach((row) => {
      const identity = computeStudentIdentity(row, sheet, sheetName, examStage);
      if (!identity.enrollmentKey) return;
      getOrCreateStudent(identity.enrollmentKey, identity);
      registerStudentName(identity.enrollmentKey, identity);
    });
  });

  const allowFallbackRoster = students.size === 0;
  if (allowFallbackRoster) {
    selectedSheetNames.forEach((sheetName) => {
      const sheet = sheets[sheetName];
      if (!sheet?.rows?.length) return;

      const headers = sheet.headers || [];
      const examStage = detectExamStage(sheet.meta?.examName || sheetName, headers);
      sheet.rows.forEach((row) => {
        const identity = computeStudentIdentity(row, sheet, sheetName, examStage);
        const studentKey = identity.enrollmentKey || identity.rollKey;
        if (!studentKey) return;
        getOrCreateStudent(studentKey, identity);
        registerStudentName(studentKey, identity);
      });
    });
  }

  function getSheetMatchCoverage(sheetName) {
    const sheet = sheets[sheetName];
    if (!sheet?.rows?.length) return 0;

    const headers = sheet.headers || [];
    const examStage = detectExamStage(sheet.meta?.examName || sheetName, headers);
    let matchedRows = 0;
    let namedRows = 0;

    sheet.rows.forEach((row) => {
      const identity = computeStudentIdentity(row, sheet, sheetName, examStage);
      if (!identity.normalizedName) return;
      namedRows += 1;
      if (identity.enrollmentKey && students.has(identity.enrollmentKey)) {
        matchedRows += 1;
        return;
      }
      if (resolveUniqueStudentKeyByName(identity.normalizedName)) {
        matchedRows += 1;
      }
    });

    return namedRows > 0 ? matchedRows / namedRows : 0;
  }

  selectedSheetNames.forEach((sheetName) => {
    const sheet = sheets[sheetName];
    if (!sheet?.rows?.length) return;

    const headers = sheet.headers || [];
    const examStage = detectExamStage(sheet.meta?.examName || sheetName, headers);
    if (examStage === 'BOARD') return;
    if (examStage === 'BASELINE' && !allowFallbackRoster && getSheetMatchCoverage(sheetName) < 0.5) return;

    sheet.rows.forEach((row) => {
      const identity = computeStudentIdentity(row, sheet, sheetName, examStage);
      const metrics = extractStudentMetrics(row, headers, sheet.meta?.examName || sheetName);
      const studentKey = identity.enrollmentKey && students.has(identity.enrollmentKey)
        ? identity.enrollmentKey
        : resolveUniqueStudentKeyByName(identity.normalizedName);
      if (!studentKey) return;

      const existing = getOrCreateStudent(studentKey, identity);
      registerStudentName(studentKey, identity);
      applyStageMetrics(existing, examStage, metrics);
      students.set(studentKey, existing);
    });
  });

  selectedSheetNames.forEach((sheetName) => {
    const sheet = sheets[sheetName];
    if (!sheet?.rows?.length) return;

    const headers = sheet.headers || [];
    const examStage = detectExamStage(sheet.meta?.examName || sheetName, headers);
    if (examStage !== 'BOARD') return;

    const boardGroups = new Map();
    sheet.rows.forEach((row) => {
      const identity = computeStudentIdentity(row, sheet, sheetName, examStage);
      const metrics = extractStudentMetrics(row, headers, sheet.meta?.examName || sheetName);
      if (identity.enrollmentKey && students.has(identity.enrollmentKey)) {
        const existing = students.get(identity.enrollmentKey);
        updatePreferredIdentityFields(existing, identity);
        applyStageMetrics(existing, 'BOARD', metrics);
        students.set(identity.enrollmentKey, existing);
        return;
      }
      if (!identity.normalizedName || metrics.examPercent === null) return;
      if (!boardGroups.has(identity.normalizedName)) {
        boardGroups.set(identity.normalizedName, []);
      }
      boardGroups.get(identity.normalizedName).push({
        identity,
        metrics,
        subjectSignature: identity.subjectSignature,
      });
    });

    boardGroups.forEach((boardRows, normalizedName) => {
      const candidateKeys = Array.from(studentsByName.get(normalizedName) || []);

      if (candidateKeys.length === 1 && boardRows.length >= 1) {
        const studentKey = candidateKeys[0];
        const existing = students.get(studentKey);
        if (!existing) return;

        boardRows
          .slice()
          .sort((a, b) => (b.metrics.examPercent ?? -1) - (a.metrics.examPercent ?? -1))
          .forEach((boardRow) => {
            if (!Number.isFinite(parseFloat(existing['Board %']))) {
              updatePreferredIdentityFields(existing, boardRow.identity);
              applyStageMetrics(existing, 'BOARD', boardRow.metrics);
            }
          });
        students.set(studentKey, existing);
        return;
      }

      if (candidateKeys.length === 0) {
        if (!allowFallbackRoster) return;
        boardRows.forEach((boardRow) => {
          const fallbackKey = boardRow.identity.rollKey || boardRow.identity.enrollmentKey;
          if (!fallbackKey) return;
          const existing = getOrCreateStudent(fallbackKey, boardRow.identity);
          registerStudentName(fallbackKey, boardRow.identity);
          applyStageMetrics(existing, 'BOARD', boardRow.metrics);
        });
        return;
      }

      const candidateEntries = candidateKeys
        .map((studentKey) => ({ studentKey, ...students.get(studentKey) }))
        .filter((entry) => entry.studentKey);

      const sortedCandidates = candidateEntries
        .slice()
        .sort((a, b) => (getLatestKnownScore(b) ?? -1) - (getLatestKnownScore(a) ?? -1));
      const assignments = assignBoardRowsToCandidates(boardRows, sortedCandidates);
      assignments.forEach(({ studentKey, row: boardRow }) => {
        const existing = students.get(studentKey);
        if (!existing) return;
        updatePreferredIdentityFields(existing, boardRow.identity);
        applyStageMetrics(existing, 'BOARD', boardRow.metrics);
        students.set(studentKey, existing);
      });
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
    const board = parseFloat(entry['Board %']);
    const pb2 = parseFloat(entry['PB2 %']);
    const pb1 = parseFloat(entry['PB1 %']);
    const hy = parseFloat(entry['HY %']);
    const latest = Number.isFinite(board) ? board : Number.isFinite(pb2) ? pb2 : Number.isFinite(pb1) ? pb1 : Number.isFinite(hy) ? hy : null;
    const target = parseFloat(entry['Target %']);
    const class9 = parseFloat(entry['Class 9 %']);
    const targetGap = latest !== null && Number.isFinite(target) ? parseFloat((latest - target).toFixed(2)) : '';
    const improvement = latest !== null && Number.isFinite(class9) ? parseFloat((latest - class9).toFixed(2)) : '';
    let status = 'Needs Review';
    if (latest !== null && Number.isFinite(target) && latest >= target) status = 'Achieved Target';
    else if (latest !== null && Number.isFinite(target) && Number.isFinite(class9) && latest > class9) status = 'Improving Toward Target';
    else if (latest !== null && Number.isFinite(class9) && latest > class9) status = 'Improved';
    else if (latest !== null && Number.isFinite(target)) status = 'Below Target';

    return {
      'Enrollment No': entry['Enrollment No'],
      'Student Name': entry['Student Name'],
      Section: entry.Section,
      'Class 9 %': entry['Class 9 %'],
      'Target %': entry['Target %'],
      'HY %': entry['HY %'],
      'PB1 %': entry['PB1 %'],
      'PB2 %': entry['PB2 %'],
      'Board %': entry['Board %'],
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

function buildCumulativeReportPayload({ uploads = [], sheets = [], students = [] }) {
  const studentComparison = students.map((student) => {
    const performances = student.performances
      .filter((performance) => performance.examPercent !== null || performance.class9Percent !== null || performance.targetPercent !== null)
      .map((performance) => ({
        examName: performance.examSheet.examName || performance.examSheet.name,
        sheetName: performance.examSheet.name,
        examPercent: performance.examPercent,
        class9Percent: performance.class9Percent,
        targetPercent: performance.targetPercent,
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
    const scores = sheet.performances.map((performance) => performance.examPercent).filter((value) => Number.isFinite(value));
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
    scores.forEach((score) => classEntry.scores.push(score));
    sheet.performances.forEach((performance) => classEntry.students.add(performance.studentId));
    if (sheet.sectionName) classEntry.sections.add(sheet.sectionName);

    const sectionKey = `${classKey}__${sheet.sectionName || 'N/A'}`;
    if (!sectionMap.has(sectionKey)) {
      sectionMap.set(sectionKey, {
        className: classKey,
        sectionName: sheet.sectionName || 'N/A',
        scores: [],
        students: new Set(),
        exams: [],
      });
    }
    const sectionEntry = sectionMap.get(sectionKey);
    scores.forEach((score) => sectionEntry.scores.push(score));
    sheet.performances.forEach((performance) => sectionEntry.students.add(performance.studentId));
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
        Section: row.section,
        'Class 9 %': row.class9Percent,
        'Target %': row.targetPercent,
        'HY %': row.hyPercent,
        'PB1 %': row.pb1Percent,
        'PB2 %': row.pb2Percent,
        'Board %': row.boardPercent,
        'Target Gap': row.targetGap,
        Improvement: row.improvement,
        Status: row.status,
      })),
    },
  };
}

module.exports = {
  buildEmptyCumulativeReport,
  average,
  buildPerformanceStatus,
  buildMasterCumulativeRows,
  buildWorkbookCumulativeSheet,
  buildCumulativeReportPayload,
};
