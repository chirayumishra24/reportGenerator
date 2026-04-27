const { parseWorkbookBuffer, detectExamStage } = require('./parser.service');
const persistenceService = require('./persistence.service');
const { buildWorkbookCumulativeResult } = require('./cumulative.service');

function badRequest(message) {
  const error = new Error(message);
  error.status = 400;
  return error;
}

function buildValidationIssues(parsedSheets = []) {
  return parsedSheets.flatMap((sheet) =>
    (sheet.validation?.issues || []).map((issue) => ({
      sheetName: sheet.sheetName,
      message: issue,
    })),
  );
}

function buildUniqueSheetName(preferredName, usedNames) {
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
}

function detectStructuredFileStage(fileName, parsed = {}) {
  // 1. Prioritize filename signal for structured imports
  const fileNameStage = detectExamStage(fileName, []);
  if (fileNameStage !== 'UNKNOWN') return fileNameStage;

  // 2. Fallback to sheet content signal
  const parsedStages = Array.from(new Set(
    (parsed.parsedSheets || [])
      .map((sheet) => {
        // Use validation stage if already set, or detect from sheet name and headers
        // Pass empty array for headers if we only want to detect from sheet name
        return sheet.validation?.examStage || detectExamStage(sheet.sheetName, sheet.headers || []);
      })
      .filter((stage) => stage && stage !== 'UNKNOWN'),
  ));

  if (parsedStages.length === 1) return parsedStages[0];
  return null;
}

async function parseWorkbook(file) {
  if (!file) {
    throw badRequest('No file uploaded');
  }

  const parsed = parseWorkbookBuffer(file.buffer, file.originalname);
  const issues = buildValidationIssues(parsed.parsedSheets);

  return {
    sheetNames: parsed.sheetNames,
    sheets: parsed.sheets,
    parsedSheets: parsed.parsedSheets,
    validationPassed: issues.length === 0,
    issues,
  };
}

async function structuredImport(files) {
  if (!files || files.length === 0) {
    throw badRequest('No files uploaded');
  }
  if (files.length !== 5) {
    throw badRequest('Structured import requires exactly 5 files: Baseline, HY, PB1, PB2, and Board.');
  }

  const usedNames = new Set();
  const mergedSheetNames = [];
  const mergedSheets = {};
  const issues = [];
  const fileStages = new Map();
  const requiredStages = ['BASELINE', 'HY', 'PB1', 'PB2', 'BOARD'];

  for (const file of files) {
    const parsed = parseWorkbookBuffer(file.buffer, file.originalname);
    const fileStage = detectStructuredFileStage(file.originalname, parsed);
    if (!fileStage || !requiredStages.includes(fileStage)) {
      throw badRequest(`Could not classify ${file.originalname} as one of Baseline, HY, PB1, PB2, or Board.`);
    }
    if (fileStages.has(fileStage)) {
      throw badRequest(`Duplicate ${fileStage} file uploaded: ${file.originalname}. Please upload exactly one file for each stage.`);
    }
    fileStages.set(fileStage, file.originalname);

    const fileIssues = buildValidationIssues(parsed.parsedSheets).map((issue) => ({
      fileName: file.originalname,
      ...issue,
    }));
    issues.push(...fileIssues);

    for (const sheetName of parsed.sheetNames || []) {
      const baseName = String(file.originalname || 'Upload').replace(/\.[^/.]+$/, '');
      const preferredName = files.length === 1 ? sheetName : `${baseName} - ${sheetName}`;
      const uniqueSheetName = buildUniqueSheetName(preferredName, usedNames);
      mergedSheetNames.push(uniqueSheetName);
      mergedSheets[uniqueSheetName] = parsed.sheets[sheetName];
    }
  }

  const missingStages = requiredStages.filter((stage) => !fileStages.has(stage));
  if (missingStages.length > 0) {
    throw badRequest(`Missing required structured file(s): ${missingStages.join(', ')}.`);
  }

  const cumulativeResult = buildWorkbookCumulativeResult(mergedSheetNames, mergedSheets);

  return {
    sheetNames: mergedSheetNames,
    sheets: mergedSheets,
    validationPassed: issues.length === 0,
    issues,
    masterCumulativeSheet: cumulativeResult.masterCumulativeSheet,
    baselineMatchReport: cumulativeResult.baselineMatchReport,
  };
}

async function importPersistent(files) {
  if (!files || files.length === 0) {
    throw badRequest('No files uploaded');
  }

  const imported = await persistenceService.importParsedFiles(files);
  const cumulativeReport = await persistenceService.getCumulativeReport();

  return {
    message: 'Files imported into cumulative database successfully.',
    imported,
    cumulativeReport,
  };
}

async function getCumulativeReport() {
  return persistenceService.getCumulativeReport();
}

async function getDbStatus() {
  return persistenceService.getDbStatus();
}

async function getStudentHistory(studentId) {
  if (!studentId) {
    throw badRequest('Student id is required');
  }
  return persistenceService.getStudentHistory(studentId);
}

module.exports = {
  parseWorkbook,
  structuredImport,
  importPersistent,
  getCumulativeReport,
  getDbStatus,
  getStudentHistory,
};
