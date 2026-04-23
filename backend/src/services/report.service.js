const { parseWorkbookBuffer } = require('./parser.service');
const persistenceService = require('./persistence.service');
const { buildWorkbookCumulativeSheet } = require('./cumulative.service');

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

  const usedNames = new Set();
  const mergedSheetNames = [];
  const mergedSheets = {};
  const issues = [];

  for (const file of files) {
    const parsed = parseWorkbookBuffer(file.buffer, file.originalname);
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

  return {
    sheetNames: mergedSheetNames,
    sheets: mergedSheets,
    validationPassed: issues.length === 0,
    issues,
    masterCumulativeSheet: buildWorkbookCumulativeSheet(mergedSheetNames, mergedSheets),
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
