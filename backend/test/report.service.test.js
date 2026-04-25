const test = require('node:test');
const assert = require('node:assert/strict');
const XLSX = require('xlsx');

const { structuredImport } = require('../src/services/report.service');

function buildWorkbookBuffer(rows = [], sheetName = 'Sheet1') {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
}

test('structuredImport supports merged HY/PB1/PB2 files and returns baseline mismatch reporting', async () => {
  const files = [
    {
      originalname: 'BASELINE_CLASS10.xlsx',
      buffer: buildWorkbookBuffer([
        ['Section', 'Enroll No', 'Student Name', '% in IX', '% in IX+30'],
        ['A', '101', 'Aarav', 80, 90],
        ['B', '201', 'Diya', 78, 88],
        ['C', '999', 'Unmatched Student', 70, 80],
      ]),
    },
    {
      originalname: 'HY_CLASS10_MERGED.xlsx',
      buffer: buildWorkbookBuffer([
        ['Section', 'Enroll No', 'Student Name', '%'],
        ['A', '101', 'Aarav', 72.5],
        ['B', '201', 'Diya', 77.5],
      ]),
    },
    {
      originalname: 'PB1_CLASS10_MERGED.xlsx',
      buffer: buildWorkbookBuffer([
        ['Section', 'Enroll No', 'Student Name', '%'],
        ['A', '101', 'Aarav', 75.5],
        ['B', '201', 'Diya', 80.5],
      ]),
    },
    {
      originalname: 'PB2_CLASS10_MERGED.xlsx',
      buffer: buildWorkbookBuffer([
        ['Section', 'Enroll No', 'Student Name', '%'],
        ['A', '101', 'Aarav', 79.5],
        ['B', '201', 'Diya', 84.5],
      ]),
    },
    {
      originalname: 'BOARD_CLASS10.xlsx',
      buffer: buildWorkbookBuffer([
        ['Section', 'Enroll No', 'Student Name', 'Percentage'],
        ['A', '101', 'Aarav', 82.5],
        ['B', '201', 'Diya', 86.5],
      ]),
    },
  ];

  const result = await structuredImport(files);
  const rowsByEnrollment = Object.fromEntries(
    (result.masterCumulativeSheet?.rows || []).map((row) => [row['Enrollment No'], row]),
  );

  assert.equal(result.validationPassed, true);
  assert.equal(result.masterCumulativeSheet.rows.length, 2);
  assert.equal(rowsByEnrollment['101'].Section, 'A');
  assert.equal(rowsByEnrollment['101']['Class 9 %'], 80);
  assert.equal(rowsByEnrollment['101']['Target %'], 90);
  assert.equal(rowsByEnrollment['101']['HY %'], 72.5);
  assert.equal(rowsByEnrollment['101']['PB1 %'], 75.5);
  assert.equal(rowsByEnrollment['101']['PB2 %'], 79.5);
  assert.equal(rowsByEnrollment['101']['Board %'], 82.5);
  assert.equal(rowsByEnrollment['201'].Section, 'B');
  assert.equal(result.baselineMatchReport.matchedCount, 2);
  assert.equal(result.baselineMatchReport.unmatchedCount, 1);
  assert.equal(result.baselineMatchReport.rows.length, 3);
  assert.equal(
    result.baselineMatchReport.rows.find((row) => row.baselineEnrollmentNo === '999')?.confidence,
    'unmatched',
  );
});

test('structuredImport surfaces validation issues when merged exam files do not provide row-level sections', async () => {
  const files = [
    {
      originalname: 'BASELINE_CLASS10.xlsx',
      buffer: buildWorkbookBuffer([
        ['Section', 'Enroll No', 'Student Name', '% in IX', '% in IX+30'],
        ['A', '101', 'Aarav', 80, 90],
      ]),
    },
    {
      originalname: 'HY_CLASS10_MERGED.xlsx',
      buffer: buildWorkbookBuffer([
        ['Enroll No', 'Student Name', '%'],
        ['101', 'Aarav', 72.5],
      ]),
    },
    {
      originalname: 'PB1_CLASS10_MERGED.xlsx',
      buffer: buildWorkbookBuffer([
        ['Section', 'Enroll No', 'Student Name', '%'],
        ['A', '101', 'Aarav', 75.5],
      ]),
    },
    {
      originalname: 'PB2_CLASS10_MERGED.xlsx',
      buffer: buildWorkbookBuffer([
        ['Section', 'Enroll No', 'Student Name', '%'],
        ['A', '101', 'Aarav', 79.5],
      ]),
    },
    {
      originalname: 'BOARD_CLASS10.xlsx',
      buffer: buildWorkbookBuffer([
        ['Section', 'Enroll No', 'Student Name', 'Percentage'],
        ['A', '101', 'Aarav', 82.5],
      ]),
    },
  ];

  const result = await structuredImport(files);

  assert.equal(result.validationPassed, false);
  assert.match(result.issues.map((issue) => issue.message).join(' | '), /Missing Section or Class Section values/);
});

test('structuredImport rejects uploads that do not contain exactly one file for each required stage', async () => {
  await assert.rejects(
    structuredImport([
      {
        originalname: 'BASELINE_CLASS10.xlsx',
        buffer: buildWorkbookBuffer([
          ['Section', 'Enroll No', 'Student Name', '% in IX', '% in IX+30'],
          ['A', '101', 'Aarav', 80, 90],
        ]),
      },
    ]),
    /requires exactly 5 files/i,
  );
});
