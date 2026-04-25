const test = require('node:test');
const assert = require('node:assert/strict');
const XLSX = require('xlsx');

const {
  detectExamStage,
  detectHeaderRowIndex,
  extractStudentMetrics,
  findNameColumn,
  isLikelyStudentDataRow,
  parseClassSection,
  parseWorkbookBuffer,
  resolveRowSection,
} = require('../src/services/parser.service');

test('detectExamStage recognizes teacher file names that use underscores and section codes', () => {
  assert.equal(detectExamStage('HY_10A_X-A', []), 'HY');
  assert.equal(detectExamStage('PB1_10B_X-B', []), 'PB1');
  assert.equal(detectExamStage('PB2_10C_X-C', []), 'PB2');
  assert.equal(detectExamStage('BOARD_CLASS10_RESULT', []), 'BOARD');
  assert.equal(detectExamStage('CBSE RESULT 2026', ['Roll No', 'Percentage']), 'BOARD');
});

test('detectHeaderRowIndex finds the real header row after a title row', () => {
  const matrix = [
    ['CLASS 10 A RESULT'],
    ['S.No', 'Enroll No', 'Student Name', 'English 80', 'Grand Total', '%'],
    [1, 101, 'Aarav', 71, 71, 88.75],
  ];

  assert.equal(detectHeaderRowIndex(matrix), 1);
});

test('parseClassSection recognizes compact and separated class-section variants', () => {
  assert.deepEqual(parseClassSection('XA'), { className: 'X', sectionName: 'A' });
  assert.deepEqual(parseClassSection('XB'), { className: 'X', sectionName: 'B' });
  assert.deepEqual(parseClassSection('10A'), { className: '10', sectionName: 'A' });
  assert.deepEqual(parseClassSection('X-A'), { className: 'X', sectionName: 'A' });
  assert.deepEqual(parseClassSection('Class X A'), { className: 'X', sectionName: 'A' });
});

test('extractStudentMetrics derives exam percentage from obtained and max marks when percent column is missing', () => {
  const headers = ['Enroll No', 'Student Name', 'English 80', 'Maths 80', 'Grand Total'];
  const row = {
    'Enroll No': '101',
    'Student Name': 'Aarav',
    'English 80': 70,
    'Maths 80': 75,
    'Grand Total': 145,
  };

  const metrics = extractStudentMetrics(row, headers, 'Half Yearly 10A');

  assert.equal(metrics.examPercent, 90.63);
  assert.equal(metrics.obtainedMarks, 145);
  assert.equal(metrics.maxMarks, 160);
});

test('parseWorkbookBuffer reports duplicate enrollment values and keeps workbook metadata', () => {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Board Result Combined'],
    ['Enroll No', 'Student Name', 'English 80', 'Maths 80', 'Grand Total'],
    [101, 'Aarav', 70, 75, 145],
    [101, 'Ira', 60, 65, 125],
  ]);

  XLSX.utils.book_append_sheet(workbook, sheet, 'Board Result');

  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  const parsed = parseWorkbookBuffer(buffer, 'Board Result.xlsx');

  assert.deepEqual(parsed.sheetNames, ['Board Result']);
  assert.equal(parsed.sheets['Board Result'].meta.headerRowIndex, 1);
  assert.equal(parsed.sheets['Board Result'].meta.examName, 'Board Result');
  assert.match(parsed.sheets['Board Result'].validation.issues.join(' | '), /Duplicate Enroll No\./);
});

test('parseWorkbookBuffer extracts XA baseline sheet metadata and preserves capped target metrics', () => {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    ['S.No', 'Enroll No', 'Student Name', '% in IX', '% in IX+30'],
    [1, '101', 'Aarav', 82, 120],
  ]);

  XLSX.utils.book_append_sheet(workbook, sheet, 'XA');
  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  const parsed = parseWorkbookBuffer(buffer, 'Target Sheet 2024.xlsx');
  const parsedSheet = parsed.sheets.XA;
  const metrics = extractStudentMetrics(parsedSheet.rows[0], parsedSheet.headers, parsedSheet.meta.examName);

  assert.equal(parsedSheet.meta.sectionName, 'A');
  assert.equal(parsedSheet.validation.examStage, 'BASELINE');
  assert.equal(metrics.class9Percent, 82);
  assert.equal(metrics.targetPercent, 100);
});

test('parseWorkbookBuffer extracts student rows from rankwise sheets and skips aggregate rows', () => {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    ['HY 10D X-D RANKWISE'],
    ['Rank', 'Enroll No', 'Student Name', 'English 80', 'Maths 80', 'Grand Total'],
    [1, 401, 'Riya', 71, 73, 144],
    [2, 402, 'Kabir', 68, 70, 138],
    ['', '', 'Average', '', '', 141],
  ]);

  XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1');
  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  const parsed = parseWorkbookBuffer(buffer, 'HY_10D_X-D-RANKWISE.xls');

  assert.equal(parsed.sheets.Sheet1.validation.examStage, 'HY');
  assert.equal(parsed.sheets.Sheet1.rows.length, 2);
  assert.deepEqual(parsed.sheets.Sheet1.rows.map((row) => row['Enroll No']), [401, 402]);
});

test('isLikelyStudentDataRow requires real student identity or measurable data', () => {
  const headers = ['Rank', 'Enroll No', 'Student Name', 'Grand Total'];

  assert.equal(isLikelyStudentDataRow({
    Rank: 1,
    'Enroll No': 501,
    'Student Name': 'Aarav',
    'Grand Total': 470,
  }, headers), true);

  assert.equal(isLikelyStudentDataRow({
    Rank: '',
    'Enroll No': '',
    'Student Name': 'Average',
    'Grand Total': 430,
  }, headers), true);
});

test('parseWorkbookBuffer accepts sheets that have enrollment and scores even when student name column is missing', () => {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    ['BOARD RESULT'],
    ['Rank', 'Enroll No', 'Best Five %'],
    [1, 901, 91.2],
    [2, 902, 88.6],
  ]);

  XLSX.utils.book_append_sheet(workbook, sheet, 'BEST FIVE');
  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  const parsed = parseWorkbookBuffer(buffer, 'BOARD_CLASS10_CBSE RESULT 2026.xlsx');

  assert.equal(parsed.sheets['BEST FIVE'].validation.ok, true);
  assert.equal(parsed.sheets['BEST FIVE'].rows.length, 2);
  assert.equal(parsed.sheets['BEST FIVE'].validation.issues.length, 0);
  assert.equal(parsed.sheets['BEST FIVE'].rows[0]['Enroll No'], 901);
});

test('findNameColumn detects board-result name columns when the header cell is blank between roll and gender', () => {
  const headers = ['Roll No', 'Column2', 'Gender', 'English', 'Percentage'];
  assert.equal(findNameColumn(headers), 'Column2');
});

test('parseWorkbookBuffer uses the title row to detect HY and PB stages for generic filenames', () => {
  const workbook = XLSX.utils.book_new();
  const hySheet = XLSX.utils.aoa_to_sheet([
    ['Cambridge Court High School Consolidated Class X- A Exam : HALF YEARLY 2025-26 RANKWISE'],
    ['S. No.', 'Enroll No.', 'Student Name', 'English (80)', '%'],
    [1, '2017-2018/0051', 'AARADHYA AGRAWAL', 58, 72.5],
  ]);
  const pbSheet = XLSX.utils.aoa_to_sheet([
    ['Cambridge Court High School Consolidated Class X- A Exam : PRE BOARD-II 2025-26'],
    ['S. No.', 'Enroll No.', 'Student Name', 'English (80)', '%'],
    [1, '2017-2018/0051', 'AARADHYA AGRAWAL', 61, 76.25],
  ]);

  XLSX.utils.book_append_sheet(workbook, hySheet, 'Sheet1');
  XLSX.utils.book_append_sheet(workbook, pbSheet, 'Sheet2');

  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  const parsed = parseWorkbookBuffer(buffer, 'X-A.xlsx');

  assert.equal(parsed.sheets.Sheet1.validation.examStage, 'HY');
  assert.equal(parsed.sheets.Sheet2.validation.examStage, 'PB2');
});

test('parseWorkbookBuffer accepts merged exam sheets with row-level Section values', () => {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    ['HALF YEARLY CLASS 10'],
    ['Section', 'Enroll No.', 'Student Name', 'English (80)', '%'],
    ['A', '2017-2018/0051', 'AARADHYA AGRAWAL', 58, 72.5],
    ['B', '2017-2018/0062', 'DIYA SHARMA', 62, 77.5],
  ]);

  XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1');

  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  const parsed = parseWorkbookBuffer(buffer, 'HY_CLASS10_MERGED.xlsx');
  const parsedSheet = parsed.sheets.Sheet1;

  assert.equal(parsedSheet.validation.ok, true);
  assert.equal(parsedSheet.meta.sectionName, null);
  assert.equal(resolveRowSection(parsedSheet.rows[0], parsedSheet.headers, parsedSheet.meta, { sheetName: 'Sheet1', sourceFileName: parsedSheet.meta.examName }), 'A');
  assert.equal(resolveRowSection(parsedSheet.rows[1], parsedSheet.headers, parsedSheet.meta, { sheetName: 'Sheet1', sourceFileName: parsedSheet.meta.examName }), 'B');
});

test('parseWorkbookBuffer reports validation issues for merged exam sheets without row-level section data', () => {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    ['HALF YEARLY CLASS 10'],
    ['Enroll No.', 'Student Name', 'English (80)', '%'],
    ['2017-2018/0051', 'AARADHYA AGRAWAL', 58, 72.5],
  ]);

  XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1');

  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  const parsed = parseWorkbookBuffer(buffer, 'HY_CLASS10_MERGED.xlsx');

  assert.match(parsed.sheets.Sheet1.validation.issues.join(' | '), /Missing Section or Class Section values/);
});
