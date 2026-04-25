const test = require('node:test');
const assert = require('node:assert/strict');
const XLSX = require('xlsx');

const {
  buildMasterCumulativeRows,
  buildWorkbookCumulativeResult,
  buildWorkbookCumulativeSheet,
  buildCumulativeReportPayload,
} = require('../src/services/cumulative.service');
const { parseWorkbookBuffer } = require('../src/services/parser.service');

test('buildMasterCumulativeRows calculates gaps, improvement, and target status', () => {
  const students = [{
    id: 'student-1',
    admissionNo: '101',
    name: 'Aarav',
    sectionName: 'A',
    performances: [
      {
        class9Percent: 80,
        targetPercent: 90,
        examPercent: null,
        examSheet: { examName: 'Target Sheet 2024', name: 'Target Sheet' },
      },
      {
        class9Percent: null,
        targetPercent: null,
        examPercent: 92,
        examSheet: { examName: 'Board Result', name: 'Board Result' },
      },
    ],
  }];

  const rows = buildMasterCumulativeRows(students);

  assert.equal(rows.length, 1);
  assert.equal(rows[0].status, 'Achieved Target');
  assert.equal(rows[0].targetGap, 2);
  assert.equal(rows[0].improvement, 12);
  assert.equal(rows[0].boardPercent, 92);
});

test('buildCumulativeReportPayload builds summary and master sheet from queried records', () => {
  const payload = buildCumulativeReportPayload({
    uploads: [{ id: 'upload-1' }],
    sheets: [{
      id: 'sheet-1',
      name: 'Board Result',
      examName: 'Board Result',
      className: '10',
      sectionName: 'A',
      createdAt: new Date('2026-04-22T10:00:00Z'),
      performances: [
        { studentId: 'student-1', examPercent: 92 },
      ],
    }],
    students: [{
      id: 'student-1',
      name: 'Aarav',
      admissionNo: '101',
      className: '10',
      sectionName: 'A',
      performances: [
        {
          class9Percent: 80,
          targetPercent: 90,
          examPercent: null,
          examSheet: { examName: 'Target Sheet 2024', name: 'Target Sheet' },
        },
        {
          class9Percent: null,
          targetPercent: null,
          examPercent: 92,
          examSheet: { examName: 'Board Result', name: 'Board Result' },
        },
      ],
    }],
  });

  assert.equal(payload.databaseEnabled, true);
  assert.equal(payload.summary.uploads, 1);
  assert.equal(payload.summary.students, 1);
  assert.equal(payload.masterCumulativeSheet.rows.length, 1);
  assert.equal(payload.masterCumulativeSheet.rows[0].Status, 'Achieved Target');
});

test('buildWorkbookCumulativeSheet builds an in-memory cumulative sheet without database records', () => {
  const sheetNames = ['Baseline', 'Board'];
  const sheets = {
    Baseline: {
      headers: ['Enroll No', 'Student Name', '% in IX', '% in IX+30'],
      rows: [{
        'Enroll No': '101',
        'Student Name': 'Aarav',
        '% in IX': 80,
        '% in IX+30': 90,
      }],
      meta: { examName: 'Target Sheet 2024', sectionName: 'A' },
    },
    Board: {
      headers: ['Enroll No', 'Student Name', 'English 80', 'Maths 80', 'Grand Total'],
      rows: [{
        'Enroll No': '101',
        'Student Name': 'Aarav Kumar',
        'English 80': 70,
        'Maths 80': 75,
        'Grand Total': 145,
      }],
      meta: { examName: 'Board Result', sectionName: 'A' },
    },
  };

  const cumulative = buildWorkbookCumulativeSheet(sheetNames, sheets);

  assert.equal(cumulative.rows.length, 1);
  assert.equal(cumulative.rows[0]['Student Name'], 'Aarav Kumar');
  assert.equal(cumulative.rows[0].Status, 'Achieved Target');
  assert.equal(cumulative.rows[0]['Board %'], 90.63);
});

test('buildWorkbookCumulativeSheet treats enrollment number as the primary merge key when names differ', () => {
  const cumulative = buildWorkbookCumulativeSheet(
    ['HY', 'PB1'],
    {
      HY: {
        headers: ['Enroll No', 'Student Name', 'English 80', 'Maths 80', 'Grand Total'],
        rows: [{
          'Enroll No': '777',
          'Student Name': 'Anaya',
          'English 80': 64,
          'Maths 80': 70,
          'Grand Total': 134,
        }],
        meta: { examName: 'HY_10A_X-A', sectionName: 'A' },
      },
      PB1: {
        headers: ['Enroll No', 'Student Name', 'English 80', 'Maths 80', 'Grand Total'],
        rows: [{
          'Enroll No': '777',
          'Student Name': 'Anaya Sharma',
          'English 80': 70,
          'Maths 80': 74,
          'Grand Total': 144,
        }],
        meta: { examName: 'PB1_10A_X-A', sectionName: 'A' },
      },
    },
  );

  assert.equal(cumulative.rows.length, 1);
  assert.equal(cumulative.rows[0]['Enrollment No'], '777');
  assert.equal(cumulative.rows[0]['Student Name'], 'Anaya Sharma');
  assert.equal(cumulative.rows[0]['HY %'], 83.75);
  assert.equal(cumulative.rows[0]['PB1 %'], 90);
});

test('buildWorkbookCumulativeSheet keeps enrollment-only rows even when student name is missing', () => {
  const cumulative = buildWorkbookCumulativeSheet(
    ['BEST FIVE'],
    {
      'BEST FIVE': {
        headers: ['Rank', 'Enroll No', 'Best Five %'],
        rows: [{
          Rank: 1,
          'Enroll No': '999',
          'Best Five %': 93.4,
        }],
        meta: { examName: 'BOARD_CLASS10_CBSE RESULT 2026', sectionName: 'A' },
      },
    },
  );

  assert.equal(cumulative.rows.length, 1);
  assert.equal(cumulative.rows[0]['Enrollment No'], '999');
  assert.equal(cumulative.rows[0]['Student Name'], '');
});

test('buildWorkbookCumulativeSheet matches blank-enrollment rows back to a uniquely known student by normalized name', () => {
  const cumulative = buildWorkbookCumulativeSheet(
    ['Baseline', 'PB1'],
    {
      Baseline: {
        headers: ['Enroll No', 'Student Name', '% in IX', '% in IX+30'],
        rows: [{
          'Enroll No': '2023-2024/0199',
          'Student Name': 'Varu.N. Moudgill',
          '% in IX': 82,
          '% in IX+30': 92,
        }],
        meta: { examName: 'Target Sheet 2024', sectionName: 'E' },
      },
      PB1: {
        headers: ['Enroll No', 'Student Name', 'English 80', 'Maths 80', 'Grand Total'],
        rows: [{
          'Enroll No': '',
          'Student Name': 'Varun Moudgill',
          'English 80': 68,
          'Maths 80': 70,
          'Grand Total': 138,
        }],
        meta: { examName: 'PB1_10E_X-E', sectionName: 'E' },
      },
    },
  );

  assert.equal(cumulative.rows.length, 1);
  assert.equal(cumulative.rows[0]['Enrollment No'], '2023-2024/0199');
  assert.equal(cumulative.rows[0]['PB1 %'], 86.25);
});

test('buildWorkbookCumulativeSheet keeps compact baseline sheet sections and capped target values', () => {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    ['S.No', 'Enroll No', 'Student Name', '% in IX', '% in IX+30'],
    [1, '101', 'Aarav', 82, 120],
  ]);

  XLSX.utils.book_append_sheet(workbook, sheet, 'XA');
  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  const parsed = parseWorkbookBuffer(buffer, 'Target Sheet 2024.xlsx');
  const cumulative = buildWorkbookCumulativeSheet(parsed.sheetNames, parsed.sheets);

  assert.equal(cumulative.rows.length, 1);
  assert.equal(cumulative.rows[0].Section, 'A');
  assert.equal(cumulative.rows[0]['Class 9 %'], 82);
  assert.equal(cumulative.rows[0]['Target %'], 100);
});

test('buildWorkbookCumulativeSheet skips unmatched name-only rows so they do not inflate the student count', () => {
  const cumulative = buildWorkbookCumulativeSheet(
    ['Baseline', 'PB1'],
    {
      Baseline: {
        headers: ['Enroll No', 'Student Name', '% in IX', '% in IX+30'],
        rows: [{
          'Enroll No': '101',
          'Student Name': 'Aarav',
          '% in IX': 80,
          '% in IX+30': 90,
        }],
        meta: { examName: 'Target Sheet 2024', sectionName: 'A' },
      },
      PB1: {
        headers: ['Enroll No', 'Student Name', 'English 80', 'Maths 80', 'Grand Total'],
        rows: [{
          'Enroll No': '',
          'Student Name': 'Unmatched Student',
          'English 80': 70,
          'Maths 80': 72,
          'Grand Total': 142,
        }],
        meta: { examName: 'PB1_10A_X-A', sectionName: 'A' },
      },
    },
  );

  assert.equal(cumulative.rows.length, 1);
  assert.equal(cumulative.rows[0]['Enrollment No'], '101');
  assert.equal(cumulative.rows[0]['PB1 %'], '');
});

test('buildWorkbookCumulativeSheet backfills matched baseline values without creating rows for unrelated baseline students', () => {
  const cumulative = buildWorkbookCumulativeSheet(
    ['HY', 'Baseline'],
    {
      HY: {
        headers: ['Enroll No', 'Student Name', 'English 80', 'Maths 80', 'Grand Total'],
        rows: [
          {
            'Enroll No': '101',
            'Student Name': 'Aarav',
            'English 80': 64,
            'Maths 80': 70,
            'Grand Total': 134,
          },
          {
            'Enroll No': '102',
            'Student Name': 'Diya',
            'English 80': 60,
            'Maths 80': 66,
            'Grand Total': 126,
          },
        ],
        meta: { examName: 'HY_10A_X-A', sectionName: 'A' },
      },
      Baseline: {
        headers: ['Enroll No', 'Student Name', '% in IX', '% in IX+30'],
        rows: [
          {
            'Enroll No': '101',
            'Student Name': 'Aarav Kumar',
            '% in IX': 81,
            '% in IX+30': 91,
          },
          {
            'Enroll No': '999',
            'Student Name': 'Unrelated Student',
            '% in IX': 70,
            '% in IX+30': 80,
          },
        ],
        meta: { examName: 'Target Sheet 2024', sectionName: 'A' },
      },
    },
  );

  assert.equal(cumulative.rows.length, 2);
  const byEnrollment = Object.fromEntries(cumulative.rows.map((row) => [row['Enrollment No'], row]));
  assert.equal(byEnrollment['101']['Class 9 %'], 81);
  assert.equal(byEnrollment['101']['Target %'], 91);
  assert.equal(byEnrollment['102']['Class 9 %'], '');
  assert.equal(byEnrollment['102']['Target %'], '');
  assert.equal(byEnrollment['999'], undefined);
});

test('buildWorkbookCumulativeResult reports matched and unmatched baseline rows for review', () => {
  const cumulative = buildWorkbookCumulativeResult(
    ['HY', 'Baseline'],
    {
      HY: {
        headers: ['Section', 'Enroll No', 'Student Name', 'English 80', 'Grand Total'],
        rows: [
          {
            Section: 'A',
            'Enroll No': '101',
            'Student Name': 'Aarav',
            'English 80': 64,
            'Grand Total': 64,
          },
          {
            Section: 'B',
            'Enroll No': '201',
            'Student Name': 'Diya',
            'English 80': 66,
            'Grand Total': 66,
          },
        ],
        meta: { examName: 'HY_CLASS10_MERGED', sectionName: '' },
      },
      Baseline: {
        headers: ['Section', 'Enroll No', 'Student Name', '% in IX', '% in IX+30'],
        rows: [
          {
            Section: 'A',
            'Enroll No': '101',
            'Student Name': 'Aarav Kumar',
            '% in IX': 81,
            '% in IX+30': 91,
          },
          {
            Section: 'B',
            'Enroll No': '999',
            'Student Name': 'Unknown Student',
            '% in IX': 70,
            '% in IX+30': 80,
          },
        ],
        meta: { examName: 'BASELINE_CLASS10_MERGED', sectionName: '' },
      },
    },
  );

  assert.equal(cumulative.masterCumulativeSheet.rows.length, 2);
  assert.equal(cumulative.baselineMatchReport.matchedCount, 1);
  assert.equal(cumulative.baselineMatchReport.unmatchedCount, 1);
  assert.equal(cumulative.baselineMatchReport.rows.length, 2);
  assert.equal(cumulative.baselineMatchReport.rows[0].confidence, 'exact');
  assert.equal(cumulative.baselineMatchReport.rows[1].confidence, 'unmatched');
});

test('buildWorkbookCumulativeResult fills exact name matches even when enrollment numbers do not overlap', () => {
  const cumulative = buildWorkbookCumulativeResult(
    ['HY', 'Baseline'],
    {
      HY: {
        headers: ['Section', 'Enroll No', 'Student Name', 'English 80', 'Grand Total'],
        rows: [{
          Section: 'A',
          'Enroll No': '2016-2017/0005',
          'Student Name': 'Aarav Jain',
          'English 80': 64,
          'Grand Total': 64,
        }],
        meta: { examName: 'HY_CLASS10_MERGED', sectionName: '' },
      },
      Baseline: {
        headers: ['Section', 'Enroll No', 'Student Name', '% in IX', '% in IX+30'],
        rows: [{
          Section: 'A',
          'Enroll No': '2015-2016/0170',
          'Student Name': 'Aarav Jain',
          '% in IX': 63.5,
          '% in IX+30': 93.5,
        }],
        meta: { examName: 'BASELINE_CLASS10_MERGED', sectionName: '' },
      },
    },
  );

  assert.equal(cumulative.masterCumulativeSheet.rows.length, 1);
  assert.equal(cumulative.masterCumulativeSheet.rows[0]['Enrollment No'], '2016-2017/0005');
  assert.equal(cumulative.masterCumulativeSheet.rows[0]['Class 9 %'], 63.5);
  assert.equal(cumulative.masterCumulativeSheet.rows[0]['Target %'], 93.5);
  assert.equal(cumulative.baselineMatchReport.matchedCount, 1);
  assert.equal(cumulative.baselineMatchReport.unmatchedCount, 0);
  assert.equal(cumulative.baselineMatchReport.rows[0].confidence, 'exact');
});

test('buildWorkbookCumulativeResult reports fuzzy baseline suggestions without applying them', () => {
  const cumulative = buildWorkbookCumulativeResult(
    ['HY', 'Baseline'],
    {
      HY: {
        headers: ['Section', 'Enroll No', 'Student Name', 'English 80', 'Grand Total'],
        rows: [{
          Section: 'A',
          'Enroll No': '2016-2017/0080',
          'Student Name': 'Aarav Gupta',
          'English 80': 64,
          'Grand Total': 64,
        }],
        meta: { examName: 'HY_CLASS10_MERGED', sectionName: '' },
      },
      Baseline: {
        headers: ['Section', 'Enroll No', 'Student Name', '% in IX', '% in IX+30'],
        rows: [{
          Section: 'A',
          'Enroll No': '2023-2024/0145',
          'Student Name': 'Arnav Gupta',
          '% in IX': 49,
          '% in IX+30': 79,
        }],
        meta: { examName: 'BASELINE_CLASS10_MERGED', sectionName: '' },
      },
    },
  );

  assert.equal(cumulative.masterCumulativeSheet.rows.length, 1);
  assert.equal(cumulative.masterCumulativeSheet.rows[0]['Class 9 %'], '');
  assert.equal(cumulative.masterCumulativeSheet.rows[0]['Target %'], '');
  assert.equal(cumulative.baselineMatchReport.matchedCount, 0);
  assert.equal(cumulative.baselineMatchReport.unmatchedCount, 1);
  assert.equal(cumulative.baselineMatchReport.rows[0].confidence, 'fuzzy');
  assert.equal(cumulative.baselineMatchReport.rows[0].reason, 'Same-section fuzzy suggestion');
  assert.equal(cumulative.baselineMatchReport.rows[0].baselineClass9Percent, 49);
  assert.equal(cumulative.baselineMatchReport.rows[0].baselineTargetPercent, 79);
  assert.equal(cumulative.baselineMatchReport.rows[0].suggestedStudentName, 'Aarav Gupta');
  assert.equal(cumulative.baselineMatchReport.rows[0].suggestedEnrollmentNo, '2016-2017/0080');
  assert.equal(cumulative.baselineMatchReport.rows[0].suggestionScore > 0.8, true);
});

test('buildWorkbookCumulativeResult rejects ambiguous fuzzy baseline suggestions', () => {
  const cumulative = buildWorkbookCumulativeResult(
    ['HY', 'Baseline'],
    {
      HY: {
        headers: ['Section', 'Enroll No', 'Student Name', 'English 80', 'Grand Total'],
        rows: [
          {
            Section: 'A',
            'Enroll No': '101',
            'Student Name': 'Aaryan Jain',
            'English 80': 64,
            'Grand Total': 64,
          },
          {
            Section: 'A',
            'Enroll No': '102',
            'Student Name': 'Arya Jain',
            'English 80': 66,
            'Grand Total': 66,
          },
        ],
        meta: { examName: 'HY_CLASS10_MERGED', sectionName: '' },
      },
      Baseline: {
        headers: ['Section', 'Enroll No', 'Student Name', '% in IX', '% in IX+30'],
        rows: [{
          Section: 'A',
          'Enroll No': '999',
          'Student Name': 'Aryan Jain',
          '% in IX': 68,
          '% in IX+30': 98,
        }],
        meta: { examName: 'BASELINE_CLASS10_MERGED', sectionName: '' },
      },
    },
  );

  const byEnrollment = Object.fromEntries(cumulative.masterCumulativeSheet.rows.map((row) => [row['Enrollment No'], row]));
  assert.equal(byEnrollment['101']['Class 9 %'], '');
  assert.equal(byEnrollment['102']['Class 9 %'], '');
  assert.equal(cumulative.baselineMatchReport.matchedCount, 0);
  assert.equal(cumulative.baselineMatchReport.unmatchedCount, 1);
  assert.equal(cumulative.baselineMatchReport.rows[0].confidence, 'unmatched');
  assert.equal(cumulative.baselineMatchReport.rows[0].reason, 'Ambiguous fuzzy candidates');
});

test('buildWorkbookCumulativeSheet still backfills matched baseline rows when overall baseline overlap is below fifty percent', () => {
  const cumulative = buildWorkbookCumulativeSheet(
    ['HY', 'Baseline'],
    {
      HY: {
        headers: ['Enroll No', 'Student Name', 'English 80', 'Maths 80', 'Grand Total'],
        rows: [
          {
            'Enroll No': '101',
            'Student Name': 'Aarav',
            'English 80': 64,
            'Maths 80': 70,
            'Grand Total': 134,
          },
          {
            'Enroll No': '102',
            'Student Name': 'Diya',
            'English 80': 60,
            'Maths 80': 66,
            'Grand Total': 126,
          },
        ],
        meta: { examName: 'HY_10A_X-A', sectionName: 'A' },
      },
      Baseline: {
        headers: ['Enroll No', 'Student Name', '% in IX', '% in IX+30'],
        rows: [
          {
            'Enroll No': '101',
            'Student Name': 'Aarav Kumar',
            '% in IX': 81,
            '% in IX+30': 91,
          },
          {
            'Enroll No': '901',
            'Student Name': 'Unrelated Student 1',
            '% in IX': 70,
            '% in IX+30': 80,
          },
          {
            'Enroll No': '902',
            'Student Name': 'Unrelated Student 2',
            '% in IX': 71,
            '% in IX+30': 81,
          },
          {
            'Enroll No': '903',
            'Student Name': 'Unrelated Student 3',
            '% in IX': 72,
            '% in IX+30': 82,
          },
        ],
        meta: { examName: 'Target Sheet 2024', sectionName: 'A' },
      },
    },
  );

  assert.equal(cumulative.rows.length, 2);
  const byEnrollment = Object.fromEntries(cumulative.rows.map((row) => [row['Enrollment No'], row]));
  assert.equal(byEnrollment['101']['Class 9 %'], 81);
  assert.equal(byEnrollment['101']['Target %'], 91);
  assert.equal(byEnrollment['102']['Class 9 %'], '');
  assert.equal(byEnrollment['102']['Target %'], '');
});

test('buildWorkbookCumulativeSheet uses same-section compact-name fallback for punctuation-split baseline names', () => {
  const cumulative = buildWorkbookCumulativeSheet(
    ['HY', 'Baseline'],
    {
      HY: {
        headers: ['Enroll No', 'Student Name', 'English 80', 'Maths 80', 'Grand Total'],
        rows: [{
          'Enroll No': '201',
          'Student Name': 'Varun Moudgill',
          'English 80': 66,
          'Maths 80': 70,
          'Grand Total': 136,
        }],
        meta: { examName: 'HY_10E_X-E', sectionName: 'E' },
      },
      Baseline: {
        headers: ['Enroll No', 'Student Name', '% in IX', '% in IX+30'],
        rows: [{
          'Enroll No': '999',
          'Student Name': 'Varu N Moudgill',
          '% in IX': 84,
          '% in IX+30': 94,
        }],
        meta: { examName: 'Target Sheet 2024', sectionName: 'E' },
      },
    },
  );

  assert.equal(cumulative.rows.length, 1);
  assert.equal(cumulative.rows[0]['Enrollment No'], '201');
  assert.equal(cumulative.rows[0]['Class 9 %'], 84);
  assert.equal(cumulative.rows[0]['Target %'], 94);
});

test('buildWorkbookCumulativeSheet does not match near names that differ beyond spacing and punctuation', () => {
  const cumulative = buildWorkbookCumulativeSheet(
    ['HY', 'Baseline'],
    {
      HY: {
        headers: ['Enroll No', 'Student Name', 'English 80', 'Maths 80', 'Grand Total'],
        rows: [{
          'Enroll No': '201',
          'Student Name': 'Aarav Gupta',
          'English 80': 66,
          'Maths 80': 70,
          'Grand Total': 136,
        }],
        meta: { examName: 'HY_10A_X-A', sectionName: 'A' },
      },
      Baseline: {
        headers: ['Enroll No', 'Student Name', '% in IX', '% in IX+30'],
        rows: [{
          'Enroll No': '999',
          'Student Name': 'Arnav Gupta',
          '% in IX': 84,
          '% in IX+30': 94,
        }],
        meta: { examName: 'Target Sheet 2024', sectionName: 'A' },
      },
    },
  );

  assert.equal(cumulative.rows.length, 1);
  assert.equal(cumulative.rows[0]['Enrollment No'], '201');
  assert.equal(cumulative.rows[0]['Class 9 %'], '');
  assert.equal(cumulative.rows[0]['Target %'], '');
});

test('buildWorkbookCumulativeSheet prefers exact enrollment matches over exact-name or fuzzy alternatives', () => {
  const cumulative = buildWorkbookCumulativeSheet(
    ['HY', 'Baseline'],
    {
      HY: {
        headers: ['Enroll No', 'Student Name', 'English 80', 'Maths 80', 'Grand Total'],
        rows: [
          {
            'Enroll No': '101',
            'Student Name': 'Aryan Bhandari',
            'English 80': 66,
            'Maths 80': 70,
            'Grand Total': 136,
          },
          {
            'Enroll No': '102',
            'Student Name': 'Aaryan Bhandari',
            'English 80': 68,
            'Maths 80': 72,
            'Grand Total': 140,
          },
        ],
        meta: { examName: 'HY_10C_X-C', sectionName: 'C' },
      },
      Baseline: {
        headers: ['Enroll No', 'Student Name', '% in IX', '% in IX+30'],
        rows: [{
          'Enroll No': '101',
          'Student Name': 'Aaryan Bhandari',
          '% in IX': 84,
          '% in IX+30': 94,
        }],
        meta: { examName: 'Target Sheet 2024', sectionName: 'C' },
      },
    },
  );

  const byEnrollment = Object.fromEntries(cumulative.rows.map((row) => [row['Enrollment No'], row]));
  assert.equal(byEnrollment['101']['Class 9 %'], 84);
  assert.equal(byEnrollment['101']['Target %'], 94);
  assert.equal(byEnrollment['102']['Class 9 %'], '');
  assert.equal(byEnrollment['102']['Target %'], '');
});

test('buildWorkbookCumulativeSheet does not use compact-name fallback for surname-only matches', () => {
  const cumulative = buildWorkbookCumulativeSheet(
    ['HY', 'Baseline'],
    {
      HY: {
        headers: ['Enroll No', 'Student Name', 'English 80', 'Maths 80', 'Grand Total'],
        rows: [{
          'Enroll No': '301',
          'Student Name': 'Varun Sharma',
          'English 80': 61,
          'Maths 80': 65,
          'Grand Total': 126,
        }],
        meta: { examName: 'HY_10A_X-A', sectionName: 'A' },
      },
      Baseline: {
        headers: ['Enroll No', 'Student Name', '% in IX', '% in IX+30'],
        rows: [{
          'Enroll No': '998',
          'Student Name': 'Gunjan Sharma',
          '% in IX': 73,
          '% in IX+30': 83,
        }],
        meta: { examName: 'Target Sheet 2024', sectionName: 'A' },
      },
    },
  );

  assert.equal(cumulative.rows.length, 1);
  assert.equal(cumulative.rows[0]['Enrollment No'], '301');
  assert.equal(cumulative.rows[0]['Class 9 %'], '');
  assert.equal(cumulative.rows[0]['Target %'], '');
});

test('buildWorkbookCumulativeSheet maps board roll-number rows onto the enrollment roster without creating duplicate students', () => {
  const cumulative = buildWorkbookCumulativeSheet(
    ['HY', 'PB2', 'BOARD RANK', 'BOARD ALL SUBJECTS'],
    {
      HY: {
        headers: ['Enroll No.', 'Student Name', 'English (80)', '%'],
        rows: [
          { 'Enroll No.': '101', 'Student Name': 'AANYA KEDIA', 'English (80)': 58, '%': 72.5 },
          { 'Enroll No.': '102', 'Student Name': 'AARADHYA AGRAWAL', 'English (80)': 62, '%': 77.5 },
        ],
        meta: { examName: 'HY_10A_X-A', sectionName: 'A' },
      },
      PB2: {
        headers: ['Enroll No.', 'Student Name', 'English (80)', '%'],
        rows: [
          { 'Enroll No.': '101', 'Student Name': 'AANYA KEDIA', 'English (80)': 64, '%': 79.5 },
          { 'Enroll No.': '102', 'Student Name': 'AARADHYA AGRAWAL', 'English (80)': 68, '%': 84.5 },
        ],
        meta: { examName: 'PB2_10A_X-A', sectionName: 'A' },
      },
      'BOARD RANK': {
        headers: ['S.NO.', 'Roll No', 'Column3', 'Gender', 'English', 'Percentage'],
        rows: [
          { 'S.NO.': 1, 'Roll No': '9001', Column3: 'AANYA KEDIA', Gender: 'F', English: 91, Percentage: 90 },
          { 'S.NO.': 2, 'Roll No': '9002', Column3: 'AARADHYA AGRAWAL', Gender: 'F', English: 95, Percentage: 94 },
        ],
        meta: { examName: 'BOARD_CLASS10_CBSE RESULT 2026', sectionName: '' },
      },
      'BOARD ALL SUBJECTS': {
        headers: ['Roll No', 'Name', 'Gender', 'English', 'Percentage'],
        rows: [
          { 'Roll No': '9001', Name: 'AANYA KEDIA 184 122 241 086 087 417 PASS', Gender: 'F', English: 91, Percentage: 90 },
          { 'Roll No': '9002', Name: 'AARADHYA AGRAWAL 184 018 041 086 087 417 PASS', Gender: 'F', English: 95, Percentage: 94 },
        ],
        meta: { examName: 'BOARD_CLASS10_CBSE RESULT 2026', sectionName: '' },
      },
    },
  );

  assert.equal(cumulative.rows.length, 2);
  const byEnrollment = Object.fromEntries(cumulative.rows.map((row) => [row['Enrollment No'], row]));
  assert.equal(byEnrollment['101']['Board %'], 90);
  assert.equal(byEnrollment['102']['Board %'], 94);
});

test('buildWorkbookCumulativeSheet resolves duplicate board names against duplicate roster names by score ordering', () => {
  const cumulative = buildWorkbookCumulativeSheet(
    ['HY', 'PB2', 'BOARD RANK'],
    {
      HY: {
        headers: ['Enroll No.', 'Student Name', 'English (80)', '%'],
        rows: [
          { 'Enroll No.': '2016-2017/0080', 'Student Name': 'AARAV GUPTA', 'English (80)': 31, '%': 44.67 },
          { 'Enroll No.': '2022-2023/0246', 'Student Name': 'AARAV GUPTA', 'English (80)': 63, '%': 84.44 },
          { 'Enroll No.': '2018-2019/0292', 'Student Name': 'AARAV GUPTA', 'English (80)': 41, '%': 45.33 },
        ],
        meta: { examName: 'HY_10A_X-A', sectionName: 'A' },
      },
      PB2: {
        headers: ['Enroll No.', 'Student Name', 'English (80)', '%'],
        rows: [
          { 'Enroll No.': '2016-2017/0080', 'Student Name': 'AARAV GUPTA', 'English (80)': 40, '%': 50.89 },
          { 'Enroll No.': '2022-2023/0246', 'Student Name': 'AARAV GUPTA', 'English (80)': 68, '%': 85.11 },
          { 'Enroll No.': '2018-2019/0292', 'Student Name': 'AARAV GUPTA', 'English (80)': 6, '%': 7.78 },
        ],
        meta: { examName: 'PB2_10A_X-A', sectionName: 'A' },
      },
      'BOARD RANK': {
        headers: ['S.NO.', 'Roll No', 'Column3', 'Gender', 'English', 'Hindi', 'Mathematics', 'BASIC MATHS', 'AI', 'Percentage'],
        rows: [
          { 'S.NO.': 1, 'Roll No': '11188112', Column3: 'AARAV GUPTA', Gender: 'M', English: 95, Hindi: 93, Mathematics: 99, 'BASIC MATHS': '', AI: 100, Percentage: 94.67 },
          { 'S.NO.': 2, 'Roll No': '11188113', Column3: 'AARAV GUPTA', Gender: 'M', English: 80, Hindi: 86, Mathematics: 60, 'BASIC MATHS': '', AI: 92, Percentage: 80.33 },
          { 'S.NO.': 3, 'Roll No': '11188114', Column3: 'AARAV GUPTA', Gender: 'M', English: 73, Hindi: 79, Mathematics: '', 'BASIC MATHS': 75, AI: 89, Percentage: 74.67 },
        ],
        meta: { examName: 'BOARD_CLASS10_CBSE RESULT 2026', sectionName: '' },
      },
    },
  );

  assert.equal(cumulative.rows.length, 3);
  const byEnrollment = Object.fromEntries(cumulative.rows.map((row) => [row['Enrollment No'], row]));
  assert.equal(byEnrollment['2022-2023/0246']['Board %'], 94.67);
  assert.equal(byEnrollment['2016-2017/0080']['Board %'], 80.33);
  assert.equal(byEnrollment['2018-2019/0292']['Board %'], 74.67);
});
