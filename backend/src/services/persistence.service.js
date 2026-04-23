const { getDatabaseContext } = require('../config/database');
const {
  parseWorkbookBuffer,
  normalizeSheetName,
  normalizeText,
  extractStudentMetrics,
  toNumber,
} = require('./parser.service');
const {
  buildEmptyCumulativeReport,
  buildCumulativeReportPayload,
} = require('./cumulative.service');

function badRequest(message) {
  const error = new Error(message);
  error.status = 400;
  return error;
}

function ensureDatabaseConfigured() {
  const { prisma, dbEnabled } = getDatabaseContext();
  if (!dbEnabled || !prisma) {
    throw badRequest('Database is not configured. Set DATABASE_URL and run Prisma migration first.');
  }
  return prisma;
}

async function upsertStudentRecord(tx, metrics, meta) {
  const admissionNo = String(metrics.admissionNo || '').trim() || null;
  const normalizedName = normalizeText(metrics.name || admissionNo || 'student');
  const className = meta.className || null;
  const sectionName = meta.sectionName || null;

  if (admissionNo) {
    const existing = await tx.student.findFirst({ where: { admissionNo } });
    if (existing) {
      return tx.student.update({
        where: { id: existing.id },
        data: {
          name: metrics.name || existing.name || admissionNo || 'Student',
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
      name: metrics.name || admissionNo || 'Student',
      normalizedName,
      admissionNo,
      className,
      sectionName,
    },
  });
}

async function importParsedFiles(files) {
  if (!files || files.length === 0) {
    throw badRequest('No files uploaded');
  }

  const prisma = ensureDatabaseConfigured();
  const importSummary = [];

  for (const file of files) {
    const parsed = parseWorkbookBuffer(file.buffer, file.originalname);
    const validationErrors = parsed.parsedSheets.flatMap((sheet) =>
      (sheet.validation?.issues || []).map((issue) => `${file.originalname} / ${sheet.sheetName}: ${issue}`),
    );
    if (validationErrors.length > 0) {
      throw badRequest(validationErrors.join(' | '));
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

async function getCumulativeReport() {
  const { prisma, dbEnabled } = getDatabaseContext();
  if (!dbEnabled || !prisma) {
    return buildEmptyCumulativeReport();
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

  return buildCumulativeReportPayload({ uploads, sheets, students });
}

async function getDbStatus() {
  const { prisma, dbEnabled } = getDatabaseContext();
  if (!dbEnabled || !prisma) {
    return { enabled: false, message: 'DATABASE_URL not configured' };
  }

  try {
    await prisma.$queryRaw`SELECT 1`;
    return { enabled: true };
  } catch (error) {
    return { enabled: false, error: error.message };
  }
}

async function getStudentHistory(studentId) {
  const prisma = ensureDatabaseConfigured();
  const student = await prisma.student.findUnique({
    where: { id: studentId },
    include: {
      performances: {
        include: { examSheet: true },
        orderBy: { createdAt: 'asc' },
      },
    },
  });

  if (!student) {
    const error = new Error('Student not found');
    error.status = 404;
    throw error;
  }

  return student;
}

module.exports = {
  importParsedFiles,
  getCumulativeReport,
  getDbStatus,
  getStudentHistory,
};
