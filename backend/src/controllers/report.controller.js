const reportService = require('../services/report.service');

async function parseWorkbook(req, res) {
  const payload = await reportService.parseWorkbook(req.file);
  res.json(payload);
}

async function structuredImport(req, res) {
  const payload = await reportService.structuredImport(req.files);
  if (!payload.validationPassed) {
    return res.status(400).json(payload);
  }
  return res.json(payload);
}

async function importPersistent(req, res) {
  const payload = await reportService.importPersistent(req.files);
  res.json(payload);
}

async function getCumulativeReport(req, res) {
  const payload = await reportService.getCumulativeReport();
  res.json(payload);
}

async function getDbStatus(req, res) {
  const payload = await reportService.getDbStatus();
  res.json(payload);
}

async function getStudentHistory(req, res) {
  const payload = await reportService.getStudentHistory(req.params.studentId);
  res.json(payload);
}

module.exports = {
  parseWorkbook,
  structuredImport,
  importPersistent,
  getCumulativeReport,
  getDbStatus,
  getStudentHistory,
};
