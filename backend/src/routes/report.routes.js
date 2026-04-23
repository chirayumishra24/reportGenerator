const express = require('express');
const multer = require('multer');

const asyncHandler = require('../middleware/async-handler');
const reportController = require('../controllers/report.controller');

const router = express.Router();
const upload = multer({ storage: multer.memoryStorage() });

router.post('/reports/parse', upload.single('file'), asyncHandler(reportController.parseWorkbook));
router.post('/reports/structured-import', upload.array('files'), asyncHandler(reportController.structuredImport));
router.post('/reports/import-persistent', upload.array('files'), asyncHandler(reportController.importPersistent));
router.get('/reports/cumulative', asyncHandler(reportController.getCumulativeReport));
router.get('/reports/db-status', asyncHandler(reportController.getDbStatus));
router.get('/reports/student-history/:studentId', asyncHandler(reportController.getStudentHistory));

module.exports = router;
