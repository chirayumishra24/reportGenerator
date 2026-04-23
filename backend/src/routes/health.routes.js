const express = require('express');

const router = express.Router();

router.get('/health', (req, res) => {
  res.status(200).json({
    ok: true,
    service: 'report-generator-api',
    architecture: 'modular-v1',
  });
});

module.exports = router;
