const express = require('express');
const cors = require('cors');

const healthRoutes = require('./routes/health.routes');
const reportRoutes = require('./routes/report.routes');
const { legacyApp } = require('./legacy/report-engine');
const { notFound, errorHandler } = require('./middleware/error-handler');

function createApp() {
  const app = express();

  app.disable('x-powered-by');
  app.use(cors());
  app.use(express.json({ limit: '50mb' }));

  app.use('/api/v1', healthRoutes);
  app.use('/api/v1', reportRoutes);

  // Preserve the current app surface while the new modular API is built out.
  app.use('/', legacyApp);

  app.use(notFound);
  app.use(errorHandler);

  return app;
}

module.exports = {
  createApp,
};
