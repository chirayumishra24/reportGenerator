function notFound(req, res) {
  res.status(404).json({
    error: `Route not found: ${req.method} ${req.originalUrl}`,
  });
}

function errorHandler(error, req, res, next) { // eslint-disable-line no-unused-vars
  const status = error.status || 500;
  res.status(status).json({
    error: error.message || 'Internal server error',
  });
}

module.exports = {
  notFound,
  errorHandler,
};
