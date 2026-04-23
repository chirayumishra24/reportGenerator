const { createApp } = require('./src/app');
const { loadEnv } = require('./src/config/env');

const app = createApp();

if (require.main === module) {
  const env = loadEnv();
  app.listen(env.port, () => {
    console.log(`Backend running on http://localhost:${env.port}`);
  });
}

module.exports = app;
