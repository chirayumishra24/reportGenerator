function loadEnv() {
  return {
    port: Number(process.env.PORT || 5000),
    nodeEnv: process.env.NODE_ENV || 'development',
  };
}

module.exports = {
  loadEnv,
};
