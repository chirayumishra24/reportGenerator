const { PrismaClient } = require('@prisma/client');

let prisma = null;
let dbEnabled = Boolean(process.env.DATABASE_URL);

if (dbEnabled) {
  try {
    prisma = new PrismaClient();
  } catch (error) {
    dbEnabled = false;
    console.warn('Prisma client could not be initialized:', error.message);
  }
}

function getDatabaseContext() {
  return {
    prisma,
    dbEnabled,
  };
}

module.exports = {
  prisma,
  dbEnabled,
  getDatabaseContext,
};
