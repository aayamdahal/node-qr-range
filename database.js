const sql = require('mssql');
const fs = require('fs');
const path = require('path');

const config = {
  user: 'sa',
  password: '1234',
  server: 'localhost',
  port: 1433,
  database: 'gdms_ctnz_another',
  options: {
    encrypt: false, // Depending on your server configuration
    enableArithAbort: true, // Depending on your server configuration
  },
};

/**
 * Function to write logs to a file
 * @param {string} logMessage - The message to be logged
 */
function writeToLogFile(logMessage) {
  const logFilePath = path.join(__dirname, 'db_logs.txt');
  const timestamp = new Date().toISOString();
  const log = `${timestamp}: ${logMessage}\n`;

  fs.appendFile(logFilePath, log, (err) => {
    if (err) {
      console.error('Error writing to log file:', err);
    }
  });
}

/**
 * Establishes a connection to the MSSQL database
 * @returns {Promise<Array<string>>} - A promise that resolves to an array of account names from the 'documents' table
 */
async function establishDBConnection() {
  try {
    const pool = await sql.connect(config);
    console.log('Connected to MSSQL Database');
    writeToLogFile('Connected to MSSQL Database');

    // Fetch all accounts from the 'documents' table
    const result = await pool.request().query('SELECT name FROM documents');
    const accountsFromDB = result.recordset.map((account) => account.name);

    return accountsFromDB;
  } catch (err) {
    const errorMessage = `Error connecting to the database: ${err}`;
    console.error(errorMessage);
    writeToLogFile(errorMessage);
    throw err;
  }
}

module.exports = {
  establishDBConnection,
};
