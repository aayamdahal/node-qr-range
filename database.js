const sql = require('mssql');

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

async function establishDBConnection() {
  try {
    const pool = await sql.connect(config);
    console.log('Connected to MSSQL Database');

    // Fetch all accounts from the 'documents' table
    const result = await pool.request().query('SELECT name FROM documents');
    const accountsFromDB = result.recordset.map((account) => account.name);

    return accountsFromDB;
  } catch (err) {
    console.error('Error connecting to the database:', err);
    throw err;
  }
}

module.exports = {
  establishDBConnection,
};
