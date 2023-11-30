const fs = require('fs');
const path = require('path');
const QRCode = require('qrcode');

/**
 * Find the Excel file in the specified directory
 * @param {string} directoryPath - The directory path to search for the Excel file
 * @returns {string|null} - The path of the Excel file if found, otherwise null
 */

function findExcelFileInDirectory(directoryPath) {
  const files = fs.readdirSync(directoryPath);
  for (const file of files) {
    if (file.endsWith('.xlsx')) {
      return path.join(directoryPath, file);
    }
  }
  return null;
}

/**
 * Generate the range of accounts between start and end account numbers
 * @param {number} start - The starting account number
 * @param {number} end - The ending account number
 * @returns {string} - A string representing the range of accounts
 */

function generateAccountRange(start, end) {
  let accountRange = '';
  for (let account = start; account <= end; account++) {
    accountRange += account.toString().padStart(14, '0') + '\n';
  }
  return accountRange;
}

/**
 * Write missing accounts to a file
 * @param {string[]} missingAccounts - An array of missing account numbers
 */

async function writeMissingAccountsToFile(missingAccounts) {
  const fileName = 'missing_accounts.txt';

  try {
    fs.writeFileSync(fileName, 'Missing Accounts:\n');
    missingAccounts.forEach((account) => {
      fs.appendFileSync(fileName, `${account}\n`);
    });
    console.log(`Missing accounts written to ${fileName}`);
  } catch (err) {
    console.error('Error writing missing accounts to file:', err);
  }
}

/**
 * Generate QR code for the given data
 * @param {string} data - The data to be encoded in the QR code
 * @returns {Promise<Buffer>} - A promise resolving to the generated QR code buffer
 */

async function generateQRCode(data) {
  try {
    const qrCode = await QRCode.toBuffer(data, { errorCorrectionLevel: 'H' });
    return qrCode;
  } catch (error) {
    throw new Error('Error generating QR code: ' + error.message);
  }
}

module.exports = {
  findExcelFileInDirectory,
  writeMissingAccountsToFile,
  generateAccountRange,
  generateQRCode,
};
