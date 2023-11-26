const fs = require('fs');
const path = require('path');
const QRCode = require('qrcode');

function findExcelFileInDirectory(directoryPath) {
  const files = fs.readdirSync(directoryPath);
  for (const file of files) {
    if (file.endsWith('.xlsx')) {
      return path.join(directoryPath, file);
    }
  }
  return null;
}

function generateAccountRange(start, end) {
  let accountRange = '';
  for (let account = start; account <= end; account++) {
    accountRange += account.toString().padStart(14, '0') + '\n';
  }
  return accountRange;
}

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

async function generateQRCode(data) {
  try {
    // Generate QR code with the given data
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
