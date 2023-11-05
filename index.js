const qr = require('qrcode');
const fs = require('fs/promises');
const XLSX = require('xlsx');

// Function to create a directory if it doesn't exist
async function createDirectoryIfNotExists(directory) {
  try {
    await fs.mkdir(directory, { recursive: true });
  } catch (error) {
    console.error(`Error creating directory "${directory}":`, error);
  }
}

// Function to generate a QR code image
async function generateQRCode(data, filename) {
  try {
    const qrCodeImage = await qr.toDataURL(data);
    await fs.writeFile(filename, qrCodeImage.split(',')[1], 'base64');
    console.log(`QR code generated for ${filename}`);
  } catch (error) {
    console.error(`Error generating QR code for ${filename}:`, error);
  }
}

// Function to generate a range of account numbers
function generateAccountRange(start, end) {
  const accountNumbers = [];
  for (let account = start; account <= end; account++) {
    accountNumbers.push(account);
  }
  return accountNumbers.join('\n');
}

// Function to process Excel data and generate QR codes
async function processExcelData(filePath) {
  try {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const excelData = XLSX.utils.sheet_to_json(worksheet);

    for (const row of excelData) {
      const startAccountNumber = row.startAccountNumber;
      const endAccountNumber = row.endAccountNumber;
      const accountRange = generateAccountRange(
        startAccountNumber,
        endAccountNumber
      );
      const qrData = `Account Range:\n${accountRange}`;
      const filename = `qr/qr_${startAccountNumber}_${endAccountNumber}.png`;

      await createDirectoryIfNotExists('qr');
      await generateQRCode(qrData, filename);
    }
  } catch (error) {
    console.error('Error processing Excel data:', error);
  }
}

// Run the processing function
processExcelData('account_ranges.xlsx');
