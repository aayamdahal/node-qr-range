// Import necessary libraries
const QRCode = require('qrcode');
const PDFDocument = require('pdfkit');
const fs = require('fs');
const XLSX = require('xlsx');

/**
 * Generate a QR code with a label.
 *
 * @param {string} fileNo - The file number.
 * @param {number} startAccount - The start account number.
 * @param {number} endAccount - The end account number.
 * @param {number} totalAccounts - The total number of accounts in the range.
 * @returns {Promise<object>} An object containing the file number, QR code, and total accounts.
 */
async function generateQRCodeWithLabel(
  fileNo,
  startAccount,
  endAccount,
  totalAccounts
) {
  // Create the data for the QR code, including an account range message
  const qrData = `Account Range:\n${generateAccountRange(
    startAccount,
    endAccount
  )}`;

  // Generate the QR code and return it as a Data URL
  const qrCode = await QRCode.toDataURL(qrData, { errorCorrectionLevel: 'H' });

  return { fileNo, qrCode, totalAccounts };
}

/**
 * Generate a PDF containing QR codes with labels.
 *
 * @param {object[]} data - An array of objects, each containing file number, QR code, and total accounts.
 */
async function generatePDFWithQRCodesAndLabels(data) {
  // Create a new PDF document
  const pdf = new PDFDocument({
    size: 'letter',
    margin: 50,
    layout: 'landscape',
  });

  // Create an output stream for the PDF
  const output = fs.createWriteStream('qr_codes.pdf');

  // Pipe the PDF document to the output stream
  pdf.pipe(output);

  // Loop through the data and add QR codes with labels to the PDF
  for (let i = 0; i < data.length; i++) {
    if (i > 0) {
      pdf.addPage(); // Add a new page for each QR code
    }

    const qrCodeWidth = 200;
    const qrCodeHeight = 200;

    const textX = (pdf.page.width - qrCodeWidth) / 2;
    const textY = (pdf.page.height - qrCodeHeight) / 2;

    // Add a label with file number and total accounts above the QR code
    pdf.text(
      `FILE NO: ${data[i].fileNo} | Total Accounts: ${data[i].totalAccounts}`,
      textX,
      textY - 20
    );

    // Add the QR code image to the PDF
    pdf.image(data[i].qrCode, textX, textY, {
      width: qrCodeWidth,
      height: qrCodeHeight,
    });
  }

  // End the PDF creation
  pdf.end();

  console.log('PDF with QR codes and labels generated.');
}

/**
 * Generate an account range message.
 *
 * @param {number} start - The start account number.
 * @param {number} end - The end account number.
 * @returns {string} The account range message.
 */
function generateAccountRange(start, end) {
  let accountRange = '';
  for (let account = start; account <= end; account++) {
    accountRange += account.toString().padStart(14, '0') + '\n';
  }
  return accountRange;
}

// Read data from an Excel file
const workbook = XLSX.readFile('account_ranges.xlsx');
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Convert the Excel data to JSON
const excelData = XLSX.utils.sheet_to_json(worksheet);

// Create an array of promises to generate QR codes with labels
const qrCodePromises = [];

// Iterate through the Excel data and generate QR codes with labels
for (const row of excelData) {
  const fileNo = row['FILE NO'];
  const startAccountNumber = row.START;
  const endAccountNumber = row.END;

  const totalAccounts = endAccountNumber - startAccountNumber + 1;

  qrCodePromises.push(
    generateQRCodeWithLabel(
      fileNo,
      startAccountNumber,
      endAccountNumber,
      totalAccounts
    )
  );
}

// Wait for all QR codes to be generated, then generate the PDF
Promise.all(qrCodePromises)
  .then((qrCodesWithLabels) => {
    generatePDFWithQRCodesAndLabels(qrCodesWithLabels);
  })
  .catch((error) => {
    console.error('Error generating QR codes and PDF:', error);
  });
