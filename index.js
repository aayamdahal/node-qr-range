const QRCode = require('qrcode');
const PDFDocument = require('pdfkit');
const fs = require('fs');
const XLSX = require('xlsx');

// Function to generate QR codes with a label
async function generateQRCodeWithLabel(
  fileNo,
  startAccount,
  endAccount,
  totalAccounts
) {
  const qrData = `Account Range:\n${generateAccountRange(
    startAccount,
    endAccount
  )}`;
  const qrCode = await QRCode.toDataURL(qrData, { errorCorrectionLevel: 'H' });
  return { fileNo, qrCode, totalAccounts };
}

async function generatePDFWithQRCodesAndLabels(data) {
  const pdf = new PDFDocument({
    size: 'letter',
    margin: 50,
    layout: 'landscape',
  });

  const output = fs.createWriteStream('qr_codes.pdf');

  pdf.pipe(output);

  for (let i = 0; i < data.length; i++) {
    if (i > 0) {
      pdf.addPage();
    }

    const qrCodeWidth = 200; // Width of the QR code image
    const qrCodeHeight = 200; // Height of the QR code image

    // Calculate the center position for the text and QR code
    const textX = (pdf.page.width - qrCodeWidth) / 2;
    const textY = (pdf.page.height - qrCodeHeight) / 2;

    pdf.text(
      `FILE NO: ${data[i].fileNo} | Total Accounts: ${data[i].totalAccounts}`,
      textX,
      textY - 20
    );
    pdf.image(data[i].qrCode, textX, textY, {
      width: qrCodeWidth,
      height: qrCodeHeight,
    });
  }

  pdf.end();

  console.log('PDF with QR codes and labels generated.');
}

// Function to generate a string containing all account numbers within a range
function generateAccountRange(start, end) {
  let accountRange = '';
  for (let account = start; account <= end; account++) {
    accountRange += account.toString().padStart(14, '0') + '\n';
  }
  return accountRange;
}

// Read the Excel file
const workbook = XLSX.readFile('account_ranges.xlsx');
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Parse the Excel data
const excelData = XLSX.utils.sheet_to_json(worksheet);

const qrCodePromises = [];

// Process each row in the Excel data and calculate the total accounts
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

Promise.all(qrCodePromises)
  .then((qrCodesWithLabels) => {
    generatePDFWithQRCodesAndLabels(qrCodesWithLabels);
  })
  .catch((error) => {
    console.error('Error generating QR codes and PDF:', error);
  });
