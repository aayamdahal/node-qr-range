const XLSX = require('xlsx');
const { PDFDocument, StandardFonts } = require('pdf-lib');
const util = require('util');
const fs = require('fs');
const writeFileAsync = util.promisify(fs.writeFile);

const { establishDBConnection } = require('./database');
const {
  findExcelFileInDirectory,
  generateAccountRange,
  generateQRCode,
} = require('./fileHandler');

// Find the Excel file in the specified directory
const directoryPath = './';
const excelFilePath = findExcelFileInDirectory(directoryPath);

if (excelFilePath) {
  establishDBConnection()
    .then(async (accountsFromDB) => {
      // Read the Excel file and convert it to JSON
      const workbook = XLSX.readFile(excelFilePath);
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const excelData = XLSX.utils.sheet_to_json(worksheet);

      // Extract unique file numbers from the Excel data
      const fileNumbers = new Set(excelData.map((row) => row['FILE NO']));
      const missingAccountsByFile = {};

      // Create a new PDF document
      const pdfDoc = await PDFDocument.create();

      // Iterate through each file number
      for (const fileNo of fileNumbers) {
        // Retrieve accounts for the current file number
        const fileAccounts = excelData.find((row) => row['FILE NO'] === fileNo);
        const startAccountNumber = fileAccounts.START;
        const endAccountNumber = fileAccounts.END;

        // Generate the range of accounts between start and end account numbers
        const accountsRange = generateAccountRange(
          startAccountNumber,
          endAccountNumber
        );
        const extractedAccountNumbers = accountsRange.match(/\d{14}/g) || [];

        // Find missing accounts by comparing with accounts from the database
        const missingAccounts = extractedAccountNumbers.filter(
          (account) => !accountsFromDB.includes(account)
        );

        if (missingAccounts.length > 0) {
          // Store missing accounts by file number
          missingAccountsByFile[fileNo] = missingAccounts;

          const remainingAccounts = extractedAccountNumbers.filter(
            (account) => !missingAccounts.includes(account)
          );

          if (remainingAccounts.length > 0) {
            // Generate QR codes for the remaining accounts and embed in PDF
            const totalCount = remainingAccounts.length;
            const remainingAccountsText = remainingAccounts.join('\n');
            const qrCodeData = await generateQRCode(remainingAccountsText);

            const qrImage = await pdfDoc.embedPng(qrCodeData);
            const page = pdfDoc.addPage();
            const { width, height } = page.getSize();

            // Set parameters for QR code placement in the PDF
            const scaleFactor = 0.5;

            // Get the size of the QR code image
            const qrWidth = qrImage.width * scaleFactor;
            const qrHeight = qrImage.height * scaleFactor;

            // Calculate the margin for text above the QR code
            const marginAboveQR = 2 * 12; // Assuming 1 rem = 12px
            const textSpacing = 15;

            // Add text related to account count and file number on the PDF page
            const fontSize = 12;
            const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
            const totalCountWidth = font.widthOfTextAtSize(
              `Total Account Count: ${totalCount}`,
              fontSize
            );
            const fileNoWidth = font.widthOfTextAtSize(
              `File Number: ${fileNo}`,
              fontSize
            );

            // Calculate the Y coordinate for text above the QR code
            const textAboveY =
              height / 2 + qrHeight / 2 + marginAboveQR + textSpacing;

            // Draw QR code image on the PDF page
            page.drawImage(qrImage, {
              x: width / 2 - qrWidth / 2,
              y: height / 2 - qrHeight / 2,
              width: qrWidth,
              height: qrHeight,
            });

            // Draw text above the QR code
            page.drawText(`Total Account Count: ${totalCount}`, {
              x: width / 2 - totalCountWidth / 2,
              y: textAboveY,
              size: fontSize,
            });

            const fileNoY = textAboveY + textSpacing;
            page.drawText(`File Number: ${fileNo}`, {
              x: width / 2 - fileNoWidth / 2,
              y: fileNoY,
              size: fontSize,
            });
          }
        }
      }

      // Save the generated PDF with QR codes
      const pdfBytes = await pdfDoc.save();
      const pdfFilePath = 'merged_qr_codes.pdf';
      await writeFileAsync(pdfFilePath, pdfBytes);

      console.log(
        `PDF with each QR code on a separate page generated: ${pdfFilePath}`
      );

      // Create a summary of missing accounts by file number and save to a text file
      let combinedMissingAccounts = '';
      for (const fileNo in missingAccountsByFile) {
        combinedMissingAccounts += `Missing accounts for FILE NO ${fileNo}:\n${missingAccountsByFile[
          fileNo
        ].join('\n')}\n\n`;
      }

      const combinedMissingAccountsFilePath = 'combined_missing_accounts.txt';
      await writeFileAsync(
        combinedMissingAccountsFilePath,
        combinedMissingAccounts
      );
      console.log(
        `Combined missing accounts dumped to: ${combinedMissingAccountsFilePath}`
      );
    })
    .catch((error) => {
      console.error('Error:', error);
    });
} else {
  console.error('No Excel file found in the directory.');
}
