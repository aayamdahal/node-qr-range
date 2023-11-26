const XLSX = require('xlsx');
const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');
const util = require('util');
const fs = require('fs');
const writeFileAsync = util.promisify(fs.writeFile);

const { establishDBConnection } = require('./database');
const {
  findExcelFileInDirectory,
  writeMissingAccountsToFile,
  generateAccountRange,
  generateQRCode,
} = require('./fileHandler');

const directoryPath = './';
const excelFilePath = findExcelFileInDirectory(directoryPath);

if (excelFilePath) {
  establishDBConnection()
    .then(async (accountsFromDB) => {
      const workbook = XLSX.readFile(excelFilePath);
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const excelData = XLSX.utils.sheet_to_json(worksheet);

      const fileNumbers = new Set(excelData.map((row) => row['FILE NO']));

      const missingAccountsByFile = {};

      const pdfDoc = await PDFDocument.create();

      for (const fileNo of fileNumbers) {
        const fileAccounts = excelData.find((row) => row['FILE NO'] === fileNo);
        const startAccountNumber = fileAccounts.START;
        const endAccountNumber = fileAccounts.END;

        const accountsRange = generateAccountRange(
          startAccountNumber,
          endAccountNumber
        );
        const extractedAccountNumbers = accountsRange.match(/\d{14}/g) || [];

        const missingAccounts = extractedAccountNumbers.filter(
          (account) => !accountsFromDB.includes(account)
        );

        if (missingAccounts.length > 0) {
          missingAccountsByFile[fileNo] = missingAccounts;

          const remainingAccounts = extractedAccountNumbers.filter(
            (account) => !missingAccounts.includes(account)
          );

          const totalCount = remainingAccounts.length;
          const remainingAccountsText = remainingAccounts.join('\n');
          const qrCodeData = await generateQRCode(remainingAccountsText);

          const qrImage = await pdfDoc.embedPng(qrCodeData);
          const page = pdfDoc.addPage();
          const { width, height } = page.getSize();

          const scaleFactor = 0.5;

          const qrYOffset = 100; // Adjust this value to move the QR code upwards

          page.drawImage(qrImage, {
            x: width / 2 - (qrImage.width * scaleFactor) / 2,
            y: height / 2 - (qrImage.height * scaleFactor) / 2 + qrYOffset, // Adjust this value as needed
            width: qrImage.width * scaleFactor,
            height: qrImage.height * scaleFactor,
          });

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

          const textSpacing = 15; // Adjust this value to create space between the text elements

          page.drawText(`Total Account Count: ${totalCount}`, {
            x: width / 2 - totalCountWidth / 2,
            y: height / 2 + qrImage.height * scaleFactor + textSpacing, // Adjust this value as needed
            size: fontSize,
          });

          const fileNoY =
            height / 2 + qrImage.height * scaleFactor + textSpacing * 2; // Adjust this value as needed
          page.drawText(`File Number: ${fileNo}`, {
            x: width / 2 - fileNoWidth / 2,
            y: fileNoY,
            size: fontSize,
          });
        }
      }

      const pdfBytes = await pdfDoc.save();
      const pdfFilePath = 'merged_qr_codes.pdf'; // Output PDF file path
      await writeFileAsync(pdfFilePath, pdfBytes);

      console.log(
        `PDF with each QR code on a separate page generated: ${pdfFilePath}`
      );

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
