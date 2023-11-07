# QR Code Generator and PDF Generator

A Node.js script that generates QR codes with labels and compiles them into a PDF document. It is designed to work with Excel files containing data to create QR codes and labels for each data entry. Right now the labels are Total Accounts in the start and end range and and File No of that particular account range.

## Prerequisites

Before using this script, ensure you have the following installed:

- [Node.js](https://nodejs.org/): The script is written in JavaScript and requires Node.js to run.
- [npm](https://www.npmjs.com/): Node Package Manager, used to install project dependencies.

## Installation

1. Clone this repository to your local machine or download the script directly.

2. Navigate to the repository's directory using the command line.

3. Install the required dependencies by running the following command:

```bash
npm install
```

## Usage

1. Place your Excel file (`.xlsx`) in the same directory as the script.

2. Open the file (`index.js`) and update the `directoryPath` variable to point to the directory where your Excel file is located.

```javascript
const directoryPath = './'; // Update this to your directory path no need to update if the excel is in the same directory as the script
```

3. Run the script using the following command:

```bash
node index.js
```

## Script Details

1. `generateQRCodeWithLabel`: Generates a QR code with a label that includes an account range message.

2. `generatePDFWithQRCodesAndLabels`: Compiles an array of QR codes with labels into a PDF document.

3. `generateAccountRange`: Generates an account range message for the label.

4. `findExcelFileInDirectory`: Searches for an Excel file in the specified directory.

## Excel File Details

The script expects an Excel file with the following structure:

1. `FILE NO`: A column containing file numbers.
2. `START`: A column containing the start account number.
3. `END`: A column containing the end account number.

## Troubleshooting

If you encounter any issues, please make sure that your Excel file is correctly formatted and that all required dependencies are installed.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details
