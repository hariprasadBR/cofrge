const XLSX = require('xlsx');
const fs = require('fs');

// Function to read the Excel file
const readExcelFile = (filePath) => {
    // Read the Excel file into a workbook
    const workbook = XLSX.readFile(filePath);

    // Get the first sheet name
    const sheetName = workbook.SheetNames[0];
    // Get the sheet object
    const sheet = workbook.Sheets[sheetName];

    return sheet;
};

// Function to find the column index for the header
const findColumnIndex = (sheet, headerName) => {
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const firstRow = range.s.r; // Get the first row number

    for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
        const cellAddress = { c: colNum, r: firstRow };
        const cellRef = XLSX.utils.encode_cell(cellAddress);
        const cell = sheet[cellRef];
        if (cell && cell.v === headerName) {
            return colNum;
        }
    }
    return -1; // Return -1 if the header is not found
};

// Function to capture all values in the specified column
const captureColumnValues = (sheet, colNum) => {
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const values = [];

    // Start from the second row to skip the header
    for (let rowNum = range.s.r + 1; rowNum <= range.e.r; rowNum++) {
        const cellAddress = { c: colNum, r: rowNum };
        const cellRef = XLSX.utils.encode_cell(cellAddress);
        const cell = sheet[cellRef];
        const cellValue = (cell ? cell.v : undefined);
        values.push(cellValue);
    }

    return values;
};

// Main function to execute the extraction
const main = (filePath, headerName) => {
    const sheet = readExcelFile(filePath);
    const colNum = findColumnIndex(sheet, headerName);

    if (colNum === -1) {
        console.error(`Column "${headerName}" not found.`);
        return;
    }

    const headerValues = captureColumnValues(sheet, colNum);

    console.log(headerValues);

    // Optionally, save the captured values to a JSON file
    fs.writeFileSync('header_values.json', JSON.stringify(headerValues, null, 2));
};

// Example usage
const filePath = 'C:/Users/Hari Prasad/Desktop/jsp/book1.xlsx'; // Replace with your actual file path
const headerName = 'method Name'; // Replace with your header name
main(filePath, headerName);
