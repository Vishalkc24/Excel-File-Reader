// src/readExcel.ts

import * as XLSX from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';

// Define the Excel file path
const excelFileName = 'iw-tech-test-retailer-data.xlsx';
const excelFilePath = path.join(__dirname, excelFileName);

// Read the Excel file
const workbook = XLSX.readFile(excelFilePath);

// Choose the sheet you want to read (assuming it's the first sheet in this example)
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

// Convert the sheet data to JSON
const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

// Extract the column headers (first row in Excel)
const headers = jsonData[0] as string[];

// Optional: If you want to write the JSON data to a file
const outputDirectory = path.join(__dirname, 'output', 'json');
const jsonOutputPath = path.join(outputDirectory, 'file.json');

// Ensure the output directory exists
if (!fs.existsSync(outputDirectory)) {
  fs.mkdirSync(outputDirectory, { recursive: true });
}

// Ensure the output file exists or create a new one
try {
  const jsonDataArray = jsonData.slice(1).map((row: any) => {
    const rowData: { [key: string]: any } = {};
    headers.forEach((header: string, index: number) => {
      // If the cell is empty, use "NULL" as the value
      rowData[header] = row[index] === undefined || row[index] === null ? "NULL" : row[index];
    });
    return rowData;
  });

  fs.writeFileSync(jsonOutputPath, JSON.stringify(jsonDataArray, null, 2));
  console.log('Excel file has been successfully processed.');
} catch (error) {
  console.error(`Error writing JSON file: ${(error as Error).message}`);
}

// Process the data
jsonData.slice(1).forEach((row: any) => {
  const rowData: { [key: string]: any } = {};

  // Map each column to its corresponding header
  headers.forEach((header: string, index: number) => {
    // If the cell is empty, use "NULL" as the value
    rowData[header] = row[index] === undefined || row[index] === null ? "NULL" : row[index];
  });

  // Output each row as an array of key-value pairs
  const rowArray = Object.entries(rowData).map(([key, value]) => `${JSON.stringify(key)}: ${JSON.stringify(value)}`);
  console.log(`[${rowArray.join(', ')}]`);
});
