const XLSX = require('xlsx');
const path = require('path');

const filePath = 'c:/Users/ASUS/OneDrive/Desktop/skilizee/report-generator/sheets/TARGET  SHEET 2024-25 NEW (1).xlsx';
const workbook = XLSX.readFile(filePath);
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1 });
const headers = matrix[0];

console.log('Sheet:', sheetName);
console.log('Headers:', JSON.stringify(headers));
console.log('First Row Data:', JSON.stringify(matrix[1]));
