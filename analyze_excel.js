
const XLSX = require('xlsx');
const path = require('path');

const files = [
    'c:\\Users\\ppc.controllers\\Downloads\\demo\\source\\SHIFT OUT PUT 05.12.25.xlsm',
    'c:\\Users\\ppc.controllers\\Downloads\\demo\\source\\C&B PENDING @ 7AM.xlsx'
];

files.forEach(file => {
    console.log(`\n--- Analyzing: ${path.basename(file)} ---`);
    try {
        const workbook = XLSX.readFile(file);
        workbook.SheetNames.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(sheet, {header: 1, range: 0});
            console.log(`Sheet: ${sheetName}`);
            console.log(`Rows: ${data.length}`);
            console.log(`Header (first 2 rows):`, data.slice(0, 2));
        });
    } catch (e) {
        console.error(`Error reading ${file}: ${e.message}`);
    }
});
