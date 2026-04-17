const XLSX = require('xlsx');

// Create a dummy workbook
const wb = XLSX.utils.book_new();
const wsData = [
  ['A', 'B', 'C'],
  [1, 2, 3],
  [4, 5, 6],
  [7, 8, 9]
];
const ws = XLSX.utils.aoa_to_sheet(wsData);

// Manually add hidden rows metadata
ws['!rows'] = [];
ws['!rows'][1] = { hidden: true };  // Hide row 2 (index 1)

// Test sheet_to_json
const rows = XLSX.utils.sheet_to_json(ws, { header: "A", defval: "", blankrows: true });
console.log("rows.length:", rows.length);
rows.forEach((row, idx) => {
    console.log(`idx: ${idx}, C: ${row['C']}`);
});
