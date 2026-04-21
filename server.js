const http = require('http');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const PORT = 3000;
const DB_PATH = '\\\\10.199.100.4\\PPC Share\\18-PPC CONTROLLER\\PPC DATABASE\\DAILY REPORT\\REPORT.xlsx';

const server = http.createServer((req, res) => {
    // Add CORS headers
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') {
        res.writeHead(204);
        res.end();
        return;
    }

    if (req.method === 'POST' && req.url === '/import') {
        let body = '';
        req.on('data', chunk => body += chunk);
        req.on('end', () => {
            try {
                const payload = JSON.parse(body);
                const { date, tables } = payload;
                
                console.log(`Received import request for: ${date}`);

                let wb;
                if (fs.existsSync(DB_PATH)) {
                    console.log(`Reading existing database at: ${DB_PATH}`);
                    wb = XLSX.readFile(DB_PATH);
                } else {
                    console.log(`Creating new database file.`);
                    wb = XLSX.utils.book_new();
                }

                // Add sheets for this day
                tables.forEach(t => {
                    const ws = XLSX.utils.aoa_to_sheet(t.rows);
                    const sheetName = `${t.suffix}_${date}`.substring(0, 31);

                    // Remove existing if duplicate
                    if (wb.SheetNames.includes(sheetName)) {
                        wb.SheetNames = wb.SheetNames.filter(n => n !== sheetName);
                        delete wb.Sheets[sheetName];
                    }
                    XLSX.utils.book_append_sheet(wb, ws, sheetName);
                });

                XLSX.writeFile(wb, DB_PATH);
                console.log(`Successfully updated: ${DB_PATH}`);

                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ success: true, message: `Report for ${date} imported successfully.` }));
            } catch (err) {
                console.error('Import Error:', err);
                res.writeHead(500, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ success: false, error: err.message }));
            }
        });
    } else {
        res.writeHead(404);
        res.end();
    }
});

server.listen(PORT, () => {
    console.log(`PPC Import Server running at http://localhost:${PORT}`);
    console.log(`Targeting Database: ${DB_PATH}`);
});
