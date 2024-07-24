const XLSX = require('xlsx');
const fs = require('fs');

const workbook = XLSX.readFile('excel.xlsx');
const sheetName = workbook.SheetNames[0];
const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '' });

const escapeXml = (unsafe) => {
    return unsafe.replace(/[<>&'"]/g, (c) => {
        switch (c) {
            case '<': return '&lt;';
            case '>': return '&gt;';
            case '&': return '&amp;';
            case '\'': return '&apos;';
            case '"': return '&quot;';
        }
    });
};

const allColumns = new Set();
sheet.forEach(row => {
    Object.keys(row).forEach(col => {
        if (col.trim() && !/^__EMPTY/.test(col)) {
            allColumns.add(col);
        }
    });
});

let xmlData = '<?xml version="1.0"?>\n<SHOP>\n';

sheet.forEach(row => {
    xmlData += '  <SHOPITEM>\n';
    allColumns.forEach(col => {
        xmlData += `    <${col}><![CDATA[${escapeXml(String(row[col] || ''))}]]></${col}>\n`;
    });
    xmlData += '  </SHOPITEM>\n';
});

xmlData += '</SHOP>\n';

fs.writeFileSync('./output/output.xml', xmlData);

console.log('XML file has been generated to /output folder in this app.');
