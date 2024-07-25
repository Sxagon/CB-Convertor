const XLSX = require('xlsx');
const fs = require('fs');

const workbook = XLSX.readFile('excel.xlsx', { cellStyles: true });
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

const jsonSheet = XLSX.utils.sheet_to_json(sheet, { header: 1 });

const headers = jsonSheet[0];
const rows = jsonSheet.slice(1);

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

const getCellValue = (headers, row, headerName) => {
    const index = headers.indexOf(headerName);
    return index !== -1 ? (row[index] || '').toString().trim() : '';
};

const extractNumber = (str) => {
    return parseFloat(str.replace(/[^\d.-]/g, '')) || 0;
};

let xmlData = '<?xml version="1.0"?>\n<SHOP>\n';

rows.forEach(row => {
    const productName = escapeXml(getCellValue(headers, row, 'nazev'));
    const code = escapeXml(getCellValue(headers, row, 'kod'));
    const description = `<p>${escapeXml(getCellValue(headers, row, 'popis'))}</p><hr><h3>Složení</h3><p>${escapeXml(getCellValue(headers, row, 'slozeni'))}</p>`;
    const annotation = escapeXml(getCellValue(headers, row, 'popis'));
    const weight = extractNumber(getCellValue(headers, row, 'gramaz')) / 1000; // Convert to kg
    const priceWithVat = parseFloat(getCellValue(headers, row, 'voc') || '0');
    const warranty = parseInt(getCellValue(headers, row, 'spotreba').replace(/[^\d]/g, ''), 10) || 0;
    const imageUrl = `ftp://b2b.cukrobaron.cz/${escapeXml(getCellValue(headers, row, 'foto') || 'example')}.jpg`;
    const allergens = escapeXml(getCellValue(headers, row, 'alergeny'));
    const freezing = escapeXml(getCellValue(headers, row, 'zamrazeni'));

    xmlData += `  <SHOPITEM>\n`;
    xmlData += `    <PRODUCTNAME><![CDATA[${productName}]]></PRODUCTNAME>\n`;
    xmlData += `    <CODE><![CDATA[${code}]]></CODE>\n`;
    xmlData += `    <DESCRIPTION><![CDATA[${description}]]></DESCRIPTION>\n`;
    xmlData += `    <ANNOTATION><![CDATA[${annotation}]]></ANNOTATION>\n`;
    xmlData += `    <VISIBLE>1</VISIBLE>\n`;
    xmlData += `    <ACTION></ACTION>\n`;
    xmlData += `    <HOMEPAGE></HOMEPAGE>\n`;
    xmlData += `    <NEW></NEW>\n`;
    xmlData += `    <TOP></TOP>\n`;
    xmlData += `    <STOCK></STOCK>\n`;
    xmlData += `    <WEIGHT>${weight.toFixed(3)}</WEIGHT>\n`;
    xmlData += `    <EAN></EAN>\n`;
    xmlData += `    <PRICE_WITH_VAT>${priceWithVat.toFixed(2)}</PRICE_WITH_VAT>\n`;
    xmlData += `    <VAT>12</VAT>\n`;
    xmlData += `    <WARRANTY>${warranty}</WARRANTY>\n`;
    xmlData += `    <CATEGORYID>1</CATEGORYID>\n`;
    xmlData += `    <IMAGES>\n`;
    xmlData += `      <IMGURL><![CDATA[${imageUrl}]]></IMGURL>\n`;
    xmlData += `    </IMAGES>\n`;
    xmlData += `    <RELATED>\n`;
    xmlData += `      <CODE></CODE>\n`;
    xmlData += `    </RELATED>\n`;
    xmlData += `    <ALTERNATIVE>\n`;
    xmlData += `      <CODE></CODE>\n`;
    xmlData += `    </ALTERNATIVE>\n`;
    xmlData += `    <CATEGOERIES>\n<CATEGORY>1</CATEGORY></CATEGOERIES>\n`;
    xmlData += `    <PARAMETERS>\n`;
    xmlData += `      <PARAM>\n`;
    xmlData += `        <NAME>Alergeny</NAME>\n`;
    xmlData += `        <VALUE><![CDATA[${allergens}]]></VALUE>\n`;
    xmlData += `        <VALUE_TEXT>Alergeny</VALUE_TEXT>\n`;
    xmlData += `      </PARAM>\n`;
    xmlData += `      <PARAM>\n`;
    xmlData += `        <NAME>Zamrazení</NAME>\n`;
    xmlData += `        <VALUE><![CDATA[${freezing}]]></VALUE>\n`;
    xmlData += `        <VALUE_TEXT>Zamrazení</VALUE_TEXT>\n`;
    xmlData += `      </PARAM>\n`;
    xmlData += `    </PARAMETERS>\n`;
    xmlData += `  </SHOPITEM>\n`;
});

xmlData += '</SHOP>\n';

fs.writeFileSync('./output/output.xml', xmlData);

console.log('XML file has been generated in the /output folder.');
