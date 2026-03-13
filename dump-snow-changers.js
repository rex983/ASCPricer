const XLSX = require('xlsx');

const wb = XLSX.readFile('C:\\Users\\Redir\\Downloads\\AZ CO UT 1 5 26.xlsx');

const sheetName = 'Snow - Changers';
const ws = wb.Sheets[sheetName];

if (!ws) {
  console.log('Sheet not found! Available sheets:');
  wb.SheetNames.forEach((n, i) => console.log(`  ${i}: "${n}"`));
  process.exit(1);
}

console.log(`Sheet: "${sheetName}"`);
console.log(`Range: ${ws['!ref']}`);
console.log('');

const range = XLSX.utils.decode_range(ws['!ref']);

for (let r = range.s.r; r <= range.e.r; r++) {
  const cells = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = ws[addr];
    if (cell !== undefined) {
      cells.push(`[${XLSX.utils.encode_col(c)}] ${cell.v}`);
    }
  }
  if (cells.length > 0) {
    console.log(`Row ${r}: ${cells.join(' | ')}`);
  }
}
