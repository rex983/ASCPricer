const XLSX = require('xlsx');
const path = require('path');

const filePath = path.resolve('C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx');
const workbook = XLSX.readFile(filePath);

const sheetName = workbook.SheetNames.find(n => n.includes('Snow') && n.includes('Math'));
if (!sheetName) {
  console.log('Sheet not found. Available sheets:', workbook.SheetNames);
  process.exit(1);
}

console.log(`Sheet: "${sheetName}"`);
const sheet = workbook.Sheets[sheetName];
const range = XLSX.utils.decode_range(sheet['!ref']);
console.log(`Range: ${sheet['!ref']} (rows ${range.s.r + 1}-${range.e.r + 1}, cols ${range.s.c + 1}-${range.e.c + 1})`);
console.log('');

// Print header row with column letters
const colLetters = [];
for (let c = range.s.c; c <= range.e.c; c++) {
  colLetters.push(XLSX.utils.encode_col(c));
}
console.log('ROW\t' + colLetters.join('\t'));
console.log('---\t' + colLetters.map(() => '---').join('\t'));

// Print every row
for (let r = range.s.r; r <= range.e.r; r++) {
  const vals = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = sheet[addr];
    if (cell) {
      // Show value (v) and formula (f) if present
      let display = cell.v !== undefined ? String(cell.v) : '';
      if (cell.f) display += ` [F:${cell.f}]`;
      vals.push(display);
    } else {
      vals.push('');
    }
  }
  console.log(`${r + 1}\t${vals.join('\t')}`);
}

// Also dump every non-empty cell individually for completeness
console.log('\n\n=== INDIVIDUAL CELL DUMP ===\n');
for (let r = range.s.r; r <= range.e.r; r++) {
  for (let c = range.s.c; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = sheet[addr];
    if (cell) {
      const parts = [`${addr} (R${r+1}C${c+1}): type=${cell.t}, value=${JSON.stringify(cell.v)}`];
      if (cell.f) parts.push(`formula=${cell.f}`);
      if (cell.w) parts.push(`formatted=${cell.w}`);
      console.log(parts.join(', '));
    }
  }
}
