/**
 * Deep analysis of "Snow - Truss Spacing" sheet
 * from "AZ CO UT 1 5 26.xlsx"
 *
 * Goals:
 *  1. Read ALL formula cells in the sheet
 *  2. Trace F54 formula and every variable it depends on
 *  3. Document enclosure lookup table rows 53-74
 *  4. Document I106 from "Snow - Changers" and how it feeds E52
 *  5. Check if F52 is used elsewhere (second output?)
 *  6. Check formula cells around rows 44-55
 *  7. Determine if the enclosure table modifies truss spacing directly
 */

import pkg from 'xlsx';
const { readFile, utils } = pkg;

const FILE = 'C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx';

const wb = readFile(FILE, { cellFormula: true, cellStyles: true, sheetStubs: true });

// ─── Helper to get cell value + formula from any sheet ───
function cell(sheetName, ref) {
  const ws = wb.Sheets[sheetName];
  if (!ws) return { v: undefined, f: undefined, note: `Sheet "${sheetName}" not found` };
  const c = ws[ref];
  if (!c) return { v: undefined, f: undefined, note: `${ref} is empty` };
  return { v: c.v, f: c.f, t: c.t, w: c.w };
}

function cellVal(sheetName, ref) {
  const c = cell(sheetName, ref);
  return c.v;
}

function cellFormula(sheetName, ref) {
  const c = cell(sheetName, ref);
  return c.f;
}

// ─── 1. ALL formula cells in "Snow - Truss Spacing" ───
console.log('='.repeat(90));
console.log('1. ALL FORMULA CELLS in "Snow - Truss Spacing"');
console.log('='.repeat(90));

const SHEET = 'Snow - Truss Spacing';
const ws = wb.Sheets[SHEET];
if (!ws) {
  console.error(`Sheet "${SHEET}" not found! Available sheets:`, wb.SheetNames);
  process.exit(1);
}

const range = utils.decode_range(ws['!ref']);
const allFormulas = [];

for (let R = range.s.r; R <= range.e.r; R++) {
  for (let C = range.s.c; C <= range.e.c; C++) {
    const addr = utils.encode_cell({ r: R, c: C });
    const c = ws[addr];
    if (c && c.f) {
      allFormulas.push({ addr, row: R + 1, col: C, formula: c.f, value: c.v, type: c.t });
    }
  }
}

console.log(`\nTotal formula cells found: ${allFormulas.length}\n`);
allFormulas.forEach((f, i) => {
  console.log(`  [${i + 1}] ${f.addr} (row ${f.row})`);
  console.log(`      Formula: =${f.formula}`);
  console.log(`      Value:   ${f.value} (type: ${f.type})`);
  console.log();
});

// ─── 2. TRACE F54 and all its dependencies ───
console.log('='.repeat(90));
console.log('2. F54 FORMULA TRACE — Complete dependency chain');
console.log('='.repeat(90));

const traceRefs = [
  // Height / config
  { sheet: SHEET, ref: 'Q46', label: 'Height value' },
  { sheet: SHEET, ref: 'Q47', label: 'Height > 12 flag' },
  { sheet: SHEET, ref: 'Q48', label: 'Q48' },
  { sheet: SHEET, ref: 'Q49', label: 'Q49' },
  // Snow code / config rows 44-55
  { sheet: SHEET, ref: 'F44', label: 'F44' },
  { sheet: SHEET, ref: 'F45', label: 'F45' },
  { sheet: SHEET, ref: 'F46', label: 'F46' },
  { sheet: SHEET, ref: 'F47', label: 'F47' },
  { sheet: SHEET, ref: 'F48', label: 'F48' },
  { sheet: SHEET, ref: 'F49', label: 'F49' },
  { sheet: SHEET, ref: 'F50', label: 'F50' },
  { sheet: SHEET, ref: 'F51', label: 'F51 — raw INDEX result' },
  { sheet: SHEET, ref: 'F52', label: 'F52 — guard / cap' },
  { sheet: SHEET, ref: 'F53', label: 'F53' },
  { sheet: SHEET, ref: 'F54', label: 'F54 — FINAL adjusted spacing' },
  { sheet: SHEET, ref: 'F55', label: 'F55' },
  // E column around those rows
  { sheet: SHEET, ref: 'E49', label: 'E49' },
  { sheet: SHEET, ref: 'E50', label: 'E50' },
  { sheet: SHEET, ref: 'E51', label: 'E51' },
  { sheet: SHEET, ref: 'E52', label: 'E52 — F51 minus I106 adjustment' },
  { sheet: SHEET, ref: 'E53', label: 'E53' },
  { sheet: SHEET, ref: 'E54', label: 'E54' },
  // H51 and related
  { sheet: SHEET, ref: 'H51', label: 'H51 — irregular × enclosure' },
  { sheet: SHEET, ref: 'H49', label: 'H49' },
  { sheet: SHEET, ref: 'H50', label: 'H50' },
  { sheet: SHEET, ref: 'H52', label: 'H52' },
  // Enclosure factor
  { sheet: SHEET, ref: 'S73', label: 'S73 — enclosure flag/code' },
  { sheet: SHEET, ref: 'S74', label: 'S74 — enclosure factor' },
  // G column
  { sheet: SHEET, ref: 'G49', label: 'G49' },
  { sheet: SHEET, ref: 'G50', label: 'G50' },
  { sheet: SHEET, ref: 'G51', label: 'G51' },
  { sheet: SHEET, ref: 'G52', label: 'G52' },
  // Also check D column
  { sheet: SHEET, ref: 'D49', label: 'D49' },
  { sheet: SHEET, ref: 'D50', label: 'D50' },
  { sheet: SHEET, ref: 'D51', label: 'D51' },
  { sheet: SHEET, ref: 'D52', label: 'D52' },
  // I106 from Snow - Changers
  { sheet: 'Snow - Changers', ref: 'I106', label: 'I106 (Snow - Changers) — adjustment' },
  { sheet: 'Snow - Changers', ref: 'I107', label: 'I107 (Snow - Changers)' },
  { sheet: 'Snow - Changers', ref: 'I105', label: 'I105 (Snow - Changers)' },
];

console.log();
for (const { sheet, ref, label } of traceRefs) {
  const c = cell(sheet, ref);
  const fStr = c.f ? `=${c.f}` : '(no formula)';
  console.log(`  ${ref.padEnd(5)} [${sheet}]  ${label}`);
  console.log(`         Formula: ${fStr}`);
  console.log(`         Value:   ${c.v}  (type: ${c.t || 'n/a'})`);
  console.log();
}

// ─── 3. Enclosure lookup table rows 53-74 ───
console.log('='.repeat(90));
console.log('3. ENCLOSURE LOOKUP TABLE — Rows 53-74 (columns A through T)');
console.log('='.repeat(90));

// First read the header row to understand columns
console.log('\n  Header labels (row 53 or nearby):');
for (let R = 52; R <= 54; R++) {
  const rowData = [];
  for (let C = 0; C <= 19; C++) { // A-T
    const addr = utils.encode_cell({ r: R, c: C });
    const c = ws[addr];
    rowData.push(c ? (c.w || String(c.v)) : '');
  }
  console.log(`    Row ${R + 1}: ${rowData.join(' | ')}`);
}

console.log('\n  Full table data:');
for (let R = 52; R <= 74; R++) {
  const rowData = [];
  for (let C = 0; C <= 19; C++) { // A-T
    const addr = utils.encode_cell({ r: R, c: C });
    const c = ws[addr];
    const val = c ? (c.f ? `[=${c.f}]→${c.v}` : (c.w || String(c.v))) : '';
    rowData.push(val.toString().substring(0, 18).padEnd(18));
  }
  console.log(`  Row ${R + 1}: ${rowData.join('|')}`);
}

// Check for formulas in this region
console.log('\n  Formula cells in rows 53-74:');
for (let R = 52; R <= 74; R++) {
  for (let C = 0; C <= 25; C++) {
    const addr = utils.encode_cell({ r: R, c: C });
    const c = ws[addr];
    if (c && c.f) {
      console.log(`    ${addr} (row ${R + 1}): =${c.f}  →  ${c.v}`);
    }
  }
}

// ─── 4. Snow - Changers I106 context ───
console.log('\n' + '='.repeat(90));
console.log('4. "Snow - Changers" I106 and surrounding context');
console.log('='.repeat(90));

const csWs = wb.Sheets['Snow - Changers'];
if (csWs) {
  console.log('\n  Rows 100-115, columns A-L from Snow - Changers:');
  for (let R = 99; R <= 114; R++) {
    const rowData = [];
    for (let C = 0; C <= 11; C++) { // A-L
      const addr = utils.encode_cell({ r: R, c: C });
      const c = csWs[addr];
      const val = c ? (c.f ? `[=${c.f}]→${c.v}` : (c.w || String(c.v))) : '';
      rowData.push(val.toString().substring(0, 22).padEnd(22));
    }
    console.log(`  Row ${R + 1}: ${rowData.join('|')}`);
  }

  // All formulas in I column of Snow - Changers
  console.log('\n  All formula cells in column I of Snow - Changers:');
  const csRange = utils.decode_range(csWs['!ref']);
  for (let R = csRange.s.r; R <= csRange.e.r; R++) {
    const addr = utils.encode_cell({ r: R, c: 8 }); // I = col 8
    const c = csWs[addr];
    if (c && c.f) {
      console.log(`    ${addr} (row ${R + 1}): =${c.f}  →  ${c.v}`);
    }
  }
} else {
  console.log('  Sheet "Snow - Changers" NOT FOUND');
}

// ─── 5. Check if F52 is referenced by other cells ───
console.log('\n' + '='.repeat(90));
console.log('5. WHERE IS F52 REFERENCED? (second output check)');
console.log('='.repeat(90));

// Search all formulas in this sheet for references to F52
console.log('\n  Formulas in "Snow - Truss Spacing" referencing F52:');
for (const f of allFormulas) {
  if (f.formula.includes('F52') && f.addr !== 'F52') {
    console.log(`    ${f.addr}: =${f.formula}  →  ${f.value}`);
  }
}

// Also check other sheets for references to this sheet's F52
console.log('\n  Formulas in OTHER sheets referencing Snow - Truss Spacing F52 or F54:');
for (const sheetName of wb.SheetNames) {
  if (sheetName === SHEET) continue;
  const sws = wb.Sheets[sheetName];
  const sRange = utils.decode_range(sws['!ref'] || 'A1');
  for (let R = sRange.s.r; R <= sRange.e.r; R++) {
    for (let C = sRange.s.c; C <= sRange.e.c; C++) {
      const addr = utils.encode_cell({ r: R, c: C });
      const c = sws[addr];
      if (c && c.f && (c.f.includes('Truss') || c.f.includes('truss'))) {
        if (c.f.includes('F52') || c.f.includes('F54') || c.f.includes('F51') || c.f.includes('E52')) {
          console.log(`    [${sheetName}] ${addr}: =${c.f.substring(0, 120)}  →  ${c.v}`);
        }
      }
    }
  }
}

// Also search for cross-sheet refs with the short pattern
console.log('\n  Searching all sheets for any reference containing "Truss Spacing":');
for (const sheetName of wb.SheetNames) {
  if (sheetName === SHEET) continue;
  const sws = wb.Sheets[sheetName];
  const sRange = utils.decode_range(sws['!ref'] || 'A1');
  let count = 0;
  for (let R = sRange.s.r; R <= sRange.e.r; R++) {
    for (let C = sRange.s.c; C <= sRange.e.c; C++) {
      const addr = utils.encode_cell({ r: R, c: C });
      const c = sws[addr];
      if (c && c.f && c.f.includes('Truss Spacing')) {
        if (count < 15) {
          console.log(`    [${sheetName}] ${addr}: =${c.f.substring(0, 150)}  →  ${c.v}`);
        }
        count++;
      }
    }
  }
  if (count > 15) console.log(`    ... and ${count - 15} more in ${sheetName}`);
  if (count > 0 && count <= 15) { /* already printed */ }
}

// ─── 6. Rows 44-55 comprehensive dump ───
console.log('\n' + '='.repeat(90));
console.log('6. ROWS 44-55 COMPREHENSIVE — All columns A-T with formulas');
console.log('='.repeat(90));

for (let R = 43; R <= 55; R++) {
  console.log(`\n  --- Row ${R + 1} ---`);
  for (let C = 0; C <= 19; C++) {
    const addr = utils.encode_cell({ r: R, c: C });
    const c = ws[addr];
    if (c) {
      const fStr = c.f ? ` [FORMULA: =${c.f}]` : '';
      console.log(`    ${addr}: ${c.w || c.v}${fStr}`);
    }
  }
}

// ─── 7. Does enclosure table modify truss spacing directly? ───
console.log('\n' + '='.repeat(90));
console.log('7. ENCLOSURE TABLE ANALYSIS — Does it modify truss spacing directly?');
console.log('='.repeat(90));

// Check S73, S74 and any cells that reference them
console.log('\n  S73 and S74 details:');
const s73 = cell(SHEET, 'S73');
const s74 = cell(SHEET, 'S74');
console.log(`    S73: formula=${s73.f ? '=' + s73.f : 'none'}, value=${s73.v}`);
console.log(`    S74: formula=${s74.f ? '=' + s74.f : 'none'}, value=${s74.v}`);

// What references S73/S74?
console.log('\n  Cells referencing S73:');
for (const f of allFormulas) {
  if (f.formula.includes('S73')) {
    console.log(`    ${f.addr}: =${f.formula}  →  ${f.value}`);
  }
}
console.log('\n  Cells referencing S74:');
for (const f of allFormulas) {
  if (f.formula.includes('S74') || f.formula.includes('$S$74')) {
    console.log(`    ${f.addr}: =${f.formula}  →  ${f.value}`);
  }
}

// Check what H51 references
console.log('\n  H51 detail:');
const h51 = cell(SHEET, 'H51');
console.log(`    H51: formula=${h51.f ? '=' + h51.f : 'none'}, value=${h51.v}`);

// Trace: what does H51 feed into?
console.log('\n  Cells referencing H51:');
for (const f of allFormulas) {
  if (f.formula.includes('H51') || f.formula.includes('$H$51')) {
    console.log(`    ${f.addr}: =${f.formula}  →  ${f.value}`);
  }
}

// Check what F51 references and what it feeds
console.log('\n  F51 detail:');
const f51 = cell(SHEET, 'F51');
console.log(`    F51: formula=${f51.f ? '=' + f51.f : 'none'}, value=${f51.v}`);

console.log('\n  Cells referencing F51:');
for (const f of allFormulas) {
  if ((f.formula.includes('F51') || f.formula.includes('$F$51')) && f.addr !== 'F51') {
    console.log(`    ${f.addr}: =${f.formula}  →  ${f.value}`);
  }
}

// Check E52
console.log('\n  E52 detail:');
const e52 = cell(SHEET, 'E52');
console.log(`    E52: formula=${e52.f ? '=' + e52.f : 'none'}, value=${e52.v}`);

console.log('\n  Cells referencing E52:');
for (const f of allFormulas) {
  if ((f.formula.includes('E52') || f.formula.includes('$E$52')) && f.addr !== 'E52') {
    console.log(`    ${f.addr}: =${f.formula}  →  ${f.value}`);
  }
}

// ─── BONUS: Check rows 1-10 for any header/label info ───
console.log('\n' + '='.repeat(90));
console.log('BONUS: Sheet layout — rows 1-10 and named ranges');
console.log('='.repeat(90));

for (let R = 0; R <= 9; R++) {
  const rowData = [];
  for (let C = 0; C <= 19; C++) {
    const addr = utils.encode_cell({ r: R, c: C });
    const c = ws[addr];
    if (c) rowData.push(`${addr}=${c.w || c.v}`);
  }
  if (rowData.length) console.log(`  Row ${R + 1}: ${rowData.join(', ')}`);
}

// Named ranges that reference this sheet
if (wb.Workbook && wb.Workbook.Names) {
  console.log('\n  Named ranges referencing Truss Spacing:');
  for (const n of wb.Workbook.Names) {
    if (n.Ref && n.Ref.includes('Truss')) {
      console.log(`    ${n.Name} = ${n.Ref}`);
    }
  }
}

// ─── BONUS 2: Full formula list for Snow - Changers too ───
console.log('\n' + '='.repeat(90));
console.log('BONUS 2: ALL formulas in Snow - Changers referencing Truss Spacing cells');
console.log('='.repeat(90));

if (csWs) {
  const csRange2 = utils.decode_range(csWs['!ref']);
  for (let R = csRange2.s.r; R <= csRange2.e.r; R++) {
    for (let C = csRange2.s.c; C <= csRange2.e.c; C++) {
      const addr = utils.encode_cell({ r: R, c: C });
      const c = csWs[addr];
      if (c && c.f && c.f.includes('Truss')) {
        console.log(`  ${addr}: =${c.f.substring(0, 200)}  →  ${c.v}`);
      }
    }
  }
}

console.log('\n' + '='.repeat(90));
console.log('DONE');
console.log('='.repeat(90));
