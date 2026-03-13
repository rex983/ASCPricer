import ExcelJS from 'exceljs';

const FILE = 'C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx';
const SHEET_NAME = 'Snow - Math Calculations';

async function main() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(FILE);

  const ws = wb.getWorksheet(SHEET_NAME);
  if (!ws) {
    console.error(`Sheet "${SHEET_NAME}" not found. Available sheets:`);
    wb.eachSheet((s) => console.error(`  - "${s.name}"`));
    process.exit(1);
  }

  console.log(`=== Sheet: "${ws.name}" ===`);
  console.log(`Dimensions: ${ws.dimensions?.toString() || 'unknown'}\n`);

  // Helper: get cell info
  function cellInfo(addr) {
    const cell = ws.getCell(addr);
    const formula = cell.formula || cell.sharedFormula || null;
    const val = cell.value;
    let displayVal = val;
    // Handle rich values
    if (val && typeof val === 'object') {
      if (val.formula) displayVal = `[formula result: ${val.result}]`;
      else if (val.sharedFormula) displayVal = `[shared formula result: ${val.result}]`;
      else if (val.richText) displayVal = val.richText.map(r => r.text).join('');
      else displayVal = JSON.stringify(val);
    }
    return { addr, formula, value: displayVal, rawValue: val };
  }

  // Helper: print cell
  function printCell(addr) {
    const info = cellInfo(addr);
    const formulaStr = info.formula ? `  FORMULA: =${info.formula}` : '';
    const result = info.rawValue?.result !== undefined ? `  RESULT: ${info.rawValue.result}` : '';
    if (info.value !== null && info.value !== undefined && info.value !== '') {
      console.log(`  ${info.addr}: value=${JSON.stringify(info.value)}${formulaStr}${result}`);
    }
    return info;
  }

  // Convert col number to letter
  function colLetter(n) {
    let s = '';
    while (n > 0) {
      n--;
      s = String.fromCharCode(65 + (n % 26)) + s;
      n = Math.floor(n / 26);
    }
    return s;
  }

  // =====================================================
  // SECTION 0: FULL GRID SCAN — Every cell A1 to AZ50
  // =====================================================
  console.log('╔══════════════════════════════════════════════════════════════╗');
  console.log('║  COMPLETE CELL SCAN: All cells A1 through AZ50             ║');
  console.log('╚══════════════════════════════════════════════════════════════╝\n');

  const maxCol = 52; // AZ = col 52
  const maxRow = 50;

  // First pass: collect ALL non-empty cells
  const allCells = [];
  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const addr = `${colLetter(c)}${r}`;
      const cell = ws.getCell(addr);
      const val = cell.value;
      if (val !== null && val !== undefined && val !== '') {
        const formula = cell.formula || cell.sharedFormula || null;
        let display = val;
        if (val && typeof val === 'object') {
          if (val.formula) display = `[F] result=${val.result}`;
          else if (val.sharedFormula) display = `[SF] result=${val.result}`;
          else if (val.richText) display = val.richText.map(x => x.text).join('');
          else display = JSON.stringify(val);
        }
        allCells.push({ addr, row: r, col: c, formula, value: display, raw: val });
      }
    }
  }

  console.log(`Total non-empty cells found: ${allCells.length}\n`);

  // Print ALL formula cells first
  console.log('─── ALL FORMULA CELLS ───');
  const formulaCells = allCells.filter(c => c.formula);
  formulaCells.forEach(c => {
    const result = c.raw?.result !== undefined ? c.raw.result : c.value;
    console.log(`  ${c.addr}: =${c.formula}  →  ${JSON.stringify(result)}`);
  });
  console.log(`\nTotal formula cells: ${formulaCells.length}\n`);

  // Print ALL non-formula cells with values
  console.log('─── ALL VALUE (non-formula) CELLS ───');
  const valueCells = allCells.filter(c => !c.formula);
  valueCells.forEach(c => {
    console.log(`  ${c.addr}: ${JSON.stringify(c.value)}`);
  });
  console.log(`\nTotal value cells: ${valueCells.length}\n`);

  // =====================================================
  // SECTION 1: HAT CHANNEL CALCULATION CHAIN (rows 9-17)
  // =====================================================
  console.log('╔══════════════════════════════════════════════════════════════╗');
  console.log('║  SECTION 1: HAT CHANNEL CALCULATION (Rows 9-17)           ║');
  console.log('╚══════════════════════════════════════════════════════════════╝\n');

  for (let r = 8; r <= 17; r++) {
    console.log(`--- Row ${r} ---`);
    for (let c = 1; c <= maxCol; c++) {
      printCell(`${colLetter(c)}${r}`);
    }
  }

  // =====================================================
  // SECTION 2: GIRT CALCULATION CHAIN (rows 19-26)
  // =====================================================
  console.log('\n╔══════════════════════════════════════════════════════════════╗');
  console.log('║  SECTION 2: GIRT CALCULATION (Rows 19-26)                 ║');
  console.log('╚══════════════════════════════════════════════════════════════╝\n');

  for (let r = 18; r <= 27; r++) {
    console.log(`--- Row ${r} ---`);
    for (let c = 1; c <= maxCol; c++) {
      printCell(`${colLetter(c)}${r}`);
    }
  }

  // =====================================================
  // SECTION 3: VERTICAL CALCULATION CHAIN (rows 28-35)
  // =====================================================
  console.log('\n╔══════════════════════════════════════════════════════════════╗');
  console.log('║  SECTION 3: VERTICAL/COLUMN CALCULATION (Rows 28-35)      ║');
  console.log('╚══════════════════════════════════════════════════════════════╝\n');

  for (let r = 27; r <= 36; r++) {
    console.log(`--- Row ${r} ---`);
    for (let c = 1; c <= maxCol; c++) {
      printCell(`${colLetter(c)}${r}`);
    }
  }

  // =====================================================
  // SECTION 4: TRUSS PRICING CHAIN (rows 1-7, P10-P22)
  // =====================================================
  console.log('\n╔══════════════════════════════════════════════════════════════╗');
  console.log('║  SECTION 4: TRUSS PRICING (Rows 1-7, P10-P22)            ║');
  console.log('╚══════════════════════════════════════════════════════════════╝\n');

  console.log('--- Header Rows 1-7 ---');
  for (let r = 1; r <= 7; r++) {
    console.log(`--- Row ${r} ---`);
    for (let c = 1; c <= maxCol; c++) {
      printCell(`${colLetter(c)}${r}`);
    }
  }

  console.log('\n--- Truss Pricing P10-P22 ---');
  for (let r = 10; r <= 22; r++) {
    for (let c = 15; c <= 26; c++) { // O through Z
      printCell(`${colLetter(c)}${r}`);
    }
  }

  // =====================================================
  // SECTION 5: GIRT PRICING CHAIN (T25-T30)
  // =====================================================
  console.log('\n╔══════════════════════════════════════════════════════════════╗');
  console.log('║  SECTION 5: GIRT PRICING (T25-T35)                        ║');
  console.log('╚══════════════════════════════════════════════════════════════╝\n');

  for (let r = 25; r <= 35; r++) {
    for (let c = 19; c <= 30; c++) { // S through AD
      printCell(`${colLetter(c)}${r}`);
    }
  }

  // =====================================================
  // SECTION 6: FINAL OUTPUT CELLS (AC, AD, AE columns)
  // =====================================================
  console.log('\n╔══════════════════════════════════════════════════════════════╗');
  console.log('║  SECTION 6: FINAL OUTPUT CELLS (AC, AD, AE)               ║');
  console.log('╚══════════════════════════════════════════════════════════════╝\n');

  for (let r = 1; r <= 35; r++) {
    for (const col of ['AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH']) {
      printCell(`${col}${r}`);
    }
  }

  // =====================================================
  // SECTION 7: Deep dive on specific cells of interest
  // =====================================================
  console.log('\n╔══════════════════════════════════════════════════════════════╗');
  console.log('║  SECTION 7: SPECIFIC CELLS DEEP DIVE                      ║');
  console.log('╚══════════════════════════════════════════════════════════════╝\n');

  const specificCells = [
    'H20', 'H21', 'H22', 'H23', 'H24',
    'U18', 'U19', 'U20',
    'I34', 'I35',
    'P22', 'X5', 'X6', 'X7',
    'AC14', 'AC19', 'AD20',
    'AE14', 'AE15', 'AE16', 'AE17', 'AE18', 'AE19',
  ];

  for (const addr of specificCells) {
    const cell = ws.getCell(addr);
    const formula = cell.formula || cell.sharedFormula || null;
    const val = cell.value;
    let display = val;
    if (val && typeof val === 'object') {
      if (val.formula) display = `result=${val.result}`;
      else if (val.richText) display = val.richText.map(x => x.text).join('');
      else display = JSON.stringify(val);
    }
    console.log(`  ${addr}:`);
    console.log(`    Value: ${JSON.stringify(display)}`);
    if (formula) console.log(`    Formula: =${formula}`);
    if (val?.result !== undefined) console.log(`    Result: ${val.result}`);
    console.log();
  }

  // =====================================================
  // SECTION 8: Also check for merged cells and named ranges
  // =====================================================
  console.log('\n╔══════════════════════════════════════════════════════════════╗');
  console.log('║  SECTION 8: MERGED CELLS & NAMED RANGES                   ║');
  console.log('╚══════════════════════════════════════════════════════════════╝\n');

  // Check merged cells
  if (ws.model?.merges?.length) {
    console.log('Merged cell ranges:');
    ws.model.merges.forEach(m => console.log(`  ${m}`));
  } else {
    console.log('No merged cells found.');
  }

  // Check defined names
  if (wb.definedNames) {
    console.log('\nDefined Names (that reference this sheet):');
    // Try to enumerate
    try {
      const names = wb.definedNames.model;
      if (names && Array.isArray(names)) {
        names.forEach(n => {
          if (n.ranges?.some(r => r.includes(SHEET_NAME) || r.includes('Snow'))) {
            console.log(`  ${n.name}: ${JSON.stringify(n.ranges)}`);
          }
        });
      }
    } catch (e) {
      console.log('  (could not enumerate defined names)');
    }
  }

  // =====================================================
  // SECTION 9: Raw dump of ExcelJS cell objects for key cells
  // =====================================================
  console.log('\n╔══════════════════════════════════════════════════════════════╗');
  console.log('║  SECTION 9: RAW CELL OBJECTS (key cells)                   ║');
  console.log('╚══════════════════════════════════════════════════════════════╝\n');

  const keyCells = [
    'A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1',
    'A9', 'B9', 'C9', 'D9', 'E9', 'F9', 'G9', 'H9', 'I9',
    'A10', 'B10', 'C10', 'D10', 'E10', 'F10', 'G10', 'H10', 'I10',
    'P10', 'P11', 'P12', 'P13', 'P14', 'P15', 'P16', 'P17', 'P18', 'P19', 'P20', 'P21', 'P22',
    'T25', 'T26', 'T27', 'T28', 'T29', 'T30',
  ];

  for (const addr of keyCells) {
    const cell = ws.getCell(addr);
    console.log(`  ${addr}: type=${cell.type}, formula=${cell.formula || 'none'}, value=${JSON.stringify(cell.value)}`);
  }

  // =====================================================
  // SECTION 10: Extended scan — rows up to 50, cols up to AZ
  // =====================================================
  console.log('\n╔══════════════════════════════════════════════════════════════╗');
  console.log('║  SECTION 10: EXTENDED SCAN (rows 36-50)                    ║');
  console.log('╚══════════════════════════════════════════════════════════════╝\n');

  for (let r = 36; r <= 50; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const addr = `${colLetter(c)}${r}`;
      const cell = ws.getCell(addr);
      const val = cell.value;
      if (val !== null && val !== undefined && val !== '') {
        const formula = cell.formula || cell.sharedFormula || null;
        let display = val;
        if (val && typeof val === 'object') {
          if (val.formula) display = `[F] result=${val.result}`;
          else if (val.richText) display = val.richText.map(x => x.text).join('');
          else display = JSON.stringify(val);
        }
        const fStr = formula ? ` FORMULA: =${formula}` : '';
        console.log(`  ${addr}: ${JSON.stringify(display)}${fStr}`);
      }
    }
  }

  console.log('\n=== ANALYSIS COMPLETE ===');
}

main().catch(e => { console.error(e); process.exit(1); });
