/**
 * Deep analysis of "Snow - Diagonal Bracing" sheet in AZ CO UT 1 5 26.xlsx
 *
 * Exhaustively documents:
 * 1. Every cell (value and formula)
 * 2. Complete structure and calculation flow
 * 3. How diagonal bracing need is determined
 * 4. How cost is calculated
 * 5. Inputs from other sheets
 * 6. Output cells read by other sheets
 * 7. J9 trace ("bracing needed" flag from Snow - Changers D115)
 * 8. K10 trace (output read by Quote Sheet Z27)
 * 9. State-specific wind thresholds
 * 10. Cost vs building dimensions
 * 11. Interaction with snow engineering total
 */

import XLSX from 'xlsx';

const FILE = 'C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx';
const wb = XLSX.readFile(FILE);

const SHEET_NAME = 'Snow - Diagonal Bracing';
const ws = wb.Sheets[SHEET_NAME];

if (!ws) {
  console.error(`Sheet "${SHEET_NAME}" not found!`);
  console.log('Available sheets:', wb.SheetNames.join(', '));
  process.exit(1);
}

function sep(title) {
  console.log('\n' + '='.repeat(90));
  console.log(`  ${title}`);
  console.log('='.repeat(90));
}

function subsep(title) {
  console.log('\n' + '-'.repeat(70));
  console.log(`  ${title}`);
  console.log('-'.repeat(70));
}

// ============================================================================
// 1. RAW SHEET DUMP - EVERY CELL WITH VALUE AND FORMULA
// ============================================================================
sep('1. COMPLETE CELL DUMP (value + formula)');

const range = XLSX.utils.decode_range(ws['!ref']);
console.log(`\nSheet range: ${ws['!ref']}`);
console.log(`Rows: ${range.s.r} to ${range.e.r} (${range.e.r - range.s.r + 1} total)`);
console.log(`Cols: ${range.s.c} to ${range.e.c} (${range.e.c - range.s.c + 1} total)`);

// Collect all non-empty cells
const allCells = [];
for (let r = range.s.r; r <= range.e.r; r++) {
  for (let c = range.s.c; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = ws[addr];
    if (cell) {
      allCells.push({
        addr,
        row: r,
        col: c,
        colLetter: XLSX.utils.encode_col(c),
        value: cell.v,
        type: cell.t,
        formula: cell.f || null,
        formatted: cell.w || null,
      });
    }
  }
}

console.log(`\nTotal non-empty cells: ${allCells.length}`);
console.log('\n--- ALL CELLS ---');
for (const c of allCells) {
  const fStr = c.formula ? `  FORMULA: =${c.formula}` : '';
  const fmtStr = c.formatted ? ` (fmt: ${c.formatted})` : '';
  console.log(`  ${c.addr}: [${c.type}] ${JSON.stringify(c.value)}${fmtStr}${fStr}`);
}

// ============================================================================
// 2. GRID VIEW (values)
// ============================================================================
sep('2. GRID VIEW (values only)');

const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
for (let i = 0; i < data.length; i++) {
  const row = data[i];
  if (row.some(v => v !== '' && v !== undefined && v !== null)) {
    const cols = row.map((v, j) => {
      const colLetter = XLSX.utils.encode_col(j);
      if (v === '' || v === undefined || v === null) return null;
      return `${colLetter}=${JSON.stringify(v)}`;
    }).filter(Boolean);
    console.log(`  Row ${i + 1}: ${cols.join(' | ')}`);
  }
}

// ============================================================================
// 3. FORMULA-ONLY VIEW (grouped by pattern)
// ============================================================================
sep('3. ALL FORMULAS (grouped)');

const formulaCells = allCells.filter(c => c.formula);
console.log(`\nTotal cells with formulas: ${formulaCells.length}`);

// Group by formula pattern
const formulaPatterns = {};
for (const c of formulaCells) {
  // Normalize formula to find patterns
  const normalized = c.formula
    .replace(/\$?[A-Z]+\$?\d+/g, 'REF')
    .replace(/\d+(\.\d+)?/g, 'NUM');
  if (!formulaPatterns[normalized]) formulaPatterns[normalized] = [];
  formulaPatterns[normalized].push(c);
}

console.log(`\nUnique formula patterns: ${Object.keys(formulaPatterns).length}`);
for (const [pattern, cells] of Object.entries(formulaPatterns)) {
  console.log(`\n  Pattern: ${pattern}`);
  for (const c of cells) {
    console.log(`    ${c.addr}: =${c.formula}  => ${JSON.stringify(c.value)}`);
  }
}

// ============================================================================
// 4. CROSS-SHEET REFERENCES (formulas referencing other sheets)
// ============================================================================
sep('4. CROSS-SHEET REFERENCES');

const crossRefs = formulaCells.filter(c => c.formula.includes("'") || c.formula.includes('!'));
console.log(`\nCells referencing other sheets: ${crossRefs.length}`);

// Parse sheet references
const sheetRefsMap = {};
for (const c of crossRefs) {
  // Extract sheet references like 'Sheet Name'!Cell or SheetName!Cell
  const matches = c.formula.matchAll(/'([^']+)'!([A-Z$]+[0-9$]+(?::[A-Z$]+[0-9$]+)?)/g);
  for (const m of matches) {
    const refSheet = m[1];
    const refCell = m[2];
    if (!sheetRefsMap[refSheet]) sheetRefsMap[refSheet] = [];
    sheetRefsMap[refSheet].push({ from: c.addr, toCell: refCell, formula: c.formula, value: c.value });
  }
  // Also non-quoted sheet refs
  const matches2 = c.formula.matchAll(/([A-Za-z_]\w*)!([A-Z$]+[0-9$]+(?::[A-Z$]+[0-9$]+)?)/g);
  for (const m of matches2) {
    const refSheet = m[1];
    const refCell = m[2];
    if (refSheet === 'IF' || refSheet === 'OR' || refSheet === 'AND' || refSheet === 'SUM' || refSheet === 'VLOOKUP' || refSheet === 'INDEX' || refSheet === 'MATCH') continue;
    if (!sheetRefsMap[refSheet]) sheetRefsMap[refSheet] = [];
    sheetRefsMap[refSheet].push({ from: c.addr, toCell: refCell, formula: c.formula, value: c.value });
  }
}

for (const [sheet, refs] of Object.entries(sheetRefsMap)) {
  console.log(`\n  From sheet "${sheet}":`);
  for (const r of refs) {
    console.log(`    ${r.from} reads ${sheet}!${r.toCell}`);
    console.log(`      Formula: =${r.formula}`);
    console.log(`      Current value: ${JSON.stringify(r.value)}`);
  }
}

// ============================================================================
// 5. TRACE J9 - "bracing needed" flag
// ============================================================================
sep('5. TRACE J9 - "Bracing Needed" flag');

const j9 = ws['J9'];
if (j9) {
  console.log(`\n  J9 value: ${JSON.stringify(j9.v)}`);
  console.log(`  J9 type: ${j9.t}`);
  console.log(`  J9 formula: ${j9.f ? '=' + j9.f : '(no formula - static value)'}`);
  console.log(`  J9 formatted: ${j9.w}`);
} else {
  console.log('\n  J9 is EMPTY');
}

// Check nearby cells for context
subsep('J9 neighborhood (rows 1-15, cols I-L)');
for (let r = 0; r < 15; r++) {
  for (let c = 8; c <= 11; c++) { // I=8, J=9, K=10, L=11
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = ws[addr];
    if (cell) {
      const fStr = cell.f ? `  FORMULA: =${cell.f}` : '';
      console.log(`  ${addr}: [${cell.t}] ${JSON.stringify(cell.v)}${fStr}`);
    }
  }
}

// Now trace what Snow - Changers D115 contains
subsep('Tracing Snow - Changers D115');
const changersSheet = wb.Sheets['Snow - Changers'];
if (changersSheet) {
  const d115 = changersSheet['D115'];
  if (d115) {
    console.log(`  Snow - Changers D115 value: ${JSON.stringify(d115.v)}`);
    console.log(`  D115 type: ${d115.t}`);
    console.log(`  D115 formula: ${d115.f ? '=' + d115.f : '(no formula - static)'}`);
  } else {
    console.log('  D115 is EMPTY in Snow - Changers');
  }

  // Check neighboring cells in Changers around D115
  console.log('\n  Snow - Changers context around D115 (rows 110-120):');
  for (let r = 109; r < 120; r++) {
    for (let c = 0; c <= 5; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = changersSheet[addr];
      if (cell) {
        const fStr = cell.f ? `  FORMULA: =${cell.f}` : '';
        console.log(`    ${addr}: [${cell.t}] ${JSON.stringify(cell.v)}${fStr}`);
      }
    }
  }
} else {
  console.log('  Sheet "Snow - Changers" not found');
}

// ============================================================================
// 6. TRACE K10 - output read by Quote Sheet Z27
// ============================================================================
sep('6. TRACE K10 - Output read by Quote Sheet Z27');

const k10 = ws['K10'];
if (k10) {
  console.log(`\n  K10 value: ${JSON.stringify(k10.v)}`);
  console.log(`  K10 type: ${k10.t}`);
  console.log(`  K10 formula: ${k10.f ? '=' + k10.f : '(no formula - static value)'}`);
  console.log(`  K10 formatted: ${k10.w}`);
} else {
  console.log('\n  K10 is EMPTY');
}

// Now verify Quote Sheet Z27
subsep('Verifying Quote Sheet Z27');
const quoteSheet = wb.Sheets['Quote Sheet'];
if (quoteSheet) {
  const z27 = quoteSheet['Z27'];
  if (z27) {
    console.log(`  Quote Sheet Z27 value: ${JSON.stringify(z27.v)}`);
    console.log(`  Z27 type: ${z27.t}`);
    console.log(`  Z27 formula: ${z27.f ? '=' + z27.f : '(static)'}`);
    console.log(`  Z27 formatted: ${z27.w}`);
  } else {
    console.log('  Z27 is EMPTY in Quote Sheet');
  }

  // Check neighboring cells in Quote Sheet around Z27
  console.log('\n  Quote Sheet context around Z27 (rows 25-30, cols X-AB):');
  for (let r = 24; r < 30; r++) {
    for (let c = 23; c <= 27; c++) { // X=23, Y=24, Z=25, AA=26, AB=27
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = quoteSheet[addr];
      if (cell) {
        const fStr = cell.f ? `  FORMULA: =${cell.f}` : '';
        console.log(`    ${addr}: [${cell.t}] ${JSON.stringify(cell.v)}${fStr}`);
      }
    }
  }
} else {
  console.log('  Sheet "Quote Sheet" not found');
}

// ============================================================================
// 7. DETERMINE: How is diagonal bracing NEED decided?
// ============================================================================
sep('7. ANALYSIS: How is diagonal bracing need determined?');

// Collect all IF formulas to find decision logic
const ifFormulas = formulaCells.filter(c => c.formula.toUpperCase().includes('IF'));
console.log(`\nCells with IF logic: ${ifFormulas.length}`);
for (const c of ifFormulas) {
  console.log(`  ${c.addr}: =${c.formula}  => ${JSON.stringify(c.value)}`);
}

// Look for threshold/comparison values
const comparisonFormulas = formulaCells.filter(c =>
  c.formula.includes('>') || c.formula.includes('<') || c.formula.includes('=')
);
console.log(`\nCells with comparisons: ${comparisonFormulas.length}`);
for (const c of comparisonFormulas) {
  console.log(`  ${c.addr}: =${c.formula}  => ${JSON.stringify(c.value)}`);
}

// ============================================================================
// 8. COST CALCULATION ANALYSIS
// ============================================================================
sep('8. COST CALCULATION ANALYSIS');

// Find cells that look like they contain dollar amounts or costs
const costCells = allCells.filter(c =>
  (typeof c.value === 'number' && c.value > 0) ||
  (c.formatted && c.formatted.includes('$')) ||
  (c.formula && (c.formula.includes('*') || c.formula.includes('VLOOKUP') || c.formula.includes('INDEX')))
);
console.log(`\nPotential cost/calculation cells: ${costCells.length}`);
for (const c of costCells) {
  const fStr = c.formula ? `  FORMULA: =${c.formula}` : '';
  console.log(`  ${c.addr}: ${JSON.stringify(c.value)} (fmt: ${c.formatted})${fStr}`);
}

// ============================================================================
// 9. STATE-SPECIFIC / WIND THRESHOLD ANALYSIS
// ============================================================================
sep('9. STATE-SPECIFIC WIND THRESHOLDS');

// Check for state references, wind speed values, etc.
const windCells = allCells.filter(c => {
  const str = JSON.stringify(c.value).toLowerCase() + (c.formula || '').toLowerCase();
  return str.includes('wind') || str.includes('state') || str.includes('mph') ||
         str.includes('speed') || str.includes('threshold') || str.includes('az') ||
         str.includes('co') || str.includes('ut') || str.includes('115') ||
         str.includes('130') || str.includes('140');
});
console.log(`\nCells possibly related to wind/state: ${windCells.length}`);
for (const c of windCells) {
  const fStr = c.formula ? `  FORMULA: =${c.formula}` : '';
  console.log(`  ${c.addr}: ${JSON.stringify(c.value)}${fStr}`);
}

// ============================================================================
// 10. SEARCH OTHER SHEETS FOR REFERENCES TO THIS SHEET
// ============================================================================
sep('10. OTHER SHEETS REFERENCING "Snow - Diagonal Bracing"');

for (const sheetName of wb.SheetNames) {
  if (sheetName === SHEET_NAME) continue;
  const otherWs = wb.Sheets[sheetName];
  const otherRange = otherWs['!ref'] ? XLSX.utils.decode_range(otherWs['!ref']) : null;
  if (!otherRange) continue;

  const refs = [];
  for (let r = otherRange.s.r; r <= otherRange.e.r; r++) {
    for (let c = otherRange.s.c; c <= otherRange.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = otherWs[addr];
      if (cell && cell.f && cell.f.includes('Diagonal Bracing')) {
        refs.push({ addr, formula: cell.f, value: cell.v, formatted: cell.w });
      }
    }
  }

  if (refs.length > 0) {
    console.log(`\n  Sheet "${sheetName}" (${refs.length} references):`);
    for (const r of refs) {
      console.log(`    ${r.addr}: =${r.formula}  => ${JSON.stringify(r.value)} (${r.formatted})`);
    }
  }
}

// ============================================================================
// 11. SEARCH FOR REFERENCES TO "Diagonal" IN Snow - Changers
// ============================================================================
sep('11. Snow - Changers references to diagonal bracing');

if (changersSheet) {
  const cRange = XLSX.utils.decode_range(changersSheet['!ref']);
  for (let r = cRange.s.r; r <= cRange.e.r; r++) {
    for (let c = cRange.s.c; c <= cRange.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = changersSheet[addr];
      if (cell) {
        const valStr = String(cell.v || '').toLowerCase();
        const fStr = (cell.f || '').toLowerCase();
        if (valStr.includes('diagonal') || valStr.includes('bracing') ||
            fStr.includes('diagonal') || fStr.includes('bracing')) {
          const formula = cell.f ? `  FORMULA: =${cell.f}` : '';
          console.log(`  ${addr}: [${cell.t}] ${JSON.stringify(cell.v)}${formula}`);
        }
      }
    }
  }
}

// ============================================================================
// 12. INTERACTION WITH SNOW ENGINEERING TOTAL
// ============================================================================
sep('12. INTERACTION WITH SNOW ENGINEERING TOTAL');

// Search all sheets for "engineering" + "diagonal" connections
console.log('\nSearching for Snow Engineering Total references to diagonal bracing...');

for (const sheetName of wb.SheetNames) {
  if (!sheetName.toLowerCase().includes('snow')) continue;
  const otherWs = wb.Sheets[sheetName];
  const otherRange = otherWs['!ref'] ? XLSX.utils.decode_range(otherWs['!ref']) : null;
  if (!otherRange) continue;

  const hits = [];
  for (let r = otherRange.s.r; r <= otherRange.e.r; r++) {
    for (let c = otherRange.s.c; c <= otherRange.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = otherWs[addr];
      if (!cell) continue;
      const valStr = String(cell.v || '').toLowerCase();
      const fStr = (cell.f || '').toLowerCase();
      if (valStr.includes('diagonal') || valStr.includes('bracing') ||
          fStr.includes('diagonal bracing')) {
        hits.push({ addr, value: cell.v, formula: cell.f, formatted: cell.w });
      }
    }
  }

  if (hits.length > 0) {
    console.log(`\n  Sheet "${sheetName}":`);
    for (const h of hits) {
      const fStr = h.formula ? `  FORMULA: =${h.formula}` : '';
      console.log(`    ${h.addr}: ${JSON.stringify(h.value)}${fStr}`);
    }
  }
}

// Also check specifically Snow - Changers for the total/sum formulas that might include diagonal bracing
subsep('Snow - Changers: cells referencing Diagonal Bracing sheet');
if (changersSheet) {
  const cRange = XLSX.utils.decode_range(changersSheet['!ref']);
  for (let r = cRange.s.r; r <= cRange.e.r; r++) {
    for (let c = cRange.s.c; c <= cRange.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = changersSheet[addr];
      if (cell && cell.f && cell.f.includes('Diagonal Bracing')) {
        console.log(`  ${addr}: =${cell.f}  => ${JSON.stringify(cell.v)}`);
      }
    }
  }
}

// ============================================================================
// 13. DIMENSION DEPENDENCY ANALYSIS
// ============================================================================
sep('13. DIMENSION DEPENDENCY (width/length/height in formulas)');

const dimCells = formulaCells.filter(c => {
  const f = c.formula.toLowerCase();
  return f.includes('width') || f.includes('length') || f.includes('height') ||
         f.includes('bay') || f.includes('dimension') || f.includes('size');
});
console.log(`\nFormulas referencing dimension keywords: ${dimCells.length}`);
for (const c of dimCells) {
  console.log(`  ${c.addr}: =${c.formula}  => ${JSON.stringify(c.value)}`);
}

// Check if any formula multiplies by a dimension-like reference
const multiplyFormulas = formulaCells.filter(c => c.formula.includes('*'));
console.log(`\nFormulas with multiplication: ${multiplyFormulas.length}`);
for (const c of multiplyFormulas) {
  console.log(`  ${c.addr}: =${c.formula}  => ${JSON.stringify(c.value)}`);
}

// ============================================================================
// 14. NAMED RANGES that might be relevant
// ============================================================================
sep('14. WORKBOOK DEFINED NAMES (relevant to diagonal bracing)');

if (wb.Workbook && wb.Workbook.Names) {
  const relevantNames = wb.Workbook.Names.filter(n => {
    const ref = (n.Ref || '').toLowerCase();
    const name = (n.Name || '').toLowerCase();
    return ref.includes('diagonal') || name.includes('diagonal') ||
           name.includes('brac') || ref.includes('brac') ||
           name.includes('wind') || name.includes('state');
  });
  console.log(`\nRelevant named ranges: ${relevantNames.length}`);
  for (const n of relevantNames) {
    console.log(`  ${n.Name} => ${n.Ref}`);
  }

  // Also dump ALL names referencing this sheet
  const sheetNames = wb.Workbook.Names.filter(n =>
    (n.Ref || '').includes('Diagonal Bracing')
  );
  console.log(`\nNamed ranges pointing to this sheet: ${sheetNames.length}`);
  for (const n of sheetNames) {
    console.log(`  ${n.Name} => ${n.Ref}`);
  }
} else {
  console.log('\nNo defined names in workbook');
}

// ============================================================================
// 15. FULL DEPENDENCY GRAPH
// ============================================================================
sep('15. FULL DEPENDENCY GRAPH');

console.log('\n--- INPUTS (this sheet reads FROM other sheets) ---');
const inputs = {};
for (const c of crossRefs) {
  const matches = [...c.formula.matchAll(/'([^']+)'!([A-Z$]+[0-9$]+)/g)];
  for (const m of matches) {
    const key = `${m[1]}!${m[2]}`;
    if (!inputs[key]) inputs[key] = [];
    inputs[key].push(c.addr);
  }
}
for (const [source, targets] of Object.entries(inputs)) {
  console.log(`  ${source} => used by ${targets.join(', ')}`);
  // Look up the actual value in the source sheet
  const parts = source.split('!');
  const srcSheet = wb.Sheets[parts[0]];
  if (srcSheet) {
    const srcCell = srcSheet[parts[1].replace(/\$/g, '')];
    if (srcCell) {
      console.log(`    (source value: ${JSON.stringify(srcCell.v)}, formula: ${srcCell.f ? '=' + srcCell.f : 'static'})`);
    }
  }
}

console.log('\n--- OUTPUTS (other sheets read FROM this sheet) ---');
// Already done in section 10, summarize
for (const sheetName of wb.SheetNames) {
  if (sheetName === SHEET_NAME) continue;
  const otherWs = wb.Sheets[sheetName];
  const otherRange = otherWs['!ref'] ? XLSX.utils.decode_range(otherWs['!ref']) : null;
  if (!otherRange) continue;

  for (let r = otherRange.s.r; r <= otherRange.e.r; r++) {
    for (let c = otherRange.s.c; c <= otherRange.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = otherWs[addr];
      if (cell && cell.f && cell.f.includes('Diagonal Bracing')) {
        // Extract what cell it reads from this sheet
        const matches = [...cell.f.matchAll(/Diagonal Bracing'!([A-Z$]+[0-9$]+)/g)];
        for (const m of matches) {
          console.log(`  ${SHEET_NAME}!${m[1]} => ${sheetName}!${addr} (=${cell.f})`);
        }
      }
    }
  }
}

// ============================================================================
// 16. MERGED CELLS
// ============================================================================
sep('16. MERGED CELLS');
if (ws['!merges'] && ws['!merges'].length > 0) {
  console.log(`\nMerged regions: ${ws['!merges'].length}`);
  for (const m of ws['!merges']) {
    const from = XLSX.utils.encode_cell(m.s);
    const to = XLSX.utils.encode_cell(m.e);
    console.log(`  ${from}:${to}`);
  }
} else {
  console.log('\nNo merged cells');
}

// ============================================================================
// 17. SUMMARY
// ============================================================================
sep('17. COMPREHENSIVE SUMMARY');

console.log(`
SHEET: "${SHEET_NAME}"
Range: ${ws['!ref']}
Total non-empty cells: ${allCells.length}
Total formula cells: ${formulaCells.length}
Cross-sheet references: ${crossRefs.length}

Key cells to examine:
  J9: ${j9 ? `${JSON.stringify(j9.v)} (formula: ${j9.f ? '=' + j9.f : 'static'})` : 'EMPTY'}
  K10: ${k10 ? `${JSON.stringify(k10.v)} (formula: ${k10.f ? '=' + k10.f : 'static'})` : 'EMPTY'}

Sheets referenced BY this sheet: ${Object.keys(sheetRefsMap).join(', ') || 'none'}
Sheets that READ this sheet: (see section 10 above)
`);
