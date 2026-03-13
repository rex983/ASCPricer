import XLSX from 'xlsx';
import { readFileSync } from 'fs';

const FILE = 'C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx';
const wb = XLSX.read(readFileSync(FILE), { cellFormula: true, cellStyles: true, cellNF: true });

console.log('='.repeat(100));
console.log('EXHAUSTIVE SNOW LOAD BREAKDOWN & QUOTE SHEET ENGINEERING ANALYSIS');
console.log('File:', FILE);
console.log('='.repeat(100));

// ─── 1. LIST ALL SHEETS ───────────────────────────────────────────────────────
console.log('\n\n██ SECTION 1: ALL SHEETS IN WORKBOOK ██');
console.log('Total sheets:', wb.SheetNames.length);
wb.SheetNames.forEach((name, i) => {
  const ws = wb.Sheets[name];
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  console.log(`  [${i}] "${name}" — range: ${ws['!ref']} (${range.e.r + 1} rows × ${range.e.c + 1} cols)`);
});

// ─── 2. EXHAUSTIVE DUMP OF SNOW LOAD BREAKDOWN ────────────────────────────────
console.log('\n\n██ SECTION 2: SNOW LOAD BREAKDOWN — EVERY CELL ██');

function dumpSheet(sheetName, rowStart, rowEnd, colStart, colEnd) {
  const ws = wb.Sheets[sheetName];
  if (!ws) { console.log(`  Sheet "${sheetName}" NOT FOUND`); return; }

  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  const rs = rowStart ?? range.s.r;
  const re = rowEnd ?? range.e.r;
  const cs = colStart ?? range.s.c;
  const ce = colEnd ?? range.e.c;

  console.log(`  Sheet range: ${ws['!ref']}`);
  console.log(`  Scanning rows ${rs}–${re}, cols ${cs}–${ce}`);
  console.log('');

  let cellCount = 0;
  for (let r = rs; r <= re; r++) {
    for (let c = cs; c <= ce; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell) {
        cellCount++;
        const parts = [`  ${addr}`];
        parts.push(`type=${cell.t}`);
        parts.push(`value=${JSON.stringify(cell.v)}`);
        if (cell.w !== undefined) parts.push(`formatted="${cell.w}"`);
        if (cell.f) parts.push(`FORMULA=[${cell.f}]`);
        if (cell.z) parts.push(`numFmt="${cell.z}"`);
        console.log(parts.join(' | '));
      }
    }
  }
  console.log(`\n  Total populated cells: ${cellCount}`);
  return ws;
}

// Find the snow load breakdown sheet (case-insensitive search)
const snowSheetName = wb.SheetNames.find(n => n.toLowerCase().includes('snow'));
if (snowSheetName) {
  console.log(`\nFound snow sheet: "${snowSheetName}"`);
  dumpSheet(snowSheetName);
} else {
  console.log('\n  *** NO sheet with "snow" in the name found ***');
  // Check for any engineering-related sheets
  const engSheets = wb.SheetNames.filter(n =>
    /snow|eng|load|breakdown/i.test(n)
  );
  console.log('  Engineering-related sheets:', engSheets.length ? engSheets : 'NONE');
}

// ─── 3. QUOTE SHEET — ENGINEERING ROWS (rows 10-30) ──────────────────────────
console.log('\n\n██ SECTION 3: QUOTE SHEET — ROWS 10–40 (ENGINEERING SECTION) ██');

const quoteSheetName = wb.SheetNames.find(n => /quote\s*sheet/i.test(n)) || 'Quote Sheet';
console.log(`Using quote sheet: "${quoteSheetName}"`);
dumpSheet(quoteSheetName, 9, 39, 0, 20); // rows 10-40 (0-indexed: 9-39), cols A-U

// ─── 4. QUOTE SHEET — FULL ENGINEERING SCAN (wider) ──────────────────────────
console.log('\n\n██ SECTION 4: QUOTE SHEET — FIRST 5 ROWS (HEADERS/CONFIG) ██');
dumpSheet(quoteSheetName, 0, 4, 0, 20);

// ─── 5. SEARCH ALL SHEETS FOR SNOW/ENGINEERING REFERENCES ────────────────────
console.log('\n\n██ SECTION 5: CROSS-SHEET SNOW/ENGINEERING REFERENCES ██');

wb.SheetNames.forEach(sheetName => {
  const ws = wb.Sheets[sheetName];
  if (!ws || !ws['!ref']) return;

  const range = XLSX.utils.decode_range(ws['!ref']);
  const hits = [];

  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (!cell) continue;

      const valStr = String(cell.v || '').toLowerCase();
      const fStr = String(cell.f || '').toLowerCase();

      if (/snow|eng(ineer)?|sl\b|psf|ground.*load|roof.*load/i.test(valStr + ' ' + fStr)) {
        hits.push({
          addr,
          value: cell.v,
          formula: cell.f || null,
          formatted: cell.w || null
        });
      }

      // Also check for references to Snow Load Breakdown sheet
      if (cell.f && /snow/i.test(cell.f)) {
        hits.push({
          addr,
          value: cell.v,
          formula: cell.f,
          formatted: cell.w || null,
          note: 'FORMULA REFERENCES SNOW SHEET'
        });
      }
    }
  }

  if (hits.length > 0) {
    console.log(`\n  Sheet "${sheetName}" — ${hits.length} snow/engineering references:`);
    hits.forEach(h => {
      console.log(`    ${h.addr}: value=${JSON.stringify(h.value)} formula=[${h.formula}] fmt="${h.formatted}" ${h.note || ''}`);
    });
  }
});

// ─── 6. MATH CALCULATIONS — SNOW/ENGINEERING SECTION ─────────────────────────
console.log('\n\n██ SECTION 6: MATH CALCULATIONS — SNOW/ENGINEERING CELLS ██');

// Find Math Calculations sheets (there may be multiple for different widths)
const mathSheets = wb.SheetNames.filter(n => /math/i.test(n));
console.log(`Found ${mathSheets.length} Math Calculation sheets:`, mathSheets);

mathSheets.forEach(sheetName => {
  const ws = wb.Sheets[sheetName];
  if (!ws || !ws['!ref']) return;

  const range = XLSX.utils.decode_range(ws['!ref']);
  console.log(`\n  --- "${sheetName}" (${ws['!ref']}) ---`);

  // Find all cells with snow/engineering references
  const snowCells = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (!cell) continue;

      const valStr = String(cell.v || '').toLowerCase();
      const fStr = String(cell.f || '').toLowerCase();

      if (/snow|eng|sl\b|psf|ground|roof.*load|truss|purl|header/i.test(valStr + ' ' + fStr)) {
        snowCells.push({ addr, cell });
      }
    }
  }

  if (snowCells.length > 0) {
    console.log(`  Found ${snowCells.length} snow/engineering-related cells:`);
    snowCells.forEach(({ addr, cell }) => {
      console.log(`    ${addr}: type=${cell.t} value=${JSON.stringify(cell.v)} formula=[${cell.f || ''}] fmt="${cell.w || ''}"`);
    });
  }

  // Also dump rows that seem to be in an "engineering" section
  // Look for rows containing "engineer" or "snow" labels and dump surrounding context
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= Math.min(range.s.c + 2, range.e.c); c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell && typeof cell.v === 'string' && /snow|engineer|upgrade/i.test(cell.v)) {
        console.log(`\n  CONTEXT around row ${r + 1} in "${sheetName}":`);
        for (let dr = Math.max(0, r - 2); dr <= Math.min(range.e.r, r + 5); dr++) {
          const rowCells = [];
          for (let dc = 0; dc <= Math.min(15, range.e.c); dc++) {
            const a = XLSX.utils.encode_cell({ r: dr, c: dc });
            const cl = ws[a];
            if (cl) rowCells.push(`${a}=${JSON.stringify(cl.v)}${cl.f ? ' [' + cl.f + ']' : ''}`);
          }
          if (rowCells.length) console.log(`    Row ${dr + 1}: ${rowCells.join(' | ')}`);
        }
      }
    }
  }
});

// ─── 7. TRACE FORMULA CHAIN: Math → Snow → Quote ────────────────────────────
console.log('\n\n██ SECTION 7: FORMULA CHAIN TRACING ██');

// Find all cells in Quote Sheet that reference other sheets
const qws = wb.Sheets[quoteSheetName];
if (qws) {
  const qrange = XLSX.utils.decode_range(qws['!ref'] || 'A1');
  console.log('\n  Quote Sheet cells with cross-sheet formula references:');

  for (let r = 0; r <= Math.min(50, qrange.e.r); r++) {
    for (let c = 0; c <= qrange.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = qws[addr];
      if (cell && cell.f && cell.f.includes('!')) {
        console.log(`    ${addr}: value=${JSON.stringify(cell.v)} formula=[${cell.f}] fmt="${cell.w || ''}"`);
      }
    }
  }
}

// ─── 8. DEEPER: DUMP QUOTE SHEET ROWS 1-60 FOR FULL PICTURE ─────────────────
console.log('\n\n██ SECTION 8: QUOTE SHEET — ROWS 1–60 COMPLETE DUMP ██');
dumpSheet(quoteSheetName, 0, 59, 0, 25);

// ─── 9. CHECK FOR CONDITIONAL LOGIC THAT COULD PRODUCE $0 ──────────────────
console.log('\n\n██ SECTION 9: CONDITIONAL FORMULAS (IF/AND/OR) IN SNOW & ENGINEERING ██');

function findConditionalFormulas(sheetName) {
  const ws = wb.Sheets[sheetName];
  if (!ws || !ws['!ref']) return;

  const range = XLSX.utils.decode_range(ws['!ref']);
  const conditionals = [];

  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell && cell.f && /IF\(|AND\(|OR\(|VLOOKUP|INDEX|MATCH/i.test(cell.f)) {
        const valStr = String(cell.v || '').toLowerCase();
        const fStr = cell.f.toLowerCase();
        // Only show if related to snow/engineering/price
        if (/snow|eng|price|cost|total|upgrade|sl|psf|truss|header|purl/i.test(fStr + ' ' + valStr)) {
          conditionals.push({ addr, value: cell.v, formula: cell.f, formatted: cell.w });
        }
      }
    }
  }

  if (conditionals.length > 0) {
    console.log(`\n  "${sheetName}" — ${conditionals.length} conditional formulas:`);
    conditionals.forEach(c => {
      console.log(`    ${c.addr}: value=${JSON.stringify(c.value)}`);
      console.log(`      formula=[${c.formula}]`);
      console.log(`      formatted="${c.formatted}"`);
    });
  }
}

wb.SheetNames.forEach(findConditionalFormulas);

// ─── 10. NAMED RANGES ───────────────────────────────────────────────────────
console.log('\n\n██ SECTION 10: DEFINED NAMES / NAMED RANGES ██');
if (wb.Workbook && wb.Workbook.Names) {
  const snowNames = wb.Workbook.Names.filter(n =>
    /snow|eng|sl|load|truss|header|purl/i.test(n.Name + ' ' + (n.Ref || ''))
  );
  console.log(`  Total defined names: ${wb.Workbook.Names.length}`);
  console.log(`  Snow/engineering related: ${snowNames.length}`);
  snowNames.forEach(n => {
    console.log(`    "${n.Name}" → ${n.Ref}`);
  });

  // Also dump ALL defined names for completeness
  console.log('\n  ALL defined names:');
  wb.Workbook.Names.forEach(n => {
    console.log(`    "${n.Name}" → ${n.Ref}`);
  });
} else {
  console.log('  No defined names found');
}

// ─── 11. SNOW LOAD BREAKDOWN — MERGE CELLS ─────────────────────────────────
console.log('\n\n██ SECTION 11: MERGED CELLS IN SNOW LOAD BREAKDOWN ██');
if (snowSheetName) {
  const sws = wb.Sheets[snowSheetName];
  if (sws['!merges']) {
    console.log(`  ${sws['!merges'].length} merged regions:`);
    sws['!merges'].forEach(m => {
      const s = XLSX.utils.encode_cell(m.s);
      const e = XLSX.utils.encode_cell(m.e);
      console.log(`    ${s}:${e}`);
    });
  } else {
    console.log('  No merged cells');
  }
}

// ─── 12. FULL DUMP OF ALL "SIDES" / PRICING SHEETS FOR ENGINEERING ROWS ────
console.log('\n\n██ SECTION 12: SIDES/PRICING SHEETS — ENGINEERING ROWS ██');

const sidesSheets = wb.SheetNames.filter(n => /side|price|option/i.test(n));
console.log(`Found ${sidesSheets.length} sides/pricing sheets:`, sidesSheets);

sidesSheets.forEach(sheetName => {
  const ws = wb.Sheets[sheetName];
  if (!ws || !ws['!ref']) return;

  const range = XLSX.utils.decode_range(ws['!ref']);

  // Search for snow/engineering rows
  for (let r = range.s.r; r <= range.e.r; r++) {
    let rowHasSnow = false;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell && /snow|eng|truss|header|purl|upgrade/i.test(String(cell.v || '') + String(cell.f || ''))) {
        rowHasSnow = true;
        break;
      }
    }

    if (rowHasSnow) {
      const rowCells = [];
      for (let c = 0; c <= Math.min(20, range.e.c); c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        if (cell) rowCells.push(`${addr}=${JSON.stringify(cell.v)}${cell.f ? ' [' + cell.f + ']' : ''}`);
      }
      console.log(`  "${sheetName}" Row ${r + 1}: ${rowCells.join(' | ')}`);
    }
  }
});

console.log('\n\n' + '='.repeat(100));
console.log('ANALYSIS COMPLETE');
console.log('='.repeat(100));
