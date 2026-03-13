import XLSX from "xlsx";

const wb = XLSX.readFile("C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx", { cellFormula: true });

function findSheet(name) {
  for (const key of Object.keys(wb.Sheets)) {
    if (key.trim() === name.trim()) return wb.Sheets[key];
  }
  return null;
}

const ws = findSheet("Snow - Changers");
if (!ws) { console.error("Sheet not found!"); process.exit(1); }

const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
console.log(`Sheet range: ${ws["!ref"]}`);
console.log(`Rows: ${range.s.r} to ${range.e.r}, Cols: ${range.s.c} to ${range.e.c}`);
console.log(`Total rows: ${range.e.r - range.s.r + 1}, Total cols: ${range.e.c - range.s.c + 1}`);

// ============================================================
// SECTION 1: COMPLETE CELL DUMP - Every cell with value or formula
// ============================================================
console.log("\n" + "=".repeat(80));
console.log("SECTION 1: COMPLETE CELL DUMP (every non-empty cell)");
console.log("=".repeat(80));

const allCells = [];
for (let r = range.s.r; r <= range.e.r; r++) {
  for (let c = range.s.c; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = ws[addr];
    if (cell && (cell.v !== undefined || cell.f)) {
      allCells.push({ r, c, addr, value: cell.v, formula: cell.f || null, type: cell.t });
    }
  }
}

// Group by row for readability
let currentRow = -1;
for (const cell of allCells) {
  if (cell.r !== currentRow) {
    currentRow = cell.r;
    console.log(`\n--- Row ${cell.r} (Excel row ${cell.r + 1}) ---`);
  }
  const fStr = cell.formula ? ` | FORMULA: ${cell.formula}` : "";
  console.log(`  ${cell.addr}: value=${JSON.stringify(cell.value)} type=${cell.type}${fStr}`);
}

// ============================================================
// SECTION 2: ROW-BY-ROW ANALYSIS WITH SECTION IDENTIFICATION
// ============================================================
console.log("\n\n" + "=".repeat(80));
console.log("SECTION 2: ROW-BY-ROW STRUCTURAL ANALYSIS");
console.log("=".repeat(80));

// Collect all cells per row
const rowMap = {};
for (const cell of allCells) {
  if (!rowMap[cell.r]) rowMap[cell.r] = [];
  rowMap[cell.r].push(cell);
}

for (const [rowStr, cells] of Object.entries(rowMap)) {
  const r = parseInt(rowStr);
  const labels = cells.filter(c => c.type === "s").map(c => `${c.addr}="${c.value}"`).join(", ");
  const numbers = cells.filter(c => c.type === "n").map(c => `${c.addr}=${c.value}`).join(", ");
  const formulas = cells.filter(c => c.formula).map(c => `${c.addr}=${c.formula}`).join(", ");

  console.log(`\nRow ${r} (Excel ${r+1}): ${cells.length} cells`);
  if (labels) console.log(`  Labels: ${labels}`);
  if (numbers) console.log(`  Numbers: ${numbers}`);
  if (formulas) console.log(`  Formulas: ${formulas}`);
}

// ============================================================
// SECTION 3: WIND SPEED BUCKETING (rows 0-4)
// ============================================================
console.log("\n\n" + "=".repeat(80));
console.log("SECTION 3: WIND SPEED BUCKETING (detailed)");
console.log("=".repeat(80));

for (let r = 0; r <= 10; r++) {
  const cells = [];
  for (let c = 0; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = ws[addr];
    if (cell && cell.v !== undefined) {
      cells.push({ col: c, colLetter: XLSX.utils.encode_col(c), value: cell.v, formula: cell.f, type: cell.t });
    }
  }
  if (cells.length > 0) {
    console.log(`\nRow ${r} (Excel ${r+1}):`);
    for (const c of cells) {
      const fStr = c.formula ? ` [FORMULA: ${c.formula}]` : "";
      console.log(`  Col ${c.colLetter} (${c.col}): ${JSON.stringify(c.value)} (${c.type})${fStr}`);
    }
  }
}

// ============================================================
// SECTION 4: HEIGHT CLASSIFICATION (rows 25-30)
// ============================================================
console.log("\n\n" + "=".repeat(80));
console.log("SECTION 4: HEIGHT CLASSIFICATION & FEET USED (rows 20-35)");
console.log("=".repeat(80));

for (let r = 20; r <= 35; r++) {
  const cells = [];
  for (let c = 0; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = ws[addr];
    if (cell && cell.v !== undefined) {
      cells.push({ col: c, colLetter: XLSX.utils.encode_col(c), value: cell.v, formula: cell.f, type: cell.t });
    }
  }
  if (cells.length > 0) {
    console.log(`\nRow ${r} (Excel ${r+1}):`);
    for (const c of cells) {
      const fStr = c.formula ? ` [FORMULA: ${c.formula}]` : "";
      console.log(`  Col ${c.colLetter} (${c.col}): ${JSON.stringify(c.value)} (${c.type})${fStr}`);
    }
  }
}

// ============================================================
// SECTION 5: PRICING TABLES (rows 35-80+)
// ============================================================
console.log("\n\n" + "=".repeat(80));
console.log("SECTION 5: PRICING TABLES (rows 35 to end)");
console.log("=".repeat(80));

for (let r = 35; r <= range.e.r; r++) {
  const cells = [];
  for (let c = 0; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = ws[addr];
    if (cell && cell.v !== undefined) {
      cells.push({ col: c, colLetter: XLSX.utils.encode_col(c), value: cell.v, formula: cell.f, type: cell.t });
    }
  }
  if (cells.length > 0) {
    console.log(`\nRow ${r} (Excel ${r+1}):`);
    for (const c of cells) {
      const fStr = c.formula ? ` [FORMULA: ${c.formula}]` : "";
      console.log(`  Col ${c.colLetter} (${c.col}): ${JSON.stringify(c.value)} (${c.type})${fStr}`);
    }
  }
}

// ============================================================
// SECTION 6: ALL CROSS-SHEET REFERENCES (formulas referencing other sheets)
// ============================================================
console.log("\n\n" + "=".repeat(80));
console.log("SECTION 6: ALL CROSS-SHEET REFERENCES IN THIS SHEET");
console.log("=".repeat(80));

const crossRefs = allCells.filter(c => c.formula && (c.formula.includes("!") || c.formula.includes("'")));
for (const cell of crossRefs) {
  console.log(`  ${cell.addr} (r${cell.r},c${cell.c}): value=${JSON.stringify(cell.value)} formula=${cell.formula}`);
}

// ============================================================
// SECTION 7: REFERENCES TO THIS SHEET FROM OTHER SHEETS
// ============================================================
console.log("\n\n" + "=".repeat(80));
console.log("SECTION 7: REFERENCES TO 'Snow - Changers' FROM OTHER SHEETS");
console.log("=".repeat(80));

for (const sheetName of wb.SheetNames) {
  if (sheetName.trim() === "Snow - Changers") continue;
  const otherWs = wb.Sheets[sheetName];
  if (!otherWs || !otherWs["!ref"]) continue;
  const otherRange = XLSX.utils.decode_range(otherWs["!ref"]);
  const refs = [];
  for (let r = otherRange.s.r; r <= otherRange.e.r; r++) {
    for (let c = otherRange.s.c; c <= otherRange.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = otherWs[addr];
      if (cell?.f && (cell.f.includes("Changers") || cell.f.includes("changers"))) {
        refs.push({ addr, r, c, value: cell.v, formula: cell.f });
      }
    }
  }
  if (refs.length > 0) {
    console.log(`\n  Sheet: "${sheetName}" (${refs.length} references):`);
    for (const ref of refs) {
      console.log(`    ${ref.addr}: value=${JSON.stringify(ref.value)} formula=${ref.formula}`);
    }
  }
}

// ============================================================
// SECTION 8: STRUCTURED TABLE EXTRACTION
// ============================================================
console.log("\n\n" + "=".repeat(80));
console.log("SECTION 8: STRUCTURED TABLE EXTRACTION");
console.log("=".repeat(80));

// Helper: get cell value
function getVal(r, c) {
  const addr = XLSX.utils.encode_cell({ r, c });
  const cell = ws[addr];
  return cell?.v ?? null;
}

function getFormula(r, c) {
  const addr = XLSX.utils.encode_cell({ r, c });
  const cell = ws[addr];
  return cell?.f ?? null;
}

// 8a: Wind bucketing table
console.log("\n--- 8a: Wind Bucketing Table ---");
console.log("Row 0 (headers):");
for (let c = 0; c <= range.e.c; c++) {
  const v = getVal(0, c);
  if (v !== null) console.log(`  Col ${XLSX.utils.encode_col(c)}: ${JSON.stringify(v)}`);
}
console.log("Row 1 (values):");
for (let c = 0; c <= range.e.c; c++) {
  const v = getVal(1, c);
  const f = getFormula(1, c);
  if (v !== null) console.log(`  Col ${XLSX.utils.encode_col(c)}: ${JSON.stringify(v)} ${f ? `[F: ${f}]` : ""}`);
}

// 8b: Height classification
console.log("\n--- 8b: Height Classification Table ---");
for (let r = 24; r <= 28; r++) {
  console.log(`Row ${r} (Excel ${r+1}):`);
  for (let c = 0; c <= range.e.c; c++) {
    const v = getVal(r, c);
    const f = getFormula(r, c);
    if (v !== null) console.log(`  Col ${XLSX.utils.encode_col(c)} (${c}): ${JSON.stringify(v)} ${f ? `[F: ${f}]` : ""}`);
  }
}

// 8c: Complete pricing grid
console.log("\n--- 8c: Pricing Grid Analysis ---");
// Find where "Truss" or "Channel" or pricing headers appear
for (let r = 35; r <= range.e.r; r++) {
  for (let c = 0; c <= 3; c++) {
    const v = getVal(r, c);
    if (v !== null && typeof v === "string" && v.length > 0) {
      console.log(`  Label at Row ${r} Col ${XLSX.utils.encode_col(c)}: "${v}"`);
    }
  }
}

// ============================================================
// SECTION 9: MERGED CELLS
// ============================================================
console.log("\n\n" + "=".repeat(80));
console.log("SECTION 9: MERGED CELLS");
console.log("=".repeat(80));

if (ws["!merges"] && ws["!merges"].length > 0) {
  for (const merge of ws["!merges"]) {
    const s = XLSX.utils.encode_cell(merge.s);
    const e = XLSX.utils.encode_cell(merge.e);
    const val = getVal(merge.s.r, merge.s.c);
    console.log(`  ${s}:${e} value=${JSON.stringify(val)}`);
  }
} else {
  console.log("  No merged cells found.");
}

// ============================================================
// SECTION 10: COLUMN SUMMARY
// ============================================================
console.log("\n\n" + "=".repeat(80));
console.log("SECTION 10: COLUMN USAGE SUMMARY");
console.log("=".repeat(80));

for (let c = range.s.c; c <= range.e.c; c++) {
  const colCells = allCells.filter(cell => cell.c === c);
  if (colCells.length > 0) {
    const types = [...new Set(colCells.map(c => c.type))].join(",");
    const hasFormulas = colCells.some(c => c.formula);
    const sampleValues = colCells.slice(0, 5).map(c => JSON.stringify(c.value)).join(", ");
    console.log(`  Col ${XLSX.utils.encode_col(c)} (${c}): ${colCells.length} cells, types=[${types}], formulas=${hasFormulas}, samples: ${sampleValues}`);
  }
}

// ============================================================
// SECTION 11: COMPLETE PRICING TABLE MATRIX FORMAT
// ============================================================
console.log("\n\n" + "=".repeat(80));
console.log("SECTION 11: PRICING TABLES AS MATRICES");
console.log("=".repeat(80));

// Print rows 36-80 as a clean grid
console.log("\nRows 36-80 grid (showing non-empty cells):");
for (let r = 36; r <= Math.min(range.e.r, 85); r++) {
  const parts = [];
  let hasData = false;
  for (let c = 0; c <= range.e.c; c++) {
    const v = getVal(r, c);
    if (v !== null) {
      hasData = true;
      const colL = XLSX.utils.encode_col(c);
      parts.push(`${colL}=${typeof v === "number" ? v : JSON.stringify(v)}`);
    }
  }
  if (hasData) {
    console.log(`  R${r}: ${parts.join(" | ")}`);
  }
}

// ============================================================
// SECTION 12: FORMULA PATTERN ANALYSIS
// ============================================================
console.log("\n\n" + "=".repeat(80));
console.log("SECTION 12: FORMULA PATTERN ANALYSIS");
console.log("=".repeat(80));

const formulaCells = allCells.filter(c => c.formula);
console.log(`\nTotal formulas in sheet: ${formulaCells.length}`);

// Group by formula pattern
const patterns = {};
for (const cell of formulaCells) {
  // Normalize: replace cell refs with placeholders
  const pattern = cell.formula.replace(/[A-Z]+\d+/g, "REF").replace(/\d+/g, "N");
  if (!patterns[pattern]) patterns[pattern] = [];
  patterns[pattern].push(cell);
}

console.log(`\nUnique formula patterns: ${Object.keys(patterns).length}`);
for (const [pattern, cells] of Object.entries(patterns)) {
  console.log(`\n  Pattern: ${pattern}`);
  console.log(`  Count: ${cells.length}`);
  console.log(`  Examples:`);
  for (const c of cells.slice(0, 3)) {
    console.log(`    ${c.addr}: ${c.formula} = ${JSON.stringify(c.value)}`);
  }
}

// ============================================================
// SECTION 13: INTERMEDIATE ROWS 5-24 (anything between wind and height)
// ============================================================
console.log("\n\n" + "=".repeat(80));
console.log("SECTION 13: INTERMEDIATE ROWS 2-24 (between wind bucketing and height)");
console.log("=".repeat(80));

for (let r = 2; r <= 24; r++) {
  const cells = [];
  for (let c = 0; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = ws[addr];
    if (cell && cell.v !== undefined) {
      cells.push({ col: c, colLetter: XLSX.utils.encode_col(c), value: cell.v, formula: cell.f, type: cell.t });
    }
  }
  if (cells.length > 0) {
    console.log(`\nRow ${r} (Excel ${r+1}):`);
    for (const c of cells) {
      const fStr = c.formula ? ` [FORMULA: ${c.formula}]` : "";
      console.log(`  Col ${c.colLetter} (${c.col}): ${JSON.stringify(c.value)} (${c.type})${fStr}`);
    }
  }
}

console.log("\n\nDONE - Complete analysis of Snow - Changers sheet");
