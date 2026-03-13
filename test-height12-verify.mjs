import XLSX from "xlsx";

const wb = XLSX.readFile("C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx", { cellFormula: true });

function findSheet(name) {
  for (const key of Object.keys(wb.Sheets)) {
    if (key.trim() === name.trim()) return wb.Sheets[key];
  }
  return null;
}

function sheetToArray(ws) {
  const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
  const rows = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      row.push(cell ? (cell.v !== undefined ? cell.v : "") : "");
    }
    rows.push(row);
  }
  return rows;
}
function num(v) { const n = Number(v); return isNaN(n) ? 0 : n; }
function cleanHeader(v) { return String(v ?? "").trim(); }

// === 1. Verify height classification from Changers ===
console.log("=== Snow - Changers: Height Classification ===");
const ch = findSheet("Snow - Changers");
// Row 26 (0-indexed=25) has heights, Row 27 (0-indexed=26) has S/M/T
// Row 28 (0-indexed=27) has feetUsed
for (let c = 1; c <= 21; c++) {
  const heightCell = ch[XLSX.utils.encode_cell({ r: 25, c })];
  const smtCell = ch[XLSX.utils.encode_cell({ r: 26, c })];
  const feetCell = ch[XLSX.utils.encode_cell({ r: 27, c })];
  const h = heightCell?.v;
  const smt = smtCell?.v;
  const ft = feetCell?.v;
  if (h !== undefined) {
    console.log(`  Col ${c}: height=${h}, S/M/T=${smt}, feetUsed=${ft}`);
  }
}

// === 2. Check the FORMULA for D31 (height classification result) ===
console.log("\n=== Changers Formula Chain for Height ===");
const d29 = ch[XLSX.utils.encode_cell({ r: 28, c: 3 })];
const d30 = ch[XLSX.utils.encode_cell({ r: 29, c: 3 })];
const d31 = ch[XLSX.utils.encode_cell({ r: 30, c: 3 })];
console.log(`D29 (height): value=${d29?.v}, formula=${d29?.f}`);
console.log(`D30 (match pos): value=${d30?.v}, formula=${d30?.f}`);
console.log(`D31 (S/M/T): value=${d31?.v}, formula=${d31?.f}`);

// === 3. Check Truss Spacing for BOTH S-30GL and T-30GL at various configs ===
console.log("\n=== Truss Spacing: S-30GL vs T-30GL ===");
const ts = findSheet("Snow - Truss Spacing");
const tsData = sheetToArray(ts);
const tsHeaders = tsData[0];

// Find columns for O-105-22-AFV and E-105-22-AFV
const targetCols = {};
for (let c = 1; c < tsHeaders.length; c++) {
  const h = cleanHeader(tsHeaders[c]);
  if (h.includes("22") && h.includes("105")) {
    targetCols[h] = c;
  }
}
console.log("Target columns:", Object.keys(targetCols));

// Find rows for S-30GL, T-30GL, S-20LL, T-20LL
const targetRows = {};
for (let r = 1; r < tsData.length; r++) {
  const key = cleanHeader(tsData[r][0]);
  if (["S-30GL", "T-30GL", "S-20LL", "T-20LL", "M-30GL", "M-20LL"].includes(key)) {
    targetRows[key] = r;
  }
}
console.log("Target rows:", Object.keys(targetRows));

for (const [snowCode, row] of Object.entries(targetRows)) {
  for (const [colName, col] of Object.entries(targetCols)) {
    console.log(`  ${snowCode} @ ${colName} = ${num(tsData[row][col])}`);
  }
}

// === 4. Also check for width 24 ===
console.log("\n=== Truss Spacing: width 24, wind 105 ===");
const targetCols24 = {};
for (let c = 1; c < tsHeaders.length; c++) {
  const h = cleanHeader(tsHeaders[c]);
  if (h.includes("24") && h.includes("105")) {
    targetCols24[h] = c;
  }
}
for (const [snowCode, row] of Object.entries(targetRows)) {
  for (const [colName, col] of Object.entries(targetCols24)) {
    console.log(`  ${snowCode} @ ${colName} = ${num(tsData[row][col])}`);
  }
}

// === 5. Original truss counts for 22-AZ and 24-AZ at length 50 ===
console.log("\n=== Original Truss Counts ===");
const tr = findSheet("Snow - Trusses");
const trData = sheetToArray(tr);
for (let c = 1; c < trData[0].length; c++) {
  const h = cleanHeader(trData[0][c]);
  if (h === "22-AZ" || h === "24-AZ") {
    for (let r = 1; r < trData.length; r++) {
      if (num(trData[r][0]) === 50) {
        console.log(`  ${h} @ length 50 = ${num(trData[r][c])} trusses`);
      }
    }
  }
}

// === 6. Check the Truss Spacing FORMULA sheet for how it builds the snow code ===
console.log("\n=== Truss Spacing: Formula cells (row 49-55) ===");
for (let r = 45; r <= 55; r++) {
  for (let c = 0; c <= 20; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = ts[addr];
    if (cell?.f) {
      console.log(`  [${addr}] r=${r} c=${c}: value=${JSON.stringify(cell.v)}, formula=${cell.f}`);
    }
  }
}

// === 7. Check what Math Calculations uses for the snow code assembly ===
console.log("\n=== Math Calculations: Snow code assembly ===");
const mc = findSheet("Snow - Math Calculations");
// The snow code is built from height prefix + snow load code
// Check cells around rows 0-5
for (let r = 0; r <= 10; r++) {
  for (let c = 0; c <= 25; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = mc[addr];
    if (cell?.f && (cell.f.includes("Changers") || cell.f.includes("D31") || cell.f.includes("D14"))) {
      console.log(`  [${addr}] r=${r} c=${c}: value=${JSON.stringify(cell.v)}, formula=${cell.f}`);
    }
  }
}

// === 8. Check Truss Spacing formula for F52 and F54 (the output cells) ===
console.log("\n=== Truss Spacing: Output formula cells ===");
for (const addr of ["F52", "F54", "F50", "F48", "H52", "Q49", "Q47"]) {
  const cell = ts[addr];
  if (cell) {
    console.log(`  ${addr}: value=${JSON.stringify(cell.v)}, formula=${cell.f ?? "(static)"}`);
  }
}

// Check what builds the config key used for truss spacing lookup
console.log("\n=== Truss Spacing: Config key construction ===");
// The F54 or similar cell uses INDEX/MATCH to find the spacing
// Let's check ALL formulas in the sheet
let formulaCount = 0;
const range = XLSX.utils.decode_range(ts["!ref"] || "A1");
for (let r = 0; r <= range.e.r; r++) {
  for (let c = 0; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = ts[addr];
    if (cell?.f) {
      formulaCount++;
      if (formulaCount <= 40) { // Print first 40
        console.log(`  [${addr}] r=${r} c=${c}: value=${JSON.stringify(cell.v)}, formula=${cell.f}`);
      }
    }
  }
}
console.log(`Total formula cells: ${formulaCount}`);
