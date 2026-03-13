import XLSX from "xlsx";

const wb = XLSX.readFile("C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx", { cellFormula: true });

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

function findSheet(name) {
  for (const key of Object.keys(wb.Sheets)) {
    if (key.trim() === name.trim()) return wb.Sheets[key];
  }
  return null;
}

// Read Truss Spacing sheet thoroughly
const tsSheet = findSheet("Snow - Truss Spacing");
const tsData = sheetToArray(tsSheet);
const headers = tsData[0];

console.log("=== Truss Spacing: Column headers containing '30' ===");
const cols30 = [];
for (let c = 0; c < headers.length; c++) {
  const h = cleanHeader(headers[c]);
  if (h.includes("30")) {
    cols30.push({ col: c, header: h });
    console.log(`  Col ${c}: "${h}"`);
  }
}

// Find T-70GL row
console.log("\n=== Row keys containing '70GL' ===");
for (let r = 0; r < tsData.length; r++) {
  const key = cleanHeader(tsData[r][0]);
  if (key.includes("70GL")) {
    console.log(`  Row ${r}: "${key}"`);
    // Print values at all 30-width columns
    for (const { col, header } of cols30) {
      console.log(`    ${header} = ${num(tsData[r][col])}`);
    }
  }
}

// Also check what config key format the columns use
console.log("\n=== Sample column headers (first 20 then around col 200) ===");
for (let c = 0; c < Math.min(20, headers.length); c++) {
  if (headers[c] !== "") console.log(`  Col ${c}: "${cleanHeader(headers[c])}"`);
}
for (let c = 195; c < Math.min(215, headers.length); c++) {
  if (headers[c] !== "") console.log(`  Col ${c}: "${cleanHeader(headers[c])}"`);
}

// Now let me also verify by reading the ACTUAL parsed data from our parser
// First check: are there duplicate columns for the same config?
console.log("\n=== All columns matching 'E-105-30-AFV' ===");
let found = 0;
for (let c = 0; c < headers.length; c++) {
  if (cleanHeader(headers[c]) === "E-105-30-AFV") {
    console.log(`  Col ${c}: value at T-70GL row = ?`);
    for (let r = 0; r < tsData.length; r++) {
      if (cleanHeader(tsData[r][0]) === "T-70GL") {
        console.log(`    T-70GL: ${num(tsData[r][c])}`);
      }
    }
    found++;
  }
}
console.log(`Found ${found} matching columns`);

// Check ALL T-70GL values at 30-width columns to see range of spacings
console.log("\n=== T-70GL all values at width-30 columns ===");
for (let r = 0; r < tsData.length; r++) {
  if (cleanHeader(tsData[r][0]) === "T-70GL") {
    for (const { col, header } of cols30) {
      const val = num(tsData[r][col]);
      if (val !== 0) console.log(`  ${header}: ${val}`);
    }
    break;
  }
}

// Also verify original truss count for 30-AZ at length 100
const trSheet = findSheet("Snow - Trusses");
const trData = sheetToArray(trSheet);
console.log("\n=== Truss Counts: 30-AZ at various lengths ===");
let trCol = -1;
for (let c = 0; c < trData[0].length; c++) {
  if (cleanHeader(trData[0][c]) === "30-AZ") { trCol = c; break; }
}
console.log("30-AZ column:", trCol);
for (let r = 1; r < trData.length; r++) {
  const len = num(trData[r][0]);
  if ([50, 75, 100].includes(len)) {
    console.log(`  Length ${len}: ${num(trData[r][trCol])} trusses`);
  }
}

// Check HC spacing for 36-70GL at 105
const hcSheet = findSheet("Snow - Hat Channels");
const hcData = sheetToArray(hcSheet);
console.log("\n=== HC spacing verification ===");
const hcHeaders = hcData[0];
let windCol105 = -1;
for (let c = 0; c < hcHeaders.length; c++) {
  if (num(hcHeaders[c]) === 105) { windCol105 = c; break; }
}
for (let r = 0; r < hcData.length; r++) {
  const key = cleanHeader(hcData[r][0]);
  if (key === "36-70GL") {
    console.log(`36-70GL at wind 105 (col ${windCol105}): ${num(hcData[r][windCol105])}`);
  }
}

// Original HC for AZ width 30
console.log("\n=== Original HC for AZ/30 ===");
let origHCCol = -1;
for (let c = 8; c < hcHeaders.length; c++) {
  if (cleanHeader(hcHeaders[c]).toLowerCase().includes("original")) {
    origHCCol = c; break;
  }
}
if (origHCCol >= 0) {
  const widthHeaders = [];
  for (let c = origHCCol + 1; c < hcHeaders.length; c++) {
    widthHeaders.push(cleanHeader(hcHeaders[c]));
  }
  console.log("Width headers after Original:", widthHeaders);
  for (let r = 1; r < 20; r++) {
    if (cleanHeader(hcData[r]?.[origHCCol]) === "AZ") {
      for (let c = origHCCol + 1; c < hcData[r].length; c++) {
        const wh = cleanHeader(hcHeaders[c]);
        if (wh === "30" || num(wh) === 30) {
          console.log(`AZ width 30: ${num(hcData[r][c])} original HCs`);
        }
      }
      // Print all AZ values
      console.log("AZ all:", hcData[r].slice(origHCCol + 1).map((v, i) => `${widthHeaders[i]}=${num(v)}`).join(", "));
    }
  }
}

// Now check: what if the TRUSS SPACING SECTION uses a DIFFERENT snow code
// Maybe the height prefix changes something?
console.log("\n=== All snow codes in Truss Spacing ===");
const snowCodes = new Set();
for (let r = 1; r < tsData.length; r++) {
  const key = cleanHeader(tsData[r][0]);
  if (key) snowCodes.add(key);
}
console.log([...snowCodes].sort().join(", "));
