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

// Check truss spacing for T-47LL at various 30-width configs
const ts = findSheet("Snow - Truss Spacing");
const tsData = sheetToArray(ts);
const tsHeaders = tsData[0];

// Find columns with "30" width
console.log("=== Truss Spacing columns with width 30 ===");
const cols30 = {};
for (let c = 1; c < tsHeaders.length; c++) {
  const h = cleanHeader(tsHeaders[c]);
  if (h.includes("-30-")) {
    cols30[h] = c;
  }
}
console.log("Columns:", Object.keys(cols30).join(", "));

// Find rows for T-47LL and related
console.log("\n=== Snow code rows ===");
const targetRows = {};
for (let r = 1; r < tsData.length; r++) {
  const key = cleanHeader(tsData[r][0]);
  if (key.includes("47")) {
    targetRows[key] = r;
    console.log(`  Row ${r}: "${key}"`);
  }
}

// Also check all T- rows
console.log("\n=== All T- rows ===");
for (let r = 1; r < tsData.length; r++) {
  const key = cleanHeader(tsData[r][0]);
  if (key.startsWith("T-")) {
    console.log(`  Row ${r}: "${key}"`);
    if (!targetRows[key]) targetRows[key] = r;
  }
}

// Look up values for T-47LL at 30-width configs (especially with 155 wind)
console.log("\n=== T-47LL values at width-30 configs ===");
for (const [colName, col] of Object.entries(cols30)) {
  for (const [rowName, row] of Object.entries(targetRows)) {
    if (rowName.includes("47")) {
      const val = num(tsData[row][col]);
      if (val > 0) console.log(`  ${rowName} @ ${colName} = ${val}`);
    }
  }
}

// Check specifically E-155-30-AFV
console.log("\n=== Specific: E-155-30-AFV column ===");
for (let c = 1; c < tsHeaders.length; c++) {
  const h = cleanHeader(tsHeaders[c]);
  if (h === "E-155-30-AFV") {
    console.log(`  Found at column ${c}`);
    // Show all T- values
    for (let r = 1; r < tsData.length; r++) {
      const key = cleanHeader(tsData[r][0]);
      if (key.startsWith("T-")) {
        console.log(`    ${key} = ${num(tsData[r][c])}`);
      }
    }
    break;
  }
}

// Also check wind 150 bucketing
console.log("\n=== Wind bucketing around 150 ===");
const ch = findSheet("Snow - Changers");
const chData = sheetToArray(ch);
// Wind buckets are in rows 0-6
for (let r = 0; r <= 6; r++) {
  const parts = [];
  for (let c = 0; c <= 20; c++) {
    const v = chData[r][c];
    if (v !== "" && v !== undefined) parts.push(`c${c}=${JSON.stringify(v)}`);
  }
  if (parts.length) console.log(`  Row ${r}: ${parts.join(" | ")}`);
}

// Check what snow code "47 Roof Load" maps to
console.log("\n=== Snow load code mapping ===");
for (let r = 8; r <= 14; r++) {
  const parts = [];
  for (let c = 0; c <= 10; c++) {
    const v = chData[r][c];
    if (v !== "" && v !== undefined) parts.push(`c${c}=${JSON.stringify(v)}`);
  }
  if (parts.length) console.log(`  Row ${r}: ${parts.join(" | ")}`);
}

// Check Changers width bucketing for width 30
console.log("\n=== Width bucketing ===");
for (let r = 48; r <= 54; r++) {
  const parts = [];
  for (let c = 0; c <= 10; c++) {
    const v = chData[r][c];
    if (v !== "" && v !== undefined) parts.push(`c${c}=${JSON.stringify(v)}`);
  }
  if (parts.length) console.log(`  Row ${r}: ${parts.join(" | ")}`);
}
