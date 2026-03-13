import XLSX from "xlsx";

const wb = XLSX.readFile("C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx");

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

// Dump Snow - Changers rows 14-25 (around S/M/T section)
const changersSheet = findSheet("Snow - Changers");
const changersData = sheetToArray(changersSheet);
console.log("=== Snow - Changers: S/M/T Section ===");

// Find S/M/T row first
let smtRow = -1;
for (let r = 0; r < changersData.length; r++) {
  const row = changersData[r];
  let smtCount = 0;
  for (let c = 0; c < (row?.length || 0); c++) {
    const v = cleanHeader(row[c]);
    if (v === "S" || v === "M" || v === "T") smtCount++;
  }
  if (smtCount >= 5) { smtRow = r; break; }
}
console.log("S/M/T row:", smtRow);

// Dump rows smtRow-3 to smtRow+5
for (let r = Math.max(0, smtRow - 3); r <= Math.min(changersData.length - 1, smtRow + 5); r++) {
  const row = changersData[r];
  const vals = [];
  for (let c = 0; c < Math.min(row.length, 25); c++) {
    const v = row[c];
    if (v !== "" && v !== null && v !== undefined) {
      vals.push(`[${c}]=${v}`);
    }
  }
  console.log(`Row ${r}: ${vals.join(" | ")}`);
}

// Now specifically dump heights (row before S/M/T) and the row after with all values
console.log("\n=== Height → S/M/T → feetUsed mapping ===");
const hRow = changersData[smtRow - 1]; // heights
const sRow = changersData[smtRow];     // S/M/T
const fRow = changersData[smtRow + 1]; // feetUsed?
const fRow2 = changersData[smtRow + 2]; // maybe another row?

console.log("Heights (row " + (smtRow - 1) + "):");
for (let c = 0; c < Math.min(hRow.length, 25); c++) {
  if (hRow[c] !== "" && hRow[c] !== null) {
    const h = num(hRow[c]);
    const prefix = cleanHeader(sRow[c]);
    const fu = num(fRow[c]);
    const fu2 = num(fRow2?.[c]);
    if (h >= 6 && h <= 20) {
      console.log(`  Height ${h}: prefix=${prefix}, row+1=${fu}, row+2=${fu2}`);
    }
  }
}

// Also check what row smtRow+1 label is
console.log("\nRow labels around S/M/T:");
console.log(`  Row ${smtRow-1} label: "${cleanHeader(changersData[smtRow-1]?.[0])}"`);
console.log(`  Row ${smtRow} label: "${cleanHeader(changersData[smtRow]?.[0])}"`);
console.log(`  Row ${smtRow+1} label: "${cleanHeader(changersData[smtRow+1]?.[0])}"`);
console.log(`  Row ${smtRow+2} label: "${cleanHeader(changersData[smtRow+2]?.[0])}"`);
console.log(`  Row ${smtRow+3} label: "${cleanHeader(changersData[smtRow+3]?.[0])}"`);

// Dump the FULL Math Calculations with ALL columns (not just non-empty)
console.log("\n\n=== Snow - Math Calculations: FULL DUMP (rows 11-34, cols 12-32) ===");
const mathSheet = findSheet("Snow - Math Calculations");
if (mathSheet) {
  const mData = sheetToArray(mathSheet);
  for (let r = 11; r <= Math.min(34, mData.length - 1); r++) {
    const vals = [];
    for (let c = 12; c <= Math.min(32, (mData[r]?.length || 0) - 1); c++) {
      const v = mData[r]?.[c];
      if (v !== "" && v !== null && v !== undefined && v !== 0) {
        vals.push(`[${c}]=${v}`);
      }
    }
    if (vals.length > 0) console.log(`Row ${r}: ${vals.join(" | ")}`);
  }

  // Also dump rows 17-21 (vertical pricing) with labels
  console.log("\n=== Vertical Pricing Detail (cols 17-21) ===");
  for (let r = 11; r <= 25; r++) {
    const label = mData[r]?.[17] || "";
    const val = mData[r]?.[20] || "";
    const val2 = mData[r]?.[19] || "";
    const val3 = mData[r]?.[21] || "";
    if (label !== "" || val !== "" || val2 !== "") {
      console.log(`  Row ${r}: label="${label}" val19=${val2} val20=${val} val21=${val3}`);
    }
  }

  // Dump the girt pricing section (cols 17-21, rows 23-30)
  console.log("\n=== Girt Pricing Detail (cols 17-21, rows 23-30) ===");
  for (let r = 23; r <= 30; r++) {
    const vals = [];
    for (let c = 17; c <= 21; c++) {
      vals.push(`[${c}]=${mData[r]?.[c] ?? ""}`);
    }
    console.log(`  Row ${r}: ${vals.join(" | ")}`);
  }

  // Dump the total section (cols 27-32, rows 27-34)
  console.log("\n=== Total Section (cols 27-32) ===");
  for (let r = 27; r <= 34; r++) {
    const vals = [];
    for (let c = 27; c <= 32; c++) {
      const v = mData[r]?.[c];
      if (v !== "" && v !== null && v !== undefined) {
        vals.push(`[${c}]=${v}`);
      }
    }
    if (vals.length > 0) console.log(`  Row ${r}: ${vals.join(" | ")}`);
  }
}
