import XLSX from "xlsx";

const wb = XLSX.readFile("C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx", { cellFormula: true });

function findSheet(name) {
  for (const key of Object.keys(wb.Sheets)) {
    if (key.trim() === name.trim()) return wb.Sheets[key];
  }
  return null;
}

const mathSheet = findSheet("Snow - Math Calculations");

// Read X6 and surrounding cells (the TRUSS CHARGE area)
console.log("=== TRUSS CHARGE area (cols W-Z = 22-25, rows 5-8) ===");
for (let r = 4; r <= 8; r++) {
  for (let c = 22; c <= 25; c++) {
    const cellAddr = XLSX.utils.encode_cell({ r, c });
    const cell = mathSheet[cellAddr];
    if (cell) {
      console.log(`[${cellAddr}] r=${r} c=${c}: value=${JSON.stringify(cell.v)}, formula=${cell.f ?? "(static)"}`);
    }
  }
}

// Also read the full chain: P22 (Extra Price for Trusses) which feeds into X6
console.log("\n=== Truss pricing chain (col P=15, rows 13-22) ===");
for (let r = 12; r <= 22; r++) {
  const cellAddr = XLSX.utils.encode_cell({ r, c: 15 });
  const cell = mathSheet[cellAddr];
  if (cell) {
    console.log(`[${cellAddr}] r=${r}: value=${JSON.stringify(cell.v)}, formula=${cell.f ?? "(static)"}`);
  }
}

// Read X column (col 23) completely
console.log("\n=== Column X (col 23) all rows ===");
const range = XLSX.utils.decode_range(mathSheet["!ref"] || "A1");
for (let r = 0; r <= range.e.r; r++) {
  const cellAddr = XLSX.utils.encode_cell({ r, c: 23 });
  const cell = mathSheet[cellAddr];
  if (cell && cell.v !== undefined && cell.v !== "") {
    console.log(`[${cellAddr}] r=${r}: value=${JSON.stringify(cell.v)}, formula=${cell.f ?? "(static)"}`);
  }
}

// Read the vertical pricing complete chain
console.log("\n=== Vertical pricing chain (col U=20, rows 11-22) ===");
for (let r = 11; r <= 22; r++) {
  const cellAddr = XLSX.utils.encode_cell({ r, c: 20 });
  const cell = mathSheet[cellAddr];
  if (cell) {
    console.log(`[${cellAddr}] r=${r}: value=${JSON.stringify(cell.v)}, formula=${cell.f ?? "(static)"}`);
  }
}

// Also check U13 (Extra Verticals value)
console.log("\n=== Extra counts chain ===");
// P13 = Extra Trusses
const p13 = mathSheet[XLSX.utils.encode_cell({ r: 12, c: 15 })];
console.log("P13 (Extra Trusses):", JSON.stringify(p13?.v), "formula:", p13?.f);

// G7 = Extra Trusses (from the main calc section)
const g7 = mathSheet[XLSX.utils.encode_cell({ r: 6, c: 6 })];
console.log("G7 (Extra Trusses Making Sure):", JSON.stringify(g7?.v), "formula:", g7?.f);

// D7 = Extra Trusses raw
const d7 = mathSheet[XLSX.utils.encode_cell({ r: 6, c: 3 })];
console.log("D7 (Extra Trusses raw):", JSON.stringify(d7?.v), "formula:", d7?.f);

// H7 (use this)
const h7 = mathSheet[XLSX.utils.encode_cell({ r: 6, c: 7 })];
console.log("H7 (Use This):", JSON.stringify(h7?.v), "formula:", h7?.f);

// Also read the girts "Use This" value
const h26 = mathSheet[XLSX.utils.encode_cell({ r: 25, c: 7 })];
console.log("H26 (Girt Use This):", JSON.stringify(h26?.v), "formula:", h26?.f);

// P13 formula
const p13cell = mathSheet[XLSX.utils.encode_cell({ r: 12, c: 15 })];
console.log("P13 formula:", p13cell?.f);

// T2 (original trusses ref from Breakdown)
const t2 = mathSheet[XLSX.utils.encode_cell({ r: 1, c: 19 })];
console.log("T2 (Original Trusses):", JSON.stringify(t2?.v), "formula:", t2?.f);

// Now find the Quote Sheet engineering value cell
const quoteSheet = findSheet("Quote Sheet");
// Search for cells referencing Snow - Math Calculations
console.log("\n=== Quote Sheet: cells referencing Math Calculations ===");
const qRange = XLSX.utils.decode_range(quoteSheet["!ref"] || "A1");
for (let r = 0; r <= qRange.e.r; r++) {
  for (let c = 0; c <= qRange.e.c; c++) {
    const cellAddr = XLSX.utils.encode_cell({ r, c });
    const cell = quoteSheet[cellAddr];
    if (cell?.f && cell.f.includes("Math")) {
      console.log(`[${cellAddr}] r=${r} c=${c}: value=${JSON.stringify(cell.v)}, formula=${cell.f}`);
    }
  }
}

// Also find any cell near "Engineering" label
console.log("\n=== Quote Sheet: Engineering area ===");
for (let r = 20; r <= 30; r++) {
  for (let c = 0; c <= 30; c++) {
    const cellAddr = XLSX.utils.encode_cell({ r, c });
    const cell = quoteSheet[cellAddr];
    if (cell && cell.v !== undefined && cell.v !== "" && cell.v !== 0) {
      console.log(`[${cellAddr}] r=${r} c=${c}: value=${JSON.stringify(cell.v)}, formula=${cell.f ?? "(static)"}`);
    }
  }
}
