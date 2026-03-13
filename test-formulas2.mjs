import XLSX from "xlsx";

const wb = XLSX.readFile("C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx", { cellFormula: true });

function findSheet(name) {
  for (const key of Object.keys(wb.Sheets)) {
    if (key.trim() === name.trim()) return wb.Sheets[key];
  }
  return null;
}

const mathSheet = findSheet("Snow - Math Calculations");
if (!mathSheet) { console.log("Sheet not found"); process.exit(1); }

// Read the FINAL total formulas: AE14:AE17 and AE19
console.log("=== FINAL TOTAL FORMULAS (col AE = 30) ===\n");
for (let r = 10; r <= 22; r++) {
  const cellAddr = XLSX.utils.encode_cell({ r, c: 30 });
  const cell = mathSheet[cellAddr];
  if (cell) {
    console.log(`[${cellAddr}] r=${r}: value=${cell.v}, formula=${cell.f ?? "(static)"}`);
  }
}

// Also check column AB-AF (28-31) for context
console.log("\n=== Columns AB-AF (28-31) rows 10-22 ===");
for (let r = 10; r <= 22; r++) {
  const parts = [];
  for (let c = 27; c <= 35; c++) {
    const cellAddr = XLSX.utils.encode_cell({ r, c });
    const cell = mathSheet[cellAddr];
    if (cell && cell.v !== undefined && cell.v !== "") {
      parts.push(`[${cellAddr}] c=${c}: ${cell.v} (${cell.f ?? "static"})`);
    }
  }
  if (parts.length > 0) console.log(`Row ${r}: ${parts.join(" | ")}`);
}

// Check the "Totals" label at col 30, row 12
console.log("\n=== Looking for Totals section ===");
const range = XLSX.utils.decode_range(mathSheet["!ref"] || "A1");
for (let r = 0; r <= range.e.r; r++) {
  for (let c = 28; c <= 35; c++) {
    const cellAddr = XLSX.utils.encode_cell({ r, c });
    const cell = mathSheet[cellAddr];
    if (cell?.f) {
      console.log(`[${cellAddr}] r=${r} c=${c}: value=${cell.v}, formula=${cell.f}`);
    }
  }
}

// Also check the Snow Breakdown sheet for the final engineering price
console.log("\n=== Looking for Snow Breakdown sheet ===");
for (const name of Object.keys(wb.Sheets)) {
  if (name.toLowerCase().includes("snow") && name.toLowerCase().includes("break")) {
    console.log("Found:", name);
    const ws = wb.Sheets[name];
    const wsRange = XLSX.utils.decode_range(ws["!ref"] || "A1");
    for (let r = 0; r <= wsRange.e.r; r++) {
      for (let c = 0; c <= wsRange.e.c; c++) {
        const cellAddr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[cellAddr];
        if (cell && cell.v !== undefined && cell.v !== "" && cell.v !== 0) {
          console.log(`[${cellAddr}] r=${r} c=${c}: value=${cell.v}, formula=${cell.f ?? "(static)"}`);
        }
      }
    }
  }
}

// Check the Quote Sheet for the engineering price cell
console.log("\n=== Quote Sheet - engineering price area ===");
const quoteSheet = findSheet("Quote Sheet");
if (quoteSheet) {
  // Find cells that reference Math Calculations
  const qRange = XLSX.utils.decode_range(quoteSheet["!ref"] || "A1");
  for (let r = 0; r <= qRange.e.r; r++) {
    for (let c = 0; c <= qRange.e.c; c++) {
      const cellAddr = XLSX.utils.encode_cell({ r, c });
      const cell = quoteSheet[cellAddr];
      if (cell?.f && cell.f.includes("Math")) {
        console.log(`[${cellAddr}] r=${r} c=${c}: value=${cell.v}, formula=${cell.f}`);
      }
    }
  }
  // Also look for "engineer" or "snow" labels
  for (let r = 0; r <= qRange.e.r; r++) {
    for (let c = 0; c <= qRange.e.c; c++) {
      const cellAddr = XLSX.utils.encode_cell({ r, c });
      const cell = quoteSheet[cellAddr];
      if (cell?.v && typeof cell.v === "string" && (cell.v.toLowerCase().includes("engineer") || cell.v.toLowerCase().includes("snow"))) {
        console.log(`[${cellAddr}] r=${r} c=${c}: value="${cell.v}"`);
      }
    }
  }
}
