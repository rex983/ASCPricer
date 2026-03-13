import XLSX from "xlsx";

const wb = XLSX.readFile("C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx", { cellFormula: true });

function findSheet(name) {
  for (const key of Object.keys(wb.Sheets)) {
    if (key.trim() === name.trim()) return wb.Sheets[key];
  }
  return null;
}

const ms = findSheet("Snow - Math Calculations");

// Read EVERY cell in the verticals section (rows 27-35, cols 0-9)
console.log("=== Verticals section (rows 27-35, all cols 0-9) ===");
for (let r = 27; r <= 35; r++) {
  const parts = [];
  for (let c = 0; c <= 9; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = ms[addr];
    if (cell) {
      parts.push(`[${addr}] c=${c}: v=${JSON.stringify(cell.v)} f=${cell.f ?? "-"}`);
    }
  }
  if (parts.length > 0) console.log(`Row ${r}:\n  ${parts.join("\n  ")}`);
}

// Read D35, H32, and the full vertical extra chain
console.log("\n=== Key vertical cells ===");
const cells = [
  "D29", "D30", "D31", "D32", "D33", "D34", "D35",
  "G30", "G31", "G32",
  "H30", "H31", "H32",
  "I33", "I34", "I35",
  "U12", "U13",
];
for (const addr of cells) {
  const cell = ms[addr];
  if (cell) {
    console.log(`${addr}: value=${JSON.stringify(cell.v)}, formula=${cell.f ?? "(static)"}`);
  } else {
    console.log(`${addr}: (empty)`);
  }
}

// Now check what the FINAL engineering price formula is on the Quote Sheet
console.log("\n=== Quote Sheet: row 25 (Engineering line) all cols ===");
const qs = findSheet("Quote Sheet");
for (let r = 22; r <= 26; r++) {
  for (let c = 0; c <= 35; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = qs[addr];
    if (cell && cell.v !== undefined && cell.v !== "" && cell.v !== 0) {
      console.log(`[${addr}] r=${r} c=${c}: value=${JSON.stringify(cell.v)}, formula=${cell.f ?? "(static)"}`);
    }
  }
}

// Also check AE column (col 30) for the subtotal that feeds into "Sales Total"
console.log("\n=== Quote Sheet: AE column (col 30) rows 10-30 ===");
for (let r = 10; r <= 30; r++) {
  const addr = XLSX.utils.encode_cell({ r, c: 30 });
  const cell = qs[addr];
  if (cell && cell.v !== undefined && cell.v !== "" && cell.v !== 0) {
    console.log(`[${addr}] r=${r}: value=${JSON.stringify(cell.v)}, formula=${cell.f ?? "(static)"}`);
  }
}

// Check if there's a ×6 multiplier anywhere
console.log("\n=== Looking for ×6 or *6 in formulas ===");
const range = XLSX.utils.decode_range(ms["!ref"] || "A1");
for (let r = 0; r <= range.e.r; r++) {
  for (let c = 0; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = ms[addr];
    if (cell?.f && (cell.f.includes("*6") || cell.f.includes("× 6") || cell.f.includes("x6"))) {
      console.log(`[${addr}] r=${r} c=${c}: value=${cell.v}, formula=${cell.f}`);
    }
  }
}

// Check X7 (TRUSS CHARGE) formula - this uses ChangersG76 × extraTrusses
// And there's a note "Returned to x6" — let me check if "x6" means cell X6
console.log("\n=== X7 and surrounding ===");
for (let r = 4; r <= 8; r++) {
  for (let c = 22; c <= 26; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = ms[addr];
    if (cell) {
      console.log(`[${addr}] r=${r} c=${c}: value=${JSON.stringify(cell.v)}, formula=${cell.f ?? "(static)"}`);
    }
  }
}

// Now, the crucial question: does the TRUSS total (AE14 = X6 = P22)
// include a height multiplier like the verticals do?
// X6 = IF(P2=0, "Contact", P22)
// P22 = P15 + P21 = (extraTrusses × basePrice) + (extraTrusses × legSurcharge)
// There's NO height multiplier on trusses in the formula chain.

// But wait - check if X7 (TRUSS CHARGE) is used somewhere instead of X6
console.log("\n=== Searching for X7 references ===");
for (let r = 0; r <= range.e.r; r++) {
  for (let c = 0; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = ms[addr];
    if (cell?.f && (cell.f.includes("X7") || cell.f.includes("$X$7"))) {
      console.log(`[${addr}] r=${r} c=${c}: value=${cell.v}, formula=${cell.f}`);
    }
  }
}

// Also look on Quote Sheet for X7 references
const qRange = XLSX.utils.decode_range(qs["!ref"] || "A1");
for (let r = 0; r <= qRange.e.r; r++) {
  for (let c = 0; c <= qRange.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = qs[addr];
    if (cell?.f && (cell.f.includes("X7") || cell.f.includes("$X$7") || cell.f.includes("X6") || cell.f.includes("$X$6"))) {
      console.log(`Quote[${addr}] r=${r} c=${c}: value=${cell.v}, formula=${cell.f}`);
    }
  }
}

// Check Snow Load Breakdown for price formulas
const slb = findSheet("Snow Load Breakdown");
if (slb) {
  console.log("\n=== Snow Load Breakdown: Price cells ===");
  const slbRange = XLSX.utils.decode_range(slb["!ref"] || "A1");
  for (let r = 0; r <= slbRange.e.r; r++) {
    for (let c = 0; c <= slbRange.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = slb[addr];
      if (cell?.f && cell.f.includes("Math")) {
        console.log(`[${addr}] r=${r} c=${c}: value=${JSON.stringify(cell.v)}, formula=${cell.f}`);
      }
    }
  }
}
