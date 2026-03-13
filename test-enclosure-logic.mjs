import XLSX from "xlsx";

const wb = XLSX.readFile("C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx", { cellFormula: true });

function findSheet(name) {
  for (const key of Object.keys(wb.Sheets)) {
    if (key.trim() === name.trim()) return wb.Sheets[key];
  }
  return null;
}

// === 1. Check Pricing - Changers D66 and surrounding ===
console.log("=== Pricing - Changers: D66 (enclosure determination) ===");
const pc = findSheet("Pricing - Changers");
for (let r = 60; r <= 72; r++) {
  for (let c = 0; c <= 25; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = pc[addr];
    if (cell && (cell.v !== undefined && cell.v !== "")) {
      console.log(`  [${addr}] r=${r} c=${c}: value=${JSON.stringify(cell.v)}, formula=${cell.f ?? "(static)"}`);
    }
  }
}

// === 2. Check what U69 is (the V-side girt flag) ===
console.log("\n=== Pricing - Changers: U69 (V-side girt flag) ===");
for (let r = 64; r <= 72; r++) {
  for (let c = 18; c <= 25; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = pc[addr];
    if (cell && (cell.v !== undefined && cell.v !== "")) {
      console.log(`  [${addr}] r=${r} c=${c}: value=${JSON.stringify(cell.v)}, formula=${cell.f ?? "(static)"}`);
    }
  }
}

// === 3. Check what Quote Sheet cells determine sides/ends/panel type ===
console.log("\n=== Quote Sheet: Sides/Ends configuration (rows 13-18) ===");
const qs = findSheet("Quote Sheet");
for (let r = 12; r <= 18; r++) {
  for (let c = 0; c <= 35; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = qs[addr];
    if (cell && cell.v !== undefined && cell.v !== "" && cell.v !== 0) {
      console.log(`  [${addr}] r=${r} c=${c}: value=${JSON.stringify(cell.v)}, formula=${cell.f ?? "(static)"}`);
    }
  }
}

// === 4. Check what Pricing - Changers B30 and B32 are (roof style) ===
console.log("\n=== Pricing - Changers: Roof/Config section (rows 28-36) ===");
for (let r = 28; r <= 36; r++) {
  for (let c = 0; c <= 10; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = pc[addr];
    if (cell && (cell.v !== undefined && cell.v !== "")) {
      console.log(`  [${addr}] r=${r} c=${c}: value=${JSON.stringify(cell.v)}, formula=${cell.f ?? "(static)"}`);
    }
  }
}

// === 5. Check what E6 is (width) ===
console.log("\n=== Pricing - Changers: Width/config (rows 4-8) ===");
for (let r = 4; r <= 8; r++) {
  for (let c = 0; c <= 10; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = pc[addr];
    if (cell && (cell.v !== undefined && cell.v !== "")) {
      console.log(`  [${addr}] r=${r} c=${c}: value=${JSON.stringify(cell.v)}, formula=${cell.f ?? "(static)"}`);
    }
  }
}

// === 6. Specifically trace the D66 formula chain ===
console.log("\n=== D66 formula chain - tracing back ===");
// D66 determines E/O. Let's see what it depends on
const d66 = pc[XLSX.utils.encode_cell({ r: 65, c: 3 })];
console.log(`D66: value=${JSON.stringify(d66?.v)}, formula=${d66?.f}`);

// Check if there's a lookup table for enclosure
console.log("\n=== Pricing - Changers: Rows 50-68 (enclosure section) ===");
for (let r = 50; r <= 68; r++) {
  for (let c = 0; c <= 12; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = pc[addr];
    if (cell && (cell.v !== undefined && cell.v !== "")) {
      console.log(`  [${addr}] r=${r} c=${c}: value=${JSON.stringify(cell.v)}, formula=${cell.f ?? "(static)"}`);
    }
  }
}

// === 7. Check Truss Spacing sheet's enclosure section (rows 53-74 referenced in it) ===
console.log("\n=== Truss Spacing: Enclosure lookup table (rows 53-65) ===");
const ts = findSheet("Snow - Truss Spacing");
for (let r = 52; r <= 75; r++) {
  for (let c = 14; c <= 21; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = ts[addr];
    if (cell && (cell.v !== undefined && cell.v !== "")) {
      console.log(`  [${addr}] r=${r} c=${c}: value=${JSON.stringify(cell.v)}, formula=${cell.f ?? "(static)"}`);
    }
  }
}

// === 8. Specifically check what "Fully Enclosed" means in context ===
// The Truss Spacing sheet at O67="Fully Enclosed" and surrounding rows
console.log("\n=== Truss Spacing: Enclosure determination (rows 63-75) ===");
for (let r = 63; r <= 75; r++) {
  for (let c = 14; c <= 21; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = ts[addr];
    if (cell && cell.v !== undefined && cell.v !== "") {
      console.log(`  [${addr}] r=${r} c=${c}: value=${JSON.stringify(cell.v)}, formula=${cell.f ?? "(static)"}`);
    }
  }
}

// === 9. Check the full "Fully Enclosed" lookup table in Truss Spacing ===
console.log("\n=== Truss Spacing: Full enclosure table (rows 53-65, cols O-U) ===");
for (let r = 52; r <= 65; r++) {
  const parts = [];
  for (let c = 14; c <= 20; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = ts[addr];
    if (cell && cell.v !== undefined) parts.push(`${String.fromCharCode(65+c)}${r+1}=${JSON.stringify(cell.v)}`);
  }
  if (parts.length) console.log(`  Row ${r+1}: ${parts.join(" | ")}`);
}
