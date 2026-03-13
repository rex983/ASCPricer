import XLSX from "xlsx";
const wb = XLSX.readFile("C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx", { cellFormula: true });
function findSheet(name) {
  for (const key of Object.keys(wb.Sheets)) {
    if (key.trim() === name.trim()) return wb.Sheets[key];
  }
  return null;
}

// Check Snow Load Breakdown K13 formula
console.log("=== Snow Load Breakdown: Contact Engineering check ===");
const slb = findSheet("Snow Load Breakdown");
for (let r = 10; r <= 25; r++) {
  for (let c = 0; c <= 15; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = slb[addr];
    if (cell?.f) {
      console.log(`  [${addr}] value=${JSON.stringify(cell.v)}, formula=${cell.f}`);
    }
  }
}

// Check Math Calculations AD20/AC19/AC14
console.log("\n=== Math Calculations: Contact Engineering formulas ===");
const mc = findSheet("Snow - Math Calculations");
// Check AC and AD columns (cols 28-29) around rows 14-22
for (let r = 12; r <= 22; r++) {
  for (let c = 26; c <= 32; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = mc[addr];
    if (cell && (cell.f || (cell.v !== undefined && cell.v !== ""))) {
      console.log(`  [${addr}] value=${JSON.stringify(cell.v)}, formula=${cell.f ?? "(static)"}`);
    }
  }
}

// Also check P2 (the truss spacing in Math Calc)
console.log("\n=== Math Calculations: P2 (truss spacing) ===");
const p2 = mc[XLSX.utils.encode_cell({ r: 1, c: 15 })];
console.log(`  P2: value=${JSON.stringify(p2?.v)}, formula=${p2?.f ?? "(static)"}`);

// Check rows 0-5 of Math Calc for truss spacing reference
for (let r = 0; r <= 5; r++) {
  for (let c = 14; c <= 20; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = mc[addr];
    if (cell && (cell.f || (cell.v !== undefined && cell.v !== ""))) {
      console.log(`  [${addr}] value=${JSON.stringify(cell.v)}, formula=${cell.f ?? "(static)"}`);
    }
  }
}
