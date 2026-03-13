import XLSX from "xlsx";

const wb = XLSX.readFile("C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx", { cellFormula: true });

function findSheet(name) {
  for (const key of Object.keys(wb.Sheets)) {
    if (key.trim() === name.trim()) return wb.Sheets[key];
  }
  return null;
}

const changers = findSheet("Snow - Changers");

// Read G76 (truss price) and surrounding cells
console.log("=== Snow - Changers: Computed result cells ===");
// Check rows 70-80, cols F-K (5-10)
for (let r = 68; r <= 82; r++) {
  const parts = [];
  for (let c = 0; c <= 12; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = changers[addr];
    if (cell && cell.v !== undefined && cell.v !== "") {
      parts.push(`[${String.fromCharCode(65+c)}${r+1}] c=${c}: v=${JSON.stringify(cell.v)} f=${cell.f ?? "-"}`);
    }
  }
  if (parts.length > 0) console.log(`Row ${r}: ${parts.join("\n  ")}`);
}

// Specifically check G76 (the truss price ref from Math Calculations)
console.log("\n=== Key formula cells ===");
const keyCells = [
  { name: "G76 (Truss Price)", r: 75, c: 6 },
  { name: "J72 (Pie Truss Price/ft)", r: 71, c: 9 },
  { name: "J75 (Channel Price)", r: 74, c: 9 },
  { name: "J76 (Tubing Price)", r: 75, c: 9 },
  { name: "G32 (feetUsed)", r: 31, c: 6 },
  { name: "D54 (width ref)", r: 53, c: 3 },
  { name: "D29 (height ref)", r: 28, c: 3 },
];
for (const { name, r, c } of keyCells) {
  const addr = XLSX.utils.encode_cell({ r, c });
  const cell = changers[addr];
  console.log(`${name} [${addr}]: value=${JSON.stringify(cell?.v)}, formula=${cell?.f ?? "(static)"}`);
}

// Check the section around row 30 (G32 formula chain)
console.log("\n=== feetUsed section (rows 28-35, cols D-H) ===");
for (let r = 28; r <= 35; r++) {
  const parts = [];
  for (let c = 3; c <= 10; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = changers[addr];
    if (cell && cell.v !== undefined && cell.v !== "") {
      parts.push(`[${String.fromCharCode(65+c)}${r+1}] c=${c}: v=${JSON.stringify(cell.v)} f=${cell.f ?? "-"}`);
    }
  }
  if (parts.length > 0) console.log(`Row ${r}: ${parts.join(" | ")}`);
}

// Also check the truss price lookup section (around G76)
// G76 references via INDEX/MATCH or similar formula
console.log("\n=== Rows 72-80, all cols 0-16 ===");
for (let r = 72; r <= 80; r++) {
  const parts = [];
  for (let c = 0; c <= 16; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = changers[addr];
    if (cell && cell.v !== undefined && cell.v !== "") {
      parts.push(`${String.fromCharCode(65+c)}${r+1}:${JSON.stringify(cell.v)}(${cell.f ?? "s"})`);
    }
  }
  if (parts.length > 0) console.log(`  Row ${r}: ${parts.join(" | ")}`);
}

// Also read the "height multiplier" area if there is one
// Check if there's a multiplier applied to truss price via G76
console.log("\n=== Any formulas with height/multiplier references ===");
const range = XLSX.utils.decode_range(changers["!ref"] || "A1");
for (let r = 0; r <= range.e.r; r++) {
  for (let c = 0; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = changers[addr];
    if (cell?.f && (cell.f.includes("Quote") && cell.f.includes("R10"))) {
      console.log(`[${addr}] r=${r} c=${c}: value=${cell.v}, formula=${cell.f}`);
    }
  }
}

// Check for any IF statements in the G70-G80 range
console.log("\n=== Column G formulas rows 70-82 ===");
for (let r = 70; r <= 82; r++) {
  const addr = XLSX.utils.encode_cell({ r, c: 6 });
  const cell = changers[addr];
  if (cell?.f) {
    console.log(`[${addr}] r=${r}: value=${cell.v}, formula=${cell.f}`);
  }
}
