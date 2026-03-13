import XLSX from "xlsx";

// Read with cellFormula option
const wb = XLSX.readFile("C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx", { cellFormula: true });

function findSheet(name) {
  for (const key of Object.keys(wb.Sheets)) {
    if (key.trim() === name.trim()) return wb.Sheets[key];
  }
  return null;
}

// Read formulas from Snow - Math Calculations
const mathSheet = findSheet("Snow - Math Calculations");
if (mathSheet) {
  console.log("=== Math Calculations FORMULAS (key cells) ===\n");

  // Key formula cells to check:
  // Row 19 (Extra Price Per Truss) = Col O (15) → row 20 in 1-based = row 19 in 0-based
  // Row 16 (Extra Leg Height) = Col P (15)
  // Row 20 (Total Price for trusses)
  // Row 18 (Total Price for Vertical)

  const keyRows = [
    { label: "Extra Trusses Need", r: 12, c: 12 },  // M13
    { label: "Truss Pricing", r: 13, c: 15 },        // P14
    { label: "Pricing for Trusses Only", r: 14, c: 15 }, // P15
    { label: "Extra Leg Height from 6'", r: 16, c: 15 }, // P17
    { label: "if Negative (leg)", r: 17, c: 15 },    // P18
    { label: "Extra Leg Height", r: 18, c: 15 },     // P19
    { label: "Extra Price Per Truss", r: 19, c: 15 }, // P20
    { label: "Total Truss Price", r: 20, c: 15 },    // P21
    { label: "Extra Price for Trusses", r: 21, c: 15 }, // P22
    { label: "Extra Verticals", r: 12, c: 17 },      // R13
    { label: "Peak Height (rise)", r: 13, c: 20 },   // U14
    { label: "Peak Height (total)", r: 15, c: 20 },  // U16
    { label: "Tubing Pricing", r: 16, c: 20 },       // U17
    { label: "Price Per Vertical", r: 17, c: 20 },   // U18
    { label: "Total Price for Vertical", r: 18, c: 20 }, // U19
    { label: "Double vert Price 13,14,15", r: 19, c: 20 }, // U20
    { label: "Price per vertx2.5 on 16-18", r: 20, c: 20 }, // U21
    { label: "price 19-20 *3", r: 21, c: 20 },       // U22
    { label: "Extra Channels Needed", r: 24, c: 15 }, // P25
    { label: "Channel Price Per Foot", r: 25, c: 15 }, // P26
    { label: "Price Per Channel", r: 27, c: 15 },     // P28
    { label: "Total Channel Price", r: 28, c: 15 },   // P29
    { label: "Girt Width", r: 24, c: 19 },            // T25
    { label: "Girt Length", r: 25, c: 19 },           // T26
    { label: "Girt Perimeter", r: 26, c: 19 },        // T27
    { label: "Total Girt Feet", r: 27, c: 19 },       // T28
    { label: "Girt Tubing Price", r: 28, c: 19 },     // T29
    { label: "Total Girt Price", r: 29, c: 19 },      // T30
    { label: "Sides Perimeter", r: 28, c: 21 },       // V29
    { label: "Ends Perimeter", r: 29, c: 21 },        // V30
    { label: "Tubing Used (for leg)", r: 13, c: 25 }, // Z14
    { label: "Leg Height", r: 12, c: 25 },            // Z13
    { label: "Total Material - Trusses", r: 20, c: 25 }, // Z21
    { label: "Total Material - Vertical", r: 21, c: 25 }, // Z22
    { label: "Total Uprights", r: 22, c: 25 },        // Z23
  ];

  for (const { label, r, c } of keyRows) {
    const cellAddr = XLSX.utils.encode_cell({ r, c });
    const cell = mathSheet[cellAddr];
    const val = cell?.v ?? "";
    const formula = cell?.f ?? "(no formula)";
    console.log(`${label} [${cellAddr}]: value=${val}, formula=${formula}`);
  }

  // Also check the "Total Price" section
  console.log("\n=== Total Price Section ===");
  for (let r = 27; r <= 34; r++) {
    for (let c = 27; c <= 35; c++) {
      const cellAddr = XLSX.utils.encode_cell({ r, c });
      const cell = mathSheet[cellAddr];
      if (cell && (cell.v !== "" && cell.v !== 0 && cell.v !== undefined)) {
        console.log(`[${cellAddr}] r=${r} c=${c}: value=${cell.v}, formula=${cell.f ?? "(static)"}`);
      }
    }
  }

  // Check rows 30-40 for any final calculation
  console.log("\n=== Rows 30-45 ALL cols ===");
  for (let r = 30; r <= 45; r++) {
    for (let c = 0; c <= 35; c++) {
      const cellAddr = XLSX.utils.encode_cell({ r, c });
      const cell = mathSheet[cellAddr];
      if (cell && cell.v !== "" && cell.v !== undefined && cell.v !== 0) {
        console.log(`[${cellAddr}] r=${r} c=${c}: value=${cell.v}, formula=${cell.f ?? "(static)"}`);
      }
    }
  }

  // Check the VERY LAST section - the return value
  console.log("\n=== Looking for TOTAL/SUM/RETURN formulas ===");
  const range = XLSX.utils.decode_range(mathSheet["!ref"] || "A1");
  for (let r = 0; r <= range.e.r; r++) {
    for (let c = 0; c <= range.e.c; c++) {
      const cellAddr = XLSX.utils.encode_cell({ r, c });
      const cell = mathSheet[cellAddr];
      if (cell?.f && (cell.f.includes("SUM") || cell.f.includes("sum"))) {
        console.log(`SUM formula at [${cellAddr}] r=${r} c=${c}: value=${cell.v}, formula=${cell.f}`);
      }
    }
  }
}
