import XLSX from "xlsx";

const wb = XLSX.readFile("C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx");
console.log("Sheets:", wb.SheetNames.map(s => `"${s}"`).join(", "));

// Inline helpers
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
  if (wb.Sheets[name]) return wb.Sheets[name];
  if (wb.Sheets[name + " "]) return wb.Sheets[name + " "];
  for (const key of Object.keys(wb.Sheets)) {
    if (key.trim() === name.trim()) return wb.Sheets[key];
  }
  return null;
}

// Read truss counts
const trussSheet = findSheet("Snow - Trusses");
console.log("\nTruss sheet found:", !!trussSheet);
if (trussSheet) {
  const data = sheetToArray(trussSheet);
  const headers = data[0] || [];
  // Find 24-AZ column
  let azCol = -1;
  for (let c = 1; c < headers.length; c++) {
    if (cleanHeader(headers[c]) === "24-AZ") { azCol = c; break; }
  }
  console.log("24-AZ column:", azCol);
  // Find length 50 row
  for (let r = 1; r < data.length; r++) {
    if (num(data[r][0]) === 50) {
      console.log("Trusses at 24-AZ/50:", num(data[r][azCol]));
      break;
    }
  }
}

// Read girts
const girtsSheet = findSheet("Snow - Girts");
console.log("\nGirts sheet found:", !!girtsSheet);
if (girtsSheet) {
  const data = sheetToArray(girtsSheet);
  console.log("Row 0:", data[0]?.slice(0,8).map(v=>cleanHeader(v)));
  console.log("Row 1 (60):", data[1]?.slice(0,8).map(v=>cleanHeader(v)));
  // Find original girts
  for (let r = 0; r < data.length; r++) {
    const label = cleanHeader(data[r]?.[0]);
    if (label === "" || label === "0") {
      // Check right section
      const rightVal = cleanHeader(data[r]?.[11]);
      if (num(rightVal) === 10) {
        console.log("Height 10 girts:", num(data[r]?.[12]));
      }
    }
  }
}

// Read verticals
const vertSheet = findSheet("Snow - Verticals");
console.log("\nVerticals sheet found:", !!vertSheet);
if (vertSheet) {
  const data = sheetToArray(vertSheet);
  // Find "Original" row
  for (let r = 0; r < data.length; r++) {
    const label = cleanHeader(data[r]?.[0]);
    if (label.toLowerCase().includes("original")) {
      console.log("Original Verticals at row:", r);
      const countRow = data[r + 1];
      if (countRow) {
        const widths = data[r];
        for (let c = 1; c < widths.length; c++) {
          if (num(widths[c]) === 24) {
            console.log("Width 24 original verts:", num(countRow[c]));
          }
        }
      }
      break;
    }
  }
  // Check spacing for height 10 / wind 105
  const headers = data[0];
  let col10 = -1;
  for (let c = 1; c < headers.length; c++) {
    if (num(headers[c]) === 10) { col10 = c; break; }
  }
  for (let r = 1; r < data.length; r++) {
    if (num(data[r][0]) === 105) {
      console.log("Vert spacing height=10 wind=105:", num(data[r][col10]));
      break;
    }
  }
}

// Read Snow - Changers for state pricing
const changersSheet = findSheet("Snow - Changers");
console.log("\nChangers sheet found:", !!changersSheet);
if (changersSheet) {
  const data = sheetToArray(changersSheet);
  // Find S/M/T row
  for (let r = 15; r < 35; r++) {
    const row = data[r];
    if (!row) continue;
    let smtCount = 0;
    for (let c = 0; c < row.length; c++) {
      const v = cleanHeader(row[c]);
      if (v === "S" || v === "M" || v === "T") smtCount++;
    }
    if (smtCount >= 5) {
      console.log("S/M/T row found at:", r);
      const hRow = data[r-1];
      const fRow = data[r+1];
      // Find height 10
      for (let c = 0; c < row.length; c++) {
        if (num(hRow?.[c]) === 10) {
          console.log("Height 10: prefix=", cleanHeader(row[c]), "feetUsed=", num(fRow?.[c]));
          break;
        }
      }
      break;
    }
  }
  // Find state code row
  for (let r = 50; r < 75; r++) {
    const row = data[r];
    if (!row) continue;
    let codes = 0;
    for (let c = 1; c < row.length; c++) {
      if (cleanHeader(row[c]).match(/^[A-Z]{2}$/)) codes++;
    }
    if (codes >= 5) {
      console.log("State code row at:", r, "- first 5 codes:",
        row.slice(1,6).map(v=>cleanHeader(v)));
      // Find AZ column
      let azCols = [];
      for (let c = 1; c < row.length; c++) {
        if (cleanHeader(row[c]) === "AZ") azCols.push(c);
      }
      console.log("AZ columns:", azCols);
      // Read truss prices for AZ
      for (let r2 = r+1; r2 < r+12; r2++) {
        const label = cleanHeader(data[r2]?.[0]).toLowerCase();
        if (label.includes("wide") || label.includes("truss") || label.includes("pie") || label.includes("channel") || label.includes("tub")) {
          const vals = azCols.map(c => num(data[r2]?.[c]));
          console.log(`  Row ${r2} "${cleanHeader(data[r2]?.[0])}": AZ vals =`, vals);
        }
      }
      break;
    }
  }
}

// Read HC original counts
const hcSheet = findSheet("Snow - Hat Channels");
console.log("\nHC sheet found:", !!hcSheet);
if (hcSheet) {
  const data = sheetToArray(hcSheet);
  const headers = data[0];
  // Find Original section
  let origCol = -1;
  for (let c = 8; c < headers.length; c++) {
    if (cleanHeader(headers[c]).toLowerCase().includes("original")) {
      origCol = c;
      break;
    }
  }
  console.log("Original HC col:", origCol);
  if (origCol >= 0) {
    // Find AZ row and width 24
    const widthHeaders = [];
    for (let c = origCol+1; c < headers.length; c++) {
      widthHeaders.push(cleanHeader(headers[c]));
    }
    console.log("Width headers:", widthHeaders);
    for (let r = 1; r < 20; r++) {
      if (cleanHeader(data[r]?.[origCol]) === "AZ") {
        const w24idx = widthHeaders.indexOf("24");
        console.log("AZ HC original at width 24:", num(data[r]?.[origCol+1+w24idx]));
        break;
      }
    }
  }
}

// Now simulate the engine calculation
console.log("\n=== SIMULATING ENGINE (24x50x10, AZ, AFV, Enclosed, 20LL, 90mph) ===");
const width = 24, length = 50, height = 10;
const bucketedWind = 105;
const snowCode = "T-20LL"; // height 10 → T
const configKey = "E-105-24-AFV";

// Truss spacing
const tsSheet = findSheet("Snow - Truss Spacing");
const tsData = sheetToArray(tsSheet);
const tsHeaders = tsData[0];
let configCol = -1;
for (let c = 1; c < tsHeaders.length; c++) {
  if (cleanHeader(tsHeaders[c]) === configKey) { configCol = c; break; }
}
let trussSpacing = 0;
for (let r = 1; r < tsData.length; r++) {
  if (cleanHeader(tsData[r][0]) === snowCode) {
    trussSpacing = num(tsData[r][configCol]);
    break;
  }
}
console.log("Truss Spacing:", trussSpacing);

// Trusses needed
const lengthInches = length * 12;
const trussesNeeded = Math.ceil(lengthInches / trussSpacing) + 1;
console.log("Trusses needed:", trussesNeeded);

// Original trusses
const trData = sheetToArray(findSheet("Snow - Trusses"));
const trHeaders = trData[0];
let trCol = -1;
for (let c = 1; c < trHeaders.length; c++) {
  if (cleanHeader(trHeaders[c]) === "24-AZ") { trCol = c; break; }
}
let origTrusses = 0;
for (let r = 1; r < trData.length; r++) {
  if (num(trData[r][0]) === length) { origTrusses = num(trData[r][trCol]); break; }
}
console.log("Original trusses:", origTrusses);
console.log("Extra trusses:", Math.max(0, trussesNeeded - origTrusses));

// HC
const hcSpacing = 54; // from 60-20LL at 105
const barInches = ((width+2)/2)*12;
const hcPerSide = Math.ceil(barInches / hcSpacing) + 1;
const totalHC = hcPerSide * 2;
console.log("\nHC needed:", totalHC, "(barInches:", barInches, ")");
console.log("HC original: 10");
console.log("Extra HC:", Math.max(0, totalHC - 10));

// Girts
const gData = sheetToArray(findSheet("Snow - Girts"));
let girtSpacing = 0;
for (let r = 1; r < gData.length; r++) {
  if (num(gData[r][0]) === 60) {
    girtSpacing = num(gData[r][1]); // col 1 = wind 105
    break;
  }
}
console.log("\nGirt spacing:", girtSpacing);
const girtsReq = Math.ceil((height*12) / girtSpacing) + 1;
console.log("Girts needed:", girtsReq);
console.log("Girts gate: enclosed=true, verticalPanels=true → girts needed");

// Verticals
console.log("\nVert spacing: 60");
const vertsNeeded = Math.ceil((width*12) / 60) + 1;
console.log("Verts needed:", vertsNeeded);
console.log("Verts original: 6");
console.log("Extra verts:", Math.max(0, vertsNeeded - 6));

console.log("\n=== ALL should be 0 extra → $0 engineering ===");
