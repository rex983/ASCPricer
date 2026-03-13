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

// Check ALL row keys in HC sheet
const hcSheet = findSheet("Snow - Hat Channels");
const hcData = sheetToArray(hcSheet);
console.log("=== HC Row Keys ===");
const hcRowKeys = new Set();
for (let r = 1; r < hcData.length; r++) {
  const key = cleanHeader(hcData[r][0]);
  if (key) hcRowKeys.add(key);
}
console.log([...hcRowKeys].sort().join(", "));

// Check what HC spacing is for different row keys with 70GL at wind 105
console.log("\n=== HC spacing at wind 105 for 70GL variants ===");
let windCol = -1;
for (let c = 1; c < hcData[0].length; c++) {
  if (num(hcData[0][c]) === 105) { windCol = c; break; }
}
for (let r = 1; r < hcData.length; r++) {
  const key = cleanHeader(hcData[r][0]);
  if (key.includes("70GL")) {
    console.log(`  ${key}: ${num(hcData[r][windCol])}`);
  }
}

// Now try BOTH STD and AFV calculations for 30x100x15 70GL 100mph
console.log("\n\n=== CALCULATION: 30x100x15, 70GL, 100mph, AZ ===");
const width = 30, length = 100, height = 15;
const bucketedWind = 105;
const snowLoad = "70GL";
const snowCode = "T-70GL"; // height 15 → T
const state = "AZ";
const feetUsed = 16;  // from changers for height 15
const piePricePerFt = 15;
const trussPrice = 330; // 30W AZ
const channelPrice = 2.5;
const tubingPrice = 3.5;
const origTrusses = 25;
const origHC = 12;
const origGirts = 4;
const origVerts = 7;

for (const roofKey of ["AFV", "STD"]) {
  console.log(`\n--- Roof: ${roofKey} ---`);

  // Truss spacing
  const tsSheet = findSheet("Snow - Truss Spacing");
  const tsData = sheetToArray(tsSheet);
  const configKey = `E-${bucketedWind}-${width}-${roofKey}`;
  let configCol = -1;
  for (let c = 1; c < tsData[0].length; c++) {
    if (cleanHeader(tsData[0][c]) === configKey) { configCol = c; break; }
  }
  let trussSpacing = 0;
  for (let r = 1; r < tsData.length; r++) {
    if (cleanHeader(tsData[r][0]) === snowCode) {
      trussSpacing = num(tsData[r][configCol]);
      break;
    }
  }
  console.log(`Config key: ${configKey}, Truss spacing: ${trussSpacing}`);

  const trussesNeeded = Math.ceil((length * 12) / trussSpacing) + 1;
  const extraTrusses = Math.max(0, trussesNeeded - origTrusses);
  const legSurcharge = feetUsed * piePricePerFt; // 16 × 15 = 240
  const trussCost = extraTrusses * (trussPrice + legSurcharge);
  console.log(`Trusses: needed=${trussesNeeded}, extra=${extraTrusses}, cost=$${trussCost}`);

  // HC
  const BUCKETS = [36, 42, 48, 54, 60];
  let bestBucket = BUCKETS[0], bestDiff = Math.abs(trussSpacing - BUCKETS[0]);
  for (const b of BUCKETS) {
    const d = Math.abs(trussSpacing - b);
    if (d < bestDiff) { bestDiff = d; bestBucket = b; }
  }
  const hcRowKey = `${bestBucket}-${snowLoad}`;
  let hcSpacing = 0;
  for (let r = 1; r < hcData.length; r++) {
    if (cleanHeader(hcData[r][0]) === hcRowKey) {
      hcSpacing = num(hcData[r][windCol]);
      break;
    }
  }
  const barInches = ((width + 2) / 2) * 12;
  const hcPerSide = Math.ceil(barInches / hcSpacing) + 1;
  const totalHC = hcPerSide * 2;
  const extraHC = Math.max(0, totalHC - origHC);
  const hcCost = extraHC * channelPrice * (length + 1);
  console.log(`HC: row=${hcRowKey}, spacing=${hcSpacing}, needed=${totalHC}, extra=${extraHC}, cost=$${hcCost}`);

  // Girts
  const gData = sheetToArray(findSheet("Snow - Girts"));
  let girtSpacing = 0;
  for (let r = 1; r < gData.length; r++) {
    if (num(gData[r][0]) === bestBucket) {
      for (let c = 1; c < gData[0].length; c++) {
        if (num(gData[0][c]) === bucketedWind) {
          girtSpacing = num(gData[r][c]);
          break;
        }
      }
      break;
    }
  }
  const girtsReq = Math.ceil((height * 12) / girtSpacing) + 1;
  const extraGirts = Math.max(0, girtsReq - origGirts);
  console.log(`Girts: spacing=${girtSpacing}, needed=${girtsReq}, extra=${extraGirts}`);

  // Verticals
  const vData = sheetToArray(findSheet("Snow - Verticals"));
  let heightCol = -1;
  for (let c = 1; c < vData[0].length; c++) {
    if (num(vData[0][c]) === height) { heightCol = c; break; }
  }
  let vertSpacing = 0;
  for (let r = 1; r < vData.length; r++) {
    if (num(vData[r][0]) === bucketedWind) {
      vertSpacing = num(vData[r][heightCol]);
      break;
    }
  }
  const vertsNeeded = Math.ceil((width * 12) / vertSpacing) + 1;
  const extraVerts = Math.max(0, vertsNeeded - origVerts);
  const endsQty = 2;
  const roofRise = roofKey === "AFV" ? (width / 2) * (3 / 12) : 0;
  const peakHeight = height + roofRise;
  // Apply height multiplier: 13-15 → ×2
  const heightMult = height >= 19 ? 3 : height >= 16 ? 2.5 : height >= 13 ? 2 : 1;
  const vertCostBase = extraVerts * endsQty * tubingPrice * peakHeight;
  const vertCost = vertCostBase * heightMult;
  console.log(`Verts: spacing=${vertSpacing}, needed=${vertsNeeded}, extra=${extraVerts}, peakH=${peakHeight}`);
  console.log(`Verts cost: base=$${vertCostBase} × ${heightMult} = $${vertCost}`);

  const total = trussCost + hcCost + extraGirts * tubingPrice * 260 + vertCost;
  console.log(`TOTAL: $${Math.round(total)}`);

  // What if we also apply height multiplier to girts?
  const girtCost = extraGirts * tubingPrice * 260;
  const girtCostMultiplied = girtCost * heightMult;
  const totalWithGirtMult = trussCost + hcCost + girtCostMultiplied + vertCost;
  console.log(`TOTAL with girt multiplier: $${Math.round(totalWithGirtMult)}`);

  // What if height multiplier applies to EVERYTHING?
  const totalAllMultiplied = (trussCost + hcCost + girtCost + vertCostBase) * heightMult;
  console.log(`TOTAL with multiplier on ALL: $${Math.round(totalAllMultiplied)}`);
}

// Also check the Math Calculations formula for how P2 (truss spacing) is set
console.log("\n=== Math Calculations: P2 formula ===");
const mathSheet = findSheet("Snow - Math Calculations");
if (mathSheet) {
  const p2 = mathSheet[XLSX.utils.encode_cell({ r: 1, c: 15 })];
  console.log("P2:", p2?.v, "formula:", p2?.f);
  // P4 (HC spacing)
  const p4 = mathSheet[XLSX.utils.encode_cell({ r: 3, c: 15 })];
  console.log("P4:", p4?.v, "formula:", p4?.f);
  // P6 (girt spacing)
  const p6 = mathSheet[XLSX.utils.encode_cell({ r: 5, c: 15 })];
  console.log("P6:", p6?.v, "formula:", p6?.f);
  // P8 (vert spacing)
  const p8 = mathSheet[XLSX.utils.encode_cell({ r: 7, c: 15 })];
  console.log("P8:", p8?.v, "formula:", p8?.f);

  // The key formula reference for HC spacing
  const d10 = mathSheet[XLSX.utils.encode_cell({ r: 9, c: 3 })];
  console.log("D10 (width for HC):", d10?.v, "formula:", d10?.f);
}
