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
  if (wb.Sheets[name]) return wb.Sheets[name];
  for (const key of Object.keys(wb.Sheets)) {
    if (key.trim() === name.trim()) return wb.Sheets[key];
  }
  return null;
}

// CONFIG: 30x100x15, 70GL, 100mph (→105), AZ, AFV, Enclosed, 2 sides V, 2 ends V
const width = 30, length = 100, height = 15;
const inputWind = 100;
const snowLoad = "70GL";
const state = "AZ";
const roofKey = "AFV";
const isEnclosed = true;
const sidesQty = 2, endsQty = 2;
const sidesOrientation = "vertical", endsOrientation = "vertical";

console.log(`=== TRACE: ${width}x${length}x${height}, ${snowLoad}, ${inputWind}mph, ${state}, ${roofKey} ===\n`);

// Step 1: Resolve inputs
// Wind bucketing
const changersSheet = findSheet("Snow - Changers");
const changersData = sheetToArray(changersSheet);
let bucketedWind = 105; // default
if (changersData[0] && changersData[1]) {
  for (let c = 1; c < changersData[0].length; c++) {
    if (num(changersData[0][c]) === inputWind) {
      bucketedWind = num(changersData[1][c]);
      break;
    }
  }
}
console.log("Bucketed wind:", bucketedWind);

// Height classification
let heightPrefix = "T";
for (let r = 0; r < changersData.length; r++) {
  const row = changersData[r];
  let smtCount = 0;
  for (let c = 0; c < (row?.length || 0); c++) {
    const v = cleanHeader(row[c]);
    if (v === "S" || v === "M" || v === "T") smtCount++;
  }
  if (smtCount >= 5) {
    const hRow = changersData[r - 1]; // heights
    for (let c = 0; c < (hRow?.length || 0); c++) {
      if (num(hRow[c]) === height) {
        heightPrefix = cleanHeader(row[c]);
        break;
      }
    }
    // Also read feetUsed
    const fRow = changersData[r + 1]; // feet used
    let feetUsed = 0;
    for (let c = 0; c < (hRow?.length || 0); c++) {
      if (num(hRow[c]) === height) {
        feetUsed = num(fRow?.[c]);
        break;
      }
    }
    console.log("Height prefix:", heightPrefix, "feetUsed:", feetUsed);
    break;
  }
}

const snowCode = `${heightPrefix}-${snowLoad}`;
const enclosure = isEnclosed ? "E" : "O";
const configKey = `${enclosure}-${bucketedWind}-${width}-${roofKey}`;
console.log("Snow code:", snowCode);
console.log("Config key:", configKey);

// Step 2: Trusses
console.log("\n--- TRUSSES ---");
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
console.log("Truss spacing:", trussSpacing, "(configCol:", configCol, ")");
if (trussSpacing === 0) {
  console.log("CONTACT ENGINEER — truss spacing is 0");
  process.exit(0);
}

const lengthInches = length * 12;
const trussesNeeded = Math.ceil(lengthInches / trussSpacing) + 1;
console.log("Trusses needed:", trussesNeeded);

// Original truss count
const trSheet = findSheet("Snow - Trusses");
const trData = sheetToArray(trSheet);
const trHeaders = trData[0];
let trCol = -1;
const widthStateKey = `${width}-${state}`;
for (let c = 1; c < trHeaders.length; c++) {
  if (cleanHeader(trHeaders[c]) === widthStateKey) { trCol = c; break; }
}
let origTrusses = 0;
for (let r = 1; r < trData.length; r++) {
  if (num(trData[r][0]) === length) { origTrusses = num(trData[r][trCol]); break; }
}
console.log("Original trusses:", origTrusses, "(" + widthStateKey + ")");
const extraTrusses = Math.max(0, trussesNeeded - origTrusses);
console.log("Extra trusses:", extraTrusses);

// Truss price
let trussPrice = 190;
let piePricePerFt = 15;
let feetUsedVal = 0;
// Find state pricing section
for (let r = 30; r < changersData.length; r++) {
  const row = changersData[r];
  if (!row) continue;
  let codes = 0;
  for (let c = 1; c < (row?.length || 0); c++) {
    if (cleanHeader(row[c]).match(/^[A-Z]{2}$/)) codes++;
  }
  if (codes >= 5) {
    // Found state header row
    let azCol = -1;
    for (let c = 1; c < row.length; c++) {
      if (cleanHeader(row[c]) === state) { azCol = c; break; }
    }
    console.log("State header row:", r, "AZ col:", azCol);
    // Read prices from subsequent rows
    for (let r2 = r + 1; r2 < Math.min(r + 15, changersData.length); r2++) {
      const label = cleanHeader(changersData[r2]?.[0]).toLowerCase();
      const val = num(changersData[r2]?.[azCol]);
      console.log(`  Row ${r2}: "${cleanHeader(changersData[r2]?.[0])}" = ${val}`);

      // Look for width-specific truss prices
      if (label.includes("26") || label.includes("30")) {
        const allVals = [];
        for (let c = 1; c < (changersData[r2]?.length || 0); c++) {
          if (num(changersData[r2][c]) > 0) allVals.push(`${cleanHeader(row[c])}=${num(changersData[r2][c])}`);
        }
        if (allVals.length > 0) console.log("    All vals:", allVals.join(", "));
      }
    }

    // Read truss price for this width
    for (let r2 = r + 1; r2 < Math.min(r + 10, changersData.length); r2++) {
      const label = cleanHeader(changersData[r2]?.[0]).toLowerCase();
      if (width <= 24 && (label.includes("12") || label.includes("small"))) {
        trussPrice = num(changersData[r2]?.[azCol]);
      }
      if (width >= 26 && (label.includes("26") || label.includes("large"))) {
        trussPrice = num(changersData[r2]?.[azCol]);
      }
      if (label.includes("pie") || label.includes("leg")) {
        piePricePerFt = num(changersData[r2]?.[azCol]);
      }
      if (label.includes("channel")) {
        console.log("Channel price:", num(changersData[r2]?.[azCol]));
      }
      if (label.includes("tub")) {
        console.log("Tubing price:", num(changersData[r2]?.[azCol]));
      }
    }
    break;
  }
}

// feetUsed
for (let r = 0; r < changersData.length; r++) {
  const row = changersData[r];
  let smtCount = 0;
  for (let c = 0; c < (row?.length || 0); c++) {
    const v = cleanHeader(row[c]);
    if (v === "S" || v === "M" || v === "T") smtCount++;
  }
  if (smtCount >= 5) {
    const hRow = changersData[r - 1];
    const fRow = changersData[r + 1];
    for (let c = 0; c < (hRow?.length || 0); c++) {
      if (num(hRow[c]) === height) {
        feetUsedVal = num(fRow?.[c]);
        break;
      }
    }
    break;
  }
}

const legSurcharge = feetUsedVal * piePricePerFt;
const trussCost = extraTrusses * (trussPrice + legSurcharge);
console.log("Truss price:", trussPrice, "legSurcharge:", legSurcharge, `(${feetUsedVal} ft × $${piePricePerFt})`);
console.log("TRUSS COST:", trussCost);

// Step 3: Hat Channels
console.log("\n--- HAT CHANNELS ---");
const TRUSS_SPACING_BUCKETS = [36, 42, 48, 54, 60];
function nearestBucket(val, buckets) {
  let best = buckets[0], bestDiff = Infinity;
  for (const b of buckets) {
    const diff = Math.abs(val - b);
    if (diff < bestDiff) { bestDiff = diff; best = b; }
  }
  return best;
}

const bucketedTrussSpacing = nearestBucket(trussSpacing, TRUSS_SPACING_BUCKETS);
const hcRowKey = `${bucketedTrussSpacing}-${snowLoad}`;
console.log("Bucketed truss spacing:", bucketedTrussSpacing, "HC row key:", hcRowKey);

const hcSheet = findSheet("Snow - Hat Channels");
const hcData = sheetToArray(hcSheet);
const hcHeaders = hcData[0];

// Find wind column
let windCol = -1;
for (let c = 1; c < hcHeaders.length; c++) {
  if (num(hcHeaders[c]) === bucketedWind) { windCol = c; break; }
}
let hcSpacing = 0;
for (let r = 1; r < hcData.length; r++) {
  if (cleanHeader(hcData[r][0]) === hcRowKey) {
    hcSpacing = num(hcData[r][windCol]);
    break;
  }
}
console.log("HC spacing:", hcSpacing);

const barSize = (width + 2) / 2;
const barInches = barSize * 12;
const hcPerSide = Math.ceil(barInches / hcSpacing) + 1;
const totalHC = hcPerSide * 2;
console.log("Bar size:", barSize, "ft, barInches:", barInches);
console.log("HC per side:", hcPerSide, "total:", totalHC);

// Original HC count
let origHC = 0;
let origCol = -1;
for (let c = 8; c < hcHeaders.length; c++) {
  if (cleanHeader(hcHeaders[c]).toLowerCase().includes("original")) {
    origCol = c; break;
  }
}
if (origCol >= 0) {
  // Find state row and width column
  const widthHeaders = [];
  for (let c = origCol + 1; c < hcHeaders.length; c++) {
    widthHeaders.push(cleanHeader(hcHeaders[c]));
  }
  for (let r = 1; r < 20; r++) {
    if (cleanHeader(hcData[r]?.[origCol]) === state) {
      const w30idx = widthHeaders.indexOf(String(width));
      if (w30idx >= 0) origHC = num(hcData[r]?.[origCol + 1 + w30idx]);
      break;
    }
  }
}
console.log("Original HC:", origHC);
const extraHC = Math.max(0, totalHC - origHC);
console.log("Extra HC:", extraHC);

// Get channel price
let channelPrice = 2;
for (let r = 30; r < changersData.length; r++) {
  const row = changersData[r];
  if (!row) continue;
  let codes = 0;
  for (let c = 1; c < (row?.length || 0); c++) {
    if (cleanHeader(row[c]).match(/^[A-Z]{2}$/)) codes++;
  }
  if (codes >= 5) {
    let azCol = -1;
    for (let c = 1; c < row.length; c++) {
      if (cleanHeader(row[c]) === state) { azCol = c; break; }
    }
    for (let r2 = r + 1; r2 < Math.min(r + 10, changersData.length); r2++) {
      const label = cleanHeader(changersData[r2]?.[0]).toLowerCase();
      if (label.includes("channel")) {
        channelPrice = num(changersData[r2]?.[azCol]);
        break;
      }
    }
    break;
  }
}

const channelLength = length + 1;
const hcCost = extraHC * channelPrice * channelLength;
console.log("Channel price:", channelPrice, "length:", channelLength);
console.log("HC COST:", hcCost);

// Step 4: Girts
console.log("\n--- GIRTS ---");
const gData = sheetToArray(findSheet("Snow - Girts"));
const gHeaders = gData[0];

// Find girt spacing for bucketed truss spacing + wind
let girtSpacing = 0;
for (let r = 1; r < gData.length; r++) {
  if (num(gData[r][0]) === bucketedTrussSpacing) {
    for (let c = 1; c < gHeaders.length; c++) {
      if (num(gHeaders[c]) === bucketedWind) {
        girtSpacing = num(gData[r][c]);
        break;
      }
    }
    break;
  }
}
console.log("Girt spacing:", girtSpacing, `(truss=${bucketedTrussSpacing}, wind=${bucketedWind})`);

// Original girts
let origGirts = 0;
// Find right section of girts sheet
for (let r = 0; r < gData.length; r++) {
  for (let c = 8; c < (gData[r]?.length || 0); c++) {
    if (cleanHeader(gData[r][c]).toLowerCase().includes("original")) {
      console.log("Original girts header at:", r, c);
      // Read height → count pairs
      for (let r2 = r + 1; r2 < gData.length; r2++) {
        const h = num(gData[r2]?.[c]);
        const cnt = num(gData[r2]?.[c + 1]);
        if (h === height || (h <= height && num(gData[r2 + 1]?.[c]) > height)) {
          origGirts = cnt;
          console.log(`Height ${h} → girts: ${cnt}`);
        }
        if (h === height) {
          origGirts = cnt;
          break;
        }
      }
      break;
    }
  }
  if (origGirts > 0) break;
}

// Try alternative: scan col 11 for heights
if (origGirts === 0) {
  for (let r = 0; r < gData.length; r++) {
    if (num(gData[r]?.[11]) === height) {
      origGirts = num(gData[r]?.[12]);
      console.log(`Height ${height} at row ${r}: girts = ${origGirts}`);
      break;
    }
  }
}

if (origGirts === 0) {
  // Default
  if (height <= 11) origGirts = 3;
  else if (height <= 17) origGirts = 4;
  else origGirts = 5;
  console.log("Using default girts:", origGirts);
}

const heightInches = height * 12;
const girtsRequired = Math.ceil(heightInches / girtSpacing) + 1;
const extraGirts = Math.max(0, girtsRequired - origGirts);
console.log("Girts required:", girtsRequired, "original:", origGirts, "extra:", extraGirts);

// Girt perimeter (only vertical surfaces)
let girtPerimeter = 0;
if (sidesOrientation === "vertical") girtPerimeter += sidesQty * length;
if (endsOrientation === "vertical") girtPerimeter += endsQty * width;
console.log("Girt perimeter:", girtPerimeter);

let tubingPrice = 3;
for (let r = 30; r < changersData.length; r++) {
  const row = changersData[r];
  if (!row) continue;
  let codes = 0;
  for (let c = 1; c < (row?.length || 0); c++) {
    if (cleanHeader(row[c]).match(/^[A-Z]{2}$/)) codes++;
  }
  if (codes >= 5) {
    let azCol = -1;
    for (let c = 1; c < row.length; c++) {
      if (cleanHeader(row[c]) === state) { azCol = c; break; }
    }
    for (let r2 = r + 1; r2 < Math.min(r + 15, changersData.length); r2++) {
      const label = cleanHeader(changersData[r2]?.[0]).toLowerCase();
      if (label.includes("tub")) {
        tubingPrice = num(changersData[r2]?.[azCol]);
        break;
      }
    }
    break;
  }
}

const girtCost = extraGirts * tubingPrice * girtPerimeter;
console.log("Tubing price:", tubingPrice);
console.log("GIRT COST:", girtCost);

// Step 5: Verticals
console.log("\n--- VERTICALS ---");
const vData = sheetToArray(findSheet("Snow - Verticals"));
const vHeaders = vData[0];

// Find spacing for height and wind
let vertSpacing = 0;
let heightCol = -1;
for (let c = 1; c < vHeaders.length; c++) {
  if (num(vHeaders[c]) === height) { heightCol = c; break; }
}
for (let r = 1; r < vData.length; r++) {
  if (num(vData[r][0]) === bucketedWind) {
    vertSpacing = num(vData[r][heightCol]);
    break;
  }
}
console.log("Vertical spacing:", vertSpacing, "(height col:", heightCol, ")");

// Original vertical count
let origVerts = 0;
for (let r = 0; r < vData.length; r++) {
  const label = cleanHeader(vData[r]?.[0]).toLowerCase();
  if (label.includes("original")) {
    const widthRow = vData[r];
    const countRow = vData[r + 1];
    for (let c = 1; c < widthRow.length; c++) {
      if (num(widthRow[c]) === width) {
        origVerts = num(countRow[c]);
        break;
      }
    }
    break;
  }
}
console.log("Original verts:", origVerts);

const widthInches = width * 12;
const vertsNeeded = Math.ceil(widthInches / vertSpacing) + 1;
const extraVerts = Math.max(0, vertsNeeded - origVerts);
console.log("Verts needed:", vertsNeeded, "extra:", extraVerts);

const roofRise = (width / 2) * (3 / 12); // 3:12 pitch
const peakHeight = height + roofRise;
const vertCost = extraVerts * endsQty * tubingPrice * peakHeight;
console.log("Peak height:", peakHeight, "ends:", endsQty);
console.log("VERTICAL COST:", vertCost);

// Total
console.log("\n--- TOTAL ---");
const rawTotal = trussCost + hcCost + girtCost + vertCost;
console.log("Trusses:", trussCost);
console.log("HC:", hcCost);
console.log("Girts:", girtCost);
console.log("Verticals:", vertCost);
console.log("Raw total:", rawTotal);
console.log("Rounded:", Math.round(rawTotal));
console.log("\nExpected from spreadsheet: $9,891");
console.log("Gap:", Math.round(rawTotal) - 9891);

// Height multiplier check
const heightMultiplier = height >= 19 ? 3.0 : height >= 16 ? 2.5 : height >= 13 ? 2.0 : 1.0;
console.log("\nHeight multiplier for height", height, ":", heightMultiplier);
console.log("Total × multiplier:", Math.round(rawTotal * heightMultiplier));

// Now dump the Math Calculations sheet for reference
console.log("\n\n=== MATH CALCULATIONS SHEET DUMP ===");
const mathSheet = findSheet("Snow - Math Calculations");
if (mathSheet) {
  const mData = sheetToArray(mathSheet);
  for (let r = 0; r < Math.min(mData.length, 35); r++) {
    const nonEmpty = mData[r].filter(v => v !== "" && v !== 0 && v !== null);
    if (nonEmpty.length > 0) {
      const vals = mData[r].map((v, i) => `[${i}]=${v}`).filter(s => !s.endsWith("=") && !s.endsWith("=0"));
      if (vals.length > 0) console.log(`Row ${r}: ${vals.join(" | ")}`);
    }
  }
}
