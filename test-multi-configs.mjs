import XLSX from "xlsx";

const wb = XLSX.readFile("C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx", {
  cellFormula: true,
});

// ── Utility functions ──
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

// ── Pre-load all sheet data ──
const changersData = sheetToArray(findSheet("Snow - Changers"));
const tsData = sheetToArray(findSheet("Snow - Truss Spacing"));
const trData = sheetToArray(findSheet("Snow - Trusses"));
const hcData = sheetToArray(findSheet("Snow - Hat Channels"));
const gData = sheetToArray(findSheet("Snow - Girts"));
const vData = sheetToArray(findSheet("Snow - Verticals"));

const TRUSS_SPACING_BUCKETS = [36, 42, 48, 54, 60];
const WIND_LOAD_CATEGORIES = [105, 115, 130, 140, 155, 165, 180];

function nearestBucketRoundDown(val, buckets) {
  let nearest = buckets[0];
  for (const b of buckets) {
    if (b <= val) nearest = b;
    else break;
  }
  return nearest;
}

function nearestBucketClosest(val, buckets) {
  let best = buckets[0], bestDiff = Infinity;
  for (const b of buckets) {
    const diff = Math.abs(val - b);
    if (diff < bestDiff) { bestDiff = diff; best = b; }
  }
  return best;
}

// ── Parse wind bucketing ──
function bucketWind(inputMph) {
  if (changersData[0] && changersData[1]) {
    for (let c = 1; c < changersData[0].length; c++) {
      if (num(changersData[0][c]) === inputMph) {
        const bucketed = num(changersData[1][c]);
        if (bucketed > 0) return bucketed;
      }
    }
  }
  return nearestBucketRoundDown(inputMph, WIND_LOAD_CATEGORIES);
}

// ── Parse height classification ──
let smtRow = -1, smtHeightRow = null, smtClassRow = null, smtFeetRow = null;
for (let r = 0; r < changersData.length; r++) {
  const row = changersData[r];
  let count = 0;
  for (let c = 0; c < (row?.length || 0); c++) {
    const v = cleanHeader(row[c]);
    if (v === "S" || v === "M" || v === "T") count++;
  }
  if (count >= 5) {
    smtRow = r;
    smtHeightRow = changersData[r - 1];
    smtClassRow = changersData[r];
    smtFeetRow = changersData[r + 1];
    break;
  }
}

function getHeightPrefix(height) {
  if (smtHeightRow) {
    for (let c = 0; c < smtHeightRow.length; c++) {
      if (num(smtHeightRow[c]) === height) return cleanHeader(smtClassRow[c]);
    }
  }
  if (height <= 6) return "S";
  if (height <= 9) return "M";
  return "T";
}

function getFeetUsed(height) {
  if (smtHeightRow && smtFeetRow) {
    for (let c = 0; c < smtHeightRow.length; c++) {
      if (num(smtHeightRow[c]) === height) return num(smtFeetRow[c]);
    }
  }
  return 0;
}

// ── Parse per-state pricing ──
let stateHeaderRow = -1, stateCodes = [];
for (let r = 30; r < changersData.length; r++) {
  const row = changersData[r];
  if (!row) continue;
  let codeCount = 0;
  const codes = [];
  for (let c = 1; c < row.length; c++) {
    const v = cleanHeader(row[c]);
    if (v.match(/^[A-Z]{2}$/)) { codeCount++; codes.push(v); }
    else codes.push("");
  }
  if (codeCount >= 5) { stateHeaderRow = r; stateCodes = codes; break; }
}

function getStateCol(state) {
  for (let i = 0; i < stateCodes.length; i++) {
    if (stateCodes[i] === state) return i + 1; // +1 because stateCodes starts at col 1
  }
  return -1;
}

function getTrussPrice(width, state) {
  const col = getStateCol(state);
  if (col < 0) return 190;
  for (let r = stateHeaderRow + 1; r < Math.min(stateHeaderRow + 12, changersData.length); r++) {
    const label = cleanHeader(changersData[r]?.[0]).toLowerCase();
    // Width-specific matching
    if (width <= 24 && (label.includes("12") || label.includes("24") || label.includes("small"))) {
      const val = num(changersData[r]?.[col]);
      if (val > 50 && val < 1000) return val;
    }
    if (width >= 26 && (label.includes("26") || label.includes("30") || label.includes("large"))) {
      const val = num(changersData[r]?.[col]);
      if (val > 50 && val < 1000) return val;
    }
  }
  return 190;
}

function getPiePricePerFt(state) {
  const col = getStateCol(state);
  if (col < 0) return 15;
  for (let r = stateHeaderRow + 1; r < Math.min(stateHeaderRow + 15, changersData.length); r++) {
    const label = cleanHeader(changersData[r]?.[0]).toLowerCase();
    if (label.includes("pie")) return num(changersData[r]?.[col]);
  }
  return 15;
}

function getChannelPrice(state) {
  const col = getStateCol(state);
  if (col < 0) return 2;
  for (let r = stateHeaderRow + 1; r < Math.min(stateHeaderRow + 15, changersData.length); r++) {
    const label = cleanHeader(changersData[r]?.[0]).toLowerCase();
    if (label.includes("channel")) return num(changersData[r]?.[col]);
  }
  return 2;
}

function getTubingPrice(state) {
  const col = getStateCol(state);
  if (col < 0) return 3;
  for (let r = stateHeaderRow + 1; r < Math.min(stateHeaderRow + 15, changersData.length); r++) {
    const label = cleanHeader(changersData[r]?.[0]).toLowerCase();
    if (label.includes("tub")) return num(changersData[r]?.[col]);
  }
  return 3;
}

// ── Truss spacing lookup ──
function lookupTrussSpacing(snowCode, configKey) {
  const headers = tsData[0];
  let col = -1;
  for (let c = 1; c < headers.length; c++) {
    if (cleanHeader(headers[c]) === configKey) { col = c; break; }
  }
  if (col < 0) return { spacing: 0, col: -1 };
  for (let r = 1; r < tsData.length; r++) {
    if (cleanHeader(tsData[r][0]) === snowCode) {
      return { spacing: num(tsData[r][col]), col };
    }
  }
  return { spacing: 0, col };
}

// ── Truss count lookup ──
function lookupOriginalTrusses(width, state, length) {
  // Try exact state first, then any state for this width
  const headers = trData[0];
  let col = -1;
  const exactKey = `${width}-${state}`;
  for (let c = 1; c < headers.length; c++) {
    if (cleanHeader(headers[c]) === exactKey) { col = c; break; }
  }
  // Fallback: find any column starting with width
  if (col < 0) {
    const prefix = `${width}-`;
    for (let c = 1; c < headers.length; c++) {
      if (cleanHeader(headers[c]).startsWith(prefix)) { col = c; break; }
    }
  }
  if (col < 0) return { count: 0, key: "NOT FOUND" };
  for (let r = 1; r < trData.length; r++) {
    if (num(trData[r][0]) === length) {
      return { count: num(trData[r][col]), key: cleanHeader(headers[col]) };
    }
  }
  return { count: 0, key: cleanHeader(headers[col]) };
}

// ── HC spacing lookup ──
function lookupHCSpacing(hcRowKey, bucketedWind) {
  const headers = hcData[0];
  let windCol = -1;
  for (let c = 1; c < headers.length; c++) {
    if (num(headers[c]) === bucketedWind) { windCol = c; break; }
  }
  if (windCol < 0) return 0;
  for (let r = 1; r < hcData.length; r++) {
    if (cleanHeader(hcData[r][0]) === hcRowKey) {
      return num(hcData[r][windCol]);
    }
  }
  return 0;
}

// ── Original HC count ──
function lookupOriginalHC(state, width) {
  const headers = hcData[0];
  let origCol = -1;
  for (let c = 8; c < headers.length; c++) {
    if (cleanHeader(headers[c]).toLowerCase().includes("original")) {
      origCol = c; break;
    }
  }
  if (origCol < 0) return 0;
  const widthHeaders = [];
  for (let c = origCol + 1; c < headers.length; c++) {
    widthHeaders.push(cleanHeader(headers[c]));
  }
  for (let r = 1; r < 20; r++) {
    if (cleanHeader(hcData[r]?.[origCol]) === state) {
      const idx = widthHeaders.indexOf(String(width));
      if (idx >= 0) return num(hcData[r]?.[origCol + 1 + idx]);
    }
  }
  return 0;
}

// ── Girt spacing lookup ──
function lookupGirtSpacing(bucketedTrussSpacing, bucketedWind) {
  const headers = gData[0];
  for (let r = 1; r < gData.length; r++) {
    if (num(gData[r][0]) === bucketedTrussSpacing) {
      for (let c = 1; c < headers.length; c++) {
        if (num(headers[c]) === bucketedWind) return num(gData[r][c]);
      }
    }
  }
  return 0;
}

// ── Original girt count ──
function lookupOriginalGirts(height) {
  // Scan for "original" section in girts sheet
  for (let r = 0; r < gData.length; r++) {
    for (let c = 8; c < (gData[r]?.length || 0); c++) {
      if (cleanHeader(gData[r][c]).toLowerCase().includes("original")) {
        for (let r2 = r + 1; r2 < gData.length; r2++) {
          const h = num(gData[r2]?.[c]);
          const cnt = num(gData[r2]?.[c + 1]);
          if (h === height) return cnt;
        }
      }
    }
  }
  // Try range-based approach from the right side
  for (let r = 1; r < gData.length; r++) {
    for (let c = 8; c < (gData[r]?.length || 0); c++) {
      const key = cleanHeader(gData[r]?.[c]);
      const parts = key.split("-").map(Number);
      if (parts.length === 2 && height >= parts[0] && height <= parts[1]) {
        return num(gData[r]?.[c + 1]);
      }
    }
  }
  // Default
  if (height <= 11) return 3;
  if (height <= 17) return 4;
  return 5;
}

// ── Vertical spacing lookup ──
function lookupVerticalSpacing(height, bucketedWind) {
  const headers = vData[0];
  let heightCol = -1;
  for (let c = 1; c < headers.length; c++) {
    if (num(headers[c]) === height) { heightCol = c; break; }
  }
  if (heightCol < 0) return 0;
  for (let r = 1; r < vData.length; r++) {
    if (num(vData[r][0]) === bucketedWind) return num(vData[r][heightCol]);
  }
  return 0;
}

// ── Original vertical count ──
function lookupOriginalVerticals(width) {
  for (let r = 0; r < vData.length; r++) {
    const label = cleanHeader(vData[r]?.[0]).toLowerCase();
    if (label.includes("original")) {
      const widthRow = vData[r];
      const countRow = vData[r + 1];
      if (countRow) {
        for (let c = 1; c < widthRow.length; c++) {
          if (num(widthRow[c]) === width) return num(countRow[c]);
        }
      }
      break;
    }
  }
  return 0;
}

// ── F54 height adjustment ──
function adjustTrussSpacingForHeight(rawSpacing, height) {
  let adjusted = rawSpacing;
  if (height >= 16) adjusted = rawSpacing - 12;
  else if (height >= 13) adjusted = rawSpacing - 6;
  if (adjusted <= 12) return 0;
  return adjusted;
}

// ── Height multiplier for verticals ──
function getVerticalHeightMultiplier(height) {
  if (height >= 19) return 3.0;
  if (height >= 16) return 2.5;
  if (height >= 13) return 2.0;
  return 1.0;
}

// ── Effective feetUsed (doubled for wider buildings at short heights) ──
function getEffectiveFeetUsed(baseFeetUsed, width, height) {
  if (width >= 26 && height < 13) return baseFeetUsed * 2;
  return baseFeetUsed;
}

// ── Roof rise for A-frame ──
function getRoofRise(width, roofKey) {
  if (roofKey === "AFV" || roofKey === "AFH") {
    return Math.ceil((width / 2) * (3 / 12));
  }
  return 0;
}

// ══════════════════════════════════════════════════════════════
// INVESTIGATION: Does HC row key use RAW or ADJUSTED truss spacing?
// ══════════════════════════════════════════════════════════════
console.log("=".repeat(90));
console.log("INVESTIGATION: HC row key — RAW vs ADJUSTED truss spacing?");
console.log("=".repeat(90));

// Check the Math Calculations sheet formulas
const mathSheet = findSheet("Snow - Math Calculations");
if (mathSheet) {
  // Look at the formulas that compute HC spacing reference
  // Key cells: the truss spacing cell and the HC lookup cell
  const cells = ["F54", "F55", "F56", "P2", "P4", "D10", "D4", "E4", "F4",
                 "G4", "H4", "I4", "J4", "K4", "L4"];
  console.log("\nKey formula cells in Snow - Math Calculations:");
  for (const addr of cells) {
    const cell = mathSheet[addr];
    if (cell) {
      console.log(`  ${addr}: value=${cell.v}, formula=${cell.f || "(none)"}`);
    }
  }

  // Dump rows 50-60 to see the F54 area
  console.log("\nRows 50-60 (F54 area):");
  for (let r = 49; r <= 60; r++) {
    const cells_in_row = [];
    for (let c = 0; c <= 20; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = mathSheet[addr];
      if (cell && cell.v !== "" && cell.v !== 0) {
        cells_in_row.push(`${XLSX.utils.encode_col(c)}${r+1}=${cell.v}${cell.f ? ` [${cell.f}]` : ""}`);
      }
    }
    if (cells_in_row.length) console.log(`  Row ${r+1}: ${cells_in_row.join(" | ")}`);
  }

  // Dump the HC computation area
  console.log("\nHC computation area (look for MATCH/INDEX on truss spacing):");
  for (let r = 0; r <= 15; r++) {
    const cells_in_row = [];
    for (let c = 0; c <= 25; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = mathSheet[addr];
      if (cell && cell.f && (cell.f.includes("Hat") || cell.f.includes("F54") || cell.f.includes("P2"))) {
        cells_in_row.push(`${XLSX.utils.encode_col(c)}${r+1}=${cell.v} [${cell.f}]`);
      }
    }
    if (cells_in_row.length) console.log(`  Row ${r+1}: ${cells_in_row.join(" | ")}`);
  }

  // Broader search: any formula referencing F54 or the truss spacing adjustment
  console.log("\nAll formulas referencing F54 or P2:");
  const range = XLSX.utils.decode_range(mathSheet["!ref"]);
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = mathSheet[addr];
      if (cell?.f && (cell.f.includes("F54") || cell.f.includes("$F$54"))) {
        console.log(`  ${addr}: value=${cell.v}, formula=${cell.f}`);
      }
    }
  }

  console.log("\nAll formulas referencing P2:");
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = mathSheet[addr];
      if (cell?.f && cell.f.includes("P2") && !cell.f.includes("P20") && !cell.f.includes("P21") && !cell.f.includes("P22") && !cell.f.includes("P23") && !cell.f.includes("P24") && !cell.f.includes("P25") && !cell.f.includes("P26") && !cell.f.includes("P27") && !cell.f.includes("P28") && !cell.f.includes("P29")) {
        console.log(`  ${addr}: value=${cell.v}, formula=${cell.f}`);
      }
    }
  }
}

// ══════════════════════════════════════════════════════════════
// MAIN CALCULATION ENGINE (raw spreadsheet-based)
// ══════════════════════════════════════════════════════════════
function calculateSnowEngineering(cfg) {
  const {
    width, length, height, snowLoad, inputWind, state, roofKey,
    sidesCoverage, sidesQty, endsQty,
    sidesOrientation, endsOrientation
  } = cfg;

  const trace = {};
  let totalCost = 0;

  // Step 1: Resolve inputs
  const bucketedWind = bucketWind(inputWind);
  const heightPrefix = getHeightPrefix(height);
  const snowCode = `${heightPrefix}-${snowLoad}`;
  const isEnclosed = sidesCoverage !== "open" && sidesQty >= 2 && endsQty >= 2;
  const enclosure = isEnclosed ? "E" : "O";
  const configKey = `${enclosure}-${bucketedWind}-${width}-${roofKey}`;

  trace.bucketedWind = bucketedWind;
  trace.heightPrefix = heightPrefix;
  trace.snowCode = snowCode;
  trace.isEnclosed = isEnclosed;
  trace.enclosure = enclosure;
  trace.configKey = configKey;

  // Step 2: Trusses
  const { spacing: rawTrussSpacing, col: tsCol } = lookupTrussSpacing(snowCode, configKey);
  const trussSpacing = adjustTrussSpacingForHeight(rawTrussSpacing, height);

  trace.rawTrussSpacing = rawTrussSpacing;
  trace.trussSpacing = trussSpacing;
  trace.tsCol = tsCol;

  if (trussSpacing === 0 || trussSpacing < 18) {
    trace.contactEngineering = true;
    trace.reason = rawTrussSpacing === 0
      ? "Raw truss spacing = 0 (no match)"
      : `Adjusted spacing ${trussSpacing} < 18 (raw=${rawTrussSpacing}, height=${height})`;
    return { total: -1, trace };
  }

  const lengthInches = length * 12;
  const trussesNeeded = Math.ceil(lengthInches / trussSpacing) + 1;
  const { count: origTrusses, key: trussKey } = lookupOriginalTrusses(width, state, length);
  const extraTrusses = Math.max(0, trussesNeeded - origTrusses);

  trace.trussesNeeded = trussesNeeded;
  trace.origTrusses = origTrusses;
  trace.trussKey = trussKey;
  trace.extraTrusses = extraTrusses;

  let trussCost = 0;
  if (extraTrusses > 0) {
    const trussPrice = getTrussPrice(width, state);
    const baseFeetUsed = getFeetUsed(height);
    const feetUsed = getEffectiveFeetUsed(baseFeetUsed, width, height);
    const piePricePerFt = getPiePricePerFt(state);
    const legSurcharge = feetUsed * piePricePerFt;
    trussCost = extraTrusses * (trussPrice + legSurcharge);

    trace.trussPrice = trussPrice;
    trace.baseFeetUsed = baseFeetUsed;
    trace.feetUsed = feetUsed;
    trace.piePricePerFt = piePricePerFt;
    trace.legSurcharge = legSurcharge;
  }
  trace.trussCost = trussCost;
  totalCost += trussCost;

  // Step 3: Hat Channels
  // KEY QUESTION: Does HC use RAW or ADJUSTED truss spacing?
  // Testing BOTH to compare
  const rawBucketed = nearestBucketClosest(rawTrussSpacing, TRUSS_SPACING_BUCKETS);
  const adjBucketed = nearestBucketClosest(trussSpacing, TRUSS_SPACING_BUCKETS);

  // The current engine code uses adjustedTrussSpacing for HC lookup
  const hcTrussSpacing = trussSpacing; // ADJUSTED
  const bucketedTrussSpacing = nearestBucketClosest(hcTrussSpacing, TRUSS_SPACING_BUCKETS);

  // HC row key: NO height prefix
  const hcRowKey = `${bucketedTrussSpacing}-${snowLoad}`;
  const hcSpacing = lookupHCSpacing(hcRowKey, bucketedWind);

  trace.rawBucketedTS = rawBucketed;
  trace.adjBucketedTS = adjBucketed;
  trace.hcTrussSpacingUsed = hcTrussSpacing;
  trace.bucketedTrussSpacing = bucketedTrussSpacing;
  trace.hcRowKey = hcRowKey;
  trace.hcSpacing = hcSpacing;

  // Also compute with RAW for comparison
  const hcRowKeyRaw = `${rawBucketed}-${snowLoad}`;
  const hcSpacingRaw = lookupHCSpacing(hcRowKeyRaw, bucketedWind);
  trace.hcRowKeyRaw = hcRowKeyRaw;
  trace.hcSpacingRaw = hcSpacingRaw;

  let hcCost = 0;
  if (hcSpacing > 0) {
    const barSize = (width + 2) / 2;
    const barInches = barSize * 12;
    const hcPerSide = Math.ceil(barInches / hcSpacing) + 1;
    const totalHC = hcPerSide * 2;
    const origHC = lookupOriginalHC(state, width);
    const extraHC = Math.max(0, totalHC - origHC);

    trace.barSize = barSize;
    trace.hcPerSide = hcPerSide;
    trace.totalHC = totalHC;
    trace.origHC = origHC;
    trace.extraHC = extraHC;

    if (extraHC > 0) {
      const channelPrice = getChannelPrice(state);
      const channelLength = length + 1;
      hcCost = extraHC * channelPrice * channelLength;
      trace.channelPrice = channelPrice;
      trace.channelLength = channelLength;
    }
  }
  trace.hcCost = hcCost;
  totalCost += hcCost;

  // Step 4: Girts (ONLY if enclosed AND vertical panels)
  const hasVerticalPanels = sidesOrientation === "vertical" || endsOrientation === "vertical";
  const girtsNeeded = isEnclosed && hasVerticalPanels;
  trace.hasVerticalPanels = hasVerticalPanels;
  trace.girtsApply = girtsNeeded;

  let girtCost = 0;
  if (girtsNeeded) {
    const girtSpacing = lookupGirtSpacing(bucketedTrussSpacing, bucketedWind);
    trace.girtSpacing = girtSpacing;

    if (girtSpacing > 0) {
      const heightInches = height * 12;
      const girtsRequired = Math.ceil(heightInches / girtSpacing) + 1;
      const origGirts = lookupOriginalGirts(height);
      const extraGirts = Math.max(0, girtsRequired - origGirts);

      trace.girtsRequired = girtsRequired;
      trace.origGirts = origGirts;
      trace.extraGirts = extraGirts;

      if (extraGirts > 0) {
        const tubingPrice = getTubingPrice(state);
        let perimeter = 0;
        if (sidesCoverage !== "open" && sidesOrientation === "vertical") perimeter += sidesQty * length;
        if (endsQty > 0 && endsOrientation === "vertical") perimeter += endsQty * width;
        girtCost = extraGirts * tubingPrice * perimeter;
        trace.tubingPrice = tubingPrice;
        trace.girtPerimeter = perimeter;
      }
    }
  }
  trace.girtCost = girtCost;
  totalCost += girtCost;

  // Step 5: Verticals
  const vertSpacing = lookupVerticalSpacing(height, bucketedWind);
  trace.vertSpacing = vertSpacing;

  let vertCost = 0;
  if (vertSpacing > 0) {
    const widthInches = width * 12;
    const vertsNeeded = Math.ceil(widthInches / vertSpacing) + 1;
    const origVerts = lookupOriginalVerticals(width);
    const extraVerts = Math.max(0, vertsNeeded - origVerts);

    trace.vertsNeeded = vertsNeeded;
    trace.origVerts = origVerts;
    trace.extraVerts = extraVerts;

    if (extraVerts > 0 && endsQty > 0) {
      const tubingPrice = getTubingPrice(state);
      const roofRise = getRoofRise(width, roofKey);
      const peakHeight = height + roofRise;
      const baseVertCost = extraVerts * endsQty * tubingPrice * peakHeight;
      const heightMult = getVerticalHeightMultiplier(height);
      vertCost = baseVertCost * heightMult;

      trace.vertTubingPrice = tubingPrice;
      trace.roofRise = roofRise;
      trace.peakHeight = peakHeight;
      trace.baseVertCost = baseVertCost;
      trace.heightMult = heightMult;
    }
  }
  trace.vertCost = vertCost;
  totalCost += vertCost;

  trace.contactEngineering = false;
  return { total: Math.round(totalCost), trace };
}

// ══════════════════════════════════════════════════════════════
// TEST CONFIGURATIONS
// ══════════════════════════════════════════════════════════════
const configs = [
  {
    label: "1. 22x50x12, AFV, Horiz, Enclosed, 30GL, 90mph",
    width: 22, length: 50, height: 12, snowLoad: "30GL", inputWind: 90,
    state: "AZ", roofKey: "AFV",
    sidesCoverage: "fully_enclosed", sidesQty: 2, endsQty: 2,
    sidesOrientation: "horizontal", endsOrientation: "horizontal",
    expected: "$0",
  },
  {
    label: "2. 30x100x15, AFV, Horiz, Enclosed, 70GL, 100mph",
    width: 30, length: 100, height: 15, snowLoad: "70GL", inputWind: 100,
    state: "AZ", roofKey: "AFV",
    sidesCoverage: "fully_enclosed", sidesQty: 2, endsQty: 2,
    sidesOrientation: "horizontal", endsOrientation: "horizontal",
    expected: "~$9,891",
  },
  {
    label: "3. 24x50x12, AFV, Horiz, Enclosed, 20LL, 90mph",
    width: 24, length: 50, height: 12, snowLoad: "20LL", inputWind: 90,
    state: "AZ", roofKey: "AFV",
    sidesCoverage: "fully_enclosed", sidesQty: 2, endsQty: 2,
    sidesOrientation: "horizontal", endsOrientation: "horizontal",
    expected: "$0",
  },
  {
    label: "4. 24x75x14, AFV, Vertical, Enclosed, 50GL, 130mph",
    width: 24, length: 75, height: 14, snowLoad: "50GL", inputWind: 130,
    state: "AZ", roofKey: "AFV",
    sidesCoverage: "fully_enclosed", sidesQty: 2, endsQty: 2,
    sidesOrientation: "vertical", endsOrientation: "vertical",
    expected: "non-zero",
  },
  {
    label: "5. 30x100x18, STD, Vertical, Enclosed, 90GL, 155mph",
    width: 30, length: 100, height: 18, snowLoad: "90GL", inputWind: 155,
    state: "AZ", roofKey: "STD",
    sidesCoverage: "fully_enclosed", sidesQty: 2, endsQty: 2,
    sidesOrientation: "vertical", endsOrientation: "vertical",
    expected: "large or Contact Engineering",
  },
  {
    label: "6. 12x25x8, AFV, Open, 30GL, 105mph",
    width: 12, length: 25, height: 8, snowLoad: "30GL", inputWind: 105,
    state: "AZ", roofKey: "AFV",
    sidesCoverage: "open", sidesQty: 0, endsQty: 0,
    sidesOrientation: "horizontal", endsOrientation: "horizontal",
    expected: "$0 or small",
  },
  {
    label: "7. 26x50x16, AFV, Horiz, Enclosed, 70GL, 140mph",
    width: 26, length: 50, height: 16, snowLoad: "70GL", inputWind: 140,
    state: "AZ", roofKey: "AFV",
    sidesCoverage: "fully_enclosed", sidesQty: 2, endsQty: 2,
    sidesOrientation: "horizontal", endsOrientation: "horizontal",
    expected: "non-zero",
  },
  {
    label: "8. 24x100x12, AFV, Horiz, Enclosed, 40GL, 115mph",
    width: 24, length: 100, height: 12, snowLoad: "40GL", inputWind: 115,
    state: "AZ", roofKey: "AFV",
    sidesCoverage: "fully_enclosed", sidesQty: 2, endsQty: 2,
    sidesOrientation: "horizontal", endsOrientation: "horizontal",
    expected: "check",
  },
  {
    label: "9. 20x50x10, STD, Open, 20LL, 105mph",
    width: 20, length: 50, height: 10, snowLoad: "20LL", inputWind: 105,
    state: "AZ", roofKey: "STD",
    sidesCoverage: "open", sidesQty: 0, endsQty: 0,
    sidesOrientation: "horizontal", endsOrientation: "horizontal",
    expected: "$0",
  },
  {
    label: "10. 28x75x20, AFV, Vertical, Enclosed, 80GL, 165mph",
    width: 28, length: 75, height: 20, snowLoad: "80GL", inputWind: 165,
    state: "AZ", roofKey: "AFV",
    sidesCoverage: "fully_enclosed", sidesQty: 2, endsQty: 2,
    sidesOrientation: "vertical", endsOrientation: "vertical",
    expected: "large or Contact Engineering",
  },
];

// ══════════════════════════════════════════════════════════════
// RUN ALL CONFIGS
// ══════════════════════════════════════════════════════════════
console.log("\n\n" + "=".repeat(90));
console.log("MULTI-CONFIG SNOW ENGINEERING TEST");
console.log("=".repeat(90));

const summaryRows = [];

for (const cfg of configs) {
  console.log("\n" + "─".repeat(90));
  console.log(`CONFIG: ${cfg.label}`);
  console.log("─".repeat(90));

  const { total, trace: t } = calculateSnowEngineering(cfg);

  console.log(`  Wind bucketing:    ${cfg.inputWind}mph → ${t.bucketedWind}mph`);
  console.log(`  Height class:      height ${cfg.height} → "${t.heightPrefix}"`);
  console.log(`  Snow code:         "${t.snowCode}"`);
  console.log(`  E/O:               ${t.enclosure} (enclosed=${t.isEnclosed})`);
  console.log(`  Config key:        "${t.configKey}"`);

  if (t.contactEngineering) {
    console.log(`  ** CONTACT ENGINEERING **`);
    console.log(`    Reason: ${t.reason}`);
    summaryRows.push({ label: cfg.label, total: "CONTACT ENGINEERING", expected: cfg.expected });
    continue;
  }

  console.log(`  Raw truss spacing: ${t.rawTrussSpacing}" (col ${t.tsCol})`);
  console.log(`  F54 adjustment:    ${t.rawTrussSpacing}" → ${t.trussSpacing}" (height ${cfg.height})`);
  console.log(`  Trusses needed:    ${t.trussesNeeded} (orig ${t.origTrusses} from "${t.trussKey}")`);
  console.log(`  Extra trusses:     ${t.extraTrusses}`);
  if (t.extraTrusses > 0) {
    console.log(`    Truss price: $${t.trussPrice}, feetUsed: ${t.feetUsed} (base ${t.baseFeetUsed}), pie: $${t.piePricePerFt}/ft, leg: $${t.legSurcharge}`);
  }
  console.log(`  TRUSS COST:        $${t.trussCost}`);

  console.log(`  HC bucket (adj):   ${t.adjBucketedTS}  |  HC bucket (raw): ${t.rawBucketedTS}`);
  console.log(`  HC row key (adj):  "${t.hcRowKey}" → spacing ${t.hcSpacing}"`);
  console.log(`  HC row key (raw):  "${t.hcRowKeyRaw}" → spacing ${t.hcSpacingRaw}"`);
  if (t.hcSpacing > 0 || t.hcSpacingRaw > 0) {
    console.log(`  HC: bar=${t.barSize}ft, perSide=${t.hcPerSide}, total=${t.totalHC}, orig=${t.origHC}, extra=${t.extraHC}`);
    if (t.extraHC > 0) console.log(`    Channel price: $${t.channelPrice}/ft × ${t.channelLength}ft`);
  }
  console.log(`  HC COST:           $${t.hcCost}`);

  console.log(`  Girts apply:       ${t.girtsApply} (enclosed=${t.isEnclosed}, vertPanels=${t.hasVerticalPanels})`);
  if (t.girtsApply && t.girtSpacing) {
    console.log(`  Girt spacing:      ${t.girtSpacing}" (truss=${t.bucketedTrussSpacing}, wind=${t.bucketedWind})`);
    console.log(`  Girts: needed=${t.girtsRequired}, orig=${t.origGirts}, extra=${t.extraGirts}`);
    if (t.extraGirts > 0) console.log(`    Tubing: $${t.tubingPrice}/ft, perimeter: ${t.girtPerimeter}ft`);
  }
  console.log(`  GIRT COST:         $${t.girtCost}`);

  console.log(`  Vert spacing:      ${t.vertSpacing}" (height=${cfg.height}, wind=${t.bucketedWind})`);
  if (t.vertSpacing > 0) {
    console.log(`  Verts: needed=${t.vertsNeeded}, orig=${t.origVerts}, extra=${t.extraVerts}`);
    if (t.extraVerts > 0) {
      console.log(`    Tubing: $${t.vertTubingPrice}/ft, rise=${t.roofRise}ft, peak=${t.peakHeight}ft, mult=${t.heightMult}`);
      console.log(`    Base vert cost: $${t.baseVertCost} × ${t.heightMult} = $${t.vertCost}`);
    }
  }
  console.log(`  VERT COST:         $${t.vertCost}`);

  console.log(`  ────────────────────────────────────────`);
  console.log(`  TOTAL:             $${total}`);
  console.log(`  EXPECTED:          ${cfg.expected}`);

  summaryRows.push({ label: cfg.label, total: `$${total}`, expected: cfg.expected });
}

// ══════════════════════════════════════════════════════════════
// SUMMARY TABLE
// ══════════════════════════════════════════════════════════════
console.log("\n\n" + "=".repeat(90));
console.log("SUMMARY TABLE");
console.log("=".repeat(90));
console.log("Config".padEnd(55) + "Result".padEnd(20) + "Expected");
console.log("-".repeat(90));
for (const row of summaryRows) {
  console.log(row.label.padEnd(55) + String(row.total).padEnd(20) + row.expected);
}

// ══════════════════════════════════════════════════════════════
// HC BUCKET ANALYSIS: RAW vs ADJUSTED
// ══════════════════════════════════════════════════════════════
console.log("\n\n" + "=".repeat(90));
console.log("HC BUCKET ANALYSIS: RAW vs ADJUSTED TRUSS SPACING");
console.log("=".repeat(90));
console.log("For configs where height >= 13 (F54 adjustment applies):");
console.log("Checking if RAW or ADJUSTED gives different HC results.\n");

for (const cfg of configs) {
  if (cfg.height < 13) continue;
  const { total, trace: t } = calculateSnowEngineering(cfg);
  if (t.contactEngineering) {
    console.log(`${cfg.label}: CONTACT ENGINEERING — skipped`);
    continue;
  }
  const rawDiffers = t.rawBucketedTS !== t.adjBucketedTS;
  const hcDiffers = t.hcSpacing !== t.hcSpacingRaw;
  console.log(`${cfg.label}:`);
  console.log(`  Raw spacing: ${t.rawTrussSpacing}" → bucket ${t.rawBucketedTS}`);
  console.log(`  Adj spacing: ${t.trussSpacing}" → bucket ${t.adjBucketedTS}`);
  console.log(`  HC row (adj): "${t.hcRowKey}" → ${t.hcSpacing}"`);
  console.log(`  HC row (raw): "${t.hcRowKeyRaw}" → ${t.hcSpacingRaw}"`);
  console.log(`  Bucket differs: ${rawDiffers} | HC spacing differs: ${hcDiffers}`);
  if (hcDiffers) {
    // Compute cost difference
    const barInches = ((cfg.width + 2) / 2) * 12;
    const hcNeededAdj = hcSpacingCalc(barInches, t.hcSpacing);
    const hcNeededRaw = hcSpacingCalc(barInches, t.hcSpacingRaw);
    const origHC = t.origHC;
    console.log(`  HC needed (adj): ${hcNeededAdj} vs (raw): ${hcNeededRaw} (orig: ${origHC})`);
  }
  console.log();
}

function hcSpacingCalc(barInches, spacing) {
  if (!spacing || spacing <= 0) return 0;
  return (Math.ceil(barInches / spacing) + 1) * 2;
}

// ══════════════════════════════════════════════════════════════
// AVAILABLE ROW KEYS IN HC SHEET (for debugging)
// ══════════════════════════════════════════════════════════════
console.log("\n" + "=".repeat(90));
console.log("AVAILABLE HC ROW KEYS IN SPREADSHEET");
console.log("=".repeat(90));
const hcRowKeys = new Set();
for (let r = 1; r < hcData.length; r++) {
  const key = cleanHeader(hcData[r][0]);
  if (key) hcRowKeys.add(key);
}
console.log([...hcRowKeys].sort().join("\n"));

console.log("\n" + "=".repeat(90));
console.log("AVAILABLE TRUSS SPACING CONFIG KEYS");
console.log("=".repeat(90));
const tsHeaders = tsData[0];
const tsConfigKeys = [];
for (let c = 1; c < tsHeaders.length; c++) {
  const h = cleanHeader(tsHeaders[c]);
  if (h) tsConfigKeys.push(h);
}
console.log(tsConfigKeys.join(", "));

console.log("\n" + "=".repeat(90));
console.log("AVAILABLE SNOW CODES (rows in Truss Spacing)");
console.log("=".repeat(90));
const snowCodes = [];
for (let r = 1; r < tsData.length; r++) {
  const key = cleanHeader(tsData[r][0]);
  if (key) snowCodes.push(key);
}
console.log(snowCodes.join(", "));

console.log("\nDone.");
